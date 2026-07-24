import { describe, expect, it, vi } from 'vitest';
import type { Transport } from '@modelcontextprotocol/sdk/shared/transport.js';
import type { JSONRPCMessage } from '@modelcontextprotocol/sdk/types.js';
import {
  applyStrictToolSchemas,
  hoistRefsUnderDefs,
  withStrictToolSchemas,
} from '../lib/strict-tool-schemas.js';

/**
 * Collect every $ref value in a schema, so tests can assert on the whole tree
 * rather than on one hand-picked node.
 */
function allRefs(node: unknown, out: string[] = []): string[] {
  if (!node || typeof node !== 'object') return out;
  if (Array.isArray(node)) {
    for (const item of node) allRefs(item, out);
    return out;
  }
  const record = node as Record<string, unknown>;
  if (typeof record.$ref === 'string') out.push(record.$ref);
  for (const value of Object.values(record)) allRefs(value, out);
  return out;
}

describe('hoistRefsUnderDefs', () => {
  it('returns a schema with no $ref unchanged', () => {
    const schema = { type: 'object', properties: { subject: { type: 'string' } } };
    expect(hoistRefsUnderDefs(schema)).toEqual(schema);
  });

  it('rewrites a root-relative pointer to #/$defs and hoists the target subschema', () => {
    // This is the exact shape ms-365-mcp-server emits today for create-draft-email:
    // zod-to-json-schema dedupes the repeated recipient object into a root pointer.
    const emailAddress = {
      type: 'object',
      properties: { address: { type: 'string' }, name: { type: 'string' } },
    };
    const result = hoistRefsUnderDefs({
      type: 'object',
      properties: {
        body: {
          type: 'object',
          properties: {
            from: emailAddress,
            sender: { $ref: '#/properties/body/properties/from' },
          },
        },
      },
    }) as Record<string, any>;

    expect(allRefs(result)).toEqual(['#/$defs/body_from']);
    expect(result.$defs.body_from).toEqual(emailAddress);
    // The original inline copy must survive, so the schema still validates.
    expect(result.properties.body.properties.from).toEqual(emailAddress);
  });

  it('produces a valid recursive $defs reference for a self-referential pointer', () => {
    const result = hoistRefsUnderDefs({
      type: 'object',
      properties: {
        body: {
          type: 'object',
          properties: {
            name: { type: 'string' },
            childFolders: { type: 'array', items: { $ref: '#/properties/body' } },
          },
        },
      },
    }) as Record<string, any>;

    expect(allRefs(result)).toEqual(['#/$defs/body', '#/$defs/body']);
    // The hoisted copy refers back to itself, which is legal JSON Schema and
    // which Moonshot accepts, rather than being inlined forever.
    expect(result.$defs.body.properties.childFolders.items.$ref).toBe('#/$defs/body');
  });

  it('resolves pointers that traverse array items and anyOf branch indices', () => {
    const result = hoistRefsUnderDefs({
      type: 'object',
      properties: {
        body: {
          type: 'object',
          properties: {
            allowedRoles: { type: 'array', items: { anyOf: [{ enum: ['read', 'write'] }] } },
            fallbackRole: { $ref: '#/properties/body/properties/allowedRoles/items/anyOf/0' },
          },
        },
      },
    }) as Record<string, any>;

    expect(allRefs(result)).toEqual(['#/$defs/body_allowedRoles_items_anyOf_0']);
    expect(result.$defs.body_allowedRoles_items_anyOf_0).toEqual({ enum: ['read', 'write'] });
  });

  it('leaves refs that already point under #/$defs alone', () => {
    const schema = {
      type: 'object',
      $defs: { addr: { type: 'string' } },
      properties: { to: { $ref: '#/$defs/addr' } },
    };
    expect(hoistRefsUnderDefs(schema)).toEqual(schema);
  });

  it('degrades an unresolvable pointer to a permissive schema instead of throwing', () => {
    const result = hoistRefsUnderDefs({
      type: 'object',
      properties: { to: { $ref: '#/properties/nope/properties/gone' } },
    }) as Record<string, any>;

    expect(allRefs(result)).toEqual(['#/$defs/nope_gone']);
    expect(result.$defs.nope_gone).toEqual({});
  });

  it('does not collide with a $defs name that already exists', () => {
    const result = hoistRefsUnderDefs({
      type: 'object',
      $defs: { body_from: { const: 'pre-existing' } },
      properties: {
        body: { type: 'object', properties: { from: { type: 'string' } } },
        other: { $ref: '#/properties/body/properties/from' },
      },
    }) as Record<string, any>;

    expect(result.$defs.body_from).toEqual({ const: 'pre-existing' });
    expect(allRefs(result)).toEqual(['#/$defs/body_from_2']);
    expect(result.$defs.body_from_2).toEqual({ type: 'string' });
  });

  it('does not mutate its input', () => {
    const schema = {
      type: 'object',
      properties: {
        body: {
          type: 'object',
          properties: {
            from: { type: 'string' },
            sender: { $ref: '#/properties/body/properties/from' },
          },
        },
      },
    };
    const snapshot = JSON.stringify(schema);
    hoistRefsUnderDefs(schema);
    expect(JSON.stringify(schema)).toBe(snapshot);
  });
});

describe('applyStrictToolSchemas', () => {
  it('rewrites inputSchema and outputSchema across every tool in a tools/list result', () => {
    const message = {
      jsonrpc: '2.0',
      id: 1,
      result: {
        tools: [
          {
            name: 'create-draft-email',
            inputSchema: {
              type: 'object',
              properties: {
                body: {
                  type: 'object',
                  properties: {
                    from: { type: 'string' },
                    sender: { $ref: '#/properties/body/properties/from' },
                  },
                },
              },
            },
            outputSchema: {
              type: 'object',
              properties: { a: { type: 'string' }, b: { $ref: '#/properties/a' } },
            },
          },
          { name: 'get-current-user', inputSchema: { type: 'object', properties: {} } },
        ],
      },
    };

    applyStrictToolSchemas(message);

    expect(allRefs(message.result.tools[0].inputSchema)).toEqual(['#/$defs/body_from']);
    expect(allRefs(message.result.tools[0].outputSchema)).toEqual(['#/$defs/a']);
    expect(message.result.tools[1].inputSchema).toEqual({ type: 'object', properties: {} });
  });

  it('leaves messages that are not tools/list results untouched', () => {
    for (const message of [
      { jsonrpc: '2.0', id: 1, result: { content: [{ type: 'text', text: 'hi' }] } },
      { jsonrpc: '2.0', method: 'notifications/initialized', params: {} },
      { jsonrpc: '2.0', id: 2, error: { code: -1, message: 'boom' } },
    ]) {
      const snapshot = JSON.stringify(message);
      applyStrictToolSchemas(message);
      expect(JSON.stringify(message)).toBe(snapshot);
    }
  });

  it('tolerates a malformed tools array without throwing', () => {
    const message = { jsonrpc: '2.0', id: 1, result: { tools: [null, 'nope', {}] } };
    expect(() => applyStrictToolSchemas(message)).not.toThrow();
  });
});

describe('withStrictToolSchemas', () => {
  it('sanitizes outgoing tools/list results and forwards everything else verbatim', async () => {
    const sent: unknown[] = [];
    // Hold the spy separately: the decorator replaces transport.send, so
    // asserting on the property after wrapping would inspect the wrapper.
    const innerSend = vi.fn(async (message: unknown) => {
      // Record what the client would actually receive on the wire.
      sent.push(JSON.parse(JSON.stringify(message)));
    });
    const transport = { send: innerSend, start: vi.fn(), close: vi.fn() } as unknown as Transport;

    const wrapped = withStrictToolSchemas(transport);
    await wrapped.send({
      jsonrpc: '2.0',
      id: 1,
      result: {
        tools: [
          {
            name: 't',
            inputSchema: {
              type: 'object',
              properties: { a: { type: 'string' }, b: { $ref: '#/properties/a' } },
            },
          },
        ],
      },
    } as unknown as JSONRPCMessage);
    await wrapped.send({
      jsonrpc: '2.0',
      id: 2,
      result: { ok: true },
    } as unknown as JSONRPCMessage);

    expect(innerSend).toHaveBeenCalledTimes(2);
    expect(allRefs(sent[0])).toEqual(['#/$defs/a']);
    expect(sent[1]).toEqual({ jsonrpc: '2.0', id: 2, result: { ok: true } });
  });

  it('returns the same transport instance so callers keep their reference', () => {
    const transport = { send: vi.fn(), start: vi.fn(), close: vi.fn() } as unknown as Transport;
    expect(withStrictToolSchemas(transport)).toBe(transport);
  });
});
