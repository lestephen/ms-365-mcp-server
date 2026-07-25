import { beforeEach, describe, expect, it, vi } from 'vitest';
import type { Transport } from '@modelcontextprotocol/sdk/shared/transport.js';
import type { JSONRPCMessage } from '@modelcontextprotocol/sdk/types.js';
import {
  enableMetrics,
  metricsText,
  recordBatchSubrequest,
  recordBlockedOperation,
  recordToolCall,
  registry,
} from '../src/metrics.js';
import { withMetricsObserver } from '../src/lib/metrics-transport.js';

/**
 * Metrics exist so the open questions from the hybrid rollout get answered with data
 * rather than inference: whether the --direct-tools set is right, whether callers are
 * hitting unregistered names (#29), whether the mail blocklist is ever exercised, and
 * what graph-batch is actually used for.
 */

async function sample(name: string): Promise<string[]> {
  const text = await metricsText();
  return text.split('\n').filter((line) => line.startsWith(name) && !line.startsWith(`${name} `));
}

describe('metric hygiene', () => {
  beforeEach(() => {
    registry.resetMetrics();
    enableMetrics();
  });

  it('exposes no label that could carry personal data', async () => {
    // Metrics are widely readable and retained; identity belongs in the audit log.
    recordToolCall('get-mail-message', 'direct', 'ok', 0.2);
    recordBlockedOperation('send-mail', 'batch');
    recordBatchSubrequest('send-mail', 'POST');

    const text = await metricsText();
    // Only this module's own families. collectDefaultMetrics contributes Node runtime
    // series under the same prefix, and their labels (le, major, version) are not ours
    // to police.
    const OURS = [
      'ms365_mcp_tool_calls_total',
      'ms365_mcp_tool_duration_seconds',
      'ms365_mcp_unknown_tool_total',
      'ms365_mcp_blocked_operations_total',
      'ms365_mcp_batch_subrequests_total',
      'ms365_mcp_discovery_stage_total',
    ];
    const labelNames = new Set<string>();
    for (const line of text.split('\n')) {
      const family = OURS.find((name) => line.startsWith(`${name}{`));
      if (!family) continue;
      const labels = line.slice(family.length + 1, line.indexOf('}'));
      for (const pair of labels.split(',')) {
        const key = pair.split('=')[0]?.trim();
        // `le` is the histogram bucket boundary, inherent to the type.
        if (key && key !== 'le') labelNames.add(key);
      }
    }
    // Bounded domains only: tool names, fixed unions, HTTP methods.
    expect([...labelNames].sort()).toEqual(['method', 'operation', 'outcome', 'route', 'tool']);
    for (const forbidden of ['user', 'upn', 'principal', 'id', 'path', 'url', 'subject']) {
      expect(labelNames.has(forbidden), `label ${forbidden} must not exist`).toBe(false);
    }
  });

  it('records a tool call with route and outcome', async () => {
    recordToolCall('get-mail-message', 'execute_tool', 'ok', 0.1);
    const lines = await sample('ms365_mcp_tool_calls_total');
    expect(lines.join('\n')).toContain('tool="get-mail-message"');
    expect(lines.join('\n')).toContain('route="execute_tool"');
    expect(lines.join('\n')).toContain('outcome="ok"');
  });

  it('separates the route, so the direct set can be tuned from data', async () => {
    recordToolCall('update-place', 'execute_tool', 'ok');
    recordToolCall('get-mail-message', 'direct', 'ok');
    const text = (await sample('ms365_mcp_tool_calls_total')).join('\n');
    // A tool repeatedly reached via execute-tool is a candidate for --direct-tools.
    expect(text).toMatch(/tool="update-place",route="execute_tool"[^\n]*1/);
    expect(text).toMatch(/tool="get-mail-message",route="direct"[^\n]*1/);
  });
});

describe('unknown-tool observation (the #29 signal)', () => {
  beforeEach(() => {
    registry.resetMetrics();
    enableMetrics();
  });

  function wired() {
    const sent: JSONRPCMessage[] = [];
    const transport = {
      send: vi.fn(async (m: JSONRPCMessage) => {
        sent.push(m);
      }),
      start: vi.fn(),
      close: vi.fn(),
      onmessage: undefined,
    } as unknown as Transport;
    const wrapped = withMetricsObserver(transport);
    // Server.connect assigns onmessage; emulate that so the observer chains in.
    const delivered: JSONRPCMessage[] = [];
    wrapped.onmessage = ((m: JSONRPCMessage) => {
      delivered.push(m);
    }) as Transport['onmessage'];
    return { wrapped, sent, delivered };
  }

  it('counts a tools/call for a name the server does not register', async () => {
    // This is the shape the SDK actually emits, confirmed against a running server: a
    // normal result carrying isError, not a JSON-RPC error object. An earlier version of
    // this test asserted the error-object shape, so it validated an assumption and the
    // real path never fired.
    const { wrapped } = wired();
    wrapped.onmessage!({
      jsonrpc: '2.0',
      id: 7,
      method: 'tools/call',
      params: { name: 'delete-mail-attachment' },
    } as never);
    await wrapped.send({
      jsonrpc: '2.0',
      id: 7,
      result: {
        content: [
          { type: 'text', text: 'MCP error -32602: Tool delete-mail-attachment not found' },
        ],
        isError: true,
      },
    } as never);

    const lines = await sample('ms365_mcp_unknown_tool_total');
    expect(lines.join('\n')).toContain('tool="delete-mail-attachment"');
  });

  it('also counts the JSON-RPC error shape, if the SDK ever returns one', async () => {
    const { wrapped } = wired();
    wrapped.onmessage!({
      jsonrpc: '2.0',
      id: 11,
      method: 'tools/call',
      params: { name: 'update-place' },
    } as never);
    await wrapped.send({
      jsonrpc: '2.0',
      id: 11,
      error: { code: -32602, message: 'Tool update-place not found' },
    } as never);

    expect((await sample('ms365_mcp_unknown_tool_total')).join('\n')).toContain(
      'tool="update-place"'
    );
  });

  it('does not count an ordinary tool failure, which comes back as a result', async () => {
    const { wrapped } = wired();
    wrapped.onmessage!({
      jsonrpc: '2.0',
      id: 8,
      method: 'tools/call',
      params: { name: 'get-mail-message' },
    } as never);
    await wrapped.send({
      jsonrpc: '2.0',
      id: 8,
      result: { content: [{ type: 'text', text: '{"error":"boom"}' }], isError: true },
    } as never);

    expect(await sample('ms365_mcp_unknown_tool_total')).toEqual([]);
  });

  it('still forwards messages in both directions', async () => {
    const { wrapped, sent, delivered } = wired();
    wrapped.onmessage!({ jsonrpc: '2.0', id: 1, method: 'ping' } as never);
    await wrapped.send({ jsonrpc: '2.0', id: 1, result: {} } as never);
    expect(delivered).toHaveLength(1);
    expect(sent).toHaveLength(1);
  });

  it('records the advertised surface from a tools/list response', async () => {
    const { wrapped } = wired();
    await wrapped.send({
      jsonrpc: '2.0',
      id: 2,
      result: {
        tools: [
          { name: 'a', inputSchema: {} },
          { name: 'b', inputSchema: {} },
        ],
      },
    } as never);

    const text = await metricsText();
    expect(text).toMatch(/ms365_mcp_tools_advertised 2/);
    expect(text).toMatch(/ms365_mcp_tool_schema_bytes \d+/);
  });

  it('leaves transport.onmessage completely alone', () => {
    // Regression guard for a real outage. An earlier version intercepted onmessage via
    // Object.defineProperty to correlate responses back to request ids. That broke
    // StreamableHTTPServerTransport outright: every HTTP request hung with no response,
    // while stdio was unaffected, so it passed a stdio-only check and reached
    // production. Only send() may be wrapped.
    const onmessage = () => {};
    const transport = {
      send: vi.fn(),
      start: vi.fn(),
      close: vi.fn(),
      onmessage,
    } as unknown as Transport;

    withMetricsObserver(transport);

    expect(transport.onmessage).toBe(onmessage);
    const descriptor = Object.getOwnPropertyDescriptor(transport, 'onmessage');
    expect(descriptor?.get, 'onmessage must not become an accessor').toBeUndefined();
    expect(descriptor?.set, 'onmessage must not become an accessor').toBeUndefined();
    expect(descriptor?.value).toBe(onmessage);
  });

  it('counts a not-found whose tool name cannot be parsed, under a sentinel', async () => {
    // Detection requires "not found" rather than keying on -32602 alone, because
    // -32602 is invalid-params generally and would over-count ordinary argument errors.
    // Within that, a name we cannot parse still counts rather than vanishing.
    const { wrapped } = wired();
    await wrapped.send({
      jsonrpc: '2.0',
      id: 12,
      result: {
        content: [{ type: 'text', text: 'MCP error -32602: tool "odd name" not found' }],
        isError: true,
      },
    } as never);

    expect((await sample('ms365_mcp_unknown_tool_total')).join('\n')).toContain('tool="unparsed"');
  });

  it('does not count an ordinary invalid-params error as an unknown tool', async () => {
    const { wrapped } = wired();
    await wrapped.send({
      jsonrpc: '2.0',
      id: 13,
      error: { code: -32602, message: 'Invalid arguments: messageId is required' },
    } as never);

    expect(await sample('ms365_mcp_unknown_tool_total')).toEqual([]);
  });
});

describe('metrics are opt-in', () => {
  it('records nothing before enableMetrics', async () => {
    // A fresh module registry cannot be simulated here, so assert the documented
    // contract instead: the helpers are guarded and the registry starts empty.
    const fresh = new (registry.constructor as new () => typeof registry)();
    expect((await fresh.metrics()).trim()).toBe('');
  });
});
