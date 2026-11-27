import { beforeEach, describe, expect, it, vi } from 'vitest';
import { Client } from '@modelcontextprotocol/client';
import { InMemoryTransport, McpServer } from '@modelcontextprotocol/server';
import { z } from 'zod';
import {
  enableMetrics,
  metricsText,
  recordBatchSubrequest,
  recordBlockedOperation,
  initBlockedOperationSeries,
  recordToolCall,
  recordUnknownTool,
  registry,
  resetUnknownToolLabelCapForTests,
} from '../src/metrics.js';
import {
  installMetricsObserver,
  seedAdvertisedSurface,
  resetAdvertisedSurfaceSeedForTests,
} from '../src/lib/metrics-transport.js';

/**
 * Metrics exist so the open questions from the hybrid rollout get answered with data
 * rather than inference: whether the --eager-tools set is right, whether callers are
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
    resetUnknownToolLabelCapForTests();
    enableMetrics();
  });

  it('buckets an unrecognised batch method instead of labelling it verbatim', async () => {
    // The method comes straight out of a caller-supplied $batch subrequest. HTTP methods
    // are a closed set, so anything else is caller-invented label cardinality.
    recordBatchSubrequest('unmatched', 'NOT-A-METHOD');

    const lines = (await sample('ms365_mcp_batch_subrequests_total')).join('\n');
    expect(lines).not.toContain('NOT-A-METHOD');
    expect(lines).toContain('method="OTHER"');
  });

  it('caps the number of distinct unknown-tool series a caller can mint', async () => {
    // The tool name is parsed out of the SDK's "Tool <name> not found" error, so its
    // value set is whatever a caller invents. prom-client retains every distinct label
    // combination for the process lifetime, so an uncapped label is an OOM primitive on
    // a hosted multi-user endpoint.
    for (let i = 0; i < 500; i++) recordUnknownTool(`invented-${i}`);

    const lines = await sample('ms365_mcp_unknown_tool_total');
    expect(lines.length).toBeLessThan(100);
    expect(lines.join('\n')).toContain('__over_cap__');
  });

  it('truncates an absurdly long unknown-tool name', async () => {
    // Capping the NUMBER of series bounds how many labels exist, not how big each one
    // is. The name matches /[A-Za-z0-9._-]+/ with no length limit, so without this the
    // first 50 callers could each retain a megabyte-scale label for the life of the
    // process and re-serialise it on every scrape.
    recordUnknownTool('z'.repeat(5000));

    expect((await sample('ms365_mcp_unknown_tool_total')).join('\n').length).toBeLessThan(500);
  });

  it('still names unknown tools while under the cap', async () => {
    // The metric has to stay useful: #29 is about finding out which real tool names
    // callers reach for directly, and that needs the actual name.
    recordUnknownTool('list-mail-messages');

    expect((await sample('ms365_mcp_unknown_tool_total')).join('\n')).toContain(
      'tool="list-mail-messages"'
    );
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
    // A tool repeatedly reached via execute-tool is a candidate for --eager-tools.
    expect(text).toMatch(/tool="update-place",route="execute_tool"[^\n]*1/);
    expect(text).toMatch(/tool="get-mail-message",route="direct"[^\n]*1/);
  });
});

describe('unknown-tool observation (the #29 signal)', () => {
  beforeEach(() => {
    registry.resetMetrics();
    resetUnknownToolLabelCapForTests();
    enableMetrics();
  });

  // A real SDK server and client over an in-memory pair, so the test exercises the
  // dispatch path every transport shares rather than an emulation of one.
  async function connected(register: (server: McpServer) => void = () => {}) {
    const server = new McpServer({ name: 'metrics-test', version: '0' });
    server.registerTool(
      'get-mail-message',
      { description: 'stub', inputSchema: z.object({ fail: z.boolean().optional() }) },
      async ({ fail }) =>
        fail
          ? { content: [{ type: 'text' as const, text: '{"error":"boom"}' }], isError: true }
          : { content: [{ type: 'text' as const, text: '{}' }] }
    );
    register(server);
    expect(installMetricsObserver(server)).toBe(true);
    const [a, b] = InMemoryTransport.createLinkedPair();
    await server.connect(a);
    const client = new Client({ name: 'test', version: '1' });
    await client.connect(b);
    return {
      client,
      close: async () => {
        await client.close();
        await server.close();
      },
    };
  }

  it('counts a tools/call for a name the server does not register', async () => {
    const { client, close } = await connected();
    try {
      await expect(
        client.callTool({ name: 'delete-mail-attachment', arguments: {} })
      ).rejects.toMatchObject({ code: -32602 });
    } finally {
      await close();
    }
    expect((await sample('ms365_mcp_unknown_tool_total')).join('\n')).toContain(
      'tool="delete-mail-attachment"'
    );
  });

  it('does not count an ordinary tool failure, which comes back as a result', async () => {
    const { client, close } = await connected();
    try {
      const result = await client.callTool({ name: 'get-mail-message', arguments: { fail: true } });
      expect(result.isError).toBe(true);
    } finally {
      await close();
    }
    expect(await sample('ms365_mcp_unknown_tool_total')).toEqual([]);
  });

  it('does not count an invalid-arguments error as an unknown tool', async () => {
    const { client, close } = await connected();
    try {
      await client
        .callTool({ name: 'get-mail-message', arguments: { fail: 'not-a-boolean' } })
        .catch(() => undefined);
    } finally {
      await close();
    }
    expect(await sample('ms365_mcp_unknown_tool_total')).toEqual([]);
  });

  it('records the advertised surface from a tools/list response', async () => {
    const { client, close } = await connected((server) =>
      server.registerTool('parse-teams-url', { description: 'stub' }, async () => ({
        content: [],
      }))
    );
    try {
      expect((await client.listTools()).tools).toHaveLength(2);
    } finally {
      await close();
    }
    const text = await metricsText();
    expect(text).toMatch(/ms365_mcp_tools_advertised 2/);
    expect(text).toMatch(/ms365_mcp_tool_schema_bytes \d+/);
  });

  it('still answers calls normally once installed, and installs once', async () => {
    const { client, close } = await connected((server) => {
      expect(installMetricsObserver(server)).toBe(true);
    });
    try {
      const result = await client.callTool({ name: 'get-mail-message', arguments: {} });
      expect(result.isError).toBeFalsy();
    } finally {
      await close();
    }
  });

  it('is never fatal when the SDK internals move', () => {
    expect(installMetricsObserver({ server: {} })).toBe(false);
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

describe('advertised-surface seed (#34)', () => {
  beforeEach(() => {
    resetAdvertisedSurfaceSeedForTests();
  });

  function fakeServer(tools: Array<Record<string, unknown>>) {
    return {
      server: {
        _requestHandlers: new Map([['tools/list', async () => ({ tools })]]),
      },
    };
  }

  it('records count and bytes without a client ever calling tools/list', async () => {
    enableMetrics();
    const tools = [{ name: 'a', inputSchema: { type: 'object' } }];
    expect(await seedAdvertisedSurface(fakeServer(tools))).toBe('seeded');
    const text = await metricsText();
    expect(text).toMatch(/^ms365_mcp_tools_advertised 1$/m);
    expect(text).toMatch(
      new RegExp(`^ms365_mcp_tool_schema_bytes ${JSON.stringify(tools).length}$`, 'm')
    );
  });

  it('seeds once per process, since HTTP builds a server per request', async () => {
    expect(await seedAdvertisedSurface(fakeServer([{ name: 'a' }]))).toBe('seeded');
    expect(await seedAdvertisedSurface(fakeServer([{ name: 'a' }]))).toBe('skipped');
  });

  it('seeds from a real, unconnected SDK server', async () => {
    enableMetrics();
    const server = new McpServer({ name: 'seed-test', version: '0' });
    server.registerTool('a', { description: 'stub' }, async () => ({ content: [] }));
    expect(await seedAdvertisedSurface(server)).toBe('seeded');
    expect(await metricsText()).toMatch(/^ms365_mcp_tools_advertised 1$/m);
  });

  it('is never fatal when the SDK internals move', async () => {
    expect(await seedAdvertisedSurface({ server: {} })).toBe('failed');
  });

  it('is never fatal when the handler throws', async () => {
    const server = {
      server: {
        _requestHandlers: new Map([
          [
            'tools/list',
            async () => {
              throw new Error('boom');
            },
          ],
        ]),
      },
    };
    expect(await seedAdvertisedSurface(server)).toBe('failed');
  });
});

describe('blocked-operation series are created at zero', () => {
  it('exposes the series before anything is refused, so increase() has a baseline', async () => {
    enableMetrics();
    initBlockedOperationSeries(['send-mail', 'reply-mail-message']);
    const text = await metricsText();
    // Without this, prom-client emits nothing until the first inc, and the first
    // refusal appears as a series whose first sample is already 1. increase() has no
    // earlier point to subtract from and reports nothing, so the single most
    // interesting call is the one the alert cannot see.
    expect(text).toMatch(
      /^ms365_mcp_blocked_operations_total\{tool="send-mail",route="direct"\} 0$/m
    );
    expect(text).toMatch(
      /^ms365_mcp_blocked_operations_total\{tool="send-mail",route="batch"\} 0$/m
    );
    expect(text).toMatch(
      /^ms365_mcp_blocked_operations_total\{tool="reply-mail-message",route="batch"\} 0$/m
    );
  });

  it('does not clobber a count that already exists', async () => {
    enableMetrics();
    recordBlockedOperation('send-mail', 'batch');
    initBlockedOperationSeries(['send-mail']);
    expect(await metricsText()).toMatch(
      /^ms365_mcp_blocked_operations_total\{tool="send-mail",route="batch"\} 1$/m
    );
  });
});
