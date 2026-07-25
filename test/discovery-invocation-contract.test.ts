import { beforeEach, describe, expect, it, vi } from 'vitest';
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { registerDiscoveryTools } from '../src/graph-tools.js';
import GraphClient from '../src/graph-client.js';

vi.mock('../src/logger.js', () => ({
  default: { info: vi.fn(), error: vi.fn(), warn: vi.fn() },
}));

vi.mock('../src/generated/client-beta.js', () => ({ api: { endpoints: [] } }));
vi.mock('../src/generated/client.js', () => ({
  api: {
    endpoints: [
      {
        alias: 'get-mail-message',
        method: 'GET',
        path: '/me/messages/{id}',
        description: 'Get a mail message',
      },
      {
        alias: 'delete-mail-attachment',
        method: 'DELETE',
        path: '/me/messages/{id}/attachments/{aid}',
        description: 'Delete a mail attachment',
      },
    ],
  },
}));

/**
 * Regression guard for the failure reported in EnviroKinetics/ms365-mcp#29.
 *
 * get-tool-schema answered with `{ name, method, path, parameters }` and nothing
 * about how to invoke the thing. In hybrid mode some tools really are callable by
 * name and some are reachable only through execute-tool, so a payload whose `name`
 * field looks like a callable tool leads a model to call it directly and get
 * `MCP error -32602: Tool ... not found`. That is what happened with
 * delete-mail-attachment: the schema loaded, the call failed, and the caller had to
 * discover the absence at call time on a destructive operation.
 *
 * The discovery payload therefore has to state the invocation route per tool.
 */
describe('discovery payload states how to invoke each tool', () => {
  let server: McpServer;
  let graphClient: GraphClient;
  let toolSpy: ReturnType<typeof vi.spyOn>;

  const DIRECT = '^(get-mail-message)$';

  function handlerFor(name: string) {
    const call = toolSpy.mock.calls.find((c) => c[0] === name);
    if (!call) throw new Error(`${name} not registered`);
    return call[call.length - 1] as (
      args: Record<string, unknown>
    ) => Promise<{ content: Array<{ text: string }>; isError?: boolean }>;
  }

  const parse = async (tool: string, args: Record<string, unknown>) =>
    JSON.parse((await handlerFor(tool)(args)).content[0].text);

  beforeEach(() => {
    server = new McpServer({ name: 'test', version: '0' });
    graphClient = {} as GraphClient;
    toolSpy = vi.spyOn(server, 'tool').mockImplementation(() => ({}) as never);
  });

  it('tells the caller to use execute-tool for a tool that is not registered by name', async () => {
    registerDiscoveryTools(
      server,
      graphClient,
      false,
      true,
      undefined,
      false,
      [],
      undefined,
      undefined,
      undefined,
      DIRECT
    );

    const schema = await parse('get-tool-schema', { tool_name: 'delete-mail-attachment' });
    expect(schema.invocation).toBeDefined();
    expect(schema.invocation.via).toBe('execute-tool');
    // The example has to carry the exact tool_name, since that is the field a caller
    // gets wrong when it guesses.
    expect(schema.invocation.example.tool_name).toBe('delete-mail-attachment');
    expect(JSON.stringify(schema.invocation)).toMatch(
      /not registered|not a named tool|cannot be called directly/i
    );
  });

  it('says a directly registered tool can be called by name', async () => {
    registerDiscoveryTools(
      server,
      graphClient,
      false,
      true,
      undefined,
      false,
      [],
      undefined,
      undefined,
      undefined,
      DIRECT
    );

    const schema = await parse('get-tool-schema', { tool_name: 'get-mail-message' });
    expect(schema.invocation.via).toBe('direct');
  });

  it('defaults to execute-tool when no direct pattern is configured', async () => {
    // Discovery-only mode: nothing is registered by name, so every tool routes
    // through execute-tool.
    registerDiscoveryTools(server, graphClient, false, true);

    for (const tool of ['get-mail-message', 'delete-mail-attachment']) {
      const schema = await parse('get-tool-schema', { tool_name: tool });
      expect(schema.invocation.via, `${tool} in discovery-only mode`).toBe('execute-tool');
    }
  });

  it('marks the route on every search-tools result, so the schema fetch is not the first hint', async () => {
    registerDiscoveryTools(
      server,
      graphClient,
      false,
      true,
      undefined,
      false,
      [],
      undefined,
      undefined,
      undefined,
      DIRECT
    );

    const found = await parse('search-tools', { query: 'mail attachment message' });
    const results = found.results ?? found.tools ?? [];
    expect(results.length).toBeGreaterThan(0);
    for (const r of results) {
      expect(r.invoke_via, `result ${r.name}`).toMatch(/^(direct|execute-tool)$/);
    }
    const byName = Object.fromEntries(results.map((r: { name: string }) => [r.name, r]));
    if (byName['get-mail-message']) expect(byName['get-mail-message'].invoke_via).toBe('direct');
    if (byName['delete-mail-attachment'])
      expect(byName['delete-mail-attachment'].invoke_via).toBe('execute-tool');
  });
});
