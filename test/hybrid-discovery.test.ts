import { beforeEach, describe, expect, it, vi } from 'vitest';
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { registerDiscoveryTools, registerGraphTools } from '../src/graph-tools.js';
import { buildMcpServerInstructions } from '../src/mcp-instructions.js';
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
        alias: 'create-draft-email',
        method: 'POST',
        path: '/me/messages',
        description: 'Create a draft email',
      },
      // Stands in for the long tail: a rarely used tool that should stay reachable
      // without its schema being loaded up front.
      {
        alias: 'update-place',
        method: 'PATCH',
        path: '/places/{id}',
        description: 'Update a place',
      },
      {
        alias: 'list-webinar-sessions',
        method: 'GET',
        path: '/solutions/virtualEvents/webinars/{id}/sessions',
        description: 'List webinar sessions',
      },
    ],
  },
}));

/**
 * Hybrid mode: register the handful of tools skills name directly, and keep
 * search-tools / get-tool-schema / execute-tool alongside them so the long tail is
 * still reachable. Loading all Graph tools costs roughly 260k tokens of schemas
 * even after the request bodies are narrowed, which does not fit a 256k context.
 * Trimming to a preset would fit but would amputate the rarely used tools, so the
 * point of this mode is full reach at a fraction of the up-front cost.
 */
describe('hybrid discovery mode', () => {
  let server: McpServer;
  let graphClient: GraphClient;
  let toolSpy: ReturnType<typeof vi.spyOn>;
  // Graph endpoints go through registerTool upstream; utilities and the discovery
  // triad still go through tool. Both must be observed.
  let registerToolSpy: ReturnType<typeof vi.spyOn>;

  const registeredNames = () => [
    ...toolSpy.mock.calls.map((call) => call[0] as string),
    ...registerToolSpy.mock.calls.map((call) => call[0] as string),
  ];
  const DISCOVERY_TRIAD = ['search-tools', 'get-tool-schema', 'execute-tool'];

  beforeEach(() => {
    server = new McpServer({ name: 'test', version: '0' });
    graphClient = {} as GraphClient;
    toolSpy = vi.spyOn(server, 'tool').mockImplementation(() => ({}) as never);
    registerToolSpy = vi.spyOn(server, 'registerTool').mockImplementation(() => ({}) as never);
  });

  it('registers the direct tools matching the pattern alongside the discovery triad', () => {
    registerDiscoveryTools(server, graphClient, false, true);
    registerGraphTools(server, graphClient, false, '^(get-mail-message|create-draft-email)$', true);

    const names = registeredNames();
    for (const tool of DISCOVERY_TRIAD) expect(names).toContain(tool);
    expect(names).toContain('get-mail-message');
    expect(names).toContain('create-draft-email');
  });

  it('does not load schemas for tools outside the direct pattern', () => {
    registerDiscoveryTools(server, graphClient, false, true);
    registerGraphTools(server, graphClient, false, '^(get-mail-message|create-draft-email)$', true);

    const names = registeredNames();
    expect(names).not.toContain('update-place');
    expect(names).not.toContain('list-webinar-sessions');
  });

  it('keeps every tool reachable through execute-tool, including ones not registered directly', async () => {
    // The discovery registry is built without a filter, so the long tail stays
    // executable even though its schema was never sent to the model.
    registerDiscoveryTools(server, graphClient, false, true);

    const executeTool = toolSpy.mock.calls.find((call) => call[0] === 'execute-tool');
    expect(executeTool).toBeDefined();
    const handler = executeTool![executeTool!.length - 1] as (args: {
      tool_name: string;
      parameters?: Record<string, unknown>;
    }) => Promise<{ content: Array<{ text: string }>; isError?: boolean }>;

    const result = await handler({ tool_name: 'update-place', parameters: {} });
    // Reached dispatch rather than being rejected as unknown. Without credentials it
    // cannot go further, which is enough to prove the tool is in the registry.
    expect(result.content[0].text).not.toContain('Tool not found');
  });

  it('registers only the triad when no direct pattern is given, preserving discovery-only behaviour', () => {
    registerDiscoveryTools(server, graphClient, false, true);

    const names = registeredNames();
    for (const tool of DISCOVERY_TRIAD) expect(names).toContain(tool);
    expect(names).not.toContain('get-mail-message');
    expect(names).not.toContain('create-draft-email');
  });

  it('tells the model that named tools and the discovery triad coexist', () => {
    const hybrid = buildMcpServerInstructions({
      discovery: true,
      directTools: true,
      orgMode: true,
      readOnly: false,
      multiAccount: false,
    });
    const discoveryOnly = buildMcpServerInstructions({
      discovery: true,
      orgMode: true,
      readOnly: false,
      multiAccount: false,
    });

    // Discovery-only must keep saying Graph is reached only via the triad.
    expect(discoveryOnly).toContain('search-tools');
    expect(discoveryOnly).not.toMatch(/already registered directly/i);

    // Hybrid must say both paths exist, or the model will route everything through
    // execute-tool and ignore the named tools it can see.
    expect(hybrid).toContain('search-tools');
    expect(hybrid).toMatch(/already registered directly/i);
  });
});
