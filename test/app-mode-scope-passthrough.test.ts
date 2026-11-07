import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { InMemoryTransport } from '@modelcontextprotocol/sdk/inMemory.js';
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import type { GraphClient } from '../src/graph-client.js';
import { registerDiscoveryTools, registerGraphTools } from '../src/graph-tools.js';
import type { GraphToolRegistrationOptions } from '../src/graph-tools.js';

vi.mock('../src/logger.js', () => ({
  default: { info: vi.fn(), warn: vi.fn(), error: vi.fn() },
}));
vi.mock('../src/lib/attachment-minting.js', () => ({ getAttachmentMinting: () => ({}) }));
const clients: Client[] = [];
const passthrough = [
  'graph-batch',
  'download-bytes',
  'download-bytes-to-file',
  'get-download-url',
  'get-upload-url',
  'upload-attachment-from-file',
];
const enabledTools = `^(${[...passthrough, 'list-users', 'parse-teams-url'].join('|')})$`;
beforeEach(() => vi.stubEnv('MS365_MCP_UPLOAD_DIRS', '/tmp'));
afterEach(async () => {
  vi.unstubAllEnvs();
  await Promise.all(clients.splice(0).map((client) => client.close()));
});

async function connect(
  mode: 'direct' | 'discovery' | 'eager',
  options: GraphToolRegistrationOptions
) {
  const graphRequest = vi.fn().mockResolvedValue({ content: [{ type: 'text', text: '{}' }] });
  const graph = { graphRequest } as unknown as GraphClient;
  const server = new McpServer({ name: 'app-policy-test', version: '1' });
  const settings = { enabledTools, ...options };
  if (mode === 'direct') registerGraphTools(server, graph, settings);
  else
    registerDiscoveryTools(server, graph, {
      ...settings,
      ...(mode === 'eager' ? { eagerTools: [...passthrough, 'list-users'] } : {}),
    });
  const [a, b] = InMemoryTransport.createLinkedPair();
  await server.connect(a);
  const client = new Client({ name: 'test-client', version: '1' });
  await client.connect(b);
  clients.push(client);
  return { client, graphRequest };
}

const batchArgs = {
  body: { requests: [{ id: '1', method: 'DELETE', url: '/users/victim' }] },
  confirm: true,
};

describe('application scope restrictions close generic Graph paths', () => {
  for (const mode of ['direct', 'discovery', 'eager'] as const) {
    it.each(['User.Read.All', ''])(
      `${mode} refuses generic requests with allowed scopes %j`,
      async (allowedScopes) => {
        const { client, graphRequest } = await connect(mode, { appMode: true, allowedScopes });
        const names = (await client.listTools()).tools.map((tool) => tool.name);
        for (const name of passthrough) {
          expect(names).not.toContain(name);
          const args =
            name === 'graph-batch'
              ? batchArgs
              : { target: '/users/victim/messages/message/$value', confirm: true };
          expect((await client.callTool({ name, arguments: args })).isError).toBe(true);
          if (mode !== 'direct') {
            expect(
              (
                await client.callTool({
                  name: 'execute-tool',
                  arguments: { tool_name: name, parameters: args },
                })
              ).isError
            ).toBe(true);
            expect(
              (await client.callTool({ name: 'get-tool-schema', arguments: { tool_name: name } }))
                .isError
            ).toBe(true);
          }
        }
        expect(graphRequest).not.toHaveBeenCalled();
        if (allowedScopes && mode !== 'discovery') expect(names).toContain('list-users');
        if (mode === 'direct') expect(names).toContain('parse-teams-url');
      }
    );
  }

  it.each([{ appMode: true }, { orgMode: true, allowedScopes: 'User.Read.All' }])(
    'preserves passthrough without app scope restrictions: %j',
    async (options) => {
      const { client } = await connect('eager', options);
      const names = (await client.listTools()).tools.map((tool) => tool.name);
      for (const name of passthrough) expect(names).toContain(name);
    }
  );
});
