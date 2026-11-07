import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { InMemoryTransport } from '@modelcontextprotocol/sdk/inMemory.js';
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { describe, expect, it, vi } from 'vitest';
import {
  buildAllowedScopeDiagnostics,
  buildScopesFromEndpoints,
  resolveAuthScopes,
  resolveAuthorizeScopes,
} from '../src/auth.js';
import type { GraphClient } from '../src/graph-client.js';
import {
  isUtilityToolAvailable,
  registerDiscoveryTools,
  registerGraphTools,
  UTILITY_TOOLS,
} from '../src/graph-tools.js';

const enabledTools = '^read-mail-text$';

describe('fixed utility scopes', () => {
  it('requests Mail.Read for a standalone read-mail-text filter through every scope entry point', () => {
    expect(buildScopesFromEndpoints(false, enabledTools, true)).toEqual(['Mail.Read']);
    expect(resolveAuthScopes({ enabledTools })).toEqual(['Mail.Read']);
    expect(resolveAuthorizeScopes({ enabledTools }, undefined)).toEqual(['Mail.Read']);
    expect(buildAllowedScopeDiagnostics({ enabledTools }).permissions).toEqual(['Mail.Read']);
  });

  it('reports inadequate scope coverage and honors blocked tools', () => {
    const options = { enabledTools, allowedScopes: 'User.Read' };
    expect(resolveAuthScopes(options)).toEqual([]);
    expect(buildAllowedScopeDiagnostics(options).disabledTools).toEqual([
      { toolName: 'read-mail-text', requiredScopes: ['Mail.Read'], missingScopes: ['Mail.Read'] },
    ]);
    expect(resolveAuthScopes({ enabledTools, blockedTools: enabledTools })).toEqual([]);
    expect(resolveAuthScopes({ enabledTools, appMode: true })).toEqual([]);
  });

  it('uses the existing hierarchy under the full tool surface', () => {
    const scopes = resolveAuthScopes({ orgMode: true });
    expect(scopes).toContain('Mail.ReadWrite');
    expect(scopes).not.toContain('Mail.Read');
    expect(resolveAuthScopes({ enabledTools, allowedScopes: 'Mail.ReadWrite' })).toEqual([
      'Mail.Read',
    ]);
  });

  it.each(['Mail.Read', 'Mail.ReadWrite'])(
    'keeps read-mail-text available with adequate scopes: %s',
    (allowedScopesValue) => {
      const utility = UTILITY_TOOLS.find((tool) => tool.name === 'read-mail-text')!;
      expect(
        isUtilityToolAvailable(utility, {
          readOnly: true,
          httpMode: true,
          scopeMode: 'org',
          allowedScopesValue,
        })
      ).toBe(true);
    }
  );

  for (const mode of ['direct', 'discovery', 'eager'] as const) {
    it.each(['User.Read', '', 'Mail.ReadBasic'])(
      `${mode} denies read-mail-text without its body-read scope: %j`,
      async (allowedScopes) => {
        const graphRequest = vi.fn();
        const graph = { graphRequest } as unknown as GraphClient;
        const server = new McpServer({ name: 'utility-scope-test', version: '1' });
        const options = { enabledTools: '^(read-mail-text|parse-teams-url)$', allowedScopes };
        if (mode === 'direct') registerGraphTools(server, graph, options);
        else
          registerDiscoveryTools(server, graph, {
            ...options,
            ...(mode === 'eager' ? { eagerTools: ['read-mail-text'] } : {}),
          });
        const [a, b] = InMemoryTransport.createLinkedPair();
        await server.connect(a);
        const client = new Client({ name: 'test', version: '1' });
        await client.connect(b);
        try {
          expect((await client.listTools()).tools.map((tool) => tool.name)).not.toContain(
            'read-mail-text'
          );
          const params = { messageIds: ['message'] };
          expect(
            (await client.callTool({ name: 'read-mail-text', arguments: params })).isError
          ).toBe(true);
          if (mode !== 'direct')
            expect(
              (
                await client.callTool({
                  name: 'execute-tool',
                  arguments: { tool_name: 'read-mail-text', parameters: params },
                })
              ).isError
            ).toBe(true);
          expect(graphRequest).not.toHaveBeenCalled();
        } finally {
          await client.close();
          await server.close();
        }
      }
    );
  }
});
