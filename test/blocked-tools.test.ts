import { beforeEach, describe, expect, it, vi } from 'vitest';
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import {
  buildToolsRegistry,
  registerDiscoveryTools,
  registerGraphTools,
} from '../src/graph-tools.js';
import { withToolBlocklist } from '../src/lib/tool-blocklist.js';
import { registerAuthTools } from '../src/auth-tools.js';
import type AuthManager from '../src/auth.js';
import GraphClient from '../src/graph-client.js';

vi.mock('../src/logger.js', () => ({
  default: { info: vi.fn(), error: vi.fn(), warn: vi.fn() },
}));

vi.mock('../src/generated/client-beta.js', () => ({ api: { endpoints: [] } }));
vi.mock('../src/generated/client.js', () => ({
  api: {
    endpoints: [
      {
        alias: 'create-draft-email',
        method: 'POST',
        path: '/me/messages',
        description: 'Create a draft email',
      },
      { alias: 'send-mail', method: 'POST', path: '/me/sendMail', description: 'Send mail' },
      {
        alias: 'reply-mail-message',
        method: 'POST',
        path: '/me/messages/{id}/reply',
        description: 'Reply to a message',
      },
      {
        alias: 'get-mail-message',
        method: 'GET',
        path: '/me/messages/{id}',
        description: 'Get a mail message',
      },
    ],
  },
}));

/**
 * `execute-tool` dispatches by name, so a client-side deny rule on a tool name
 * (for example a policy of drafting mail but never sending it) is bypassed the
 * moment discovery mode is on: the call the client sees is `execute-tool`, not
 * `send-mail`. A blocklist enforced inside the server closes that, and it binds
 * every client rather than only the one that happens to ship the deny rule.
 */
const BLOCKED = '^(send-mail|send-draft-message|reply-mail-message|forward-mail-message)$';

describe('blocked tools', () => {
  let server: McpServer;
  let graphClient: GraphClient;
  let toolSpy: ReturnType<typeof vi.spyOn>;

  const registeredNames = () => toolSpy.mock.calls.map((call) => call[0] as string);

  function handlerFor(name: string) {
    const call = toolSpy.mock.calls.find((c) => c[0] === name);
    if (!call) throw new Error(`${name} was not registered`);
    return call[call.length - 1] as (args: Record<string, unknown>) => Promise<{
      content: Array<{ text: string }>;
      isError?: boolean;
    }>;
  }

  beforeEach(() => {
    server = new McpServer({ name: 'test', version: '0' });
    graphClient = {} as GraphClient;
    toolSpy = vi.spyOn(server, 'tool').mockImplementation(() => ({}) as never);
  });

  it('keeps a blocked tool out of the direct registration', () => {
    registerGraphTools(
      server,
      graphClient,
      false,
      undefined,
      true,
      undefined,
      false,
      [],
      undefined,
      BLOCKED
    );

    const names = registeredNames();
    expect(names).toContain('create-draft-email');
    expect(names).toContain('get-mail-message');
    expect(names).not.toContain('send-mail');
    expect(names).not.toContain('reply-mail-message');
  });

  it('keeps a blocked tool out of the discovery registry, so execute-tool cannot reach it', () => {
    const registry = buildToolsRegistry(
      false,
      true,
      undefined,
      undefined,
      [],
      new RegExp(BLOCKED, 'i')
    );

    expect(registry.has('create-draft-email')).toBe(true);
    expect(registry.has('send-mail')).toBe(false);
    expect(registry.has('reply-mail-message')).toBe(false);
  });

  it('reports a blocked tool as not found through execute-tool', async () => {
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
      BLOCKED
    );

    const result = await handlerFor('execute-tool')({ tool_name: 'send-mail', parameters: {} });
    expect(result.content[0].text).toContain('Tool not found');
    expect(result.isError).toBe(true);
  });

  it('reports a blocked tool as not found through get-tool-schema', async () => {
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
      BLOCKED
    );

    const result = await handlerFor('get-tool-schema')({ tool_name: 'send-mail' });
    expect(result.content[0].text).toContain('Tool not found');
  });

  it('does not surface a blocked tool from search-tools', async () => {
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
      BLOCKED
    );

    const result = await handlerFor('search-tools')({ query: 'send mail reply message' });

    // Assert on the tool names the payload actually offers, not on its raw text. An
    // unblocked tool's description may legitimately name a blocked tool: upstream's
    // create-draft-email description cross-references send-mail ("use send-mail to send
    // a message directly", v0.133.2 / #594). A substring match over the whole payload
    // fails on that prose while nothing blocked is reachable, which is a false alarm.
    // What matters is that no blocked name is offered as a callable result.
    const names = (JSON.parse(result.content[0].text).tools as { name: string }[]).map(
      (t) => t.name
    );
    // Positive assertion first, so the check below cannot pass merely because the
    // search returned nothing at all.
    expect(names).toContain('create-draft-email');
    expect(names).not.toContain('send-mail');
    expect(names).not.toContain('reply-mail-message');
  });

  it('still allows an unblocked tool through execute-tool', async () => {
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
      BLOCKED
    );

    const result = await handlerFor('execute-tool')({
      tool_name: 'create-draft-email',
      parameters: {},
    });
    // Reaches dispatch instead of being rejected as unknown.
    expect(result.content[0].text).not.toContain('Tool not found');
  });

  it('blocks a tool even when an enable pattern explicitly selects it', () => {
    registerGraphTools(
      server,
      graphClient,
      false,
      '^(send-mail|create-draft-email)$',
      true,
      undefined,
      false,
      [],
      undefined,
      BLOCKED
    );

    const names = registeredNames();
    expect(names).toContain('create-draft-email');
    expect(names).not.toContain('send-mail');
  });

  it('blocks auth tools too, so the contract holds on every registration path', () => {
    // Four reviewers independently flagged that registerAuthTools was outside the
    // guardrail, so --blocked-tools '^logout$' silently did nothing.
    const guarded = withToolBlocklist(server, '^(logout|remove-account)$');
    registerAuthTools(guarded, {} as AuthManager);

    const names = registeredNames();
    expect(names).toContain('login');
    expect(names).toContain('verify-login');
    expect(names).not.toContain('logout');
    expect(names).not.toContain('remove-account');
  });

  it('can block the generic dispatcher itself', () => {
    // An operator may want the named tools without generic Graph execution.
    const guarded = withToolBlocklist(server, '^execute-tool$');
    registerDiscoveryTools(guarded, graphClient, false, true);

    const names = registeredNames();
    expect(names).toContain('search-tools');
    expect(names).toContain('get-tool-schema');
    expect(names).not.toContain('execute-tool');
  });

  it('registers everything when no blocklist is given', () => {
    const guarded = withToolBlocklist(server, undefined);
    registerAuthTools(guarded, {} as AuthManager);

    expect(registeredNames()).toContain('logout');
  });

  it('refuses to start on an invalid blocklist regex instead of failing open', () => {
    // --enabled-tools logs and ignores a bad pattern, which is fine for a filter.
    // A blocklist is a guardrail: ignoring a typo would silently unblock send-mail,
    // so this has to be loud and fatal.
    expect(() =>
      registerGraphTools(
        server,
        graphClient,
        false,
        undefined,
        true,
        undefined,
        false,
        [],
        undefined,
        '([bad'
      )
    ).toThrow(/blocked-tools/i);
  });
});
