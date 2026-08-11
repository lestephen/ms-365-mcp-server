import { beforeEach, describe, expect, it, vi } from 'vitest';
import MicrosoftGraphServer from '../src/server.js';
import type AuthManager from '../src/auth.js';
import GraphClient from '../src/graph-client.js';
import { getCombinedPresetPattern } from '../src/tool-categories.js';

/**
 * Guards the WIRING of registerGraphTools, not its behaviour.
 *
 * test/batch-subrequest-guard.test.ts and test/blocked-tools.test.ts both call the
 * registrars directly with hand-written argument lists, so they verify the guard works
 * when it is handed a blocklist. Neither notices when the caller fails to hand it one.
 *
 * That gap shipped. registerGraphTools takes eleven positional parameters, and upstream
 * v0.132 inserted httpMode at position ten, where our blockedTools had been. The
 * non-hybrid call site conflicted during the rebase and was updated; the hybrid branch
 * is EKI-only code, did not conflict, and silently kept passing ten arguments. So
 * blockedTools landed in the httpMode slot and blockedToolsPattern became undefined,
 * which makes buildBlockedOperationMatchers return [] and skips the graph-batch
 * subrequest guard entirely (#24) in precisely the mode production runs.
 *
 * Nothing caught it: tools were still unregistered by name via withToolBlocklist, so
 * every behavioural assertion still passed, and tsup does not typecheck.
 *
 * These assert on argument positions on purpose. For a function with eleven positional
 * parameters, the position IS the contract.
 */

vi.mock('../src/logger.js', () => ({
  default: { info: vi.fn(), error: vi.fn(), warn: vi.fn(), debug: vi.fn() },
}));

const registerGraphTools = vi.fn().mockReturnValue(0);
const registerDiscoveryTools = vi.fn().mockReturnValue(0);
vi.mock('../src/graph-tools.js', async (importOriginal) => {
  const actual = await importOriginal<typeof import('../src/graph-tools.js')>();
  return {
    ...actual,
    registerGraphTools: (...args: unknown[]) => registerGraphTools(...args),
    registerDiscoveryTools: (...args: unknown[]) => registerDiscoveryTools(...args),
  };
});

const BLOCKED = '^(send-mail|send-draft-message|reply-mail-message)$';
const DIRECT = '^(graph-batch|create-draft-email|get-current-user)$';

// registerGraphTools(server, graphClient, readOnly, enabledTools, orgMode, authManager,
//                    multiAccount, accountNames, allowedScopes, httpMode, blockedTools,
//                    publicBaseUrl)
const HTTP_MODE = 9;
const ENABLED_TOOLS = 3;
const BLOCKED_TOOLS = 10;
const PUBLIC_BASE_URL = 11;
const ARITY = 12;

function buildServer(options: Record<string, unknown>) {
  const authManager = { getToken: vi.fn() } as unknown as AuthManager;
  const server = new MicrosoftGraphServer(authManager, options);
  // createMcpServer is private and normally reached through initialize(), which needs
  // secrets and a real Graph client. The wiring under test is independent of both.
  const internals = server as unknown as {
    graphClient: GraphClient | null;
    createMcpServer: () => unknown;
  };
  internals.graphClient = {} as GraphClient;
  internals.createMcpServer();
}

describe('registerGraphTools wiring', () => {
  beforeEach(() => {
    registerGraphTools.mockClear();
    registerDiscoveryTools.mockClear();
  });

  it('passes the blocklist to the batch guard in hybrid mode', () => {
    buildServer({ http: '3000', discovery: true, directTools: DIRECT, blockedTools: BLOCKED });

    expect(registerGraphTools).toHaveBeenCalledTimes(1);
    const args = registerGraphTools.mock.calls[0];
    // The regression: ten arguments instead of eleven, shifting blockedTools into
    // httpMode and leaving the batch guard with nothing to match against.
    expect(args).toHaveLength(ARITY);
    expect(args[BLOCKED_TOOLS]).toBe(BLOCKED);
    expect(args[HTTP_MODE]).toBe(true);
  });

  it('intersects hybrid direct tools with the preset-enabled surface', () => {
    const mailPreset = getCombinedPresetPattern(['mail']);
    buildServer({
      discovery: true,
      enabledTools: mailPreset,
      directTools: '^get-',
      orgMode: true,
    });

    expect(registerGraphTools).toHaveBeenCalledTimes(1);
    const effectiveDirectTools = registerGraphTools.mock.calls[0][ENABLED_TOOLS];
    expect(typeof effectiveDirectTools).toBe('string');
    const effective = new RegExp(effectiveDirectTools, 'i');
    expect(effective.test('get-mail-message')).toBe(true);
    expect(effective.test('get-calendar-event')).toBe(false);
    expect(effective.test('get-drive-item')).toBe(false);

    // Discovery uses the same intersection for its invocation hints, so it cannot
    // advertise a tool outside the preset as directly callable.
    expect(registerDiscoveryTools.mock.calls[0][11]).toBe(effectiveDirectTools);
  });

  it('passes the blocklist in non-hybrid mode', () => {
    buildServer({ http: '3000', enabledTools: DIRECT, blockedTools: BLOCKED });

    const args = registerGraphTools.mock.calls[0];
    expect(args).toHaveLength(ARITY);
    expect(args[BLOCKED_TOOLS]).toBe(BLOCKED);
    expect(args[HTTP_MODE]).toBe(true);
  });

  it('reports httpMode as a boolean, not a truthy option string', () => {
    // The shifted argument was a regex string, which is truthy, so httpMode was
    // accidentally correct in production HTTP and wrong for stdio with a blocklist:
    // stdioOnly tools would have been filtered out of a stdio server.
    buildServer({ discovery: true, directTools: DIRECT, blockedTools: BLOCKED });

    const args = registerGraphTools.mock.calls[0];
    expect(args[HTTP_MODE]).toBe(false);
    expect(typeof args[HTTP_MODE]).toBe('boolean');
  });

  it('passes the CLI public URL into broker-aware tools', () => {
    buildServer({
      http: '3000',
      publicUrl: 'https://cli.example.com/',
      enabledTools: DIRECT,
    });

    const args = registerGraphTools.mock.calls[0];
    expect(args).toHaveLength(ARITY);
    expect(args[PUBLIC_BASE_URL]).toBe('https://cli.example.com');
  });

  it('passes the environment public URL when the CLI option is absent', () => {
    const previous = process.env.MS365_MCP_PUBLIC_URL;
    process.env.MS365_MCP_PUBLIC_URL = 'https://env.example.com/';
    try {
      buildServer({ http: '3000', enabledTools: DIRECT });

      const args = registerGraphTools.mock.calls[0];
      expect(args).toHaveLength(ARITY);
      expect(args[PUBLIC_BASE_URL]).toBe('https://env.example.com');
    } finally {
      if (previous === undefined) delete process.env.MS365_MCP_PUBLIC_URL;
      else process.env.MS365_MCP_PUBLIC_URL = previous;
    }
  });
});
