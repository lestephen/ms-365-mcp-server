import { beforeEach, describe, expect, it, vi } from 'vitest';
import MicrosoftGraphServer, { resolvePublicBaseUrl } from '../src/server.js';
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
 * That gap shipped. registerGraphTools used to take eleven positional parameters, and upstream
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
 * Upstream v0.160 moved both registrars to an options object, which retires the
 * positional hazard itself. These still assert that the call sites hand the blocklist
 * and httpMode through, because omitting a property from an object is the
 * same silent failure as omitting an argument: the type makes every one optional.
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

// registerGraphTools(server, graphClient, options) and registerDiscoveryTools likewise.
type RegistrationOptions = Record<string, unknown>;
const optionsOf = (mock: typeof registerGraphTools, call = 0): RegistrationOptions =>
  mock.mock.calls[call][2] as RegistrationOptions;

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
    const options = optionsOf(registerGraphTools);
    // The regression: the hybrid call site dropping the blocklist and leaving the
    // batch guard with nothing to match against.
    expect(options.blockedTools).toBe(BLOCKED);
    expect(options.httpMode).toBe(true);
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
    const effectiveDirectTools = optionsOf(registerGraphTools).enabledTools as (
      name: string
    ) => boolean;
    expect(typeof effectiveDirectTools).toBe('function');
    expect(effectiveDirectTools('get-mail-message')).toBe(true);
    expect(effectiveDirectTools('get-calendar-event')).toBe(false);
    expect(effectiveDirectTools('get-drive-item')).toBe(false);

    // Discovery uses the same intersection for its invocation hints, so it cannot
    // advertise a tool outside the preset as directly callable.
    expect(optionsOf(registerDiscoveryTools).directTools).toBe(effectiveDirectTools);
  });

  it('intersects independently valid regexes with duplicate named captures', () => {
    buildServer({
      discovery: true,
      enabledTools: '^(?<kind>get)-mail',
      directTools: '^(?<kind>get)-',
      orgMode: true,
    });

    const effective = optionsOf(registerGraphTools).enabledTools as (name: string) => boolean;
    expect(typeof effective).toBe('function');
    expect(effective('get-mail-message')).toBe(true);
    expect(effective('get-calendar-event')).toBe(false);
    expect(optionsOf(registerDiscoveryTools).directTools).toBe(effective);
  });

  it('passes the blocklist in non-hybrid mode', () => {
    buildServer({ http: '3000', enabledTools: DIRECT, blockedTools: BLOCKED });

    const options = optionsOf(registerGraphTools);
    expect(options.blockedTools).toBe(BLOCKED);
    expect(options.httpMode).toBe(true);
  });

  it('reports httpMode as a boolean, not a truthy option string', () => {
    // The shifted argument was a regex string, which is truthy, so httpMode was
    // accidentally correct in production HTTP and wrong for stdio with a blocklist:
    // stdioOnly tools would have been filtered out of a stdio server.
    buildServer({ discovery: true, directTools: DIRECT, blockedTools: BLOCKED });

    const options = optionsOf(registerGraphTools);
    expect(options.httpMode).toBe(false);
    expect(typeof options.httpMode).toBe('boolean');
  });

  it('resolves the OAuth public URL from the CLI option or the environment', () => {
    const previous = process.env.MS365_MCP_PUBLIC_URL;
    process.env.MS365_MCP_PUBLIC_URL = 'https://oauth.example.com/';
    try {
      expect(resolvePublicBaseUrl({})).toBe('https://oauth.example.com');
      expect(resolvePublicBaseUrl({ publicUrl: 'https://cli.example.com/' })).toBe(
        'https://cli.example.com'
      );
      delete process.env.MS365_MCP_PUBLIC_URL;
      expect(resolvePublicBaseUrl({})).toBeUndefined();
    } finally {
      if (previous === undefined) delete process.env.MS365_MCP_PUBLIC_URL;
      else process.env.MS365_MCP_PUBLIC_URL = previous;
    }
  });
});
