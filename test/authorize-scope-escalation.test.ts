import { afterAll, beforeAll, describe, expect, it } from 'vitest';
import { spawn, type ChildProcess } from 'child_process';
import path from 'path';
import { fileURLToPath } from 'url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const repoRoot = path.join(__dirname, '..');

/**
 * Guards the WIRING of the authorize route's scope resolution.
 *
 * test/blocked-tools-scopes.test.ts proves resolveAuthorizeScopes refuses to widen past
 * the permitted set. It cannot notice when /authorize fails to call it. That is the gap
 * the bug lived in: scope derivation subtracted blocked tools correctly, and the route
 * simply bypassed the derivation whenever the client supplied its own `scope`.
 *
 * The bearer token the client receives is usable directly against Graph, outside every
 * guarded MCP tool path, so a scope the operator blocked must never reach the redirect.
 * That makes this an end-to-end assertion on the real spawned server rather than a unit
 * test with a stubbed route.
 *
 * Requires `npm run build` first, which CI does before `npm test`.
 */

const PORT = 3198;
const BASE = `http://127.0.0.1:${PORT}`;
const BLOCKED =
  '^(send-mail|send-draft-message|send-shared-mailbox-mail|reply-mail-message|reply-all-mail-message|forward-mail-message)$';

let child: ChildProcess;

async function authorizeScopes(requested: string | null): Promise<string[]> {
  const qs = new URLSearchParams({ redirect_uri: 'http://localhost', state: 'test-state' });
  if (requested !== null) qs.set('scope', requested);

  const res = await fetch(`${BASE}/authorize?${qs.toString()}`, { redirect: 'manual' });
  const location = res.headers.get('location');
  if (!location) throw new Error(`no redirect from /authorize (status ${res.status})`);

  const scope = new URL(location).searchParams.get('scope') ?? '';
  return scope.split(/\s+/).filter(Boolean);
}

beforeAll(async () => {
  child = spawn(
    'node',
    [
      path.join(repoRoot, 'dist', 'index.js'),
      '--http',
      String(PORT),
      '--org-mode',
      '--preset',
      'all',
      '--allow-unauthenticated-discovery',
      '--blocked-tools',
      BLOCKED,
    ],
    { cwd: repoRoot, stdio: 'ignore', env: { ...process.env, MS365_MCP_CLIENT_ID: 'test-client' } }
  );

  const deadline = Date.now() + 30_000;
  for (;;) {
    try {
      const res = await fetch(`${BASE}/.well-known/oauth-authorization-server`);
      if (res.ok) break;
    } catch {
      /* not listening yet */
    }
    if (Date.now() > deadline) throw new Error('server did not start within 30s');
    await new Promise((r) => setTimeout(r, 300));
  }
}, 45_000);

afterAll(() => {
  child?.kill('SIGKILL');
});

describe('/authorize cannot be talked into a blocked scope', () => {
  it('strips a client-requested Mail.Send when every send tool is blocked', async () => {
    const scopes = await authorizeScopes('Mail.Send Mail.ReadWrite');

    expect(scopes).not.toContain('Mail.Send');
    expect(scopes).toContain('Mail.ReadWrite');
  });

  it('still omits Mail.Send when the client requests no scope at all', async () => {
    const scopes = await authorizeScopes(null);

    expect(scopes).not.toContain('Mail.Send');
    expect(scopes.length).toBeGreaterThan(0);
  });
});
