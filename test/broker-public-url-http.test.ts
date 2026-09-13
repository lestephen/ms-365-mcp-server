import { afterAll, beforeAll, describe, expect, it } from 'vitest';
import { spawn, type ChildProcess } from 'child_process';
import { createServer, type AddressInfo } from 'net';
import path from 'path';
import { fileURLToPath } from 'url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const repoRoot = path.join(__dirname, '..');

/**
 * MS365_MCP_BROKER_PUBLIC_URL against the real spawned HTTP server.
 *
 * The unit tests prove the resolver. These prove the wiring the resolver feeds: the
 * /download route is registered from the override, OAuth metadata ignores it, and a
 * malformed value stops startup instead of producing a server that mints dead links.
 *
 * Requires `npm run build` first, which CI does before `npm test`.
 */

const children: ChildProcess[] = [];

async function freePort(): Promise<number> {
  const server = createServer();
  await new Promise<void>((resolve, reject) => {
    server.once('error', reject);
    server.listen(0, '127.0.0.1', resolve);
  });
  const { port } = server.address() as AddressInfo;
  await new Promise<void>((resolve) => server.close(() => resolve()));
  return port;
}

function spawnServer(port: number, env: Record<string, string | undefined>): ChildProcess {
  const childEnv: Record<string, string | undefined> = {
    ...process.env,
    MS365_MCP_CLIENT_ID: 'test-client',
  };
  delete childEnv.MS365_MCP_PUBLIC_URL;
  delete childEnv.MS365_MCP_BASE_URL;
  delete childEnv.MS365_MCP_BROKER_PUBLIC_URL;
  for (const [key, value] of Object.entries(env)) {
    if (value !== undefined) childEnv[key] = value;
  }
  const child = spawn(
    'node',
    [
      path.join(repoRoot, 'dist', 'index.js'),
      '--http',
      String(port),
      '--org-mode',
      '--preset',
      'mail',
      '--allow-unauthenticated-discovery',
    ],
    { cwd: repoRoot, stdio: 'ignore', env: childEnv }
  );
  children.push(child);
  return child;
}

async function startServer(env: Record<string, string | undefined>): Promise<string> {
  const port = await freePort();
  const base = `http://127.0.0.1:${port}`;
  spawnServer(port, env);
  const deadline = Date.now() + 30_000;
  for (;;) {
    try {
      const res = await fetch(`${base}/.well-known/oauth-authorization-server`);
      if (res.ok) return base;
    } catch {
      /* not listening yet */
    }
    if (Date.now() > deadline) throw new Error('server did not start within 30s');
    await new Promise((r) => setTimeout(r, 300));
  }
}

let splitBase: string;
let brokerOnlyBase: string;
let neitherBase: string;

beforeAll(async () => {
  [splitBase, brokerOnlyBase, neitherBase] = await Promise.all([
    startServer({
      MS365_MCP_PUBLIC_URL: 'https://oauth.example.com',
      MS365_MCP_BROKER_PUBLIC_URL: 'https://broker.example.com',
    }),
    startServer({ MS365_MCP_BROKER_PUBLIC_URL: 'https://broker.example.com' }),
    startServer({}),
  ]);
}, 45_000);

afterAll(() => {
  for (const child of children) child.kill('SIGKILL');
});

async function metadata(base: string): Promise<Record<string, unknown>> {
  const res = await fetch(`${base}/.well-known/oauth-authorization-server`);
  expect(res.status).toBe(200);
  return (await res.json()) as Record<string, unknown>;
}

describe('MS365_MCP_BROKER_PUBLIC_URL over HTTP', () => {
  it('leaves OAuth metadata on the OAuth public URL', async () => {
    const meta = await metadata(splitBase);
    expect(meta.issuer).toMatch(/^https:\/\/oauth\.example\.com\/?$/);
    expect(String(meta.authorization_endpoint)).toMatch(/^https:\/\/oauth\.example\.com\//);
    expect(JSON.stringify(meta)).not.toContain('broker.example.com');
  });

  it('registers the download route when both URLs are set', async () => {
    const res = await fetch(`${splitBase}/download/unknown-handle`);
    expect(res.status).toBe(404);
    expect(await res.json()).toMatchObject({
      error: expect.stringMatching(/invalid or has expired/),
    });
  });

  it('registers the download route from the override alone, without touching OAuth', async () => {
    const res = await fetch(`${brokerOnlyBase}/download/unknown-handle`);
    expect(res.status).toBe(404);
    expect(await res.json()).toMatchObject({
      error: expect.stringMatching(/invalid or has expired/),
    });

    const meta = await metadata(brokerOnlyBase);
    expect(String(meta.authorization_endpoint)).toMatch(/^http:\/\/127\.0\.0\.1:\d+\//);
    expect(JSON.stringify(meta)).not.toContain('broker.example.com');
  });

  it('does not register the download route when neither URL is set', async () => {
    const res = await fetch(`${neitherBase}/download/unknown-handle`);
    expect(res.status).toBe(404);
    // Express's own fallthrough, not the broker's JSON refusal.
    expect(res.headers.get('content-type') ?? '').not.toMatch(/application\/json/);
  });

  it('exits non-zero at startup on a malformed value', async () => {
    const port = await freePort();
    const child = spawnServer(port, {
      MS365_MCP_PUBLIC_URL: 'https://oauth.example.com',
      MS365_MCP_BROKER_PUBLIC_URL: 'https://broker.example.com/?leak=1',
    });
    const code = await new Promise<number | null>((resolve) => {
      const timer = setTimeout(() => resolve(null), 30_000);
      child.once('exit', (exitCode) => {
        clearTimeout(timer);
        resolve(exitCode);
      });
    });
    expect(code).toBe(1);
  }, 35_000);
});
