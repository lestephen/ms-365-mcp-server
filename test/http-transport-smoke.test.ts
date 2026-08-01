import { afterAll, beforeAll, describe, expect, it } from 'vitest';
import { spawn, type ChildProcess } from 'child_process';
import path from 'path';
import { fileURLToPath } from 'url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const repoRoot = path.join(__dirname, '..');

/**
 * Boot the server over HTTP and make a real request.
 *
 * This exists because a change shipped that broke every authenticated HTTP call and
 * reached production. The suite passed and a manual end-to-end check passed, both over
 * STDIO. Production runs streamable HTTP, and the fault was specific to that transport:
 * wrapping `transport.onmessage` hung `StreamableHTTPServerTransport` while leaving
 * stdio untouched.
 *
 * Unit tests with a fake transport cannot catch that class of bug by construction. This
 * spawns the built server and speaks HTTP to it, so anything that stalls the real
 * transport fails here instead of in the cluster.
 *
 * Requires `npm run build` first, which CI does before `npm test`.
 */

const PORT = 3199;
const METRICS_PORT = 9499;
const BASE = `http://127.0.0.1:${PORT}`;

let child: ChildProcess;

async function post(body: unknown, timeoutMs = 8000): Promise<{ status: number; ms: number }> {
  const started = Date.now();
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), timeoutMs);
  try {
    const res = await fetch(`${BASE}/mcp`, {
      method: 'POST',
      headers: {
        'content-type': 'application/json',
        accept: 'application/json, text/event-stream',
      },
      body: JSON.stringify(body),
      signal: controller.signal,
    });
    // Drain, so a transport that returns headers but never finishes the body is caught.
    await res.text();
    return { status: res.status, ms: Date.now() - started };
  } finally {
    clearTimeout(timer);
  }
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
      'mail',
      '--allow-unauthenticated-discovery',
      // Metrics on, because the transport decorators are the risk this guards.
      '--metrics',
      String(METRICS_PORT),
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

describe('HTTP transport answers real requests', () => {
  it('responds to initialize promptly', async () => {
    const { status, ms } = await post({
      jsonrpc: '2.0',
      id: 1,
      method: 'initialize',
      params: {
        protocolVersion: '2025-06-18',
        capabilities: {},
        clientInfo: { name: 't', version: '0' },
      },
    });
    expect(status).toBe(200);
    // The failure mode was a hang, not a wrong answer, so assert on latency too. A
    // healthy local response is tens of milliseconds; 5s means something stalled.
    expect(ms).toBeLessThan(5000);
  });

  it('responds to tools/list promptly', async () => {
    const { status, ms } = await post({ jsonrpc: '2.0', id: 2, method: 'tools/list', params: {} });
    expect(status).toBe(200);
    expect(ms).toBeLessThan(5000);
  });

  it('serves metrics on the separate port', async () => {
    const res = await fetch(`http://127.0.0.1:${METRICS_PORT}/metrics`);
    expect(res.status).toBe(200);
    expect(await res.text()).toContain('ms365_mcp_');
  });

  it('does not serve metrics on the MCP port', async () => {
    // The MCP port is publicly published; operational detail must not be reachable there.
    const res = await fetch(`${BASE}/metrics`);
    expect(res.status).toBe(404);
  });
});
