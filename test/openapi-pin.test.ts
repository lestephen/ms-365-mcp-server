import { describe, expect, it, vi } from 'vitest';
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from 'fs';
import { createHash } from 'crypto';
import { tmpdir } from 'os';
import path from 'path';
import { fileURLToPath } from 'url';
// @ts-expect-error - bin modules are plain ESM JavaScript with no type declarations
import {
  downloadGraphOpenAPI,
  readSpecPin,
  resolveSpecRef,
  specUrl,
  verifyDownload,
} from '../bin/modules/download-openapi.mjs';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const repoRoot = path.join(__dirname, '..');

/**
 * EnviroKinetics/ms365-mcp#27.
 *
 * `generate` fetched refs/heads/master, so CI built against whatever Microsoft had
 * published that minute while a developer checkout silently reused an older copy on
 * disk, and nothing recorded which spec was used. Worse, the download wrote
 * `response.text()` straight out with no integrity check, so a truncated body became a
 * malformed spec that only surfaced later as a generated client missing schemas.
 */
describe('the Graph spec is pinned', () => {
  it('pins an immutable upstream commit, not a moving branch', () => {
    const pin = readSpecPin(repoRoot);
    expect(pin.repo).toBe('microsoftgraph/msgraph-metadata');
    // A 40-char hex commit, so the fetch is reproducible.
    expect(pin.ref).toMatch(/^[0-9a-f]{40}$/);
    expect(pin.ref).not.toMatch(/master|main|refs\/heads/);
  });

  it('records the size and digest of both specs', () => {
    const pin = readSpecPin(repoRoot);
    for (const version of ['v1.0', 'beta']) {
      const spec = pin.specs[version];
      expect(spec, version).toBeDefined();
      expect(spec.sha256, version).toMatch(/^[0-9a-f]{64}$/);
      expect(spec.bytes, version).toBeGreaterThan(1_000_000);
    }
  });

  it('builds a commit-pinned URL rather than a branch URL', () => {
    const pin = readSpecPin(repoRoot);
    const url = specUrl(pin, 'v1.0');
    expect(url).toContain(pin.ref);
    expect(url).not.toContain('refs/heads');
  });
});

describe('a corrupt or unexpected download is rejected', () => {
  const pin = readSpecPin(repoRoot);
  const expected = pin.specs['v1.0'];

  it('accepts content matching the pinned size and digest', () => {
    const body = 'hello spec';
    const good = {
      bytes: Buffer.byteLength(body),
      sha256: createHash('sha256').update(body).digest('hex'),
    };
    expect(() => verifyDownload('v1.0', Buffer.from(body), good)).not.toThrow();
  });

  it('rejects a truncated body, which is the silent failure mode', () => {
    // A short read used to be written to disk and only surfaced later as a generated
    // client missing schemas.
    expect(() => verifyDownload('v1.0', Buffer.from('too short'), expected)).toThrow(/bytes/i);
  });

  it('rejects content of the right length but the wrong digest', () => {
    const body = Buffer.alloc(expected.bytes, 0x20);
    expect(() => verifyDownload('v1.0', body, expected)).toThrow(/sha256|digest/i);
  });

  it('names the version in the failure, so the message points at the right spec', () => {
    expect(() => verifyDownload('beta', Buffer.from('x'), expected)).toThrow(/beta/);
  });
});

describe('refreshing the Graph specs', () => {
  it('resolves master once and downloads both specs from that immutable SHA', async () => {
    const sha = '1234567890abcdef1234567890abcdef12345678';
    const resolveFetch = vi.fn().mockResolvedValue(
      new Response(JSON.stringify({ sha }), {
        status: 200,
        headers: { 'content-type': 'application/json' },
      })
    );
    const resolved = await resolveSpecRef(
      'microsoftgraph/msgraph-metadata',
      'master',
      resolveFetch
    );

    expect(resolved).toBe(sha);
    expect(resolveFetch).toHaveBeenCalledTimes(1);

    const scratch = mkdtempSync(path.join(tmpdir(), 'openapi-refresh-'));
    const specFetch = vi.fn(
      async (url: string) =>
        new Response(url.includes('/v1.0/') ? 'v1 refreshed' : 'beta refreshed', { status: 200 })
    );
    const log = vi.spyOn(console, 'log').mockImplementation(() => undefined);
    try {
      writeFileSync(
        path.join(scratch, 'openapi-pin.json'),
        JSON.stringify({
          repo: 'microsoftgraph/msgraph-metadata',
          ref: 'a'.repeat(40),
          specs: {
            'v1.0': { path: 'openapi/v1.0/openapi.yaml' },
            beta: { path: 'openapi/beta/openapi.yaml' },
          },
        })
      );
      const targetDir = path.join(scratch, 'openapi');

      await downloadGraphOpenAPI(targetDir, path.join(targetDir, 'v1.yaml'), 'v1.0', {
        repoRoot: scratch,
        refreshSpec: true,
        refreshRef: resolved,
        fetchImpl: specFetch,
      });
      await downloadGraphOpenAPI(targetDir, path.join(targetDir, 'beta.yaml'), 'beta', {
        repoRoot: scratch,
        refreshSpec: true,
        refreshRef: resolved,
        fetchImpl: specFetch,
      });

      expect(specFetch).toHaveBeenCalledTimes(2);
      for (const [url] of specFetch.mock.calls) {
        expect(url).toContain(`/${sha}/`);
        expect(url).not.toContain('master');
      }
      expect(log.mock.calls.flat().join('\n')).toContain(`ref ${sha}`);
    } finally {
      log.mockRestore();
      rmSync(scratch, { recursive: true, force: true });
    }
  });

  it('refuses refresh downloads without an immutable resolved SHA', async () => {
    const scratch = mkdtempSync(path.join(tmpdir(), 'openapi-refresh-invalid-'));
    try {
      writeFileSync(
        path.join(scratch, 'openapi-pin.json'),
        JSON.stringify({
          repo: 'microsoftgraph/msgraph-metadata',
          ref: 'a'.repeat(40),
          specs: { 'v1.0': { path: 'openapi/v1.0/openapi.yaml' } },
        })
      );

      await expect(
        downloadGraphOpenAPI(scratch, path.join(scratch, 'v1.yaml'), 'v1.0', {
          repoRoot: scratch,
          refreshSpec: true,
          refreshRef: 'master',
          fetchImpl: vi.fn(),
        })
      ).rejects.toThrow(/immutable.*commit SHA/i);
    } finally {
      rmSync(scratch, { recursive: true, force: true });
    }
  });
});
