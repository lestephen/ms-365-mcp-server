import { describe, expect, it } from 'vitest';
import { readFileSync } from 'fs';
import { createHash } from 'crypto';
import path from 'path';
import { fileURLToPath } from 'url';
// @ts-expect-error - bin modules are plain ESM JavaScript with no type declarations
import { readSpecPin, specUrl, verifyDownload } from '../bin/modules/download-openapi.mjs';

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
