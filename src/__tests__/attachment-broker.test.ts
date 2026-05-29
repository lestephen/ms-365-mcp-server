import { describe, it, expect, vi, beforeEach } from 'vitest';

vi.mock('../logger.js', () => ({
  default: { info: vi.fn(), warn: vi.fn(), error: vi.fn(), debug: vi.fn() },
}));
vi.mock('../audit-log.js', () => ({ auditLog: vi.fn() }));

import {
  mintDownloadUrl,
  downloadRouteHandler,
  parseRange,
  isBrokerEnabled,
  __testing,
} from '../attachment-broker.js';

function mockRes() {
  const res: any = { headers: {}, statusCode: 200, body: undefined, ended: false };
  res.setHeader = (k: string, v: string) => {
    res.headers[k.toLowerCase()] = v;
    return res;
  };
  res.status = (c: number) => {
    res.statusCode = c;
    return res;
  };
  res.json = (o: unknown) => {
    res.body = o;
    res.ended = true;
    return res;
  };
  res.end = (b?: unknown) => {
    res.body = b;
    res.ended = true;
    return res;
  };
  return res;
}

function handleFromUrl(url: string): string {
  return url.slice(url.lastIndexOf('/') + 1);
}

describe('attachment-broker', () => {
  beforeEach(() => {
    __testing.reset();
    process.env.MS365_MCP_PUBLIC_URL = 'https://mcp.example.com';
    delete process.env.MS365_MCP_BROKER_TTL_MS;
    delete process.env.MS365_MCP_BROKER_MAX_BYTES;
    delete process.env.MS365_MCP_BROKER_MAX_TOTAL_BYTES;
  });

  describe('parseRange', () => {
    it('returns null when there is no range header', () => {
      expect(parseRange(undefined, 100)).toBeNull();
    });
    it('parses a closed range', () => {
      expect(parseRange('bytes=0-9', 100)).toEqual({ start: 0, end: 9 });
    });
    it('clamps an open-ended range to the last byte', () => {
      expect(parseRange('bytes=10-', 100)).toEqual({ start: 10, end: 99 });
    });
    it('parses a suffix range', () => {
      expect(parseRange('bytes=-20', 100)).toEqual({ start: 80, end: 99 });
    });
    it('flags an unsatisfiable range', () => {
      expect(parseRange('bytes=200-300', 100)).toBe('invalid');
    });
  });

  describe('mintDownloadUrl', () => {
    it('returns undefined when no public base URL is configured', () => {
      delete process.env.MS365_MCP_PUBLIC_URL;
      expect(isBrokerEnabled()).toBe(false);
      expect(
        mintDownloadUrl({ bytes: Buffer.from('x'), contentType: 'text/plain', resourcePath: '/x' })
      ).toBeUndefined();
    });

    it('mints a tokenless URL on the public origin', () => {
      const url = mintDownloadUrl({
        bytes: Buffer.from('hello'),
        contentType: 'text/plain',
        resourcePath: '/me/messages/m1/attachments/a1/$value',
      });
      expect(url).toMatch(/^https:\/\/mcp\.example\.com\/download\/[A-Za-z0-9_-]+$/);
    });

    it('throws when content exceeds the per-item byte cap', () => {
      process.env.MS365_MCP_BROKER_MAX_BYTES = '4';
      expect(() =>
        mintDownloadUrl({ bytes: Buffer.from('toolong'), contentType: 'text/plain', resourcePath: '/x' })
      ).toThrow(/per-item broker limit/);
    });

    it('enforces an aggregate memory budget across live items', () => {
      process.env.MS365_MCP_BROKER_MAX_TOTAL_BYTES = '10';
      const first = mintDownloadUrl({
        bytes: Buffer.from('123456'),
        contentType: 'text/plain',
        resourcePath: '/a',
      });
      expect(first).toBeTruthy();
      expect(__testing.totalBytes()).toBe(6);
      // 6 + 6 > 10 -> rejected
      expect(() =>
        mintDownloadUrl({ bytes: Buffer.from('789012'), contentType: 'text/plain', resourcePath: '/b' })
      ).toThrow(/memory budget exceeded/);
    });

    it('frees budget when an entry is consumed/expired before minting', () => {
      process.env.MS365_MCP_BROKER_MAX_TOTAL_BYTES = '10';
      process.env.MS365_MCP_BROKER_TTL_MS = '1';
      const url = mintDownloadUrl({
        bytes: Buffer.from('123456'),
        contentType: 'text/plain',
        resourcePath: '/a',
      })!;
      // Expire the first entry so the sweep at next mint frees its budget.
      __testing.store.get(handleFromUrl(url))!.expiresAt = Date.now() - 1;
      const second = mintDownloadUrl({
        bytes: Buffer.from('789012'),
        contentType: 'text/plain',
        resourcePath: '/b',
      });
      expect(second).toBeTruthy();
      expect(__testing.totalBytes()).toBe(6);
    });
  });

  describe('downloadRouteHandler', () => {
    it('serves the full bytes (200) for a valid handle', () => {
      const url = mintDownloadUrl({
        bytes: Buffer.from('hello world'),
        contentType: 'text/plain',
        name: 'greeting.txt',
        resourcePath: '/x',
      })!;
      const res = mockRes();
      downloadRouteHandler({ params: { handle: handleFromUrl(url) }, headers: {} } as any, res);

      expect(res.statusCode).toBe(200);
      expect(res.headers['cache-control']).toBe('no-store');
      expect(res.headers['content-type']).toBe('text/plain');
      expect(res.headers['content-length']).toBe(String('hello world'.length));
      expect(Buffer.isBuffer(res.body) ? res.body.toString() : res.body).toBe('hello world');
    });

    it('serves a partial range (206) with Content-Range', () => {
      const url = mintDownloadUrl({
        bytes: Buffer.from('hello world'),
        contentType: 'text/plain',
        resourcePath: '/x',
      })!;
      const res = mockRes();
      downloadRouteHandler(
        { params: { handle: handleFromUrl(url) }, headers: { range: 'bytes=0-4' } } as any,
        res
      );

      expect(res.statusCode).toBe(206);
      expect(res.headers['content-range']).toBe('bytes 0-4/11');
      expect((res.body as Buffer).toString()).toBe('hello');
    });

    it('returns 416 for an unsatisfiable range', () => {
      const url = mintDownloadUrl({
        bytes: Buffer.from('hello'),
        contentType: 'text/plain',
        resourcePath: '/x',
      })!;
      const res = mockRes();
      downloadRouteHandler(
        { params: { handle: handleFromUrl(url) }, headers: { range: 'bytes=99-200' } } as any,
        res
      );
      expect(res.statusCode).toBe(416);
      expect(res.headers['content-range']).toBe('bytes */5');
    });

    it('returns 404 for an unknown handle', () => {
      const res = mockRes();
      downloadRouteHandler({ params: { handle: 'nope' }, headers: {} } as any, res);
      expect(res.statusCode).toBe(404);
      expect(res.body).toMatchObject({ error: expect.stringMatching(/invalid or has expired/) });
    });

    it('returns 404 once the capability has expired', () => {
      process.env.MS365_MCP_BROKER_TTL_MS = '1';
      const url = mintDownloadUrl({
        bytes: Buffer.from('x'),
        contentType: 'text/plain',
        resourcePath: '/x',
      })!;
      // Force expiry deterministically rather than waiting on a timer.
      const handle = handleFromUrl(url);
      const cap = __testing.store.get(handle)!;
      cap.expiresAt = Date.now() - 1;
      const res = mockRes();
      downloadRouteHandler({ params: { handle }, headers: {} } as any, res);
      expect(res.statusCode).toBe(404);
    });
  });
});
