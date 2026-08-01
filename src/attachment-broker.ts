import { randomBytes, randomUUID } from 'node:crypto';
import type { Request, Response } from 'express';
import { auditLog } from './audit-log.js';
import logger from './logger.js';

/**
 * Out-of-band binary broker for Microsoft Graph content that exposes NO native
 * pre-authenticated download URL (mail file attachments, /$value byte endpoints).
 *
 * Flow (EnviroKinetics ms365-mcp#5, Tier 2):
 *  1. Mint — inside an authenticated MCP tool call (which carries the user's
 *     delegated bearer server-side), the bytes are fetched from Graph and held
 *     here under a high-entropy, unguessable handle with a short TTL.
 *  2. Redeem — the client pulls `${MS365_MCP_PUBLIC_URL}/download/{handle}`
 *     with NO Authorization header (the handle is the capability), so bytes
 *     reach local disk without round-tripping base64 through the agent context.
 *
 * Security model: byte custody is captured at mint for the specific authorized
 * request, so redemption needs no credential and there is no confused-deputy at
 * pull time. Exposure rests on the 256-bit unguessable handle + short TTL +
 * `Cache-Control: no-store`; redemption is audited. This is the same trust shape
 * as Graph's own short-lived `@microsoft.graph.downloadUrl` (used for Tier 1).
 *
 * Range/resume is supported within the TTL (multiple GETs per handle), matching
 * Graph's own download URLs; the handle expires by time rather than on first use.
 *
 * State is in-memory and assumes a single replica (the deployment runs
 * `replicas: 1`). If ever scaled out, this needs a shared ephemeral store or
 * sticky routing.
 */

const DEFAULT_TTL_MS = 120_000; // 2 minutes
const DEFAULT_MAX_BYTES = 50 * 1024 * 1024; // 50 MB safety cap on a single brokered item
const DEFAULT_MAX_TOTAL_BYTES = 512 * 1024 * 1024; // 512 MB aggregate cap across all live items
const SWEEP_INTERVAL_MS = 30_000;

function readPositiveEnv(name: string, fallback: number): number {
  const v = Number(process.env[name]);
  return Number.isFinite(v) && v > 0 ? v : fallback;
}

function ttlMs(): number {
  return readPositiveEnv('MS365_MCP_BROKER_TTL_MS', DEFAULT_TTL_MS);
}

function maxBytes(): number {
  return readPositiveEnv('MS365_MCP_BROKER_MAX_BYTES', DEFAULT_MAX_BYTES);
}

function maxTotalBytes(): number {
  return readPositiveEnv('MS365_MCP_BROKER_MAX_TOTAL_BYTES', DEFAULT_MAX_TOTAL_BYTES);
}

/** The externally routable origin clients use to pull brokered bytes. */
export function getPublicBaseUrl(): string | undefined {
  const url = process.env.MS365_MCP_PUBLIC_URL;
  return url && url.trim() ? url.trim().replace(/\/+$/, '') : undefined;
}

/** The broker only works in HTTP mode with a public base URL configured. */
export function isBrokerEnabled(): boolean {
  return !!getPublicBaseUrl();
}

interface Capability {
  bytes: Buffer;
  contentType: string;
  name?: string;
  userPrincipalName?: string;
  resourcePath: string;
  expiresAt: number;
}

const store = new Map<string, Capability>();
let totalBytes = 0;
let sweepTimer: ReturnType<typeof setInterval> | undefined;

function deleteEntry(handle: string): void {
  const cap = store.get(handle);
  if (cap) {
    totalBytes -= cap.bytes.length;
    store.delete(handle);
  }
}

function sweep(): void {
  const now = Date.now();
  for (const [handle, cap] of store) {
    if (cap.expiresAt <= now) {
      deleteEntry(handle);
    }
  }
  if (store.size === 0 && sweepTimer) {
    clearInterval(sweepTimer);
    sweepTimer = undefined;
  }
}

function ensureSweeper(): void {
  if (!sweepTimer) {
    sweepTimer = setInterval(sweep, SWEEP_INTERVAL_MS);
    // Do not keep the process alive solely for the sweep timer.
    if (typeof sweepTimer.unref === 'function') sweepTimer.unref();
  }
}

export interface MintInput {
  bytes: Buffer;
  contentType: string;
  name?: string;
  userPrincipalName?: string;
  resourcePath: string;
}

/**
 * Hold already-fetched bytes and return a short-lived, tokenless download URL.
 * Returns undefined if the broker is not configured. Throws if the content
 * exceeds the configured per-item limit.
 */
export function mintDownloadUrl(input: MintInput): string | undefined {
  const base = getPublicBaseUrl();
  if (!base) return undefined;

  const limit = maxBytes();
  if (input.bytes.length > limit) {
    throw new Error(
      `Content is ${input.bytes.length} bytes, exceeding the per-item broker limit of ${limit} bytes (set MS365_MCP_BROKER_MAX_BYTES to raise it).`
    );
  }

  // Free expired entries first, then enforce an aggregate in-memory budget so a burst of large
  // mints cannot exhaust the single replica's memory (a security-relevant availability backstop
  // on the tokenless download path).
  sweep();
  const totalLimit = maxTotalBytes();
  if (totalBytes + input.bytes.length > totalLimit) {
    throw new Error(
      `Broker memory budget exceeded (${totalBytes} + ${input.bytes.length} > ${totalLimit} bytes; set MS365_MCP_BROKER_MAX_TOTAL_BYTES to raise it). Retry shortly.`
    );
  }

  const handle = randomBytes(32).toString('base64url');
  store.set(handle, {
    bytes: input.bytes,
    contentType: input.contentType || 'application/octet-stream',
    name: input.name,
    userPrincipalName: input.userPrincipalName,
    resourcePath: input.resourcePath,
    expiresAt: Date.now() + ttlMs(),
  });
  totalBytes += input.bytes.length;
  ensureSweeper();

  auditLog({
    event: 'attachment.mint',
    request_id: randomUUID(),
    user_principal_name: input.userPrincipalName,
    tool: 'get-download-url',
    status: 'success',
    target_resource: { type: 'graph.binary', id: input.resourcePath },
  });

  return `${base}/download/${handle}`;
}

/** Look up a live capability without consuming it (range/resume issue many GETs). */
function lookupCapability(handle: string): Capability | undefined {
  const cap = store.get(handle);
  if (!cap) return undefined;
  if (cap.expiresAt <= Date.now()) {
    deleteEntry(handle);
    return undefined;
  }
  return cap;
}

interface ParsedRange {
  start: number;
  end: number;
}

/** Parse a single-range `bytes=` header against a known total length. */
export function parseRange(
  header: string | undefined,
  total: number
): ParsedRange | 'invalid' | null {
  if (typeof header !== 'string') return null;
  const m = /^bytes=(\d*)-(\d*)$/.exec(header.trim());
  if (!m) return null;
  const [, rawStart, rawEnd] = m;
  if (rawStart === '' && rawEnd === '') return 'invalid';
  let start: number;
  let end: number;
  if (rawStart === '') {
    // suffix range: final N bytes
    const n = parseInt(rawEnd, 10);
    if (Number.isNaN(n) || n === 0) return 'invalid';
    start = Math.max(0, total - n);
    end = total - 1;
  } else {
    start = parseInt(rawStart, 10);
    end = rawEnd === '' ? total - 1 : Math.min(parseInt(rawEnd, 10), total - 1);
  }
  if (Number.isNaN(start) || Number.isNaN(end) || start > end || start >= total || start < 0) {
    return 'invalid';
  }
  return { start, end };
}

/** Express handler for `GET /download/:handle`. Tokenless by design. */
export function downloadRouteHandler(req: Request, res: Response): void {
  const handle = req.params.handle;
  const cap = lookupCapability(handle);
  if (!cap) {
    res.status(404).json({ error: 'Download link is invalid or has expired.' });
    return;
  }

  const total = cap.bytes.length;
  res.setHeader('Cache-Control', 'no-store');
  res.setHeader('Content-Type', cap.contentType);
  res.setHeader('Accept-Ranges', 'bytes');
  if (cap.name) {
    res.setHeader('Content-Disposition', `attachment; filename="${encodeURIComponent(cap.name)}"`);
  }

  const parsed = parseRange(req.headers.range, total);
  if (parsed === 'invalid') {
    res.setHeader('Content-Range', `bytes */${total}`);
    res.status(416).end();
    return;
  }

  let status = 200;
  let slice = cap.bytes;
  if (parsed) {
    status = 206;
    slice = cap.bytes.subarray(parsed.start, parsed.end + 1);
    res.setHeader('Content-Range', `bytes ${parsed.start}-${parsed.end}/${total}`);
  }

  res.setHeader('Content-Length', String(slice.length));
  res.status(status);

  // The redeem path is intentionally tokenless, so the redeemer's identity is unknown:
  // user_principal_name records the MINTING user (context), and we capture best-effort
  // network attribution (source IP / user-agent) for forensics should a handle ever leak.
  auditLog({
    event: 'attachment.redeem',
    request_id: randomUUID(),
    user_principal_name: cap.userPrincipalName,
    tool: 'get-download-url',
    http_method: 'GET',
    status: 'success',
    target_resource: { type: 'graph.binary', id: cap.resourcePath },
    source_ip: req.ip,
    user_agent:
      typeof req.headers['user-agent'] === 'string' ? req.headers['user-agent'] : undefined,
  });

  logger.info(`Broker served ${slice.length}/${total} bytes for ${cap.resourcePath}`);
  res.end(slice);
}

// Exposed for tests only.
export const __testing = {
  store,
  sweep,
  lookupCapability,
  reset: () => {
    store.clear();
    totalBytes = 0;
  },
  totalBytes: () => totalBytes,
};
