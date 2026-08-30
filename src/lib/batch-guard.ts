import { readFileSync } from 'fs';
import { stripGraphVersionSegment } from './graph-version-prefix.js';
import path from 'path';
import { fileURLToPath } from 'url';
import logger from '../logger.js';

const __dirname = path.dirname(fileURLToPath(import.meta.url));

/**
 * Stop `graph-batch` being used to reach an operation the blocklist prohibits.
 *
 * `--blocked-tools` matches tool NAMES. `graph-batch` is `POST /$batch` carrying
 * arbitrary `{method, url, body}` subrequests, so `POST /me/sendMail` inside a batch
 * reached Graph even with every send tool unreachable by name, and `graph-batch`
 * declares no scopes of its own so `--allowed-scopes` was no backstop either. The
 * drafts-only policy was therefore enforceable against an honest model but not against
 * an injected one, which matters because these skills read customer mail and vendor
 * documents. See EnviroKinetics/ms365-mcp#24.
 *
 * This is deliberately narrow. Rather than authorize method/path pairs on every request,
 * it checks batch subrequests against the operations the blocklist already names.
 * `endpoints.json` carries `method` and `pathPattern` per tool, so the blocked names give
 * the matchers directly. Blocking `graph-batch` outright would have closed the same hole
 * but broken batched Planner work in four skills.
 */

interface EndpointRecord {
  toolName: string;
  method: string;
  pathPattern: string;
}

export interface BlockedOperationMatcher {
  toolName: string;
  method: string;
  pattern: RegExp;
}

export interface BlockedSubrequest {
  id: string;
  method: string;
  url: string;
  toolName: string;
  reason?: string;
}

const endpointsData = JSON.parse(
  readFileSync(path.join(__dirname, '..', 'endpoints.json'), 'utf8')
) as EndpointRecord[];

function escapeRegexLiteral(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

/** Expand placeholders whether they fill a segment or sit inside a Graph function segment. */
function pathSegmentToRegex(segment: string): string {
  const placeholder = /\{[^{}]+\}/g;
  let cursor = 0;
  let expanded = '';
  for (const match of segment.matchAll(placeholder)) {
    const index = match.index ?? 0;
    expanded += escapeRegexLiteral(segment.slice(cursor, index));
    const name = match[0].slice(1, -1);
    // Graph's site-by-path endpoint explicitly accepts a slash-delimited subpath in
    // {path}; other placeholders remain confined to one URL segment.
    expanded += name === 'path' ? '.+' : '[^/]+';
    cursor = index + match[0].length;
  }
  return expanded + escapeRegexLiteral(segment.slice(cursor));
}

/** `/me/messages/{message-id}/reply` -> anchored regex with one segment per parameter. */
function pathPatternToRegex(pathPattern: string): RegExp {
  const escaped = pathPattern.split('/').map(pathSegmentToRegex).join('/');
  // Anchored at both ends so a deeper path under a blocked one is not mistaken for it.
  return new RegExp(`^${escaped}$`, 'i');
}

/**
 * Build matchers for the operations named by a --blocked-tools pattern. Returns an empty
 * list when nothing is blocked, so the guard costs nothing in the default configuration.
 */
export function buildBlockedOperationMatchers(
  blockedToolsPattern?: string
): BlockedOperationMatcher[] {
  if (!blockedToolsPattern) return [];
  let blocked: RegExp;
  try {
    blocked = new RegExp(blockedToolsPattern, 'i');
  } catch (error) {
    // parseArgs and compileBlockedToolsRegex already reject an invalid pattern at
    // startup, so reaching here means a direct caller passed one.
    logger.error(
      `Invalid blocked-tools pattern for batch guard: ${blockedToolsPattern}. ` +
        `Batch subrequests will not be checked: ${(error as Error).message}`
    );
    return [];
  }

  return endpointsData
    .filter((endpoint) => blocked.test(endpoint.toolName))
    .map((endpoint) => ({
      toolName: endpoint.toolName,
      method: endpoint.method.toUpperCase(),
      pattern: pathPatternToRegex(endpoint.pathPattern),
    }));
}

type NormalizedSubrequestUrl = { resource: string } | { error: string };

/** Decode and normalize a batch path without letting encoded separators change its structure. */
function normalizeSubrequestUrl(url: string): NormalizedSubrequestUrl {
  let path: string;
  let absolute = false;
  const trimmed = url.trim();
  try {
    if (/^https?:\/\//i.test(trimmed)) {
      path = new URL(trimmed).pathname;
      absolute = true;
    } else {
      path = trimmed.split(/[?#]/)[0];
    }
  } catch {
    return { error: 'the batch URL is malformed' };
  }

  if (!path.startsWith('/')) path = `/${path}`;
  const canonicalSegments: string[] = [];
  for (const rawSegment of path.split('/')) {
    if (rawSegment === '') continue;
    let decoded: string;
    try {
      decoded = decodeURIComponent(rawSegment);
    } catch {
      return { error: 'the batch URL contains malformed percent-encoding' };
    }
    if (decoded.includes('/') || decoded.includes('\\')) {
      return { error: 'the batch URL contains an encoded path separator' };
    }
    // Fail closed on a second encoded layer. Graph and intermediaries have not always agreed on
    // decode count, so matching a partially decoded path would recreate the same bypass.
    if (/%[0-9a-f]{2}/i.test(decoded)) {
      return { error: 'the batch URL contains nested percent-encoding' };
    }
    if (decoded === '.') continue;
    if (decoded === '..') {
      canonicalSegments.pop();
      continue;
    }
    canonicalSegments.push(decoded);
  }

  // Unconditionally, not just for absolute URLs. A relative `/v1.0/me/sendMail` is a
  // valid batch subrequest URL and used to slip past the blocklist that both
  // `/me/sendMail` and the absolute form hit.
  const versionless = stripGraphVersionSegment(canonicalSegments);
  return { resource: versionless.length > 0 ? `/${versionless.join('/')}` : '/' };
}

/**
 * Return every subrequest in a `$batch` body that matches a blocked operation. Tolerates a
 * malformed body: a caller sending nonsense should get Graph's own error, not a crash here.
 */
export function findBlockedSubrequests(
  body: unknown,
  matchers: BlockedOperationMatcher[]
): BlockedSubrequest[] {
  if (matchers.length === 0) return [];
  if (!body || typeof body !== 'object') return [];
  const requests = (body as { requests?: unknown }).requests;
  if (!Array.isArray(requests)) return [];

  const hits: BlockedSubrequest[] = [];
  for (const [index, entry] of requests.entries()) {
    if (!entry || typeof entry !== 'object') continue;
    const sub = entry as { id?: unknown; method?: unknown; url?: unknown };
    if (typeof sub.url !== 'string' || typeof sub.method !== 'string') continue;

    const method = sub.method.toUpperCase();
    const normalized = normalizeSubrequestUrl(sub.url);
    if ('error' in normalized) {
      hits.push({
        id: typeof sub.id === 'string' ? sub.id : String(sub.id ?? index),
        method,
        url: sub.url,
        toolName: 'invalid-batch-url',
        reason: normalized.error,
      });
      continue;
    }
    const resource = normalized.resource;
    const match = matchers.find((m) => m.method === method && m.pattern.test(resource));
    if (match) {
      hits.push({
        id: typeof sub.id === 'string' ? sub.id : String(sub.id ?? index),
        method,
        url: sub.url,
        toolName: match.toolName,
      });
    }
  }
  return hits;
}

/** Operator-facing message naming each offending subrequest and why it was refused. */
export function describeBlockedSubrequests(hits: BlockedSubrequest[]): string {
  const detail = hits
    .map((h) =>
      h.reason
        ? `  request ${h.id}: ${h.method} ${h.url} refused because ${h.reason}`
        : `  request ${h.id}: ${h.method} ${h.url} matches blocked tool ${h.toolName}`
    )
    .join('\n');
  return (
    'Refused: this batch contains subrequests for operations the server blocks.\n' +
    `${detail}\n` +
    'The blocklist applies to the operation, not just the tool name, so routing it ' +
    'through graph-batch does not bypass it. Remove those subrequests and retry.'
  );
}

/**
 * The matcher list for the WHOLE catalogue, built once.
 *
 * `buildBlockedOperationMatchers('.*')` compiles a path regex for all 330 endpoints, and
 * endpointsData is read once at module load, so the result can never change between
 * calls. Rebuilding it per graph-batch call was pure waste
 * (EnviroKinetics/ms365-mcp#54).
 */
let allMatchersMemo: BlockedOperationMatcher[] | undefined;
export function allOperationMatchers(): BlockedOperationMatcher[] {
  return (allMatchersMemo ??= buildBlockedOperationMatchers('.*'));
}

/** Test seam: drop the memo so a test can observe a rebuild. */
export function resetOperationMatcherMemoForTests(): void {
  allMatchersMemo = undefined;
}
