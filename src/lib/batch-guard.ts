import { readFileSync } from 'fs';
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
}

const endpointsData = JSON.parse(
  readFileSync(path.join(__dirname, '..', 'endpoints.json'), 'utf8')
) as EndpointRecord[];

/** `/me/messages/{message-id}/reply` -> anchored regex with one segment per parameter. */
function pathPatternToRegex(pathPattern: string): RegExp {
  const escaped = pathPattern
    .split('/')
    .map((segment) =>
      /^\{.+\}$/.test(segment) ? '[^/]+' : segment.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
    )
    .join('/');
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

/** Strip an absolute Graph prefix and any query string, leaving the resource path. */
function normalizeSubrequestUrl(url: string): string {
  let out = url.trim();
  out = out.replace(/^https?:\/\/[^/]+\/(v1\.0|beta)/i, '');
  out = out.split(/[?#]/)[0];
  if (!out.startsWith('/')) out = `/${out}`;
  return out.replace(/\/+$/, '') || '/';
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
    const resource = normalizeSubrequestUrl(sub.url);
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
    .map((h) => `  request ${h.id}: ${h.method} ${h.url} matches blocked tool ${h.toolName}`)
    .join('\n');
  return (
    'Refused: this batch contains subrequests for operations the server blocks.\n' +
    `${detail}\n` +
    'The blocklist applies to the operation, not just the tool name, so routing it ' +
    'through graph-batch does not bypass it. Remove those subrequests and retry.'
  );
}
