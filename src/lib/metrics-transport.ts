import type { McpServer } from '@modelcontextprotocol/server';
import { recordAdvertisedSurface, recordUnknownTool } from '../metrics.js';
import logger from '../logger.js';

/**
 * Observe MCP requests at the server's own request handlers to capture two things the
 * tool handlers cannot see for themselves.
 *
 * `unknown_tool_total`: a `tools/call` for a name the server does not register never
 * reaches a tool handler, so the SDK's `tools/call` dispatcher is the only place to see
 * it. This is the signal behind #29, where a caller read a schema via get-tool-schema
 * and then called the name directly.
 *
 * `tools_advertised` / `tool_schema_bytes`: the up-front context cost of the current
 * configuration, measured from what `tools/list` actually returns rather than
 * recomputed.
 *
 * SDK v1 builds observed the transport's `send`. SDK v2 serves 2026-07-28 clients
 * through `createMcpHandler`, which owns its transport, so a transport wrapper would
 * only have seen 2025-era HTTP and stdio traffic and undercounted silently as clients
 * upgraded. Every serving path (legacy HTTP, modern HTTP, stdio) dispatches through
 * the server's request handler map, so decorating the two handlers there covers all
 * of them with one seam. The map is SDK-internal; if it disappears the observer logs
 * and does nothing rather than breaking the server, and the metrics tests fail.
 */

type RequestHandler = (request: unknown, ctx: unknown) => Promise<unknown>;

function requestHandlers(server: { server: unknown }): Map<string, RequestHandler> | undefined {
  const handlers = (server.server as { _requestHandlers?: unknown })._requestHandlers;
  return handlers instanceof Map ? (handlers as Map<string, RequestHandler>) : undefined;
}

const OBSERVED = Symbol('ms365-metrics-observed');

/**
 * Decorate `tools/call` and `tools/list` on a fully registered server. Must run after
 * the last tool is registered, because the SDK installs those handlers lazily on the
 * first registration. Idempotent per server.
 */
export function installMetricsObserver(server: McpServer | { server: unknown }): boolean {
  const handlers = requestHandlers(server as { server: unknown });
  if (!handlers) {
    logger.debug('Skipping metrics observer: request handler map not found');
    return false;
  }

  const call = handlers.get('tools/call');
  if (call && !(call as unknown as Record<symbol, boolean>)[OBSERVED]) {
    const observedCall: RequestHandler = async (request, ctx) => {
      try {
        const result = await call(request, ctx);
        const unknown = unknownToolName({ result });
        if (unknown) recordUnknownTool(unknown);
        return result;
      } catch (error) {
        const unknown = unknownToolName({ error: error as { code?: number; message?: string } });
        if (unknown) recordUnknownTool(unknown);
        throw error;
      }
    };
    (observedCall as unknown as Record<symbol, boolean>)[OBSERVED] = true;
    handlers.set('tools/call', observedCall);
  }

  const list = handlers.get('tools/list');
  if (list && !(list as unknown as Record<symbol, boolean>)[OBSERVED]) {
    const observedList: RequestHandler = async (request, ctx) => {
      const result = await list(request, ctx);
      const tools = (result as { tools?: unknown } | undefined)?.tools;
      if (Array.isArray(tools)) recordAdvertisedSurface(tools.length, JSON.stringify(tools).length);
      return result;
    };
    (observedList as unknown as Record<symbol, boolean>)[OBSERVED] = true;
    handlers.set('tools/list', observedList);
  }
  return true;
}

/**
 * Did this outcome mean "no such tool", and for which tool?
 *
 * SDK v2 throws a ProtocolError with code -32602 and the message
 * "Tool <name> not found" (it also uses -32602 for a disabled tool, which is not an
 * unknown name). SDK v1 returned a result with `isError: true` carrying the same text.
 * Both shapes are accepted so a future SDK change either way still counts.
 *
 * The name is parsed from the message. That is fragile to a rewording, which is why an
 * unrecognised but clearly -32602 not-found failure still counts under an "unparsed"
 * sentinel rather than vanishing.
 */
function unknownToolName(outcome: {
  error?: { code?: number; message?: string };
  result?: unknown;
}): string | undefined {
  const texts: string[] = [];
  if (outcome.error?.code === -32602 && typeof outcome.error.message === 'string') {
    texts.push(outcome.error.message);
  }
  const result = outcome.result as
    | { isError?: boolean; content?: Array<{ text?: unknown }> }
    | undefined;
  if (result?.isError && Array.isArray(result.content)) {
    for (const part of result.content) if (typeof part?.text === 'string') texts.push(part.text);
  }
  for (const text of texts) {
    if (!/not found/i.test(text)) continue;
    const named = /Tool ([A-Za-z0-9._-]+) not found/i.exec(text);
    if (named) return named[1];
    if (text.includes('-32602') || outcome.error?.code === -32602) return 'unparsed';
  }
  return undefined;
}

/**
 * Populate `tools_advertised` / `tool_schema_bytes` without waiting for a client.
 *
 * Both gauges used to be set only by observing an outgoing tools/list, so a pod that
 * had served real traffic but no listing reported 0, which is indistinguishable from
 * advertising nothing (#34). They answer "what does this configuration cost a client
 * before any work happens", which is knowable at startup, so measure it then.
 *
 * Calls the SDK's tools/list handler directly, after the $ref normalizer is installed,
 * so the byte count is what a client actually receives rather than the pre-normalized
 * form. The transport observer still refines both on every real listing.
 *
 * Once per process. In HTTP mode a server is constructed per request, so seeding on
 * every construction would re-serialize the whole surface on each one; the tool
 * surface does not vary between them.
 */
let surfaceSeeded = false;

export async function seedAdvertisedSurface(server: {
  server: unknown;
}): Promise<'seeded' | 'skipped' | 'failed'> {
  if (surfaceSeeded) return 'skipped';
  surfaceSeeded = true;
  const handlers = (
    server.server as {
      _requestHandlers?: Map<
        string,
        (request: unknown, extra: unknown) => Promise<{ tools?: Array<Record<string, unknown>> }>
      >;
    }
  )._requestHandlers;
  const list = handlers?.get('tools/list');
  if (!list) {
    logger.debug('Skipping advertised-surface seed: tools/list handler not found');
    return 'failed';
  }
  try {
    const result = await list(
      { method: 'tools/list', params: {} },
      { signal: new AbortController().signal }
    );
    const tools = result.tools ?? [];
    recordAdvertisedSurface(tools.length, JSON.stringify(tools).length);
    return 'seeded';
  } catch (error) {
    // Never fatal: the observer path still populates these on the first real listing,
    // which is exactly the behaviour this replaces.
    logger.debug('Advertised-surface seed failed', { error: String(error) });
    return 'failed';
  }
}

/** Test seam: the once-guard is process-wide by design. */
export function resetAdvertisedSurfaceSeedForTests(): void {
  surfaceSeeded = false;
}
