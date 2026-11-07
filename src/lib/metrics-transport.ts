import type { Transport } from '@modelcontextprotocol/sdk/shared/transport.js';
import type { JSONRPCMessage } from '@modelcontextprotocol/sdk/types.js';
import { recordAdvertisedSurface, recordUnknownTool } from '../metrics.js';
import logger from '../logger.js';

/**
 * Observe MCP traffic at the transport to capture two things the tool handlers cannot
 * see for themselves.
 *
 * `unknown_tool_total`: a `tools/call` for a name the server does not register never
 * reaches a handler, so the only place to see it is the wire. This is the signal behind
 * #29, where a caller read a schema via get-tool-schema and then called the name
 * directly.
 *
 * ONLY `send` is wrapped. An earlier version also intercepted `onmessage` via
 * Object.defineProperty to correlate the response back to the request id, which avoided
 * parsing the tool name out of an error string. It also broke
 * StreamableHTTPServerTransport outright: every HTTP request hung with no response,
 * while stdio was unaffected. Production runs HTTP. Wrapping `send` alone is the pattern
 * already proven there by withStrictToolSchemas, so the tool name is parsed from the
 * message and falls back to an "unparsed" sentinel rather than being lost.
 *
 * `tools_advertised` / `tool_schema_bytes`: the up-front context cost of the current
 * configuration, measured from what actually goes out rather than recomputed.
 */

export function withMetricsObserver<T extends Transport>(transport: T): T {
  const originalSend = transport.send.bind(transport);
  transport.send = (message: JSONRPCMessage, options?: Parameters<Transport['send']>[1]) => {
    observeOutgoing(message);
    return originalSend(message, options);
  };
  return transport;
}

/**
 * Did this response mean "no such tool", and for which tool?
 *
 * The SDK does not return a JSON-RPC error object here: it returns a normal result
 * carrying `isError: true` with the text "MCP error -32602: Tool <name> not found".
 * Verified against a running server; an earlier version checked `error.code` only and
 * silently never fired. Both shapes are accepted so a future SDK change either way
 * still counts.
 *
 * The name is parsed from the message. That is fragile to a rewording, which is why an
 * unrecognised but clearly -32602 failure still counts under an "unparsed" sentinel
 * rather than vanishing. The alternative, correlating against the incoming request id,
 * required intercepting `onmessage` and broke the HTTP transport.
 */
function unknownToolName(message: {
  error?: { code?: number; message?: string };
  result?: unknown;
}): string | undefined {
  const texts: string[] = [];
  if (message.error?.code === -32602 && typeof message.error.message === 'string') {
    texts.push(message.error.message);
  }
  const result = message.result as
    | { isError?: boolean; content?: Array<{ text?: unknown }> }
    | undefined;
  if (result?.isError && Array.isArray(result.content)) {
    for (const part of result.content) if (typeof part?.text === 'string') texts.push(part.text);
  }
  for (const text of texts) {
    if (!/not found/i.test(text)) continue;
    const named = /Tool ([A-Za-z0-9._-]+) not found/i.exec(text);
    if (named) return named[1];
    // Recognisable as an unknown-tool failure but not parseable: still counted, under a
    // sentinel, so the signal is not lost silently.
    if (text.includes('-32602')) return 'unparsed';
  }
  return undefined;
}

function observeOutgoing(message: unknown): void {
  if (!message || typeof message !== 'object') return;
  const m = message as {
    id?: string | number;
    error?: { code?: number; message?: string };
    result?: { tools?: unknown };
  };

  const unknown = unknownToolName(m);
  if (unknown) recordUnknownTool(unknown);

  if (m.result && Array.isArray((m.result as { tools?: unknown[] }).tools)) {
    const tools = (m.result as { tools: unknown[] }).tools;
    recordAdvertisedSurface(tools.length, JSON.stringify(tools).length);
  }
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
