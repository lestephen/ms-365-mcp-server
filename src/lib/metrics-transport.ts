import type { Transport } from '@modelcontextprotocol/sdk/shared/transport.js';
import type { JSONRPCMessage } from '@modelcontextprotocol/sdk/types.js';
import { recordAdvertisedSurface, recordUnknownTool } from '../metrics.js';

/**
 * Observe MCP traffic at the transport to capture two things the tool handlers cannot
 * see for themselves.
 *
 * `unknown_tool_total`: a `tools/call` for a name the server does not register is
 * rejected by the SDK before any handler runs, so the only place to see it is the wire.
 * This is the signal behind #29, where a caller read a schema via get-tool-schema and
 * then called the name directly. Correlating the outgoing error back to the incoming
 * request id gives the tool name without parsing it out of an error string, which would
 * break the moment the SDK reworded the message.
 *
 * `tools_advertised` / `tool_schema_bytes`: the up-front context cost of the current
 * configuration, measured from what actually goes out rather than recomputed.
 */

const MAX_TRACKED = 512;

export function withMetricsObserver<T extends Transport>(transport: T): T {
  // Request id -> tool name, for calls still in flight.
  const pending = new Map<string | number, string>();

  const originalSend = transport.send.bind(transport);
  transport.send = (message: JSONRPCMessage, options?: Parameters<Transport['send']>[1]) => {
    observeOutgoing(message, pending);
    return originalSend(message, options);
  };

  // onmessage is assigned by Server.connect, so wrap it lazily: capture whatever the
  // SDK sets and chain ours in front of it.
  let handler: Transport['onmessage'];
  Object.defineProperty(transport, 'onmessage', {
    configurable: true,
    get: () => handler,
    set: (next: Transport['onmessage']) => {
      handler = ((message: JSONRPCMessage, extra?: unknown) => {
        observeIncoming(message, pending);
        return (next as ((m: JSONRPCMessage, e?: unknown) => void) | undefined)?.(message, extra);
      }) as Transport['onmessage'];
    },
  });

  return transport;
}

function observeIncoming(message: unknown, pending: Map<string | number, string>): void {
  if (!message || typeof message !== 'object') return;
  const m = message as { id?: string | number; method?: string; params?: { name?: unknown } };
  if (m.method !== 'tools/call' || m.id === undefined) return;
  const name = m.params?.name;
  if (typeof name !== 'string') return;

  // Bound the map: a client that never gets responses must not leak memory.
  if (pending.size >= MAX_TRACKED) {
    const oldest = pending.keys().next();
    if (!oldest.done) pending.delete(oldest.value);
  }
  pending.set(m.id, name);
}

/**
 * Did this response mean "no such tool"?
 *
 * The SDK does NOT return a JSON-RPC error object for a tools/call naming an unknown
 * tool. It catches internally and returns a normal result carrying
 * `isError: true` with the text "MCP error -32602: Tool <name> not found". Verified
 * against a running server; an earlier version of this file checked `error.code` and
 * silently never fired.
 *
 * Both shapes are accepted so a future SDK change in either direction still counts. The
 * tool name comes from the correlated request rather than from parsing this string,
 * which would break the moment the wording changed.
 */
function isUnknownToolFailure(message: { error?: { code?: number }; result?: unknown }): boolean {
  if (message.error?.code === -32602) return true;
  const result = message.result as
    | { isError?: boolean; content?: Array<{ text?: unknown }> }
    | undefined;
  if (!result?.isError || !Array.isArray(result.content)) return false;
  return result.content.some(
    (part) =>
      typeof part?.text === 'string' && part.text.includes('-32602') && /not found/i.test(part.text)
  );
}

function observeOutgoing(message: unknown, pending: Map<string | number, string>): void {
  if (!message || typeof message !== 'object') return;
  const m = message as {
    id?: string | number;
    error?: { code?: number };
    result?: { tools?: unknown };
  };

  if (m.id !== undefined && pending.has(m.id)) {
    const tool = pending.get(m.id)!;
    pending.delete(m.id);
    if (isUnknownToolFailure(m)) recordUnknownTool(tool);
  }

  if (m.result && Array.isArray((m.result as { tools?: unknown[] }).tools)) {
    const tools = (m.result as { tools: unknown[] }).tools;
    recordAdvertisedSurface(tools.length, JSON.stringify(tools).length);
  }
}
