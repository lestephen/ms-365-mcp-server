import { Counter, Gauge, Histogram, Registry, collectDefaultMetrics } from 'prom-client';

/**
 * Prometheus metrics for the MCP server.
 *
 * Deliberately separate from the audit log (`./audit-log.ts`), which records identity
 * per event for DSAR and forensics. Metrics are aggregate and must stay free of personal
 * data: they are widely readable, retained for a long time, and every distinct label
 * value costs a time series forever.
 *
 * PII and cardinality rules, both load-bearing:
 *  - NEVER label with a user, message id, drive item id, file name, subject, or a
 *    resolved URL. Use the endpoint's tool name or pathPattern, which are bounded sets.
 *  - `target_resource.id` exists in the audit log for a reason and must not cross here.
 *  - Every label below has a bounded domain: tool names come from endpoints.json,
 *    routes and outcomes from fixed unions.
 *
 * These exist to answer specific questions with data instead of inference:
 *  - is the --direct-tools set right? (tool_calls_total by route)
 *  - are callers hitting tools that are not registered? (unknown_tool_total; this is
 *    the signal that would have surfaced #29 before a human tripped over it)
 *  - is the mail blocklist ever exercised? (blocked_operations_total)
 *  - is graph-batch actually used, and for what? (batch_subrequests_total)
 */

export type ToolRoute = 'direct' | 'execute_tool' | 'batch';
export type ToolOutcome = 'ok' | 'error' | 'blocked' | 'not_found';
export type DiscoveryStage = 'search_tools' | 'get_tool_schema' | 'execute_tool';

export const registry = new Registry();

const PREFIX = 'ms365_mcp_';

export const toolCalls = new Counter({
  name: `${PREFIX}tool_calls_total`,
  help: 'Tool invocations by tool, the route the caller used, and the outcome.',
  labelNames: ['tool', 'route', 'outcome'] as const,
  registers: [registry],
});

export const toolDuration = new Histogram({
  name: `${PREFIX}tool_duration_seconds`,
  help: 'Wall-clock duration of a tool invocation.',
  labelNames: ['tool'] as const,
  // Graph calls are typically well under a second; the long buckets catch batch and
  // large list operations without needing a second histogram.
  buckets: [0.05, 0.1, 0.25, 0.5, 1, 2.5, 5, 10, 30],
  registers: [registry],
});

export const unknownToolCalls = new Counter({
  name: `${PREFIX}unknown_tool_total`,
  help: 'tools/call for a name this server does not register. In hybrid mode this usually means a caller invoked a discovery-only tool directly instead of via execute-tool.',
  labelNames: ['tool'] as const,
  registers: [registry],
});

export const blockedOperations = new Counter({
  name: `${PREFIX}blocked_operations_total`,
  help: 'Operations refused by --blocked-tools. A nonzero rate is worth alerting on: it means either a caller bug or an attempt to reach a prohibited operation.',
  labelNames: ['tool', 'route'] as const,
  registers: [registry],
});

export const batchSubrequests = new Counter({
  name: `${PREFIX}batch_subrequests_total`,
  help: 'Subrequests carried by graph-batch, labelled with the tool whose operation they match (or "unmatched"). Shows what batching is actually used for.',
  labelNames: ['operation', 'method'] as const,
  registers: [registry],
});

export const discoveryStages = new Counter({
  name: `${PREFIX}discovery_stage_total`,
  help: 'Discovery funnel stages. Searches that never reach execute mean discovery is not landing.',
  labelNames: ['stage'] as const,
  registers: [registry],
});

export const toolsAdvertised = new Gauge({
  name: `${PREFIX}tools_advertised`,
  help: 'Tools registered by name in this configuration, i.e. what a client sees in tools/list.',
  registers: [registry],
});

export const toolSchemaBytes = new Gauge({
  name: `${PREFIX}tool_schema_bytes`,
  help: 'Approximate serialized size of the advertised tool schemas, the up-front context cost of this configuration.',
  registers: [registry],
});

let enabled = false;

/** Metrics are opt-in, so the default configuration carries no overhead. */
export function enableMetrics(): void {
  if (enabled) return;
  enabled = true;
  collectDefaultMetrics({ register: registry, prefix: PREFIX });
}

export function metricsEnabled(): boolean {
  return enabled;
}

/** Recording helpers are no-ops while metrics are off, so call sites need no guard. */
export function recordToolCall(
  tool: string,
  route: ToolRoute,
  outcome: ToolOutcome,
  durationSeconds?: number
): void {
  if (!enabled) return;
  toolCalls.inc({ tool, route, outcome });
  if (durationSeconds !== undefined) toolDuration.observe({ tool }, durationSeconds);
}

/**
 * The unknown-tool name is parsed out of the SDK's "Tool <name> not found" error, so its
 * value set is whatever a caller invents. prom-client retains every distinct label
 * combination for the life of the process, which makes an uncapped label an
 * out-of-memory primitive on a hosted multi-user endpoint: a client looping tools/call
 * with random names grows the registry without bound and inflates every scrape.
 *
 * Cap the distinct names rather than dropping the label. The point of this metric (#29)
 * is learning WHICH real tool names callers reach for directly, so the name has to
 * survive in the normal case, and a real deployment stays far below the cap. Past it,
 * everything folds into one sentinel series that still shows the volume.
 */
const UNKNOWN_TOOL_LABEL_CAP = 50;
const UNKNOWN_TOOL_OVER_CAP = '__over_cap__';
// Capping the count bounds how many labels exist, not how big each one is: the name is
// matched with /[A-Za-z0-9._-]+/ and has no length limit of its own, so without this a
// caller could retain 50 enormous strings and re-serialise them on every scrape. Real
// tool names are far shorter than this.
const UNKNOWN_TOOL_LABEL_MAX_CHARS = 64;
const seenUnknownTools = new Set<string>();

/** Test seam, mirroring resetAdvertisedSurfaceSeedForTests. */
export function resetUnknownToolLabelCapForTests(): void {
  seenUnknownTools.clear();
}

export function recordUnknownTool(tool: string): void {
  if (!enabled) return;

  let label = tool.slice(0, UNKNOWN_TOOL_LABEL_MAX_CHARS);
  if (!seenUnknownTools.has(label)) {
    if (seenUnknownTools.size >= UNKNOWN_TOOL_LABEL_CAP) {
      label = UNKNOWN_TOOL_OVER_CAP;
    } else {
      seenUnknownTools.add(label);
    }
  }

  unknownToolCalls.inc({ tool: label });
}

export function recordBlockedOperation(tool: string, route: ToolRoute): void {
  if (!enabled) return;
  blockedOperations.inc({ tool, route });
}

/**
 * Create the blocked-operation series at zero before anything is refused.
 *
 * prom-client only emits a labelled counter once it is incremented, so the first
 * refusal for a tool/route appears as a series whose first sample is already 1.
 * increase() has no earlier point to subtract and reports nothing, which makes the
 * very first refused call, and the first after every restart, invisible to an alert
 * watching the rate. That is the one call most worth seeing.
 *
 * Every workaround for this in PromQL trades away something else: comparing against
 * an offset loses reset-awareness, and an instant selector drops a terminated pod's
 * contribution while its samples are still inside the window. Giving the series a
 * zero baseline removes the need for any of them, and plain increase() is then
 * correct across resets, restarts and rolling replacements.
 *
 * Only possible because this label set is bounded: the tool names come from the
 * blocklist pattern and route is direct or batch. unknown_tool_total cannot do this,
 * since its label is whatever name a caller invented.
 */
export function initBlockedOperationSeries(toolNames: readonly string[]): void {
  if (!enabled) return;
  for (const tool of toolNames) {
    for (const route of ['direct', 'batch'] as const) {
      blockedOperations.inc({ tool, route }, 0);
    }
  }
}

/**
 * HTTP methods are a closed set, but this one arrives inside a caller-supplied $batch
 * subrequest, so an unrecognised value is caller-invented cardinality on the same
 * unbounded-label footing as the tool name above. `operation` needs no such treatment:
 * it is resolved against our own endpoint catalogue, or the fixed string "unmatched".
 */
const BATCH_METHODS = new Set(['GET', 'POST', 'PATCH', 'PUT', 'DELETE', 'HEAD', 'OPTIONS']);

export function recordBatchSubrequest(operation: string, method: string): void {
  if (!enabled) return;
  const upper = method.toUpperCase();
  batchSubrequests.inc({ operation, method: BATCH_METHODS.has(upper) ? upper : 'OTHER' });
}

export function recordDiscoveryStage(stage: DiscoveryStage): void {
  if (!enabled) return;
  discoveryStages.inc({ stage });
}

export function recordAdvertisedSurface(toolCount: number, schemaBytes: number): void {
  if (!enabled) return;
  toolsAdvertised.set(toolCount);
  toolSchemaBytes.set(schemaBytes);
}

export async function metricsText(): Promise<string> {
  return registry.metrics();
}

export function contentType(): string {
  return registry.contentType;
}
