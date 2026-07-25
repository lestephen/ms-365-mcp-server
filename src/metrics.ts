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

export function recordUnknownTool(tool: string): void {
  if (!enabled) return;
  unknownToolCalls.inc({ tool });
}

export function recordBlockedOperation(tool: string, route: ToolRoute): void {
  if (!enabled) return;
  blockedOperations.inc({ tool, route });
}

export function recordBatchSubrequest(operation: string, method: string): void {
  if (!enabled) return;
  batchSubrequests.inc({ operation, method: method.toUpperCase() });
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
