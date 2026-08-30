import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { CallToolRequestSchema, type ServerResult } from '@modelcontextprotocol/sdk/types.js';
import { randomUUID } from 'crypto';
import logger from './logger.js';
import { compileBlockedToolsRegex } from './lib/tool-blocklist.js';
import {
  recordBatchSubrequest,
  recordBlockedOperation,
  recordDiscoveryStage,
  recordToolCall,
  metricsEnabled,
  type ToolRoute,
  type ToolOutcome,
  initBlockedOperationSeries,
} from './metrics.js';
import { describeInlineBytes, findInlineByteFields } from './lib/inline-bytes-guard.js';
import {
  allOperationMatchers,
  buildBlockedOperationMatchers,
  describeBlockedSubrequests,
  findBlockedSubrequests,
  type BlockedOperationMatcher,
} from './lib/batch-guard.js';
import { auditLog, getUserIdentityForAudit, type AuditEvent } from './audit-log.js';
import GraphClient, { GraphDownloadSizeLimitError } from './graph-client.js';
import { isDestructiveOperation } from './lib/destructive-ops.js';
import { describePathParam } from './lib/path-params.js';
import AuthManager, {
  getEndpointScopeGroups,
  getMissingAllowedScopesForGroups,
  parseAllowedScopes,
} from './auth.js';
import { api } from './generated/client.js';
import { api as betaApi } from './generated/client-beta.js';

// Tools from every Graph API version share one registry. Each tool's version is carried
// by its endpoints.json config (apiVersion), so the generated clients stay version-agnostic
// and the runtime picks the URL prefix per request. v1.0 endpoints are unchanged.
const allEndpoints = [...api.endpoints, ...betaApi.endpoints];
import { z } from 'zod';
import { readFileSync } from 'fs';
import { access } from 'fs/promises';
import path from 'path';
import { fileURLToPath } from 'url';
import { TOOL_CATEGORIES } from './tool-categories.js';
import { getRequestTokens } from './request-context.js';
import {
  getBrokerMaxBytes,
  isBrokerEnabled,
  mintDownloadUrl,
  releaseBrokerCapacity,
  reserveBrokerCapacity,
} from './attachment-broker.js';
import { parseTeamsUrl } from './lib/teams-url-parser.js';
import { buildBM25Index, scoreQuery, tokenize, type BM25Index } from './lib/bm25.js';
import { deriveTargetResource, type AuditTargetResource } from './audit-target-resource.js';
import {
  prepareUnencodedPathParameter,
  refineUnencodedPathParameterSchema,
} from './lib/unencoded-path-params.js';
export interface DiscoverySearchIndex {
  bm25: BM25Index;
  nameTokens: Map<string, Set<string>>;
}
import { describeToolSchema, describeUtilityToolSchema } from './lib/tool-schema.js';
import {
  TOP_UNSUPPORTED_DELTA_TOOLS,
  shouldOmitTopParam,
  paginationAllowed,
  positiveIntFromEnv,
  DEFAULT_MAX_PAGES,
  getMaxPages,
  isFetchAllPagesApplicable,
  FILTER_PARAM_DESCRIPTION,
  SEARCH_PARAM_DESCRIPTION,
  SELECT_PARAM_DESCRIPTION,
  EXPAND_PARAM_DESCRIPTION,
  ORDERBY_PARAM_DESCRIPTION,
  TOP_PARAM_DESCRIPTION,
  SKIP_PARAM_DESCRIPTION,
  COUNT_PARAM_DESCRIPTION,
  CONFIRM_PARAM_DESCRIPTION,
  TIMEZONE_PARAM_DESCRIPTION,
  EXPAND_EXTENDED_PROPERTIES_PARAM_DESCRIPTION,
  getAccountParamDescription,
  getFetchAllPagesParamDescription,
} from './lib/param-descriptions.js';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

interface EndpointConfig {
  pathPattern: string;
  method: string;
  toolName: string;
  scopes?: string[] | string[][];
  workScopes?: string[] | string[][];
  apiVersion?: string; // Graph API version ('v1.0' default, or 'beta'). Selects spec + URL prefix.
  returnDownloadUrl?: boolean;
  supportsTimezone?: boolean;
  supportsExpandExtendedProperties?: boolean;
  llmTip?: string;
  // Replaces the Microsoft-supplied base description everywhere it is surfaced (tool
  // registration, BM25 discovery index, search-tools, get-tool-schema). Use when the
  // generated description leads with the wrong Graph operation. llmTip is still appended after.
  descriptionOverride?: string;
  skipEncoding?: string[]; // Parameter names that should NOT be URL-encoded (for function-style API calls)
  contentType?: string;
  acceptType?: string; // Custom Accept header for endpoints returning non-JSON content (e.g., text/vtt)
  readOnly?: boolean; // When true, allow this endpoint in read-only mode even if method is not GET
  presets?: string[]; // Presets this endpoint belongs to (mail, outlook, personal, ...)
  // JSON Schema for the request body of an endpoint that Microsoft has NOT published
  // in its OpenAPI metadata. Consumed at generate time by bin/modules/simplified-openapi.mjs
  // to synthesize a typed requestBody (instead of a generic object), so the generated client
  // exposes a validated `body` param. Ignored for endpoints already present in the spec.
  requestBodySchema?: Record<string, unknown>;
}

const endpointsData = JSON.parse(
  readFileSync(path.join(__dirname, 'endpoints.json'), 'utf8')
) as EndpointConfig[];

/**
 * Prefix beta-version tools with a [beta] marker so the instability is visible in the
 * tool description itself, regardless of what (if anything) the llmTip says. Tools on
 * v1.0 (the default) are returned unchanged.
 */
function withApiVersionPrefix(description: string, config?: EndpointConfig): string {
  return config?.apiVersion === 'beta' ? `[beta] ${description}` : description;
}

/** When set to a positive integer, caps Graph `$top` on list requests (see README). */
function maxTopFromEnv(): number | undefined {
  const raw = process.env.MS365_MCP_MAX_TOP;
  if (raw === undefined || raw === '') return undefined;
  const n = Number.parseInt(raw, 10);
  if (!Number.isFinite(n) || n < 1) {
    logger.warn(
      `Ignoring invalid MS365_MCP_MAX_TOP=${JSON.stringify(raw)} (use a positive integer)`
    );
    return undefined;
  }
  return n;
}

function clampTopQueryParam(queryParams: Record<string, string>): void {
  const cap = maxTopFromEnv();
  if (cap === undefined || queryParams['$top'] === undefined) return;
  const requested = Number.parseInt(queryParams['$top'], 10);
  if (!Number.isFinite(requested) || requested <= cap) return;
  logger.info(`Clamping $top from ${requested} to ${cap} (MS365_MCP_MAX_TOP)`);
  queryParams['$top'] = String(cap);
}

const DEFAULT_MAX_ITEMS = 10_000;

// Canonical definitions of TOP_UNSUPPORTED_DELTA_TOOLS, paginationAllowed, and
// positiveIntFromEnv live in lib/param-descriptions.ts so tool-schema.ts can use
// them without circling back through graph-tools.ts, and so the description text
// they parameterize can't drift between the two registration paths (see that
// file's header comment).

// Canonical definition lives in lib/destructive-ops.ts so tool-schema.ts can
// use it without circling back through graph-tools.ts; re-exported here for
// external callers (tests, etc.) that imported it from this module.
export { isDestructiveOperation };

/**
 * Defense-in-depth: destructive tools require an explicit `confirm: true` from
 * the caller before they reach Microsoft Graph. Mitigates accidental
 * sendMail / deleteEvent / etc. when an LLM misroutes a request or follows an
 * injected instruction. Opt in per-deployment via MS365_MCP_REQUIRE_CONFIRM=true
 * (default off, so the gate is a non-breaking, additive opt-in that can coexist
 * with client-side elicitation prompts).
 */
function isConfirmGateEnabled(): boolean {
  return process.env.MS365_MCP_REQUIRE_CONFIRM === 'true';
}

type TextContent = {
  type: 'text';
  text: string;
  [key: string]: unknown;
};

type ImageContent = {
  type: 'image';
  data: string;
  mimeType: string;
  [key: string]: unknown;
};

type AudioContent = {
  type: 'audio';
  data: string;
  mimeType: string;
  [key: string]: unknown;
};

type ResourceTextContent = {
  type: 'resource';
  resource: {
    text: string;
    uri: string;
    mimeType?: string;
    [key: string]: unknown;
  };
  [key: string]: unknown;
};

type ResourceBlobContent = {
  type: 'resource';
  resource: {
    blob: string;
    uri: string;
    mimeType?: string;
    [key: string]: unknown;
  };
  [key: string]: unknown;
};

type ResourceContent = ResourceTextContent | ResourceBlobContent;

type ContentItem = TextContent | ImageContent | AudioContent | ResourceContent;

interface CallToolResult {
  content: ContentItem[];
  _meta?: Record<string, unknown>;
  isError?: boolean;

  [key: string]: unknown;
}

function auditHttpStatus(value: unknown): number | undefined {
  if (typeof value === 'number' && Number.isInteger(value)) return value;
  if (typeof value === 'string' && /^\d+$/.test(value)) return Number(value);
  return undefined;
}

function auditErrorCode(value: unknown): string | number | undefined {
  return typeof value === 'string' || typeof value === 'number' ? value : undefined;
}

function auditNonNegativeInteger(value: unknown): number | undefined {
  return typeof value === 'number' && Number.isInteger(value) && value >= 0 ? value : undefined;
}

function auditStringNumberMap(value: unknown): Record<string, number> | undefined {
  if (typeof value !== 'object' || value === null || Array.isArray(value)) return undefined;

  const entries = Object.entries(value).filter(
    (entry): entry is [string, number] =>
      typeof entry[1] === 'number' && Number.isInteger(entry[1]) && entry[1] >= 0
  );

  return entries.length > 0 ? Object.fromEntries(entries) : undefined;
}

function graphResponseAuditFields(
  response: Pick<CallToolResult, '_meta' | 'isError'>
): Pick<
  AuditEvent,
  | 'http_status'
  | 'error_code'
  | 'graph_batch_subrequest_count'
  | 'graph_batch_http_status_counts'
  | 'graph_batch_error_code_counts'
> {
  const httpStatus = auditHttpStatus(response._meta?.http_status);
  const errorCode = response.isError ? auditErrorCode(response._meta?.error_code) : undefined;
  const graphBatchSubrequestCount = auditNonNegativeInteger(
    response._meta?.graph_batch_subrequest_count
  );
  const graphBatchHttpStatusCounts = auditStringNumberMap(
    response._meta?.graph_batch_http_status_counts
  );
  const graphBatchErrorCodeCounts = auditStringNumberMap(
    response._meta?.graph_batch_error_code_counts
  );

  return {
    ...(httpStatus !== undefined ? { http_status: httpStatus } : {}),
    ...(errorCode !== undefined ? { error_code: errorCode } : {}),
    ...(graphBatchSubrequestCount !== undefined
      ? { graph_batch_subrequest_count: graphBatchSubrequestCount }
      : {}),
    ...(graphBatchHttpStatusCounts !== undefined
      ? { graph_batch_http_status_counts: graphBatchHttpStatusCounts }
      : {}),
    ...(graphBatchErrorCodeCounts !== undefined
      ? { graph_batch_error_code_counts: graphBatchErrorCodeCounts }
      : {}),
  };
}

// Graph is inconsistent about casing: entity creation (POST /me/messages) uses
// camelCase body fields, while action endpoints (POST .../forward, /me/sendMail)
// use PascalCase. Matched case-insensitively so both are covered.
const RECIPIENT_FIELDS = new Set([
  'torecipients',
  'ccrecipients',
  'bccrecipients',
  'attendees',
  // driveItem /invite mails an outsider a link to the file
  'recipients',
]);
// Worst shape the docstring below promises to cover is 7, not 4: a graph-batch
// sub-request carrying an itemAttachment lands at requests -> request -> body -> message
// -> attachments -> attachment -> item, and the attached message's own toRecipients match
// there. That leaves one level spare, so this can't be trimmed without giving that case
// up - there's a test pinning it. Arrays charge depth too - traversing them for free
// leaves an array-only path unbounded, and this runs inside the catch handler where a
// stack overflow would take the audit record with it.
const MAX_BODY_DEPTH = 8;
// A big distribution list would otherwise dump every domain into one audit line, at
// a length the caller picks. recipient_count is untouched, so we never lose how many.
const MAX_RECIPIENT_DOMAINS = 50;

// Graph matches property names case-insensitively, so we have to as well - or a
// PascalCase payload sends mail the log never sees
function lookupCaseInsensitive(node: unknown, lowerName: string): unknown {
  if (!node || typeof node !== 'object') return undefined;
  for (const [key, value] of Object.entries(node as Record<string, unknown>)) {
    if (key.toLowerCase() === lowerName) return value;
  }
  return undefined;
}

// Recipients are not one shape: mail and events use emailAddress.address, driveItem
// /invite uses email, meeting participants use upn. alias/objectId name someone with
// no domain at all - those still count, they just don't add one.
function readAddress(entry: unknown): string | undefined {
  // Graph 400s a bare string in a recipient array, but the attempt is the signal and
  // everything else here reads high - this shouldn't be the one place that reads low
  if (typeof entry === 'string') return entry;

  const emailAddress = lookupCaseInsensitive(entry, 'emailaddress');
  const candidates = [
    typeof emailAddress === 'string'
      ? emailAddress
      : lookupCaseInsensitive(emailAddress, 'address'),
    lookupCaseInsensitive(entry, 'email'),
    lookupCaseInsensitive(entry, 'upn'),
  ];
  for (const candidate of candidates) {
    if (typeof candidate === 'string' && candidate) return candidate;
  }
  return undefined;
}

// Positive check, not suffix-stripping. Everything after the last @ is caller-controlled
// and Graph tolerates enough junk that subtracting kept losing - "example.com/path" and
// "evil<script" both got through. Anything that isn't a plain dotted hostname gets dropped.
const DOMAIN_PATTERN = /^[a-z0-9]([a-z0-9-]*[a-z0-9])?(\.[a-z0-9]([a-z0-9-]*[a-z0-9])?)+$/;
// Longest a domain can legally be (RFC 1035). The pattern accepts any length, and the
// 50-entry cap only bounds how many domains land in a record, not how long each one is -
// without this one address picks the size of the audit line. Tested before the pattern so
// a caller can't make us scan a huge string either.
const MAX_DOMAIN_LENGTH = 253;

function addRecipient(entry: unknown, domains: Set<string>, counter: { count: number }): void {
  // Count the entry, not our ability to parse it
  counter.count += 1;

  const address = readAddress(entry);
  if (address === undefined) return;
  const at = address.lastIndexOf('@');
  if (at <= 0 || at >= address.length - 1) return;
  // Peel the wrappers Graph tolerates, then let the pattern decide
  const domain = address
    .slice(at + 1)
    .trim()
    .split(/\s/)[0]
    .replace(/[>.]+$/, '')
    .toLowerCase();
  if (domain.length <= MAX_DOMAIN_LENGTH && DOMAIN_PATTERN.test(domain)) domains.add(domain);
}

function collectRecipients(
  node: unknown,
  domains: Set<string>,
  counter: { count: number },
  depth = 0
): void {
  if (!node || typeof node !== 'object' || depth > MAX_BODY_DEPTH) return;

  if (Array.isArray(node)) {
    for (const item of node) collectRecipients(item, domains, counter, depth + 1);
    return;
  }

  for (const [key, value] of Object.entries(node as Record<string, unknown>)) {
    if (RECIPIENT_FIELDS.has(key.toLowerCase()) && Array.isArray(value)) {
      for (const entry of value) addRecipient(entry, domains, counter);
      continue;
    }
    // Walk everything, not an allowlist of container names - a recipient list
    // nested somewhere we didn't think of still sends real mail
    collectRecipients(value, domains, counter, depth + 1);
  }
}

// A body that failed schema parsing goes to Graph as a raw string, and mail sent that way
// would record nothing. Only a JSON-shaped string can carry recipients though: the binary
// upload tools send base64, and parsing that threw on every single upload.
function bodyForRecipientWalk(body: unknown): unknown {
  if (typeof body !== 'string') return body;
  if (!/^\s*[{[]/.test(body)) return undefined;
  return JSON.parse(body);
}

/**
 * Derives recipient metadata from an outgoing request body for the audit trail.
 *
 * A send, forward or meeting invite records that the tool ran but not who it
 * reached, so an instruction injected via message content that quietly addresses
 * something outside the organisation looks identical in the log to a legitimate
 * reply.
 *
 * Domains, not addresses. The detection question is "did this leave the
 * organisation", which a domain answers; the full address is message content and
 * logging it by default would be a heavier privacy cost than the signal
 * justifies. Counts alone cannot distinguish internal from external.
 *
 * Covers mail recipients, event attendees and driveItem /invite recipients, at
 * any casing and up to MAX_BODY_DEPTH levels down, including inside a
 * graph-batch sub-request and inside a body forwarded as a raw JSON string.
 *
 * recipient_count is entries in a recipient-shaped array; a driveRecipient given
 * only as alias or objectId counts but yields no domain. Keyed on body shape,
 * not on the endpoint, so drafts and event edits count like real sends, and so
 * does an attached message's own toRecipients inside an itemAttachment. Reads
 * high rather than low, which is the safe direction for a detection signal.
 *
 * Gaps remain, so absence of these fields is NOT evidence that nothing left
 * the organisation:
 *  - On `reply` / `replyAll`, recipients added via the optional `Message` are
 *    recorded, but the original thread's are resolved server-side by Graph and
 *    never appear. A plain reply-all to a wide external thread records nothing.
 *  - `POST /me/messages/{id}/send` carries no body. Its recipients were logged
 *    when the draft was created, but only if the draft was created through this
 *    server; one composed in Outlook and sent here records nothing.
 *
 * Policy-denied attempts (`tool.denied`) also record none, but nothing was sent.
 */
function recipientAuditFields(
  body: unknown
): Pick<AuditEvent, 'recipient_count' | 'recipient_domains' | 'recipient_domains_truncated'> {
  const domains = new Set<string>();
  const counter = { count: 0 };
  // Broader than a recursion guard on purpose: this also runs on the catch path, where
  // throwing would cost the audit record AND the caller's error response. Losing the
  // fields beats losing both.
  try {
    collectRecipients(bodyForRecipientWalk(body), domains, counter);
  } catch (error) {
    logger.warn(
      `Skipped recipient audit metadata: ${error instanceof Error ? error.message : 'unknown error'}`
    );
    return {};
  }

  if (counter.count === 0) return {};
  const sorted = [...domains].sort();
  const capped = sorted.slice(0, MAX_RECIPIENT_DOMAINS);
  return {
    recipient_count: counter.count,
    ...(capped.length > 0 ? { recipient_domains: capped } : {}),
    ...(sorted.length > capped.length ? { recipient_domains_truncated: true } : {}),
  };
}

function thrownErrorAuditFields(error: unknown): Pick<AuditEvent, 'http_status' | 'error_code'> {
  const err = error as {
    code?: string | number;
    status?: string | number;
    httpStatus?: string | number;
    graphErrorCode?: string | number;
  };
  const httpStatus = auditHttpStatus(err?.httpStatus ?? err?.status);
  const errorCode = auditErrorCode(err?.graphErrorCode ?? err?.code ?? err?.status);

  return {
    ...(httpStatus !== undefined ? { http_status: httpStatus } : {}),
    ...(errorCode !== undefined ? { error_code: errorCode } : {}),
  };
}

interface UtilityToolContext {
  graphClient: GraphClient;
  authManager?: AuthManager;
  multiAccount: boolean;
  accountNames: string[];
  httpMode: boolean;
  publicBaseUrl?: string;
}

interface UtilityTool {
  name: string;
  // Synthetic for display in search-tools / get-tool-schema. The `tool:` prefix
  // marks these as non-Graph so an LLM doesn't try to construct a Graph URL from them.
  method: string;
  path: string;
  description: string;
  searchKeywords?: string;
  buildSchema: (ctx: UtilityToolContext) => Record<string, z.ZodTypeAny>;
  execute: (params: Record<string, unknown>, ctx: UtilityToolContext) => Promise<CallToolResult>;
  // This is policy, not merely an MCP annotation. Registration, read-only
  // filtering, confirmation, and annotations must all derive from this value.
  mutatesState: boolean;
  openWorldHint?: boolean;
  // When true, this tool writes to the server's local filesystem and is only
  // registered in stdio mode — never in HTTP/OAuth mode, where a remote client
  // must not be able to write arbitrary files onto the host.
  stdioOnly?: boolean;
}

export type ToolNameMatcher = string | ((name: string) => boolean);
type CompiledToolNameMatcher = RegExp | ((name: string) => boolean);

function toolNameMatches(matcher: CompiledToolNameMatcher, name: string): boolean {
  return typeof matcher === 'function' ? matcher(name) : matcher.test(name);
}

interface DisabledToolScope {
  toolName: string;
  missingScopes: string[];
}

type ToolDeniedReason = 'allowed_scopes' | 'tool_allowlist';

interface DeniedToolPolicy {
  toolName: string;
  reason: ToolDeniedReason;
  missingScopes?: string[];
  pathPattern?: string;
}

function formatDisabledToolsForLog(disabledTools: DisabledToolScope[]): string {
  const shown = disabledTools
    .slice(0, 20)
    .map((tool) => `${tool.toolName} (missing: ${tool.missingScopes.join(', ')})`);
  const suffix =
    disabledTools.length > shown.length ? `, ... +${disabledTools.length - shown.length} more` : '';
  return `${shown.join('; ')}${suffix}`;
}

function deniedToolPolicyForGraphTool(
  tool: (typeof allEndpoints)[number],
  config: EndpointConfig | undefined,
  reason: ToolDeniedReason,
  missingScopes?: string[]
): DeniedToolPolicy {
  return {
    toolName: tool.alias,
    reason,
    ...(missingScopes && missingScopes.length > 0 ? { missingScopes } : {}),
    pathPattern: config?.pathPattern ?? tool.path,
  };
}

function collectDeniedToolPolicies(options: {
  readOnly: boolean;
  orgMode: boolean;
  // A predicate, not a RegExp: --direct-tools intersection yields a function.
  enabledToolsMatches?: (name: string) => boolean;
  allowedScopesValue?: string;
  httpMode: boolean;
}): Map<string, DeniedToolPolicy> {
  const deniedTools = new Map<string, DeniedToolPolicy>();
  const allowedScopes = parseAllowedScopes(options.allowedScopesValue);

  for (const tool of allEndpoints) {
    const endpointConfig = endpointsData.find((e) => e.toolName === tool.alias);
    if (!options.orgMode && endpointConfig && !endpointConfig.scopes && endpointConfig.workScopes) {
      continue;
    }

    const method = tool.method.toUpperCase();
    if (options.readOnly && method !== 'GET' && !(method === 'POST' && endpointConfig?.readOnly)) {
      continue;
    }

    if (options.enabledToolsMatches && !options.enabledToolsMatches(tool.alias)) {
      deniedTools.set(
        tool.alias,
        deniedToolPolicyForGraphTool(tool, endpointConfig, 'tool_allowlist')
      );
      continue;
    }

    const missingScopes =
      allowedScopes !== undefined && !endpointConfig
        ? ['endpoint scope metadata']
        : getMissingAllowedScopesForGroups(
            getEndpointScopeGroups(endpointConfig, options.orgMode),
            allowedScopes
          );
    if (missingScopes.length > 0) {
      deniedTools.set(
        tool.alias,
        deniedToolPolicyForGraphTool(tool, endpointConfig, 'allowed_scopes', missingScopes)
      );
    }
  }

  for (const utility of UTILITY_TOOLS) {
    if (options.readOnly && utility.mutatesState) continue;
    if (options.httpMode && utility.stdioOnly) continue;
    if (options.enabledToolsMatches && !options.enabledToolsMatches(utility.name)) {
      deniedTools.set(utility.name, {
        toolName: utility.name,
        reason: 'tool_allowlist',
      });
    }
  }

  return deniedTools;
}

function auditToolDenied(policy: DeniedToolPolicy, params: Record<string, unknown> = {}): void {
  const targetResource = policy.pathPattern
    ? deriveTargetResource({ pathPattern: policy.pathPattern, params })
    : undefined;

  auditLog({
    event: 'tool.denied',
    request_id: randomUUID(),
    user_principal_name: getUserIdentityForAudit(getRequestTokens()?.accessToken),
    tool: policy.toolName,
    status: 'denied',
    reason: policy.reason,
    ...(policy.missingScopes ? { missing_scopes: policy.missingScopes } : {}),
    ...(targetResource ? { target_resource: targetResource } : {}),
  });
}

function installDeniedToolAuditHandler(
  server: McpServer,
  deniedTools: ReadonlyMap<string, DeniedToolPolicy>
): void {
  if (deniedTools.size === 0) return;

  const lowLevel = server.server;
  const handlers = (
    lowLevel as unknown as {
      _requestHandlers?: Map<
        string,
        (request: unknown, extra: unknown) => Promise<ServerResult> | ServerResult
      >;
    }
  )._requestHandlers;
  const original = handlers?.get('tools/call');
  if (!original) {
    logger.warn('Skipping denied-tool audit hook: tools/call handler not found');
    return;
  }

  lowLevel.setRequestHandler(CallToolRequestSchema, async (request, extra) => {
    const policy = deniedTools.get(request.params.name);
    if (policy) {
      const params =
        request.params.arguments && typeof request.params.arguments === 'object'
          ? (request.params.arguments as Record<string, unknown>)
          : {};
      auditToolDenied(policy, params);
    }

    return original(request, extra);
  });
}

const DEFAULT_DOWNLOAD_BYTES_MAX_INLINE = 256 * 1024;

function downloadBytesMaxInline(): number {
  const raw = process.env.MS365_MCP_DOWNLOAD_BYTES_MAX_INLINE;
  if (raw === undefined || raw === '') return DEFAULT_DOWNLOAD_BYTES_MAX_INLINE;
  const n = Number.parseInt(raw, 10);
  if (!Number.isFinite(n) || n < 0) {
    logger.warn(
      `Ignoring invalid MS365_MCP_DOWNLOAD_BYTES_MAX_INLINE=${JSON.stringify(raw)} (use 0 or a positive integer)`
    );
    return DEFAULT_DOWNLOAD_BYTES_MAX_INLINE;
  }
  return n;
}

type DownloadUrlTargetKind = 'meeting-recording' | 'brokerable' | 'drive-item' | 'unsupported';

interface DownloadUrlTargetClassification {
  kind: DownloadUrlTargetKind;
  pathPart: string;
  isDriveContentEndpoint: boolean;
}

interface CanonicalBinaryTarget {
  target: string;
  classification: DownloadUrlTargetClassification;
}

const ENCODED_BINARY_ROUTE_DELIMITER = /%(?:2f|3a|3f|5c)/i;
const ENCODED_FRAGMENT_DELIMITER = /%23/i;

function encodedHashesAreDrivePathData(target: string): boolean {
  const drivePath = /\/root:\/(.+):(?:\/content)?\/?$/i.exec(target);
  if (!drivePath) return !ENCODED_FRAGMENT_DELIMITER.test(target);

  const filenameStart = drivePath.index + '/root:/'.length;
  const filenameEnd = filenameStart + drivePath[1].length;
  for (const match of target.matchAll(/%23/gi)) {
    const index = match.index;
    if (index < filenameStart || index >= filenameEnd) return false;
  }
  return true;
}

function validateBinaryTargetSegments(target: string): void {
  if (target.split('/').some((segment) => segment === '.' || segment === '..')) {
    throw new Error('target must not contain dot path segments.');
  }
}

function validateBinaryTargetEncoding(target: string): void {
  let layer = target;
  for (let depth = 0; depth < 16; depth++) {
    if (ENCODED_BINARY_ROUTE_DELIMITER.test(layer) || !encodedHashesAreDrivePathData(layer)) {
      throw new Error('target must not contain encoded route, query, or fragment delimiters.');
    }

    // Encoded dots are ordinary filename data unless decoding makes the whole segment
    // `.` or `..`, which changes routing semantics. Do not decode the dispatched path.
    const dotsDecoded = layer.replace(/%2e/gi, '.');
    if (dotsDecoded.split('/').some((segment) => segment === '.' || segment === '..')) {
      throw new Error('target must not contain an encoded dot path segment.');
    }

    // Peel one layer of encoded percent signs so double-encoded delimiters cannot hide.
    // A literal `%25` filename is fine: it becomes `%` on the next pass and stabilizes.
    const next = layer.replace(/%25/gi, '%');
    if (next === layer) return;
    layer = next;
  }
  throw new Error('target contains excessive nested percent encoding.');
}

/**
 * Validate and normalize the Graph path once before either classification or dispatch.
 * URL fragments are not sent by fetch, and encoded route delimiters may be decoded by
 * an intermediary, so accepting either would let the classifier inspect a different
 * effective resource from the one Graph receives.
 */
function canonicalizeBinaryTarget(target: string): CanonicalBinaryTarget {
  if (!target.startsWith('/')) {
    throw new Error('target must be a relative Microsoft Graph path starting with "/".');
  }
  if (target.includes('?')) throw new Error('target must not include query parameters.');
  if (target.includes('#')) throw new Error('target must not include a URL fragment.');
  if (target.includes('\\')) throw new Error('target must not include a backslash.');
  // URL implementations normalize raw dot segments before dispatch. Reject them
  // explicitly so classification and the effective Graph resource cannot diverge.
  validateBinaryTargetSegments(target);
  validateBinaryTargetEncoding(target);
  for (let index = 0; index < target.length; index++) {
    const code = target.charCodeAt(index);
    if (code <= 0x1f || code === 0x7f) {
      throw new Error('target must not contain control characters.');
    }
  }

  const canonicalTarget = target.replace(/\/+$/, '');
  if (canonicalTarget === '') throw new Error('target must identify a Microsoft Graph resource.');
  return {
    target: canonicalTarget,
    classification: classifyDownloadUrlTarget(canonicalTarget),
  };
}

/**
 * Classify the exact target shapes get-download-url can serve. This is shared with
 * download-bytes so its HTTP inline cutoff is only applied when the suggested
 * out-of-band route really exists.
 */
function classifyDownloadUrlTarget(target: string): DownloadUrlTargetClassification {
  const pathPart = target;
  const isMeetingRecording =
    /^(\/me|\/users\/[^/]+)\/onlineMeetings\/[^/]+\/recordings\/[^/]+(?:\/content)?$/.test(
      pathPart
    ) || /^\/communications\/calls\/[^/]+\/recordings\/[^/]+(?:\/content)?$/.test(pathPart);
  if (isMeetingRecording) {
    return { kind: 'meeting-recording', pathPart, isDriveContentEndpoint: false };
  }

  // Match only real Graph mail/calendar attachment resources so driveItem path addressing
  // with folders named messages/events/attachments is not falsely treated as an attachment.
  const isAttachmentEndpoint =
    /^(\/me|\/users\/[^/]+)\/messages\/[^/]+\/attachments\/[^/]+(?:\/\$value)?$/.test(pathPart) ||
    /^(\/me|\/users\/[^/]+)\/events\/[^/]+\/attachments\/[^/]+(?:\/\$value)?$/.test(pathPart) ||
    /^\/groups\/[^/]+\/messages\/[^/]+\/attachments\/[^/]+(?:\/\$value)?$/.test(pathPart) ||
    /^\/groups\/[^/]+\/events\/[^/]+\/attachments\/[^/]+(?:\/\$value)?$/.test(pathPart);
  if (isAttachmentEndpoint || pathPart.endsWith('/$value')) {
    return { kind: 'brokerable', pathPart, isDriveContentEndpoint: false };
  }

  const isDriveItemById =
    /^\/drives\/[^/]+\/items\/[^/]+(?:\/content)?$/.test(pathPart) ||
    /^\/(?:me|users\/[^/]+|groups\/[^/]+|sites\/[^/]+)\/drive\/items\/[^/]+(?:\/content)?$/.test(
      pathPart
    ) ||
    /^\/(?:groups\/[^/]+|sites\/[^/]+)\/drives\/[^/]+\/items\/[^/]+(?:\/content)?$/.test(pathPart);
  const isDriveItemByPath =
    /^\/drives\/[^/]+\/root:\/.+:(?:\/content)?$/.test(pathPart) ||
    /^\/(?:me|users\/[^/]+|groups\/[^/]+|sites\/[^/]+)\/drive\/root:\/.+:(?:\/content)?$/.test(
      pathPart
    ) ||
    /^\/(?:groups\/[^/]+|sites\/[^/]+)\/drives\/[^/]+\/root:\/.+:(?:\/content)?$/.test(pathPart);
  const isDriveContentEndpoint =
    /\/items\/[^/]+\/content$/.test(pathPart) || pathPart.endsWith(':/content');
  return {
    kind: isDriveItemById || isDriveItemByPath ? 'drive-item' : 'unsupported',
    pathPart,
    isDriveContentEndpoint,
  };
}

/**
 * In OAuth/HTTP bearer mode the `account` parameter cannot switch identities —
 * every Graph call uses the connecting client's bearer token. Previously a
 * provided `account` was silently ignored and the bearer user's data returned
 * (discussion #467). Returns an error message when an `account` param is
 * provided that the bearer identity cannot honor; a param matching the bearer's
 * own identity passes through. Returns null when account routing via the MSAL
 * cache is available (stdio mode, or HTTP with --trust-proxy-auth).
 */
async function checkAccountParamInBearerMode(
  accountParam: string | undefined,
  authManager?: AuthManager
): Promise<string | null> {
  if (!accountParam || !authManager) return null;
  const contextToken = getRequestTokens()?.accessToken;
  if (!contextToken && !authManager.isOAuthModeEnabled()) return null;
  const bearerToken = contextToken ?? (await authManager.getToken().catch(() => null)) ?? undefined;
  const bearerIdentity = getUserIdentityForAudit(bearerToken);
  if (bearerIdentity && bearerIdentity.toLowerCase() === accountParam.toLowerCase()) return null;
  return (
    `The 'account' parameter is not supported in HTTP/OAuth mode: every request uses the identity ` +
    `of the connecting client's bearer token` +
    (bearerIdentity ? ` ('${bearerIdentity}')` : '') +
    `, so account switching is not possible. To act as '${accountParam}', reconnect the MCP client ` +
    `authenticated as that account, or run the server in stdio mode (or HTTP with --trust-proxy-auth) ` +
    `where cached accounts are available.`
  );
}

export const UTILITY_TOOLS: readonly UtilityTool[] = [
  {
    name: 'parse-teams-url',
    method: 'POST',
    path: 'tool:parse-teams-url',
    description:
      'Converts any Teams meeting URL format (short /meet/, full /meetup-join/, or recap ?threadId=) into a standard joinWebUrl. Use this before list-online-meetings when the user provides a recap or short URL.',
    mutatesState: false,
    openWorldHint: false,
    buildSchema: () => ({
      url: z.string().describe('Teams meeting URL in any format'),
    }),
    execute: async (params) => {
      const url = params.url;
      if (typeof url !== 'string') {
        return {
          content: [{ type: 'text', text: JSON.stringify({ error: 'url is required.' }) }],
          isError: true,
        };
      }
      try {
        const joinWebUrl = parseTeamsUrl(url);
        return { content: [{ type: 'text', text: joinWebUrl }] };
      } catch (error) {
        return {
          content: [{ type: 'text', text: JSON.stringify({ error: (error as Error).message }) }],
          isError: true,
        };
      }
    },
  },
  {
    name: 'download-bytes',
    method: 'GET',
    path: 'tool:download-bytes',
    description:
      'Download binary content from Microsoft Graph and return it as base64. Single tool for any binary read: drive file content, mail attachment, profile photo, Teams hosted content, meeting recording. Returns { contentType, encoding: "base64", contentLength, contentBytes }. For large content, prefer get-download-url, which returns native pre-authenticated URLs for drive/SharePoint files and brokered URLs for supported attachments when the broker is configured.',
    mutatesState: false,
    openWorldHint: true,
    buildSchema: (ctx) => {
      const schema: Record<string, z.ZodTypeAny> = {
        target: z
          .string()
          .describe(
            'Relative Microsoft Graph path starting with "/". Common paths: ' +
              '/drives/{drive-id}/items/{driveItem-id}/content (drive file content); ' +
              '/me/messages/{message-id}/attachments/{attachment-id}/$value (mail attachment, list-mail-attachments returns the IDs); ' +
              '/me/photo/$value or /users/{user-id}/photo/$value (profile photo); ' +
              '/chats/{chat-id}/messages/{chatMessage-id}/hostedContents/{chatMessageHostedContent-id}/$value (Teams chat hosted content, list-chat-message-hosted-contents returns the IDs); ' +
              '/teams/{team-id}/channels/{channel-id}/messages/{chatMessage-id}/hostedContents/{chatMessageHostedContent-id}/$value (Teams channel hosted content). ' +
              'For meeting recordings, use get-meeting-recording-content where available; Microsoft Graph returns authenticated recording bytes, not a pre-authenticated download URL.'
          ),
      };
      if (ctx.multiAccount) {
        schema['account'] = z
          .string()
          .optional()
          .describe(
            'Account to use when multiple Microsoft accounts are configured. Required when multiple accounts exist (see list-accounts).'
          );
      }
      return schema;
    },
    execute: async (params, { graphClient, authManager, httpMode, publicBaseUrl }) => {
      const target = params.target;
      const accountParam = params.account as string | undefined;
      if (typeof target !== 'string' || target.length === 0) {
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify({ error: 'target is required and must be a non-empty string.' }),
            },
          ],
          isError: true,
        };
      }
      let canonical: CanonicalBinaryTarget;
      try {
        canonical = canonicalizeBinaryTarget(target);
      } catch (error) {
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify({ error: (error as Error).message }),
            },
          ],
          isError: true,
        };
      }
      const { target: canonicalTarget, classification } = canonical;
      try {
        const accountModeError = await checkAccountParamInBearerMode(accountParam, authManager);
        if (accountModeError) {
          return {
            content: [{ type: 'text', text: JSON.stringify({ error: accountModeError }) }],
            isError: true,
          };
        }
        let accountAccessToken: string | undefined;
        if (authManager && !authManager.isOAuthModeEnabled() && !getRequestTokens()) {
          accountAccessToken = await authManager.getTokenForAccount(accountParam);
        }
        // Enforce the inline limit while Graph is streaming, before constructing a base64
        // payload. Only do so when get-download-url can actually provide the out-of-band
        // alternative named by the error. Meeting recordings and other authenticated
        // /content targets retain the historical response path because HTTP mode does not
        // expose download-bytes-to-file. Set MS365_MCP_DOWNLOAD_BYTES_MAX_INLINE=0 to
        // disable the cutoff even for supported targets.
        const downloadUrlKind = classification.kind;
        const hasOutOfBandAlternative =
          downloadUrlKind === 'brokerable' || downloadUrlKind === 'drive-item';
        const maxInline =
          isBrokerEnabled(httpMode, publicBaseUrl) && hasOutOfBandAlternative
            ? downloadBytesMaxInline()
            : 0;
        if (maxInline > 0) {
          try {
            const download = await graphClient.downloadToBuffer(canonicalTarget, maxInline, {
              accessToken: accountAccessToken,
            });
            return {
              content: [
                {
                  type: 'text',
                  text: JSON.stringify({
                    contentType: download.contentType,
                    encoding: 'base64',
                    contentLength: download.contentLength,
                    contentBytes: download.bytes.toString('base64'),
                  }),
                },
              ],
            };
          } catch (error) {
            if (error instanceof GraphDownloadSizeLimitError) {
              return {
                content: [
                  {
                    type: 'text',
                    text: JSON.stringify({
                      error: `Content exceeds the inline limit of ${maxInline} bytes (MS365_MCP_DOWNLOAD_BYTES_MAX_INLINE). Use get-download-url to fetch it out-of-band instead of base64 through the agent context.`,
                      ...(error.actualBytes === undefined
                        ? {}
                        : { contentLength: error.actualBytes }),
                    }),
                  },
                ],
                isError: true,
              };
            }
            throw error;
          }
        }

        // Without a usable broker, retain the historical byte-faithful response path.
        const response = await graphClient.graphRequest(canonicalTarget, {
          accessToken: accountAccessToken,
          rawResponse: true,
        });
        return response;
      } catch (error) {
        return {
          content: [{ type: 'text', text: JSON.stringify({ error: (error as Error).message }) }],
          isError: true,
        };
      }
    },
  },
  {
    name: 'download-bytes-to-file',
    method: 'GET',
    path: 'tool:download-bytes-to-file',
    searchKeywords:
      'save to disk save to file write file to disk save attachment to disk save recording to disk write bytes to local file output path',
    // Front-loaded on purpose: the discovery search index caps a tool's
    // description at ~40 tokens, so the OneDrive/SharePoint guidance below
    // sits past the cap. That keeps the hint for the reading LLM while letting
    // get-download-url own the high-signal "drive"/"sharepoint" search terms.
    description:
      'Write authenticated Microsoft Graph byte content to a local file on the server, returning { path, contentType, bytesWritten } instead of base64. Handles mail attachments, meeting recordings, profile photos, and Teams hosted content. Writes to an absolute outputPath and never overwrites an existing file. Stdio mode only, not available over HTTP. Prefer get-download-url when available because it returns native or brokered URLs for fully out-of-band download; download-bytes-to-file remains the out-of-band path for meeting recordings.',
    mutatesState: true,
    openWorldHint: true,
    stdioOnly: true,
    buildSchema: (ctx) => {
      const schema: Record<string, z.ZodTypeAny> = {
        target: z
          .string()
          .describe(
            'Relative Microsoft Graph path starting with "/". Common paths: ' +
              '/drives/{drive-id}/items/{driveItem-id}/content (drive file content); ' +
              '/me/messages/{message-id}/attachments/{attachment-id}/$value (mail attachment, list-mail-attachments returns the IDs); ' +
              '/me/photo/$value or /users/{user-id}/photo/$value (profile photo); ' +
              '/chats/{chat-id}/messages/{chatMessage-id}/hostedContents/{chatMessageHostedContent-id}/$value (Teams chat hosted content, list-chat-message-hosted-contents returns the IDs); ' +
              '/teams/{team-id}/channels/{channel-id}/messages/{chatMessage-id}/hostedContents/{chatMessageHostedContent-id}/$value (Teams channel hosted content). ' +
              'For meeting recordings, use get-meeting-recording-content where available; Microsoft Graph returns authenticated recording bytes, not a pre-authenticated download URL.'
          ),
        outputPath: z
          .string()
          .describe(
            "Absolute path on the server's filesystem where the bytes are written, e.g. /Users/me/downloads/invoice.pdf. Must be absolute; relative paths are rejected. The parent directory must already exist, and an existing file is never overwritten (the call errors if outputPath already exists)."
          ),
      };
      if (ctx.multiAccount) {
        schema['account'] = z
          .string()
          .optional()
          .describe(
            'Account to use when multiple Microsoft accounts are configured. Required when multiple accounts exist (see list-accounts).'
          );
      }
      return schema;
    },
    execute: async (params, { graphClient, authManager }) => {
      const target = params.target;
      const outputPath = params.outputPath;
      const accountParam = params.account as string | undefined;
      if (typeof target !== 'string' || target.length === 0) {
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify({ error: 'target is required and must be a non-empty string.' }),
            },
          ],
          isError: true,
        };
      }
      let canonicalTarget: string;
      try {
        canonicalTarget = canonicalizeBinaryTarget(target).target;
      } catch (error) {
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify({ error: (error as Error).message }),
            },
          ],
          isError: true,
        };
      }
      if (typeof outputPath !== 'string' || outputPath.length === 0) {
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify({
                error: 'outputPath is required and must be a non-empty string.',
              }),
            },
          ],
          isError: true,
        };
      }
      if (!path.isAbsolute(outputPath)) {
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify({
                error: `outputPath must be an absolute path, e.g. /Users/me/downloads/file.ext. Received: ${outputPath}`,
              }),
            },
          ],
          isError: true,
        };
      }
      // downloadToFile's wx is the real no-overwrite guard; this just gives a
      // friendlier "already exists" error before we bother calling Graph.
      let fileExists = false;
      try {
        await access(outputPath);
        fileExists = true;
      } catch {
        // ENOENT (and any other access error) means the file isn't readable/there;
        // let the write attempt surface the real problem.
      }
      if (fileExists) {
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify({ error: `file already exists at ${outputPath}` }),
            },
          ],
          isError: true,
        };
      }
      try {
        const accountModeError = await checkAccountParamInBearerMode(accountParam, authManager);
        if (accountModeError) {
          return {
            content: [{ type: 'text', text: JSON.stringify({ error: accountModeError }) }],
            isError: true,
          };
        }
        let accountAccessToken: string | undefined;
        if (authManager && !authManager.isOAuthModeEnabled() && !getRequestTokens()) {
          accountAccessToken = await authManager.getTokenForAccount(accountParam);
        }
        // Stream to disk instead of buffering: makeRequest holds the whole file
        // in memory as base64, which dies on big recordings (V8 max string length).
        const result = await graphClient.downloadToFile(canonicalTarget, outputPath, {
          accessToken: accountAccessToken,
        });
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify({
                path: outputPath,
                contentType: result.contentType,
                bytesWritten: result.contentLength,
              }),
            },
          ],
          ...(result.httpStatus !== undefined ? { _meta: { http_status: result.httpStatus } } : {}),
        };
      } catch (error) {
        const metadata = thrownErrorAuditFields(error);
        return {
          content: [{ type: 'text', text: JSON.stringify({ error: (error as Error).message }) }],
          isError: true,
          ...(Object.keys(metadata).length > 0 ? { _meta: metadata } : {}),
        };
      }
    },
  },
  {
    name: 'get-download-url',
    method: 'GET',
    path: 'tool:get-download-url',
    searchKeywords:
      'download file download drive file download onedrive file sharepoint file download large drive file large sharepoint file large file out-of-band download pre-authenticated url',
    description:
      "Resolve a short-lived download URL for Microsoft Graph binary content. Drive/SharePoint file content returns Graph's native pre-authenticated URL. When the EKI out-of-band broker is configured, mail attachments and other /$value byte endpoints are fetched server-side and served through a short-lived tokenless broker URL. Prefer this over download-bytes for any file above a few KB or any bulk download. Returns { downloadUrl, name?, size?, contentType?, brokered? }. NOTE: meeting recordings do NOT expose a pre-authenticated URL — Graph offers no such link for them; use download-bytes for small ones or a recording-specific tool where available.",
    mutatesState: false,
    openWorldHint: true,
    buildSchema: (ctx) => {
      const schema: Record<string, z.ZodTypeAny> = {
        target: z
          .string()
          .describe(
            'Relative Microsoft Graph path starting with "/". Either a driveItem content path or the item path itself, e.g. ' +
              '/drives/{drive-id}/items/{driveItem-id}/content, /me/drive/items/{driveItem-id}/content, ' +
              'or /sites/{site-id}/drive/items/{driveItem-id}. ' +
              'A trailing /content is optional and is stripped automatically for drive items. Mail attachment $value paths and other /$value byte endpoints require the EKI broker; meeting recordings are not supported because Graph exposes authenticated bytes rather than a pre-authenticated URL.'
          ),
      };
      if (ctx.multiAccount) {
        schema['account'] = z
          .string()
          .optional()
          .describe(
            'Account to use when multiple Microsoft accounts are configured. Required when multiple accounts exist (see list-accounts).'
          );
      }
      return schema;
    },
    execute: async (params, { graphClient, authManager, httpMode, publicBaseUrl }) => {
      const target = params.target;
      const accountParam = params.account as string | undefined;
      if (typeof target !== 'string' || target.length === 0) {
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify({ error: 'target is required and must be a non-empty string.' }),
            },
          ],
          isError: true,
        };
      }
      let canonical: CanonicalBinaryTarget;
      try {
        canonical = canonicalizeBinaryTarget(target);
      } catch (error) {
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify({ error: (error as Error).message }),
            },
          ],
          isError: true,
        };
      }
      const { classification } = canonical;
      const { pathPart } = classification;
      // Recording content endpoints return authenticated bytes, not a pre-authenticated URL.
      if (classification.kind === 'meeting-recording') {
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify({
                error:
                  'Meeting recordings do not expose a pre-authenticated download URL. Use download-bytes for small recordings or get-meeting-recording-content where available.',
              }),
            },
          ],
          isError: true,
        };
      }
      // The downloadUrl lives on driveItem metadata, not the /content sub-resource.
      // Only strip true Graph content endpoints: ID-addressed /items/{id}/content
      // and path-addressed root:/path/file:/content. A drive item can itself be
      // named "content", so a plain trailing /content is not enough.
      const itemPath = classification.isDriveContentEndpoint
        ? pathPart.slice(0, -'/content'.length)
        : pathPart;
      try {
        const accountModeError = await checkAccountParamInBearerMode(accountParam, authManager);
        if (accountModeError) {
          return {
            content: [{ type: 'text', text: JSON.stringify({ error: accountModeError }) }],
            isError: true,
          };
        }

        let accountAccessToken: string | undefined;
        if (authManager && !authManager.isOAuthModeEnabled() && !getRequestTokens()) {
          accountAccessToken = await authManager.getTokenForAccount(accountParam);
        }

        if (classification.kind === 'brokerable') {
          if (!isBrokerEnabled(httpMode, publicBaseUrl)) {
            return {
              content: [
                {
                  type: 'text',
                  text: JSON.stringify({
                    error:
                      'This resource does not expose a pre-authenticated download URL and the out-of-band broker is not configured. Use download-bytes to read these bytes.',
                  }),
                },
              ],
              isError: true,
            };
          }
          const fetchPath = pathPart.endsWith('/$value') ? pathPart : `${pathPart}/$value`;
          const maximumBytes = getBrokerMaxBytes();
          const reservation = reserveBrokerCapacity(maximumBytes, httpMode, publicBaseUrl);
          try {
            const download = await graphClient.downloadToBuffer(fetchPath, maximumBytes, {
              accessToken: accountAccessToken,
            });
            const { bytes, contentType } = download;
            const downloadUrl = mintDownloadUrl(
              {
                bytes,
                memoryBytes: download.allocatedBytes,
                contentType,
                userPrincipalName: getUserIdentityForAudit(getRequestTokens()?.accessToken),
                resourcePath: fetchPath,
              },
              httpMode,
              publicBaseUrl,
              reservation
            );
            return {
              content: [
                {
                  type: 'text',
                  text: JSON.stringify({
                    downloadUrl,
                    size: bytes.length,
                    contentType,
                    brokered: true,
                  }),
                },
              ],
            };
          } finally {
            releaseBrokerCapacity(reservation);
          }
        }

        if (classification.kind !== 'drive-item') {
          return {
            content: [
              {
                type: 'text',
                text: JSON.stringify({
                  error:
                    'target must identify a driveItem in OneDrive or SharePoint. Use a drive item metadata path or /content path, such as /drives/{drive-id}/items/{driveItem-id}/content, /me/drive/items/{driveItem-id}, /sites/{site-id}/drive/items/{driveItem-id}, or /me/drive/root:/path/file.ext:/content. Other Graph byte resources must use download-bytes.',
                }),
              },
            ],
            isError: true,
          };
        }

        const response = await graphClient.graphRequest(itemPath, {
          accessToken: accountAccessToken,
          // We JSON.parse the metadata below, so force JSON - under --toon it'd be
          // TOON and the parse would fail, masking a real item as "no download url".
          forceJsonOutput: true,
        });
        // graphRequest swallows Graph HTTP errors and returns { isError: true } (see
        // graph-client.ts); surface the real error (401/403/404/429/...) instead of masking
        // it as "no download URL available".
        if (response?.isError) {
          return response;
        }
        const text = response?.content?.[0]?.text;
        let item: Record<string, unknown> | undefined;
        if (typeof text === 'string') {
          try {
            item = JSON.parse(text);
          } catch {
            item = undefined;
          }
        }
        const downloadUrl = item?.['@microsoft.graph.downloadUrl'] as string | undefined;
        if (!downloadUrl) {
          return {
            content: [
              {
                type: 'text',
                text: JSON.stringify({
                  error:
                    'No pre-authenticated download URL is available for this resource. It may not be a drive item, or it exposes bytes only via download-bytes.',
                }),
              },
            ],
            isError: true,
            _meta: response._meta,
          };
        }
        const file = item?.file as { mimeType?: string } | undefined;
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify({
                downloadUrl,
                name: item?.name,
                size: item?.size,
                contentType: file?.mimeType,
              }),
            },
          ],
          _meta: response._meta,
        };
      } catch (error) {
        const metadata = thrownErrorAuditFields(error);
        return {
          content: [{ type: 'text', text: JSON.stringify({ error: (error as Error).message }) }],
          isError: true,
          ...(Object.keys(metadata).length > 0 ? { _meta: metadata } : {}),
        };
      }
    },
  },
];

function buildUtilitySchema(
  utility: UtilityTool,
  ctx: UtilityToolContext
): Record<string, z.ZodTypeAny> {
  const schema = utility.buildSchema(ctx);
  if (utility.mutatesState) {
    schema.confirm = z.boolean().describe(CONFIRM_PARAM_DESCRIPTION).optional();
  }
  return schema;
}

function registerUtilityToolWithMcp(
  server: McpServer,
  utility: UtilityTool,
  ctx: UtilityToolContext
): void {
  server.tool(
    utility.name,
    utility.description,
    buildUtilitySchema(utility, ctx),
    {
      title: utility.name,
      readOnlyHint: !utility.mutatesState,
      destructiveHint: utility.mutatesState,
      openWorldHint: utility.openWorldHint ?? true,
    },
    async (params) => executeUtilityTool(utility, params, ctx, 'direct')
  );
}

async function executeUtilityTool(
  utility: UtilityTool,
  params: Record<string, unknown>,
  ctx: UtilityToolContext,
  route: ToolRoute
): Promise<CallToolResult> {
  // Merged at the v0.148.0 rebase: upstream grew its own executeUtilityTool for audit
  // logging while this fork had one for Prometheus metrics, both wrapping the same
  // utility.execute. This is the union, on EKI's signature because `route` must survive
  // (a utility reached through execute-tool is not a direct call).
  const requestId = randomUUID();
  const startedAt = Date.now();
  const elapsed = () => (Date.now() - startedAt) / 1000;
  const upn = getUserIdentityForAudit(getRequestTokens()?.accessToken);
  if (isConfirmGateEnabled() && utility.mutatesState && params.confirm !== true) {
    logger.warn(`Refusing destructive utility ${utility.name}: missing confirm: true`);
    const response: CallToolResult = {
      content: [
        {
          type: 'text',
          text: JSON.stringify({
            error: 'confirmation_required',
            tool: utility.name,
            method: utility.method.toUpperCase(),
            destructive: true,
            message:
              'This tool modifies user data or local state. Re-call with parameter "confirm": true after the user has explicitly approved the operation.',
          }),
        },
      ],
      isError: true,
    };
    recordToolCall(utility.name, route, 'blocked', elapsed());
    return response;
  }
  try {
    const response = await utility.execute(params, ctx);
    recordToolCall(utility.name, route, response.isError ? 'error' : 'ok', elapsed());
    auditLog({
      event: 'tool.call',
      request_id: requestId,
      user_principal_name: upn,
      tool: utility.name,
      http_method: utility.method.toUpperCase(),
      status: response.isError ? 'error' : 'success',
      duration_ms: Date.now() - startedAt,
      ...graphResponseAuditFields(response),
    });
    return response;
  } catch (error) {
    const err = error as { name?: string };
    recordToolCall(utility.name, route, 'error', elapsed());
    auditLog({
      event: 'tool.call',
      request_id: requestId,
      user_principal_name: upn,
      tool: utility.name,
      http_method: utility.method.toUpperCase(),
      status: 'error',
      duration_ms: Date.now() - startedAt,
      error_type: err?.name || 'Error',
      ...thrownErrorAuditFields(error),
    });
    throw error;
  }
}

// Every nested `body` field in the generated clients is an itemBody, so an @odata.type
// naming it is the only one that belongs inside a body we just moved fields into
function namesNestedBodyType(value: unknown): boolean {
  return typeof value === 'string' && value.toLowerCase().endsWith('itembody');
}

// Dig out the object shape of a Body schema so flattened top-level params can be
// matched against it (#569). z.lazy (chatMessage etc.) hides it behind _def.getter
function bodySchemaShape(schema: z.ZodTypeAny | undefined): Record<string, unknown> | null {
  let current: z.ZodTypeAny | undefined = schema;
  for (let i = 0; i < 10 && current; i++) {
    if (current instanceof z.ZodObject) {
      return current.shape as Record<string, unknown>;
    }
    const def = (
      current as {
        _def?: { innerType?: z.ZodTypeAny; schema?: z.ZodTypeAny; getter?: () => z.ZodTypeAny };
      }
    )._def;
    current = def?.innerType ?? def?.schema ?? def?.getter?.();
  }
  return null;
}

// SDK validation hands the handler the PARSED value, and strip-mode objects silently
// drop unknown keys - passthrough keeps whatever the client sent
function lenientBodySchema(schema: z.ZodTypeAny): z.ZodTypeAny {
  if (schema instanceof z.ZodObject) {
    return schema.passthrough();
  }
  if (schema instanceof z.ZodOptional) {
    return lenientBodySchema(schema.unwrap()).optional();
  }
  if (schema instanceof z.ZodNullable) {
    return lenientBodySchema(schema.unwrap()).nullable();
  }
  if (schema instanceof z.ZodLazy) {
    return z.lazy(() => lenientBodySchema(schema.schema));
  }
  return schema;
}

// Read-only in Graph - merging an echoed id/timestamp into a POST/PATCH body can 400
const READ_ONLY_BODY_FIELDS = new Set([
  'id',
  'createdDateTime',
  'lastModifiedDateTime',
  'changeKey',
]);

function isPlainObject(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value);
}

// Object.hasOwn, but tsconfig targets ES2020. Not `in` - that would match
// toString/constructor through the prototype
function hasOwn(obj: Record<string, unknown>, key: string): boolean {
  return Object.prototype.hasOwnProperty.call(obj, key);
}

// JSON.stringify recurses, so a params object nested deeply enough overflows the stack -
// on Node 20 well before Node 26. It runs before the try below, so an unguarded throw
// escapes executeGraphTool entirely: the caller gets a protocol error instead of a tool
// error, and no audit record is written at all. Nesting depth is caller-controlled.
function describeParamsForLog(params: Record<string, unknown>): string {
  try {
    return JSON.stringify(params);
  } catch (error) {
    return `[unserializable: ${error instanceof Error ? error.name : 'unknown error'}]`;
  }
}

/**
 * Label each subrequest in a batch with the tool whose operation it matches, so the
 * metric shows what batching is used for. Unmatched subrequests are counted as
 * "unmatched" rather than by path, which would be unbounded cardinality and could carry
 * resource ids.
 */
function recordBatchSubrequestsFor(body: unknown, blocked: BlockedOperationMatcher[]): void {
  // recordBatchSubrequest already no-ops when metrics are off, but it does so AFTER this
  // function has labelled every subrequest, so all of that work was thrown away in the
  // default configuration (EnviroKinetics/ms365-mcp#54). Check before doing it.
  if (!metricsEnabled()) return;
  if (!body || typeof body !== 'object') return;
  const requests = (body as { requests?: unknown }).requests;
  if (!Array.isArray(requests)) return;
  const all = allOperationMatchers();
  for (const entry of requests) {
    if (!entry || typeof entry !== 'object') continue;
    const sub = entry as { method?: unknown; url?: unknown };
    if (typeof sub.method !== 'string' || typeof sub.url !== 'string') continue;
    const hit = findBlockedSubrequests({ requests: [entry] }, all.length > 0 ? all : blocked);
    recordBatchSubrequest(hit[0]?.toolName ?? 'unmatched', sub.method);
  }
}

async function executeGraphTool(
  tool: (typeof api.endpoints)[0],
  config: EndpointConfig | undefined,
  graphClient: GraphClient,
  params: Record<string, unknown>,
  authManager?: AuthManager,
  blockedOperations: BlockedOperationMatcher[] = [],
  route: ToolRoute = 'direct'
): Promise<CallToolResult> {
  logger.info(`Tool ${tool.alias} called with params: ${describeParamsForLog(params)}`);
  const startedAt = Date.now();
  const elapsed = () => (Date.now() - startedAt) / 1000;
  let metricsRecorded = false;
  const finish = (result: CallToolResult, outcome?: ToolOutcome): CallToolResult => {
    if (!metricsRecorded) {
      recordToolCall(tool.alias, route, outcome ?? (result.isError ? 'error' : 'ok'), elapsed());
      metricsRecorded = true;
    }
    return result;
  };

  if (
    isConfirmGateEnabled() &&
    isDestructiveOperation(tool.method, config) &&
    params.confirm !== true
  ) {
    logger.warn(
      `Refusing destructive tool ${tool.alias} (${tool.method.toUpperCase()}): missing confirm: true`
    );
    return finish(
      {
        content: [
          {
            type: 'text',
            text: JSON.stringify({
              error: 'confirmation_required',
              tool: tool.alias,
              method: tool.method.toUpperCase(),
              destructive: true,
              message:
                'This tool modifies user data. Re-call with parameter "confirm": true after the user has explicitly approved the operation.',
            }),
          },
        ],
        isError: true,
      },
      'blocked'
    );
  }

  const requestId = randomUUID();
  const startTime = Date.now();
  const upn = getUserIdentityForAudit(getRequestTokens()?.accessToken);
  const httpMethod = tool.method.toUpperCase();
  let targetResource: AuditTargetResource | undefined;
  // Hoisted alongside targetResource so the catch-path audit can still report
  // recipients. A send that times out or trips the breaker is exactly the
  // ambiguous case: a thrown request is not proof that nothing was delivered.
  let body: unknown = null;

  try {
    const accountParam = params.account as string | undefined;

    // In OAuth/HTTP bearer mode, refuse an `account` param that doesn't match the bearer
    // identity instead of silently returning the bearer user's data (discussion #467).
    const accountModeError = await checkAccountParamInBearerMode(accountParam, authManager);
    if (accountModeError) {
      return finish({
        content: [{ type: 'text', text: JSON.stringify({ error: accountModeError }) }],
        isError: true,
      });
    }

    // Resolve account-specific token if `account` parameter is provided (or auto-resolve for single account).
    // Skip in OAuth/HTTP mode — let the request context drive token selection via GraphClient.
    // Also skip when a request-context token exists (HTTP/OAuth flow where token comes from middleware).
    let accountAccessToken: string | undefined;
    if (authManager && !authManager.isOAuthModeEnabled() && !getRequestTokens()) {
      try {
        accountAccessToken = await authManager.getTokenForAccount(accountParam);
      } catch (err) {
        return finish({
          content: [
            {
              type: 'text',
              text: JSON.stringify({ error: (err as Error).message }),
            },
          ],
          isError: true,
        });
      }
    }

    const parameterDefinitions = tool.parameters || [];

    let path = tool.path;
    const queryParams: Record<string, string> = {};
    const headers: Record<string, string> = {};

    // Body fields the client passed as top-level params (#569) - merged into the
    // request body after the loop
    const bodyShape = bodySchemaShape(
      parameterDefinitions.find((p) => p.type === 'Body')?.schema as z.ZodTypeAny | undefined
    );
    const strayBodyFields: Record<string, unknown> = {};

    for (const [paramName, paramValue] of Object.entries(params)) {
      // Skip control parameters - not part of the Microsoft Graph API
      if (
        [
          'account',
          'confirm',
          'fetchAllPages',
          'includeHeaders',
          'excludeResponse',
          'timezone',
          'expandExtendedProperties',
        ].includes(paramName)
      ) {
        continue;
      }

      // Ok, so, MCP clients (such as claude code) doesn't support $ in parameter names,
      // and others might not support __, so we strip them in hack.ts and restore them here
      const odataParams = [
        'filter',
        'select',
        'expand',
        'orderby',
        'skip',
        'top',
        'count',
        'search',
        'format',
      ];
      // Handle both "top" and "$top" formats - strip $ if present, then re-add it
      const normalizedParamName = paramName.startsWith('$') ? paramName.slice(1) : paramName;
      const isOdataParam = odataParams.includes(normalizedParamName.toLowerCase());
      const fixedParamName = isOdataParam ? `$${normalizedParamName.toLowerCase()}` : paramName;
      // Convert kebab-case param names to camelCase for path param matching.
      // endpoints.json uses {message-id} but hack.ts extracts :messageId (camelCase) from the path.
      // LLMs may pass "message-id" (kebab) — we normalize so both forms work.
      const camelCaseParamName = paramName.replace(/-([a-z])/g, (_, c: string) => c.toUpperCase());
      const skipEncodingParamName = config?.skipEncoding?.find(
        (configuredName) =>
          configuredName === paramName ||
          configuredName.replace(/-([a-z])/g, (_, c: string) => c.toUpperCase()) ===
            camelCaseParamName
      );

      // Look up param definition using normalized name (without $) for OData params,
      // or camelCase equivalent for kebab-case path params
      const paramDef = parameterDefinitions.find(
        (p) =>
          p.name === paramName ||
          p.name === camelCaseParamName ||
          (isOdataParam && p.name === normalizedParamName)
      );

      if (paramDef) {
        switch (paramDef.type) {
          case 'Path': {
            // Check if this parameter should skip URL encoding (for function-style API calls)
            const shouldSkipEncoding = skipEncodingParamName !== undefined;
            // Use encodeURIComponent but preserve '=' which is valid in path segments (RFC 3986)
            // and commonly appears in Microsoft Graph base64-encoded resource IDs.
            // Without this, IDs like "AAMk...AAA=" become "AAMk...AAA%3D" causing 404 errors.
            // First we encode, then unencode. Crazy, check out https://github.com/Softeria/ms-365-mcp-server/issues/245
            const encodedValue = shouldSkipEncoding
              ? prepareUnencodedPathParameter(
                  config!.pathPattern,
                  skipEncodingParamName,
                  paramValue
                )
              : encodeURIComponent(paramValue as string).replace(/%3D/g, '=');

            // Replace both the original param name and the camelCase variant
            // to handle {message-id} (endpoints.json) and :messageId (generated client) formats
            path = path
              .replace(`{${paramName}}`, encodedValue)
              .replace(`:${paramName}`, encodedValue)
              .replace(`{${camelCaseParamName}}`, encodedValue)
              .replace(`:${camelCaseParamName}`, encodedValue);
            break;
          }

          case 'Query':
            if (paramValue !== '' && paramValue != null) {
              queryParams[fixedParamName] = `${paramValue}`;
            }
            break;

          case 'Body':
            if (paramDef.schema) {
              const parseResult = paramDef.schema.safeParse(paramValue);
              if (!parseResult.success) {
                const wrapped = { [paramName]: paramValue };
                const wrappedResult = paramDef.schema.safeParse(wrapped);
                if (wrappedResult.success) {
                  logger.info(
                    `Auto-corrected parameter '${paramName}': AI passed nested field directly, wrapped it as {${paramName}: ...}`
                  );
                  body = wrapped;
                } else {
                  body = paramValue;
                }
              } else {
                body = paramValue;
              }
            } else {
              body = paramValue;
            }
            break;

          case 'Header':
            headers[fixedParamName] = `${paramValue}`;
            break;
        }
      } else if (paramName === 'body') {
        body = paramValue;
        logger.info(`Set body param: ${JSON.stringify(body)}`);
      } else if (
        path.includes(`:${paramName}`) ||
        path.includes(`{${paramName}}`) ||
        path.includes(`:${camelCaseParamName}`) ||
        path.includes(`{${camelCaseParamName}}`)
      ) {
        // Fallback: path param not declared in tool.parameters (generated client omits them).
        // Replace placeholder directly so the URL is valid.
        const encodedValue = skipEncodingParamName
          ? prepareUnencodedPathParameter(config!.pathPattern, skipEncodingParamName, paramValue)
          : encodeURIComponent(paramValue as string).replace(/%3D/g, '=');
        path = path
          .replace(`{${paramName}}`, encodedValue)
          .replace(`:${paramName}`, encodedValue)
          .replace(`{${camelCaseParamName}}`, encodedValue)
          .replace(`:${camelCaseParamName}`, encodedValue);
        logger.info(`Path param fallback: replaced :${camelCaseParamName} with encoded value`);
      } else if (isOdataParam) {
        // Fallback: OData param recognised by name but absent from generated client's parameter
        // list — forward it as a query param rather than silently dropping it.
        queryParams[fixedParamName] = `${paramValue}`;
        logger.info(`OData param fallback: forwarded ${fixedParamName}=${paramValue}`);
      } else if (tool.path === '/$batch' && paramName === 'requests') {
        // The generated graph-batch Body schema is an empty passthrough object, so its
        // shape cannot identify `requests` for the generic flattened-body fallback below.
        // Preserve the advertised top-level calling convention, then let the normalized
        // body guard inspect exactly what will be sent to Graph.
        strayBodyFields.requests = paramValue;
        logger.info("Body field fallback: merging top-level param 'requests' into graph-batch");
      } else if (
        bodyShape &&
        (hasOwn(bodyShape, paramName) || hasOwn(bodyShape, camelCaseParamName)) &&
        !READ_ONLY_BODY_FIELDS.has(hasOwn(bodyShape, paramName) ? paramName : camelCaseParamName)
      ) {
        // Client flattened the body object into top-level params - rescue instead of
        // dropping. The read-only check uses the resolved name so kebab-case variants
        // can't sneak past
        const fieldName = hasOwn(bodyShape, paramName) ? paramName : camelCaseParamName;
        strayBodyFields[fieldName] = paramValue;
        logger.info(
          `Body field fallback: merging top-level param '${fieldName}' into request body`
        );
      } else {
        logger.warn(`Dropping unrecognized parameter '${paramName}' for tool ${tool.alias}`);
      }
    }

    // The client passed the nested itemBody's own fields as the whole request body - move
    // them under the schema's `body` field (#620). Graph body schemas are all-optional, so
    // the safeParse wrap in `case 'Body'` can't catch this: the inner itemBody parses clean
    // as the outer type. A key only moves if it belongs to the nested field's own shape;
    // being unknown to the outer shape means nothing, because generated schemas are trimmed
    // subsets that passthrough real fields (message.isRead isn't in the shape). Keys are
    // matched case-insensitively, and anything left behind stays top-level so a stray
    // sibling can't cost us the repair - or get buried in the body and silently lost
    if (
      isPlainObject(body) &&
      bodyShape != null &&
      hasOwn(bodyShape, 'body') &&
      !Object.keys(body).some((k) => k.toLowerCase() === 'body')
    ) {
      const nestedShape = bodySchemaShape(bodyShape.body as z.ZodTypeAny);
      if (nestedShape != null) {
        const nestedKeys = new Set(Object.keys(nestedShape).map((k) => k.toLowerCase()));
        const outerKeys = new Set(Object.keys(bodyShape).map((k) => k.toLowerCase()));
        // Null-prototype accumulators: a literal would route a '__proto__' key through the
        // legacy setter, dropping the field instead of storing it
        const nested: Record<string, unknown> = Object.create(null);
        const kept: Record<string, unknown> = Object.create(null);
        const annotations: Record<string, unknown> = Object.create(null);

        for (const [key, value] of Object.entries(body)) {
          const lower = key.toLowerCase();
          if (key.startsWith('@')) {
            // An annotation belongs to whatever it names, so only an @odata.type naming the
            // nested type travels with the moved fields. An outer @odata.type (or an
            // @odata.etag) describes the entity it is already on and stays put
            if (lower === '@odata.type' && namesNestedBodyType(value)) {
              annotations[key] = value;
            } else {
              kept[key] = value;
            }
          } else if (nestedKeys.has(lower) && !outerKeys.has(lower)) {
            nested[key] = value;
          } else {
            kept[key] = value;
          }
        }

        if (Object.keys(nested).length > 0) {
          body = { ...kept, body: { ...nested, ...annotations } };
          logger.info(
            `Moved misplaced fields into nested 'body' for ${tool.alias}: ${Object.keys(nested).join(', ')}`
          );
        }
      }
    }

    if (Object.keys(strayBodyFields).length > 0) {
      if (isPlainObject(body)) {
        // Spread order lets an explicit body win over stray duplicates
        body = { ...strayBodyFields, ...body };
        logger.info(`Merged flattened body fields: ${Object.keys(strayBodyFields).join(', ')}`);
      } else if (body == null) {
        body = strayBodyFields;
        logger.info(`Merged flattened body fields: ${Object.keys(strayBodyFields).join(', ')}`);
      } else {
        logger.warn(
          `Cannot merge flattened body fields (${Object.keys(strayBodyFields).join(', ')}) into non-object request body; dropping them`
        );
      }
    }

    // A blocked tool is unreachable by name, but graph-batch carries arbitrary
    // method/url subrequests, so the operation itself has to be checked (#24). Inspect
    // the normalized body, not params.body: passthrough clients may flatten requests to
    // the top level, and the fallback above is what turns that shape into the payload
    // Graph will actually receive.
    // Inline file bytes are refused at any size, on the same normalized body and for the
    // same reason as the batch check below: the payload is what the policy is about, not
    // the tool name. add-mail-attachment carries contentBytes only for a fileAttachment;
    // a referenceAttachment has none and stays allowed, which name-blocking could not
    // express. Recursive, so a contentBytes inside a graph-batch subrequest is caught here
    // too.
    const inlineBytes = findInlineByteFields(body);
    if (inlineBytes.length > 0) {
      recordBlockedOperation(tool.alias, route);
      logger.warn(
        `Refusing ${tool.alias}: request carries inline file bytes at ${describeInlineBytes(inlineBytes)}`
      );
      return finish(
        {
          content: [
            {
              type: 'text',
              text: JSON.stringify({
                error: 'inline_bytes_refused',
                tool: tool.alias,
                fields: inlineBytes.map((hit) => hit.path),
                message:
                  'This server refuses file bytes passed as a tool argument, at any size. ' +
                  'Base64 in the transcript has caused corruption, truncation, and automated ' +
                  'misuse misclassification. Create an upload session instead and PUT the ' +
                  'bytes from disk: create-mail-attachment-upload-session for Outlook ' +
                  'attachments, create-upload-session for drive files. A reference ' +
                  'attachment (sourceUrl, no contentBytes) is also allowed.',
              }),
            },
          ],
          isError: true,
        },
        'blocked'
      );
    }

    if (tool.path === '/$batch') {
      // Count what batching is actually used for, so the question of whether graph-batch
      // earns its keep can be answered from data rather than by grepping skill text.
      recordBatchSubrequestsFor(body, blockedOperations);
    }
    if (blockedOperations.length > 0 && tool.path === '/$batch') {
      const hits = findBlockedSubrequests(body, blockedOperations);
      if (hits.length > 0) {
        for (const hit of hits) recordBlockedOperation(hit.toolName, 'batch');
        logger.warn(
          `Refusing graph-batch: ${hits.length} subrequest(s) match blocked operations: ` +
            hits.map((h) => `${h.method} ${h.url} (${h.toolName})`).join(', ')
        );
        return finish(
          {
            content: [
              {
                type: 'text',
                text: JSON.stringify({
                  error: 'blocked_operation',
                  message: describeBlockedSubrequests(hits),
                  blocked: hits,
                }),
              },
            ],
            isError: true,
          },
          'blocked'
        );
      }
    }

    // Defense-in-depth: the calendar delta tools don't support $top (see
    // TOP_UNSUPPORTED_DELTA_TOOLS). Their user-facing schema strips top/$top, so
    // freshly-connected clients can't send it. Cached/stale clients (and ad-hoc
    // callers) might still try — drop it server-side before clamping or sending.
    if (TOP_UNSUPPORTED_DELTA_TOOLS.has(tool.alias)) {
      delete queryParams['$top'];
    }

    clampTopQueryParam(queryParams);

    const preferValues: string[] = [];

    // Handle timezone parameter for calendar endpoints
    if (config?.supportsTimezone && params.timezone) {
      preferValues.push(`outlook.timezone="${params.timezone}"`);
      logger.info(`Setting timezone preference: outlook.timezone="${params.timezone}"`);
    }

    const bodyFormat = process.env.MS365_MCP_BODY_FORMAT || 'text';
    if (bodyFormat !== 'html' && tool.method.toUpperCase() === 'GET') {
      preferValues.push(`outlook.body-content-type="${bodyFormat}"`);
    }

    if (preferValues.length > 0) {
      headers['Prefer'] = preferValues.join(', ');
    }

    // Handle expandExtendedProperties parameter for calendar endpoints
    if (config?.supportsExpandExtendedProperties && params.expandExtendedProperties === true) {
      const expandValue = 'singleValueExtendedProperties';
      if (queryParams['$expand']) {
        queryParams['$expand'] += `,${expandValue}`;
      } else {
        queryParams['$expand'] = expandValue;
      }
      logger.info(`Adding $expand=${expandValue} for extended properties`);
    }

    if (config?.contentType) {
      headers['Content-Type'] = config.contentType;
      logger.info(`Setting custom Content-Type: ${config.contentType}`);
    }

    if (config?.acceptType) {
      headers['Accept'] = config.acceptType;
      logger.info(`Setting custom Accept: ${config.acceptType}`);
    }

    if (Object.keys(queryParams).length > 0) {
      const queryString = Object.entries(queryParams)
        .map(([key, value]) => `${key}=${encodeURIComponent(value).replace(/%2C/gi, ',')}`)
        .join('&');
      path = `${path}${path.includes('?') ? '&' : '?'}${queryString}`;
    }

    const options: {
      method: string;
      headers: Record<string, string>;
      body?: string | Buffer | Uint8Array;
      rawResponse?: boolean;
      includeHeaders?: boolean;
      excludeResponse?: boolean;
      queryParams?: Record<string, string>;
      accessToken?: string;
      apiVersion?: string;
      forceJsonOutput?: boolean;
    } = {
      method: tool.method.toUpperCase(),
      headers,
    };

    // Route beta-flagged endpoints to the /beta surface; everything else stays on v1.0.
    if (config?.apiVersion) {
      options.apiVersion = config.apiVersion;
    }

    if (options.method !== 'GET' && body) {
      if (tool.requestFormat === 'binary' && typeof body === 'string') {
        options.body = Buffer.from(body, 'base64');
        if (!config?.contentType) {
          headers['Content-Type'] = 'application/octet-stream';
        }
      } else if (config?.contentType === 'text/html') {
        if (typeof body === 'string') {
          options.body = body;
        } else if (typeof body === 'object' && 'content' in body) {
          options.body = (body as { content: string }).content;
        } else {
          options.body = String(body);
        }
      } else {
        options.body = typeof body === 'string' ? body : JSON.stringify(body);
      }
    }

    const isProbablyMediaContent =
      tool.errors?.some((error) => error.description === 'Retrieved media content') ||
      path.endsWith('/content');

    if (config?.returnDownloadUrl && path.endsWith('/content')) {
      path = path.replace(/\/content$/, '');
      logger.info(
        `Auto-returning download URL for ${tool.alias} (returnDownloadUrl=true in endpoints.json)`
      );
    } else if (isProbablyMediaContent) {
      options.rawResponse = true;
    }

    // Set includeHeaders if requested
    if (params.includeHeaders === true) {
      options.includeHeaders = true;
    }

    // Set excludeResponse if requested
    if (params.excludeResponse === true) {
      options.excludeResponse = true;
    }

    // Pass account-resolved token if available
    if (accountAccessToken) {
      options.accessToken = accountAccessToken;
    }

    targetResource = deriveTargetResource({
      pathPattern: config?.pathPattern ?? tool.path,
      params,
    });

    // Redact accessToken from log output to prevent credential leakage
    const { accessToken: _redacted, ...safeOptions } = options;
    logger.info(
      `Making graph request to ${path} with options: ${JSON.stringify(safeOptions)}${_redacted ? ' [accessToken=REDACTED]' : ''}`
    );

    const fetchAllPages = params.fetchAllPages === true;
    const paginationEnabled = paginationAllowed();
    if (fetchAllPages && !paginationEnabled) {
      logger.info(
        'fetchAllPages requested but MS365_MCP_ALLOW_PAGINATION is disabled; returning first page only'
      );
    }
    // Force every page to JSON so the merge loop can parse them. Under --toon they'd
    // be TOON and JSON.parse would throw, silently returning only page one (#560).
    // The merged result gets re-encoded once at the end.
    const mergePages = fetchAllPages && paginationEnabled;
    if (mergePages) {
      options.forceJsonOutput = true;
    }

    let response = await graphClient.graphRequest(path, options);

    if (mergePages && response?.content?.[0]?.text) {
      type ODataPage = {
        value?: unknown[];
        '@odata.nextLink'?: string;
        '@odata.deltaLink'?: string;
        '@odata.count'?: number;
        [key: string]: unknown;
      };
      let combinedResponse: ODataPage | undefined;
      try {
        combinedResponse = JSON.parse(response.content[0].text) as ODataPage;

        // Only merge if page one is actually a collection. fetchAllPages can be set
        // on a single-object GET too, and we'd otherwise graft a bogus value:[] on it.
        const firstValue = combinedResponse.value;
        if (Array.isArray(firstValue)) {
          let allItems: unknown[] = firstValue;
          let nextLink = combinedResponse['@odata.nextLink'];
          let pageCount = 1;
          const maxPages = positiveIntFromEnv('MS365_MCP_MAX_PAGES', DEFAULT_MAX_PAGES);
          const maxItems = positiveIntFromEnv('MS365_MCP_MAX_ITEMS', DEFAULT_MAX_ITEMS);
          // Graph only emits @odata.deltaLink on the final page of a /delta query.
          // Track it across the pagination loop so we can stamp it on the combined
          // response — otherwise fetchAllPages on a /delta endpoint silently drops
          // the resume token and forces callers to re-list from scratch.
          let deltaLink = combinedResponse['@odata.deltaLink'];

          while (nextLink && pageCount < maxPages && allItems.length < maxItems) {
            logger.info(`Fetching page ${pageCount + 1} from: ${nextLink}`);

            // Extract path + query string from the nextLink URL.
            // Pass the full path (with query string) as the endpoint so that
            // $skiptoken and other pagination params are preserved.
            // Previously, query params were extracted into nextOptions.queryParams
            // but graphRequest/performRequest never read that field — they were lost.
            const url = new URL(nextLink);
            // nextLink is absolute and version-qualified (/v1.0/... or /beta/...). Strip the
            // version segment so performRequest can re-apply the request's own apiVersion.
            const nextPath = url.pathname.replace(/^\/(v1\.0|beta)/, '') + url.search;
            const nextOptions = { ...options };

            const nextResponse = await graphClient.graphRequest(nextPath, nextOptions);
            if (nextResponse?.isError) {
              response = nextResponse;
              combinedResponse = undefined;
              break;
            }
            if (nextResponse?.content?.[0]?.text) {
              const nextJsonResponse = JSON.parse(nextResponse.content[0].text) as ODataPage;
              if (Array.isArray(nextJsonResponse.value)) {
                allItems = allItems.concat(nextJsonResponse.value);
              }
              nextLink = nextJsonResponse['@odata.nextLink'];
              if (nextJsonResponse['@odata.deltaLink']) {
                deltaLink = nextJsonResponse['@odata.deltaLink'];
              }
              pageCount++;
            } else {
              break;
            }
          }

          if (combinedResponse !== undefined) {
            if (pageCount >= maxPages) {
              logger.warn(`Reached maximum page limit (${maxPages}) for pagination`);
            }
            if (allItems.length >= maxItems) {
              logger.warn(
                `Reached maximum item limit (${maxItems}) for pagination — truncated at ${allItems.length} items`
              );
            }

            combinedResponse.value = allItems;
            if (combinedResponse['@odata.count']) {
              combinedResponse['@odata.count'] = allItems.length;
            }
            delete combinedResponse['@odata.nextLink'];
            if (deltaLink) {
              combinedResponse['@odata.deltaLink'] = deltaLink;
            }

            logger.info(
              `Pagination complete: collected ${allItems.length} items across ${pageCount} pages`
            );
          }
        }
      } catch (e) {
        logger.error(`Error during pagination: ${e}`);
      }

      // Re-encode once in the configured format. Runs whenever page one parsed
      // (non-collection skip and mid-loop abort included), so a --toon client
      // never gets handed the forced-JSON body.
      if (combinedResponse !== undefined) {
        response.content[0].text = graphClient.serialize(combinedResponse);
      }
    }

    if (response?.content?.[0]?.text) {
      const responseText = response.content[0].text;
      logger.info(`Response size: ${responseText.length} characters`);

      try {
        const jsonResponse = JSON.parse(responseText);
        if (jsonResponse.value && Array.isArray(jsonResponse.value)) {
          logger.info(`Response contains ${jsonResponse.value.length} items`);
        }
        if (jsonResponse['@odata.nextLink']) {
          logger.info(`Response has pagination nextLink: ${jsonResponse['@odata.nextLink']}`);
        }
      } catch {
        // Non-JSON response
      }
    }

    // Convert McpResponse to CallToolResult with the correct structure
    const content: ContentItem[] = response.content.map((item) => ({
      type: 'text' as const,
      text: item.text,
    }));

    auditLog({
      event: 'tool.call',
      request_id: requestId,
      user_principal_name: upn,
      tool: tool.alias,
      http_method: httpMethod,
      status: response.isError ? 'error' : 'success',
      duration_ms: Date.now() - startTime,
      ...(targetResource ? { target_resource: targetResource } : {}),
      ...graphResponseAuditFields(response),
      ...recipientAuditFields(body),
    });
    return finish({
      content,
      _meta: response._meta,
      isError: response.isError,
    });
  } catch (error) {
    const err = error as { name?: string; code?: string | number; status?: string | number };
    logger.error(`Error in tool ${tool.alias}: ${(error as Error).message}`);
    auditLog({
      event: 'tool.call',
      request_id: requestId,
      user_principal_name: upn,
      tool: tool.alias,
      http_method: httpMethod,
      status: 'error',
      duration_ms: Date.now() - startTime,
      ...(targetResource ? { target_resource: targetResource } : {}),
      error_type: err?.name || 'Error',
      ...thrownErrorAuditFields(error),
      ...recipientAuditFields(body),
    });
    return finish({
      content: [
        {
          type: 'text',
          text: JSON.stringify({
            error: `Error in tool ${tool.alias}: ${(error as Error).message}`,
          }),
        },
      ],
      isError: true,
    });
  }
}

export function registerGraphTools(
  server: McpServer,
  graphClient: GraphClient,
  readOnly: boolean = false,
  enabledToolsPattern?: ToolNameMatcher,
  orgMode: boolean = false,
  authManager?: AuthManager,
  multiAccount: boolean = false,
  accountNames: string[] = [],
  allowedScopesValue?: string,
  httpMode: boolean = false,
  blockedToolsPattern?: string,
  publicBaseUrl?: string
): number {
  const blockedToolsRegex = compileBlockedToolsRegex(blockedToolsPattern);
  // Operations the blocklist prohibits, so graph-batch cannot carry one as a
  // subrequest (#24).
  const blockedOperations = buildBlockedOperationMatchers(blockedToolsPattern);
  // Give those series a zero baseline now, so increase() can see the first refusal.
  initBlockedOperationSeries([...new Set(blockedOperations.map((m) => m.toolName))]);
  let enabledToolsMatches: ((name: string) => boolean) | undefined;
  if (typeof enabledToolsPattern === 'function') {
    enabledToolsMatches = enabledToolsPattern;
  } else if (enabledToolsPattern) {
    try {
      const enabledToolsRegex = new RegExp(enabledToolsPattern, 'i');
      enabledToolsMatches = (name) => enabledToolsRegex.test(name);
      logger.info(`Tool filtering enabled with pattern: ${enabledToolsPattern}`);
    } catch {
      logger.error(`Invalid tool filter regex pattern: ${enabledToolsPattern}. Ignoring filter.`);
    }
  }

  let registeredCount = 0;
  let skippedCount = 0;
  let failedCount = 0;
  const allowedScopes = parseAllowedScopes(allowedScopesValue);
  const disabledByAllowedScopes: DisabledToolScope[] = [];
  const deniedTools = collectDeniedToolPolicies({
    readOnly,
    orgMode,
    enabledToolsMatches,
    allowedScopesValue,
    httpMode,
  });

  for (const tool of allEndpoints) {
    const endpointConfig = endpointsData.find((e) => e.toolName === tool.alias);
    if (!orgMode && endpointConfig && !endpointConfig.scopes && endpointConfig.workScopes) {
      logger.info(`Skipping work account tool ${tool.alias} - not in org mode`);
      skippedCount++;
      continue;
    }

    const method = tool.method.toUpperCase();
    if (readOnly && method !== 'GET') {
      // Allow POST endpoints that are explicitly marked as readOnly in endpoints.json
      // (e.g. get-schedule, find-meeting-times which are read-only queries via POST).
      // PATCH/DELETE are always blocked in read-only mode.
      if (!(method === 'POST' && endpointConfig?.readOnly)) {
        logger.info(`Skipping write operation ${tool.alias} in read-only mode`);
        skippedCount++;
        continue;
      }
    }

    if (enabledToolsMatches && !enabledToolsMatches(tool.alias)) {
      logger.info(`Skipping tool ${tool.alias} - doesn't match filter pattern`);
      skippedCount++;
      continue;
    }

    // The blocklist wins over any enable pattern, including an explicit --direct-tools
    // selection, so a blocked tool is unreachable by every path.
    if (blockedToolsRegex && blockedToolsRegex.test(tool.alias)) {
      logger.info(`Blocking tool ${tool.alias} - matches blocklist pattern`);
      skippedCount++;
      continue;
    }

    const missingScopes =
      allowedScopes !== undefined && !endpointConfig
        ? ['endpoint scope metadata']
        : getMissingAllowedScopesForGroups(
            getEndpointScopeGroups(endpointConfig, orgMode),
            allowedScopes
          );
    if (missingScopes.length > 0) {
      disabledByAllowedScopes.push({ toolName: tool.alias, missingScopes });
      skippedCount++;
      continue;
    }

    const paramSchema: Record<string, z.ZodTypeAny> = {};
    if (tool.parameters && tool.parameters.length > 0) {
      for (const param of tool.parameters) {
        // Lenient Body validation, or the SDK strips a flattened body value to {} (#569)
        paramSchema[param.name] =
          param.type === 'Body' && param.schema
            ? lenientBodySchema(param.schema as z.ZodTypeAny)
            : param.schema || z.any();
      }
    }

    // Extract path parameters from the path pattern (e.g., :todoTaskListId from /me/todo/lists/:todoTaskListId/tasks)
    // The generated client omits these from tool.parameters, so we add them manually.
    const pathParamMatches = tool.path.matchAll(/:([a-zA-Z]+)/g);
    for (const match of pathParamMatches) {
      const pathParamName = match[1];
      if (!(pathParamName in paramSchema)) {
        paramSchema[pathParamName] = z.string().describe(describePathParam(pathParamName));
      }
    }

    // Raw path interpolation is needed for a few Graph function and colon-addressing
    // routes. Refine their public schemas, then enforce the same rules again during
    // execution because execute-tool and stale clients do not necessarily pass through
    // the registered direct-tool schema.
    for (const paramName of endpointConfig?.skipEncoding ?? []) {
      const camelCaseParamName = paramName.replace(/-([a-z])/g, (_, c: string) => c.toUpperCase());
      const schemaName = paramSchema[paramName] !== undefined ? paramName : camelCaseParamName;
      const schema = paramSchema[schemaName];
      if (!schema) {
        throw new Error(
          `skipEncoding parameter ${JSON.stringify(paramName)} has no schema for ${tool.alias}`
        );
      }
      paramSchema[schemaName] = refineUnencodedPathParameterSchema(
        schema,
        endpointConfig!.pathPattern,
        paramName
      );
    }

    if (isFetchAllPagesApplicable(tool)) {
      const maxPages = getMaxPages();
      paramSchema['fetchAllPages'] = z
        .boolean()
        .describe(getFetchAllPagesParamDescription(maxPages))
        .optional();
    }

    // Override OData parameter descriptions with spec-gap guidance. Text lives in
    // lib/param-descriptions.ts, shared with describeToolSchema (--discovery mode),
    // so the two paths cannot describe the same parameter differently.
    if (paramSchema['filter'] !== undefined || paramSchema['$filter'] !== undefined) {
      const key = paramSchema['$filter'] !== undefined ? '$filter' : 'filter';
      paramSchema[key] = z.string().describe(FILTER_PARAM_DESCRIPTION).optional();
    }
    if (paramSchema['search'] !== undefined || paramSchema['$search'] !== undefined) {
      const key = paramSchema['$search'] !== undefined ? '$search' : 'search';
      paramSchema[key] = z.string().describe(SEARCH_PARAM_DESCRIPTION).optional();
    }
    // Accept an array as well as a comma-separated string (EnviroKinetics/ms365-mcp#48).
    // The query serializer interpolates the value, and JS joins an array with commas, so
    // both shapes already reach Graph correctly; forcing z.string() here only meant the
    // SDK rejected a real array locally, and the model's workaround was to send the JSON
    // TEXT `["id","subject"]`, which passed through verbatim and produced Graph's
    // 400 "An identifier was expected at position 0". $expand below is already an array,
    // so this makes the OData params consistent rather than introducing a new shape.
    if (paramSchema['select'] !== undefined || paramSchema['$select'] !== undefined) {
      const key = paramSchema['$select'] !== undefined ? '$select' : 'select';
      paramSchema[key] = z
        .union([z.string(), z.array(z.string())])
        .describe(SELECT_PARAM_DESCRIPTION)
        .optional();
    }
    // The spec describes every $expand as "Expand related entities", which says nothing about
    // what is expandable. Models pass non-navigation properties — message body is the one I
    // hit repeatedly — and Graph answers 400 "Parsing OData Select and Expand failed".
    // Restated as the override rather than a new schema: $expand is already array<string>
    // everywhere, so the type is unchanged in practice.
    if (paramSchema['expand'] !== undefined || paramSchema['$expand'] !== undefined) {
      const key = paramSchema['$expand'] !== undefined ? '$expand' : 'expand';
      paramSchema[key] = z.array(z.string()).describe(EXPAND_PARAM_DESCRIPTION).optional();
    }
    if (paramSchema['orderby'] !== undefined || paramSchema['$orderby'] !== undefined) {
      const key = paramSchema['$orderby'] !== undefined ? '$orderby' : 'orderby';
      paramSchema[key] = z
        .union([z.string(), z.array(z.string())])
        .describe(ORDERBY_PARAM_DESCRIPTION)
        .optional();
    }
    // The calendar delta tools don't support $top (see TOP_UNSUPPORTED_DELTA_TOOLS) —
    // page size is controlled via Prefer: odata.maxpagesize. Strip top/$top from
    // their schemas so callers can't reach for a parameter that won't work. Other
    // delta tools (message/driveItem/site) do support $top, so leave them alone.
    // Server-side defense-in-depth in executeGraphTool handles stale clients.
    if (shouldOmitTopParam(tool.alias)) {
      delete paramSchema['top'];
      delete paramSchema['$top'];
    } else if (paramSchema['top'] !== undefined || paramSchema['$top'] !== undefined) {
      const key = paramSchema['$top'] !== undefined ? '$top' : 'top';
      paramSchema[key] = z.number().describe(TOP_PARAM_DESCRIPTION).optional();
    }
    if (paramSchema['skip'] !== undefined || paramSchema['$skip'] !== undefined) {
      const key = paramSchema['$skip'] !== undefined ? '$skip' : 'skip';
      paramSchema[key] = z.number().describe(SKIP_PARAM_DESCRIPTION).optional();
    }
    if (paramSchema['count'] !== undefined || paramSchema['$count'] !== undefined) {
      const countKey = paramSchema['$count'] !== undefined ? '$count' : 'count';
      paramSchema[countKey] = z.boolean().describe(COUNT_PARAM_DESCRIPTION).optional();
    }

    // Add account parameter for multi-account mode.
    // Layer 2: Account names are surfaced in the description (not as a strict enum) so the LLM
    // sees available accounts upfront without a round-trip, but accounts added mid-session via
    // --login are still accepted — getTokenForAccount() handles validation at runtime.
    if (multiAccount) {
      paramSchema['account'] = z
        .string()
        .describe(getAccountParamDescription(accountNames))
        .optional();
    }

    // Add includeHeaders parameter for all tools to capture ETags and other headers
    paramSchema['includeHeaders'] = z
      .boolean()
      .describe('Include response headers (including ETag) in the response metadata')
      .optional();

    // Add excludeResponse parameter to only return success/failure indication
    paramSchema['excludeResponse'] = z
      .boolean()
      .describe('Exclude the full response body and only return success or failure indication')
      .optional();

    // Destructive tools (POST except readOnly, PATCH, PUT, DELETE) require an
    // explicit `confirm: true` server-side gate. See isDestructiveOperation +
    // executeGraphTool for the enforcement; surface the param in the schema so
    // the LLM/agent sees it upfront.
    const destructive = isDestructiveOperation(tool.method, endpointConfig);
    if (destructive) {
      paramSchema['confirm'] = z.boolean().describe(CONFIRM_PARAM_DESCRIPTION).optional();
    }

    // Add timezone parameter for calendar endpoints that support it
    if (endpointConfig?.supportsTimezone) {
      paramSchema['timezone'] = z.string().describe(TIMEZONE_PARAM_DESCRIPTION).optional();
    }

    // Add expandExtendedProperties parameter for calendar endpoints that support it
    if (endpointConfig?.supportsExpandExtendedProperties) {
      paramSchema['expandExtendedProperties'] = z
        .boolean()
        .describe(EXPAND_EXTENDED_PROPERTIES_PARAM_DESCRIPTION)
        .optional();
    }

    // Build the tool description, optionally appending LLM tips
    let toolDescription = withApiVersionPrefix(
      (endpointConfig?.descriptionOverride ?? tool.description) ||
        `Execute ${tool.method.toUpperCase()} request to ${tool.path}`,
      endpointConfig
    );
    if (endpointConfig?.llmTip) {
      toolDescription += `\n\n💡 TIP: ${endpointConfig.llmTip}`;
    }

    // An endpoint marked readOnly in endpoints.json (e.g. a POST query like
    // copilot-retrieve) is a read-only operation despite its write verb, so derive
    // the hints from that flag rather than the HTTP method alone — otherwise a
    // read-only query lands as destructiveHint:true and clients mis-rank it.
    const isReadOnlyTool = tool.method.toUpperCase() === 'GET' || endpointConfig?.readOnly === true;

    try {
      // .passthrough() object, not a raw shape - the SDK wraps raw shapes in z.object()
      // and strips unknown keys before the handler runs, which is exactly how #569's
      // flattened subject/toRecipients got lost
      server.registerTool(
        tool.alias,
        {
          title: tool.alias,
          description: toolDescription,
          inputSchema: z.object(paramSchema).passthrough(),
          annotations: {
            title: tool.alias,
            readOnlyHint: isReadOnlyTool,
            destructiveHint: destructive,
            openWorldHint: true, // All tools call Microsoft Graph API
          },
        },
        async (params: Record<string, unknown>) =>
          executeGraphTool(
            tool,
            endpointConfig,
            graphClient,
            params,
            authManager,
            blockedOperations
          )
      );
      registeredCount++;
    } catch (error) {
      logger.error(`Failed to register tool ${tool.alias}: ${(error as Error).message}`);
      failedCount++;
    }
  }

  if (multiAccount) {
    logger.info('Multi-account mode: "account" parameter injected into all tool schemas');
  }

  if (disabledByAllowedScopes.length > 0) {
    logger.info(
      `Allowed scopes disabled ${disabledByAllowedScopes.length} Graph tools: ${formatDisabledToolsForLog(disabledByAllowedScopes)}`
    );
  }

  const utilityCtx: UtilityToolContext = {
    graphClient,
    authManager,
    multiAccount,
    accountNames,
    httpMode,
    publicBaseUrl,
  };
  for (const utility of UTILITY_TOOLS) {
    if (readOnly && utility.mutatesState) continue;
    if (httpMode && utility.stdioOnly) continue;
    if (enabledToolsMatches && !enabledToolsMatches(utility.name)) continue;
    if (blockedToolsRegex && blockedToolsRegex.test(utility.name)) continue;
    try {
      registerUtilityToolWithMcp(server, utility, utilityCtx);
      registeredCount++;
    } catch (error) {
      logger.error(`Failed to register tool ${utility.name}: ${(error as Error).message}`);
      failedCount++;
    }
  }

  // Layer 3 (list-accounts tool) is registered by registerAuthTools in auth-tools.ts.
  // It is the canonical owner of account discovery — no duplicate registration here.

  logger.info(
    `Tool registration complete: ${registeredCount} registered, ${skippedCount} skipped, ${failedCount} failed`
  );
  installDeniedToolAuditHandler(server, deniedTools);
  return registeredCount;
}

export function buildToolsRegistry(
  readOnly: boolean,
  orgMode: boolean,
  enabledToolsMatches?: CompiledToolNameMatcher,
  allowedScopesValue?: string,
  disabledByAllowedScopes: Array<{ toolName: string; missingScopes: string[] }> = [],
  blockedToolsRegex?: RegExp
): Map<string, { tool: (typeof api.endpoints)[0]; config: EndpointConfig | undefined }> {
  const toolsMap = new Map<
    string,
    { tool: (typeof api.endpoints)[0]; config: EndpointConfig | undefined }
  >();
  const allowedScopes = parseAllowedScopes(allowedScopesValue);

  for (const tool of allEndpoints) {
    const endpointConfig = endpointsData.find((e) => e.toolName === tool.alias);

    if (!orgMode && endpointConfig && !endpointConfig.scopes && endpointConfig.workScopes) {
      continue;
    }

    const method = tool.method.toUpperCase();
    if (readOnly && method !== 'GET') {
      if (!(method === 'POST' && endpointConfig?.readOnly)) {
        continue;
      }
    }

    if (enabledToolsMatches && !toolNameMatches(enabledToolsMatches, tool.alias)) {
      continue;
    }

    // A blocked tool is unreachable by every path: direct registration, the
    // discovery registry, search-tools, get-tool-schema and execute-tool.
    if (blockedToolsRegex && blockedToolsRegex.test(tool.alias)) {
      continue;
    }

    const missingScopes =
      allowedScopes !== undefined && !endpointConfig
        ? ['endpoint scope metadata']
        : getMissingAllowedScopesForGroups(
            getEndpointScopeGroups(endpointConfig, orgMode),
            allowedScopes
          );
    if (missingScopes.length > 0) {
      disabledByAllowedScopes.push({ toolName: tool.alias, missingScopes });
      continue;
    }

    toolsMap.set(tool.alias, { tool, config: endpointConfig });
  }

  return toolsMap;
}

/**
 * Builds a BM25 index over the tool registry. Name tokens are weighted 3x and llmTip
 * tokens 2x via repetition, so a tool whose name matches the query outranks one that
 * merely mentions the query term in its Microsoft-supplied description.
 */
export function buildDiscoverySearchIndex(
  toolsRegistry: ReturnType<typeof buildToolsRegistry>,
  utilityTools: readonly UtilityTool[] = []
): DiscoverySearchIndex {
  // Cap contribution from the `description` and `llmTip` fields so a verbose llmTip
  // (e.g. the KQL search-syntax guide on list-mail-messages, ~300 tokens) doesn't
  // inflate a tool's doc length and crush BM25's length normalization. Names and
  // paths are short and reliable, so they stay uncapped and are repeated to carry
  // the bulk of the ranking signal. Tip excerpt (12 tokens) is enough to capture
  // the first "what this tool does" phrase without swamping the doc.
  const TIP_EXCERPT_TOKENS = 12;
  const DESC_CAP_TOKENS = 40;
  const docs: Array<{ id: string; tokens: string[] }> = [];
  const nameTokens = new Map<string, Set<string>>();
  for (const [name, { tool, config }] of toolsRegistry) {
    const nt = tokenize(name);
    nameTokens.set(name, new Set(nt));
    const pathTokens = tokenize(tool.path);
    const descTokens = tokenize(config?.descriptionOverride ?? tool.description).slice(
      0,
      DESC_CAP_TOKENS
    );
    const tipTokens = tokenize(config?.llmTip).slice(0, TIP_EXCERPT_TOKENS);
    const tokens = [
      ...nt,
      ...nt,
      ...nt,
      ...nt,
      ...nt,
      ...pathTokens,
      ...pathTokens,
      ...tipTokens,
      ...descTokens,
    ];
    docs.push({ id: name, tokens });
  }
  for (const utility of utilityTools) {
    const nt = tokenize(utility.name);
    nameTokens.set(utility.name, new Set(nt));
    const pathTokens = tokenize(utility.path);
    const keywordTokens = tokenize(utility.searchKeywords);
    const descTokens = tokenize(utility.description).slice(0, DESC_CAP_TOKENS);
    const tokens = [
      ...nt,
      ...nt,
      ...nt,
      ...nt,
      ...nt,
      ...pathTokens,
      ...pathTokens,
      ...keywordTokens,
      ...keywordTokens,
      ...descTokens,
    ];
    docs.push({ id: utility.name, tokens });
  }
  return { bm25: buildBM25Index(docs), nameTokens };
}

/**
 * BM25 + a "name precision" bonus: reward tools whose names contain a high fraction
 * of the query tokens (and consist mostly of query-matching tokens). This counteracts
 * cases where a tool with a longer or more off-topic description outranks a tool
 * whose name directly matches — a common problem because many endpoint descriptions
 * are the wrong Graph prose pasted in.
 */
export function scoreDiscoveryQuery(
  query: string,
  index: DiscoverySearchIndex
): Array<{ id: string; score: number }> {
  const queryTokenSet = new Set(tokenize(query));
  if (queryTokenSet.size === 0) return [];
  const ranked = scoreQuery(query, index.bm25);
  const NAME_BONUS_WEIGHT = 2;
  for (const r of ranked) {
    const nt = index.nameTokens.get(r.id);
    if (!nt || nt.size === 0) continue;
    let matchedIdf = 0;
    let matchedCount = 0;
    for (const qt of queryTokenSet) {
      if (nt.has(qt)) {
        matchedCount++;
        matchedIdf += index.bm25.idf.get(qt) ?? 0;
      }
    }
    if (matchedCount === 0) continue;
    const precision = matchedCount / nt.size;
    r.score += precision * matchedIdf * NAME_BONUS_WEIGHT;
  }
  ranked.sort((a, b) => b.score - a.score);
  return ranked;
}

export function registerDiscoveryTools(
  server: McpServer,
  graphClient: GraphClient,
  readOnly: boolean = false,
  orgMode: boolean = false,
  authManager?: AuthManager,
  multiAccount: boolean = false,
  accountNames: string[] = [],
  enabledTools?: string,
  allowedScopesValue?: string,
  httpMode: boolean = false,
  blockedToolsPattern?: string,
  directToolsPattern?: ToolNameMatcher,
  publicBaseUrl?: string
): void {
  const blockedToolsRegex = compileBlockedToolsRegex(blockedToolsPattern);
  // execute-tool can dispatch graph-batch, so the same operation check applies here (#24).
  const blockedOperations = buildBlockedOperationMatchers(blockedToolsPattern);
  // Give those series a zero baseline now, so increase() can see the first refusal.
  initBlockedOperationSeries([...new Set(blockedOperations.map((m) => m.toolName))]);

  // Hybrid mode registers some tools by name and leaves the rest reachable only
  // through execute-tool. Discovery output must say which, per tool: a payload whose
  // `name` field looks callable leads a model to call it directly and get
  // "Tool ... not found" (see EnviroKinetics/ms365-mcp#29). An invalid pattern here is
  // not fatal, unlike the blocklist: the consequence is a wrong hint, not a policy hole.
  let directToolsMatches: ((name: string) => boolean) | undefined;
  if (typeof directToolsPattern === 'function') {
    directToolsMatches = directToolsPattern;
  } else if (directToolsPattern) {
    try {
      const directToolsRegex = new RegExp(directToolsPattern, 'i');
      directToolsMatches = (name) => directToolsRegex.test(name);
    } catch (error) {
      logger.error(
        `Invalid --direct-tools regex ${JSON.stringify(directToolsPattern)} for discovery hints; ` +
          `treating every tool as execute-tool only: ${(error as Error).message}`
      );
    }
  }

  /** Route a caller must use to invoke `name` in this configuration. */
  const invokeVia = (name: string): 'direct' | 'execute-tool' =>
    directToolsMatches?.(name) ? 'direct' : 'execute-tool';

  const invocationFor = (name: string) =>
    invokeVia(name) === 'direct'
      ? {
          via: 'direct' as const,
          note: `${name} is registered as a named tool in this configuration; call it directly.`,
        }
      : {
          via: 'execute-tool' as const,
          note: `${name} is not registered as a named tool here, so it cannot be called directly. Invoke it through execute-tool.`,
          example: { tool_name: name, parameters: {} as Record<string, unknown> },
        };
  let enabledToolsRegex: RegExp | undefined;
  if (enabledTools) {
    try {
      enabledToolsRegex = new RegExp(enabledTools, 'i');
      logger.info(`Discovery mode: filtering tools with pattern ${enabledTools}`);
    } catch (error) {
      logger.error(
        `Invalid --enabled-tools regex ${JSON.stringify(enabledTools)} — ignoring filter: ${(error as Error).message}`
      );
    }
  }

  const disabledByAllowedScopes: Array<{ toolName: string; missingScopes: string[] }> = [];
  const deniedTools = collectDeniedToolPolicies({
    readOnly,
    orgMode,
    enabledToolsMatches: enabledToolsRegex ? (name) => enabledToolsRegex.test(name) : undefined,
    allowedScopesValue,
    httpMode,
  });
  const toolsRegistry = buildToolsRegistry(
    readOnly,
    orgMode,
    enabledToolsRegex ? (name) => enabledToolsRegex.test(name) : undefined,
    allowedScopesValue,
    disabledByAllowedScopes,
    // Keeps blocked tools out of the registry itself, which is what execute-tool,
    // get-tool-schema and search-tools all read from.
    blockedToolsRegex
  );
  if (disabledByAllowedScopes.length > 0) {
    logger.info(
      `Discovery mode: allowed scopes disabled ${disabledByAllowedScopes.length} Graph tools: ${formatDisabledToolsForLog(disabledByAllowedScopes)}`
    );
  }
  const utilityTools = UTILITY_TOOLS.filter((u) => {
    if (readOnly && u.mutatesState) return false;
    if (httpMode && u.stdioOnly) return false;
    if (enabledToolsRegex && !enabledToolsRegex.test(u.name)) return false;
    if (blockedToolsRegex && blockedToolsRegex.test(u.name)) return false;
    return true;
  });
  const searchIndex = buildDiscoverySearchIndex(toolsRegistry, utilityTools);
  const totalCount = toolsRegistry.size + utilityTools.length;
  logger.info(
    `Discovery mode: ${totalCount} tools (${toolsRegistry.size} Graph + ${utilityTools.length} utility)`
  );

  const utilityCtx: UtilityToolContext = {
    graphClient,
    authManager,
    multiAccount,
    accountNames,
    httpMode,
    publicBaseUrl,
  };
  const utilityByName = new Map(utilityTools.map((u) => [u.name, u]));

  const categoryNames = Object.keys(TOOL_CATEGORIES).join(', ');

  const toResultEntry = (name: string) => {
    const entry = toolsRegistry.get(name);
    if (entry) {
      const { tool, config } = entry;
      return {
        name,
        method: tool.method.toUpperCase(),
        path: tool.path,
        description: withApiVersionPrefix(
          (config?.descriptionOverride ?? tool.description) ||
            `${tool.method.toUpperCase()} ${tool.path}`,
          config
        ),
        ...(config?.llmTip ? { llmTip: config.llmTip } : {}),
        invoke_via: invokeVia(name),
      };
    }
    const utility = utilityByName.get(name);
    if (utility) {
      return {
        name: utility.name,
        method: utility.method,
        path: utility.path,
        description: utility.description,
        invoke_via: invokeVia(utility.name),
      };
    }
    return null;
  };

  server.tool(
    'search-tools',
    `Search through ${totalCount} tools (${toolsRegistry.size} Microsoft Graph API operations + ${utilityTools.length} server utilities like download-bytes). Ranks results by BM25 over tool name, llmTip, description, and path. After picking a tool, call get-tool-schema for parameters, then execute-tool.`,
    {
      query: z
        .string()
        .describe(
          'Natural-language query. Tokenized and BM25-ranked. E.g. "send email", "download photo", "list unread messages".'
        )
        .optional(),
      category: z.string().describe(`Optional pre-filter by category: ${categoryNames}`).optional(),
      limit: z.number().describe('Maximum results (default: 10, max: 50)').optional(),
    },
    {
      title: 'search-tools',
      readOnlyHint: true,
      openWorldHint: true,
    },
    async ({ query, category, limit = 10 }) => {
      recordDiscoveryStage('search_tools');
      const maxLimit = Math.min(Math.max(limit, 1), 50);
      const categoryDef = category ? TOOL_CATEGORIES[category] : undefined;
      const categoryFilter = (name: string) => !categoryDef || categoryDef.pattern.test(name);

      let orderedNames: string[];
      if (query && query.trim().length > 0) {
        const ranked = scoreDiscoveryQuery(query, searchIndex);
        orderedNames = ranked.map((r) => r.id).filter(categoryFilter);
      } else {
        orderedNames = [...toolsRegistry.keys(), ...utilityTools.map((u) => u.name)].filter(
          categoryFilter
        );
      }

      const tools = orderedNames.slice(0, maxLimit).map(toResultEntry).filter(Boolean);

      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify(
              {
                found: tools.length,
                total: totalCount,
                tools,
                tip: 'Call get-tool-schema(tool_name) to see parameters before invoking execute-tool.',
              },
              null,
              2
            ),
          },
        ],
      };
    }
  );

  server.tool(
    'get-tool-schema',
    'Returns the full parameter schema (name, placement, required, JSON Schema) for a tool discovered via search-tools. Call this before execute-tool so you know what parameters to pass and what enum values are valid.',
    {
      tool_name: z.string().describe('Exact tool name from search-tools (e.g. "send-mail")'),
    },
    {
      title: 'get-tool-schema',
      readOnlyHint: true,
      openWorldHint: false,
    },
    async ({ tool_name }) => {
      recordDiscoveryStage('get_tool_schema');
      const entry = toolsRegistry.get(tool_name);
      if (entry) {
        const schema = describeToolSchema(entry.tool, entry.config, { multiAccount, accountNames });
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify({ ...schema, invocation: invocationFor(tool_name) }, null, 2),
            },
          ],
        };
      }
      const utility = utilityByName.get(tool_name);
      if (utility) {
        const schema = describeUtilityToolSchema(
          { ...utility, buildSchema: () => buildUtilitySchema(utility, utilityCtx) },
          utilityCtx
        );
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify({ ...schema, invocation: invocationFor(tool_name) }, null, 2),
            },
          ],
        };
      }
      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify({
              error: `Tool not found: ${tool_name}`,
              tip: 'Use search-tools to find available tools.',
            }),
          },
        ],
        isError: true,
      };
    }
  );

  server.tool(
    'execute-tool',
    'Execute a Microsoft Graph API tool by name. Workflow: search-tools → get-tool-schema → execute-tool. Call get-tool-schema first for any tool you have not seen before — passing the wrong shape to parameters will fail validation or return a Graph 400. For list endpoints, prefer modest $top plus $select.',
    {
      tool_name: z.string().describe('Name of the tool to execute (e.g., "list-mail-messages")'),
      parameters: z
        .record(z.any())
        .describe(
          'Parameters shaped per get-tool-schema. Path/query/header params go at the top level; request bodies go under "body".'
        )
        .optional(),
    },
    {
      title: 'execute-tool',
      readOnlyHint: false,
      destructiveHint: true,
      openWorldHint: true,
    },
    async ({ tool_name, parameters = {} }) => {
      recordDiscoveryStage('execute_tool');
      const toolData = toolsRegistry.get(tool_name);
      if (toolData) {
        return executeGraphTool(
          toolData.tool,
          toolData.config,
          graphClient,
          parameters,
          authManager,
          blockedOperations,
          'execute_tool'
        );
      }
      const utility = utilityByName.get(tool_name);
      if (utility) {
        return executeUtilityTool(utility, parameters, utilityCtx, 'execute_tool');
      }
      const deniedPolicy = deniedTools.get(tool_name);
      if (deniedPolicy) {
        auditToolDenied(deniedPolicy, parameters);
      }
      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify({
              error: `Tool not found: ${tool_name}`,
              tip: 'Use search-tools to find available tools.',
            }),
          },
        ],
        isError: true,
      };
    }
  );

  installDeniedToolAuditHandler(server, deniedTools);

  // Layer 3 (list-accounts) is registered by registerAuthTools — no duplicate here.
}

// Re-exported so existing importers keep working after the helper moved to lib/.
export { compileBlockedToolsRegex };
