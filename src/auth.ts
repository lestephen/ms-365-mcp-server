import type { AccountInfo, Configuration, ICachePlugin, TokenCacheContext } from '@azure/msal-node';
import { AuthError, PublicClientApplication } from '@azure/msal-node';
import logger from './logger.js';
import { readFileSync } from 'fs';
import { fileURLToPath } from 'url';
import path from 'path';
import { getSecrets, type AppSecrets } from './secrets.js';
import { getCloudEndpoints, getDefaultClientId } from './cloud-config.js';
import {
  createTokenCacheStorage,
  DefaultTokenCacheStorage,
  getSelectedAccountPath,
  getTokenCachePath,
  pickNewest,
  type TokenCacheStorage,
  unwrapCache,
  wrapCache,
} from './token-cache-storage.js';

interface EndpointConfig {
  pathPattern: string;
  method: string;
  toolName: string;
  // A flat string[] is a single AND-group (all scopes required). A nested string[][]
  // expresses alternatives: the endpoint is satisfied if ALL scopes in ANY one group are
  // held (e.g. copilot-retrieve needs Files.Read.All + Sites.Read.All, OR ExternalItem.Read.All).
  scopes?: string[] | string[][];
  workScopes?: string[] | string[][];
  llmTip?: string;
  readOnly?: boolean;
  presets?: string[]; // Presets this endpoint belongs to (mail, outlook, personal, ...)
}

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const endpointsData = JSON.parse(
  readFileSync(path.join(__dirname, 'endpoints.json'), 'utf8')
) as EndpointConfig[];

const endpoints = {
  default: endpointsData,
};

/**
 * Creates MSAL configuration from secrets.
 * This is called during AuthManager initialization.
 */
function createMsalConfig(secrets: AppSecrets): Configuration {
  const cloudEndpoints = getCloudEndpoints(secrets.cloudType);
  return {
    auth: {
      clientId: secrets.clientId || getDefaultClientId(secrets.cloudType),
      authority: `${cloudEndpoints.authority}/${secrets.tenantId || 'common'}`,
    },
  };
}

/**
 * Builds an MSAL cache plugin that keeps the file-backed token cache coherent across
 * concurrent processes. In the common stdio deployment several MCP server processes share
 * one token cache file. Microsoft rotates refresh tokens on silent refresh, so without this
 * plugin a process holds whatever it loaded at startup and fails once a sibling rotates the
 * refresh token on disk (invalid_grant / no_tokens_found). See issue #545.
 *
 * MSAL invokes beforeCacheAccess/afterCacheAccess around every cache operation, so:
 *  - beforeCacheAccess reloads the newest persisted cache into MSAL right before each access
 *  - afterCacheAccess persists (atomically, via storage) only when MSAL changed the cache
 * This preserves the existing cache envelope (wrapCache) and the storage provider's
 * fail-closed semantics.
 *
 * This is best-effort, last-writer-wins reconciliation (the storage layer breaks ties via
 * the savedAt stamp), not a cross-process lock. It closes the dominant #545 window - a
 * long-lived sibling refreshing against a token another process already rotated on disk.
 * Two known limits remain, both accepted as out of scope in the #545 discussion:
 *  - Two processes refreshing the very same refresh token at the same instant can still race.
 *  - A sibling logout is not reflected in an already-running process: load() returns nothing
 *    so the deletion is not picked up (the deserialize is guarded on a present cache), and a
 *    later successful silent acquire here persists the in-memory cache, recreating the file.
 */
export function buildDiskCoherencyCachePlugin(storage: TokenCacheStorage): ICachePlugin {
  return {
    beforeCacheAccess: async (context: TokenCacheContext) => {
      try {
        const cacheRaw = await storage.load('token-cache');
        if (cacheRaw) {
          context.tokenCache.deserialize(unwrapCache(cacheRaw).data);
        }
      } catch (error) {
        logger.error(`Error reloading token cache: ${(error as Error).message}`);
        if (storage.failClosed) {
          throw error;
        }
      }
    },
    afterCacheAccess: async (context: TokenCacheContext) => {
      if (!context.cacheHasChanged) {
        return;
      }
      try {
        await storage.save('token-cache', wrapCache(context.tokenCache.serialize()));
      } catch (error) {
        logger.error(`Error saving token cache: ${(error as Error).message}`);
        if (storage.failClosed) {
          throw error;
        }
      }
    },
  };
}

interface ScopeHierarchy {
  [key: string]: string[];
}

const SCOPE_HIERARCHY: ScopeHierarchy = {
  'Mail.ReadWrite': ['Mail.Read'],
  'Calendars.ReadWrite': ['Calendars.Read'],
  'Files.ReadWrite': ['Files.Read'],
  'Tasks.ReadWrite': ['Tasks.Read'],
  'Contacts.ReadWrite': ['Contacts.Read'],
};

interface AllowedScopeOptions {
  orgMode?: boolean;
  enabledTools?: string;
  readOnly?: boolean;
  allowedScopes?: string;
  extraScopes?: string;
  blockedTools?: string;
}

interface DisabledToolScope {
  toolName: string;
  requiredScopes: string[];
  missingScopes: string[];
}

interface ScopeDiagnostics {
  permissions: string[];
  toolPermissions: string[];
  effectivePermissions: string[];
  allowedScopes?: string[];
  disabledTools: DisabledToolScope[];
  missingAllowedScopesForTools: string[];
  extraAllowedScopesNotUsedByTools: string[];
}

function parseAllowedScopes(value?: string): string[] | undefined {
  if (value === undefined) {
    return undefined;
  }

  return Array.from(new Set(value.trim().split(/\s+/).filter(Boolean)));
}

function getEndpointRequiredScopes(
  endpoint: Pick<EndpointConfig, 'scopes' | 'workScopes'> | undefined,
  includeWorkAccountScopes: boolean = false
): string[] {
  if (!endpoint) {
    return [];
  }

  const scopes = new Set<string>();
  getEndpointScopeGroups(endpoint, includeWorkAccountScopes).forEach((group) =>
    group.forEach((scope) => scopes.add(scope))
  );
  return Array.from(scopes);
}

/**
 * Normalizes a scopes/workScopes value into a list of alternative AND-groups.
 * A flat string[] becomes a single group; a nested string[][] is already groups.
 */
function toScopeGroups(value?: string[] | string[][]): string[][] {
  if (!value || value.length === 0) {
    return [];
  }
  return Array.isArray(value[0]) ? (value as string[][]) : [value as string[]];
}

/**
 * Returns the alternative scope groups for an endpoint. The endpoint is satisfied if ALL
 * scopes in ANY single group are held. scopes and workScopes are mutually exclusive per
 * endpoints.json validation, so in practice only one side contributes groups.
 */
function getEndpointScopeGroups(
  endpoint: Pick<EndpointConfig, 'scopes' | 'workScopes'> | undefined,
  includeWorkAccountScopes: boolean = false
): string[][] {
  if (!endpoint) {
    return [];
  }
  const groups = [...toScopeGroups(endpoint.scopes)];
  if (includeWorkAccountScopes) {
    groups.push(...toScopeGroups(endpoint.workScopes));
  }
  return groups;
}

/**
 * The scopes to request at login for an endpoint: the primary (first) group only.
 * Microsoft's guidance is to consent to least-privileged scopes and add higher-privileged
 * ones on demand, so for OR-group endpoints we request the first group and leave the rest
 * to --extra-scopes. Flat (single-group) endpoints are unaffected.
 */
function getEndpointLoginScopes(
  endpoint: Pick<EndpointConfig, 'scopes' | 'workScopes'> | undefined,
  includeWorkAccountScopes: boolean = false
): string[] {
  const groups = getEndpointScopeGroups(endpoint, includeWorkAccountScopes);
  return groups.length > 0 ? groups[0] : [];
}

/**
 * Gate check for OR-group endpoints. Returns [] (allowed) if any group is fully covered by
 * allowedScopes; otherwise the missing scopes of the closest group (fewest missing), for
 * diagnostics. With a single group this matches getMissingAllowedScopes.
 */
function getMissingAllowedScopesForGroups(
  scopeGroups: string[][],
  allowedScopes?: string[]
): string[] {
  if (allowedScopes === undefined || scopeGroups.length === 0) {
    return [];
  }
  const coveredAllowedScopes = new Set(collapseScopeHierarchy(allowedScopes));
  let closest: string[] | undefined;
  for (const group of scopeGroups) {
    const missing = group.filter((scope) => !coveredAllowedScopes.has(scope));
    if (missing.length === 0) {
      return [];
    }
    if (!closest || missing.length < closest.length) {
      closest = missing;
    }
  }
  return closest ?? [];
}

/**
 * The scopes actually requested at login for an endpoint, honoring an optional allowlist.
 *
 * Without an allowlist this is the primary (first) group, per least-privilege (matches
 * getEndpointLoginScopes). With an allowlist it is the first group fully covered by the
 * allowlist, so an OR-group endpoint enabled via a non-primary alternative requests that
 * alternative's scopes and never scopes outside the allowlist. Returns [] when no group is
 * satisfied - the same allowlist disables the endpoint in that case (see
 * getMissingAllowedScopesForGroups), so it contributes no scopes.
 */
function getEndpointEffectiveLoginScopes(
  scopeGroups: string[][],
  allowedScopes?: string[]
): string[] {
  if (scopeGroups.length === 0) {
    return [];
  }
  if (allowedScopes === undefined) {
    return scopeGroups[0];
  }
  const coveredAllowedScopes = new Set(collapseScopeHierarchy(allowedScopes));
  const satisfied = scopeGroups.find((group) =>
    group.every((scope) => coveredAllowedScopes.has(scope))
  );
  return satisfied ?? [];
}

function collapseRedundantScopes(scopes: string[]): string[] {
  const scopesSet = new Set(scopes);

  // Scope hierarchy: if we have BOTH a higher scope (ReadWrite) AND lower scopes (Read),
  // keep only the higher scope since it includes the permissions of the lower scopes.
  // Do NOT upgrade Read to ReadWrite if we only have Read scopes.
  Object.entries(SCOPE_HIERARCHY).forEach(([higherScope, lowerScopes]) => {
    if (scopesSet.has(higherScope) && lowerScopes.every((scope) => scopesSet.has(scope))) {
      lowerScopes.forEach((scope) => scopesSet.delete(scope));
    }
  });

  return Array.from(scopesSet);
}

/**
 * Compile a --blocked-tools pattern for scope derivation.
 *
 * Deliberately non-fatal, unlike the registration-time guardrail: failing closed here
 * would request no scopes at all and break sign-in. By the time scopes are built the
 * pattern has already been validated at startup (parseArgs, and
 * compileBlockedToolsRegex during registration), so an unparseable value at this point
 * means a caller passed one directly.
 */
function compileBlockedForScopes(pattern?: string): RegExp | undefined {
  if (!pattern) return undefined;
  try {
    return new RegExp(pattern, 'i');
  } catch {
    logger.error(
      `Invalid blocked-tools regex pattern for scope derivation: ${pattern}. Deriving scopes without it.`
    );
    return undefined;
  }
}

function buildScopesFromEndpoints(
  includeWorkAccountScopes: boolean = false,
  enabledToolsPattern?: string,
  readOnly: boolean = false,
  blockedToolsPattern?: string
): string[] {
  const scopesSet = new Set<string>();
  const blockedToolsRegex = compileBlockedForScopes(blockedToolsPattern);

  // Create regex for tool filtering if pattern is provided
  let enabledToolsRegex: RegExp | undefined;
  if (enabledToolsPattern) {
    try {
      enabledToolsRegex = new RegExp(enabledToolsPattern, 'i');
      logger.info(`Building scopes with tool filter pattern: ${enabledToolsPattern}`);
    } catch {
      logger.error(
        `Invalid tool filter regex pattern: ${enabledToolsPattern}. Building scopes without filter.`
      );
    }
  }

  endpoints.default.forEach((endpoint) => {
    // Skip write operations in read-only mode
    if (readOnly && endpoint.method.toUpperCase() !== 'GET') {
      if (!(endpoint.method.toUpperCase() === 'POST' && endpoint.readOnly)) {
        return;
      }
    }

    // A blocked tool is unreachable, so requesting its scopes would grant the token a
    // capability the operator prohibited (#24).
    if (blockedToolsRegex && blockedToolsRegex.test(endpoint.toolName)) {
      return;
    }

    // Skip endpoints that don't match the tool filter
    if (enabledToolsRegex && !enabledToolsRegex.test(endpoint.toolName)) {
      return;
    }

    // Skip endpoints that only have workScopes if not in work mode
    if (!includeWorkAccountScopes && !endpoint.scopes && endpoint.workScopes) {
      return;
    }

    getEndpointLoginScopes(endpoint, includeWorkAccountScopes).forEach((scope) =>
      scopesSet.add(scope)
    );
  });

  const scopes = collapseRedundantScopes(Array.from(scopesSet));
  if (enabledToolsPattern) {
    logger.info(`Built ${scopes.length} scopes for filtered tools: ${scopes.join(', ')}`);
  }

  return scopes;
}

function lowerScopesFor(scope: string): string[] {
  const lowerScopes = new Set(SCOPE_HIERARCHY[scope] ?? []);

  if (scope.endsWith('.ReadWrite.All')) {
    const readAllScope = scope.replace(/\.ReadWrite\.All$/, '.Read.All');
    const readWriteScope = scope.replace(/\.ReadWrite\.All$/, '.ReadWrite');
    const readScope = scope.replace(/\.ReadWrite\.All$/, '.Read');
    lowerScopes.add(readAllScope);
    lowerScopes.add(readWriteScope);
    lowerScopes.add(readScope);
  } else if (scope.endsWith('.ReadWrite.Shared')) {
    lowerScopes.add(scope.replace(/\.ReadWrite\.Shared$/, '.Read.Shared'));
  } else if (scope.endsWith('.ReadWrite')) {
    lowerScopes.add(scope.replace(/\.ReadWrite$/, '.Read'));
  } else if (scope.endsWith('.Read.All')) {
    lowerScopes.add(scope.replace(/\.Read\.All$/, '.Read'));
  }

  return Array.from(lowerScopes);
}

function addImpliedScopes(scope: string, scopesSet: Set<string>): void {
  for (const lowerScope of lowerScopesFor(scope)) {
    if (!scopesSet.has(lowerScope)) {
      scopesSet.add(lowerScope);
      addImpliedScopes(lowerScope, scopesSet);
    }
  }
}

function collapseScopeHierarchy(scopes: string[]): string[] {
  const scopesSet = new Set(scopes);
  for (const scope of scopes) {
    addImpliedScopes(scope, scopesSet);
  }
  return Array.from(scopesSet);
}

function getMissingAllowedScopes(requiredScopes: string[], allowedScopes?: string[]): string[] {
  if (allowedScopes === undefined) {
    return [];
  }

  const coveredAllowedScopes = new Set(collapseScopeHierarchy(allowedScopes));
  return requiredScopes.filter((scope) => !coveredAllowedScopes.has(scope));
}

function isScopeUsedByTools(allowedScope: string, toolScopes: string[]): boolean {
  const coveredByAllowedScope = new Set(collapseScopeHierarchy([allowedScope]));
  return toolScopes.some((scope) => coveredByAllowedScope.has(scope));
}

function endpointMatchesNormalToolSurface(
  endpoint: EndpointConfig,
  includeWorkAccountScopes: boolean,
  enabledToolsRegex?: RegExp,
  readOnly: boolean = false
): boolean {
  if (readOnly && endpoint.method.toUpperCase() !== 'GET') {
    if (!(endpoint.method.toUpperCase() === 'POST' && endpoint.readOnly)) {
      return false;
    }
  }

  if (enabledToolsRegex && !enabledToolsRegex.test(endpoint.toolName)) {
    return false;
  }

  if (!includeWorkAccountScopes && !endpoint.scopes && endpoint.workScopes) {
    return false;
  }

  return true;
}

function buildAllowedScopeDiagnostics(options: AllowedScopeOptions = {}): ScopeDiagnostics {
  const allowedScopes = parseAllowedScopes(options.allowedScopes);
  let enabledToolsRegex: RegExp | undefined;
  if (options.enabledTools) {
    try {
      enabledToolsRegex = new RegExp(options.enabledTools, 'i');
    } catch {
      logger.error(
        `Invalid tool filter regex pattern: ${options.enabledTools}. Building diagnostics without filter.`
      );
    }
  }

  const blockedToolsRegex = compileBlockedForScopes(options.blockedTools);

  const normalToolScopes = new Set<string>();
  const effectiveToolScopes = new Set<string>();
  // Union of every group's scopes for passing tools, used only to judge whether an
  // allowed scope is used by some tool (an OR-group's non-primary scopes still count).
  const effectiveToolScopesAllGroups = new Set<string>();
  const disabledTools: DisabledToolScope[] = [];

  for (const endpoint of endpoints.default) {
    // Same reasoning as buildScopesFromEndpoints: a blocked tool must not pull its
    // scopes onto the token.
    if (blockedToolsRegex && blockedToolsRegex.test(endpoint.toolName)) continue;
    if (
      !endpointMatchesNormalToolSurface(
        endpoint,
        Boolean(options.orgMode),
        enabledToolsRegex,
        Boolean(options.readOnly)
      )
    ) {
      continue;
    }

    const scopeGroups = getEndpointScopeGroups(endpoint, Boolean(options.orgMode));
    const loginScopes = getEndpointLoginScopes(endpoint, Boolean(options.orgMode));
    const allScopes = getEndpointRequiredScopes(endpoint, Boolean(options.orgMode));
    loginScopes.forEach((scope) => normalToolScopes.add(scope));

    const missingScopes = getMissingAllowedScopesForGroups(scopeGroups, allowedScopes);
    if (missingScopes.length > 0) {
      disabledTools.push({
        toolName: endpoint.toolName,
        requiredScopes: allScopes.sort((a, b) => a.localeCompare(b)),
        missingScopes: missingScopes.sort((a, b) => a.localeCompare(b)),
      });
      continue;
    }

    // Request the group that actually satisfied the allowlist, not unconditionally the
    // primary group. For an OR-group endpoint enabled via a non-primary alternative, requesting
    // the primary group would both leak scopes outside the allowlist and omit the scope the
    // tool was enabled for. Without an allowlist this is the primary group, unchanged.
    getEndpointEffectiveLoginScopes(scopeGroups, allowedScopes).forEach((scope) =>
      effectiveToolScopes.add(scope)
    );
    allScopes.forEach((scope) => effectiveToolScopesAllGroups.add(scope));
  }

  const toolPermissions = collapseRedundantScopes(Array.from(normalToolScopes)).sort((a, b) =>
    a.localeCompare(b)
  );
  const effectivePermissions = collapseRedundantScopes(Array.from(effectiveToolScopes)).sort(
    (a, b) => a.localeCompare(b)
  );
  const sortedAllowedScopes = allowedScopes
    ? [...allowedScopes].sort((a, b) => a.localeCompare(b))
    : undefined;
  const missingAllowedScopesForTools = Array.from(
    new Set(disabledTools.flatMap((tool) => tool.missingScopes))
  ).sort((a, b) => a.localeCompare(b));
  const allEffectiveToolScopes = Array.from(effectiveToolScopesAllGroups);
  const extraAllowedScopesNotUsedByTools =
    sortedAllowedScopes?.filter((scope) => !isScopeUsedByTools(scope, allEffectiveToolScopes)) ??
    [];

  return {
    permissions: effectivePermissions,
    toolPermissions,
    effectivePermissions,
    ...(sortedAllowedScopes ? { allowedScopes: sortedAllowedScopes } : {}),
    disabledTools,
    missingAllowedScopesForTools,
    extraAllowedScopesNotUsedByTools,
  };
}

function resolveAuthScopes(options: AllowedScopeOptions = {}): string[] {
  const toolScopes = buildAllowedScopeDiagnostics(options).effectivePermissions;
  // Extra scopes are appended verbatim to the token request, independent of the tool
  // surface and the allowed-scopes filter. They let a user on their own app registration
  // request scopes no bundled tool needs (e.g. CopilotPackages.ReadWrite.All) and then
  // drive the matching endpoints via graph-batch.
  const extraScopes = parseAllowedScopes(options.extraScopes);
  if (!extraScopes || extraScopes.length === 0) {
    return toolScopes;
  }
  return Array.from(new Set([...toolScopes, ...extraScopes]));
}

/**
 * Reduce a scope to the permission name Entra will actually read.
 *
 * A scope may name its resource in several ways: `Mail.Send`,
 * `https://graph.microsoft.com/Mail.Send`, the same on a sovereign-cloud host, or the
 * resource app-id GUID form `00000003-0000-0000-c000-000000000000/Mail.Send`. In every
 * case the permission is the segment after the LAST '/', and no Graph permission name
 * contains a slash. Taking that segment therefore handles the resource URI, the GUID
 * form, a doubled prefix and a stray double slash uniformly.
 *
 * Stripping a single `scheme://host/` prefix instead, which is what this did first, let
 * the GUID form and `https://graph.microsoft.com//Mail.Send` straight through.
 *
 * Case is preserved here; canonicalScope lowercases for comparison.
 */
function stripScopeResourceUri(scope: string): string {
  // Trailing slashes come off first. Without this, `Mail.Send/` puts the last slash at
  // the end and reduces to an empty permission name, which matches nothing prohibited
  // and sails through.
  const trimmed = scope.trim().replace(/\/+$/, '');
  const lastSlash = trimmed.lastIndexOf('/');
  return lastSlash === -1 ? trimmed : trimmed.slice(lastSlash + 1);
}

/** Comparison form for a scope: no resource, no case, no surrounding space. */
function canonicalScope(scope: string): string {
  return stripScopeResourceUri(scope).toLowerCase();
}

/**
 * Expand a canonical (lowercased) scope to everything it implies.
 *
 * `lowerScopesFor` matches `.ReadWrite`/`.Read.All` suffixes case-sensitively and reads
 * SCOPE_HIERARCHY by exact key, so it silently expands to nothing on a lowercased
 * string. Re-casing via the endpoint catalogue only helps for scopes the catalogue
 * carries: `Files.ReadWrite.All` is a real delegated permission that is not in
 * endpoints.json, so `files.readwrite.all` expanded to nothing, matched nothing
 * prohibited, and was granted while the correctly-cased spelling was refused.
 *
 * These rules mirror lowerScopesFor, applied to canonical strings so casing cannot
 * defeat them.
 */
function canonicalImpliedScopes(canonical: string): Set<string> {
  const seen = new Set<string>([canonical]);
  const pending = [canonical];

  while (pending.length > 0) {
    const scope = pending.pop()!;
    const lower: string[] = [];

    if (scope.endsWith('.readwrite.all')) {
      const stem = scope.slice(0, -'.readwrite.all'.length);
      lower.push(`${stem}.read.all`, `${stem}.readwrite`, `${stem}.read`);
    } else if (scope.endsWith('.readwrite.shared')) {
      lower.push(`${scope.slice(0, -'.readwrite.shared'.length)}.read.shared`);
    } else if (scope.endsWith('.readwrite')) {
      lower.push(`${scope.slice(0, -'.readwrite'.length)}.read`);
    } else if (scope.endsWith('.read.all')) {
      lower.push(`${scope.slice(0, -'.read.all'.length)}.read`);
    }

    for (const implied of lower) {
      if (!seen.has(implied)) {
        seen.add(implied);
        pending.push(implied);
      }
    }
  }

  return seen;
}

/**
 * Canonical scopes that exist only because a tool the operator blocked would have
 * needed them. The difference between deriving with and without the blocklist IS the
 * prohibited set, so this stays correct as the endpoint catalogue changes.
 *
 * Two things this got wrong, both of which quietly emptied the difference:
 *
 * `--read-only` was applied to BOTH sides. Read-only already strips write scopes from
 * the unblocked side, so the delta cancelled and a tightening flag made the filter
 * LOOSER: Mail.Send was refused under `--blocked-tools` alone and granted once
 * `--read-only` was added. Both sides are now derived with read-only off, so the
 * blocklist delta is whatever the blocklist itself removes, independent of it.
 *
 * Only the permitted side was hierarchy-expanded. buildScopesFromEndpoints emits the
 * collapsed higher form (`Mail.ReadWrite`) and never the bare lower one, so `Mail.Read`
 * was never in the unblocked side and never came out prohibited: with the entire
 * catalogue blocked a client could still be granted Mail.Read, Files.Read and
 * Calendars.Read. Both sides are expanded now.
 *
 * `--enabled-tools` stays on both sides deliberately. It selects which tools to load
 * rather than expressing a prohibition, and upstream's allowed-scopes test pins that a
 * client may request a scope outside the loaded preset.
 */
function blockedToolScopes(options: AllowedScopeOptions): Set<string> {
  if (!options.blockedTools) {
    return new Set();
  }

  const READ_ONLY_OFF = false;

  const permitted = new Set(
    collapseScopeHierarchy(
      buildScopesFromEndpoints(
        options.orgMode,
        options.enabledTools,
        READ_ONLY_OFF,
        options.blockedTools
      )
    ).map(canonicalScope)
  );

  return new Set(
    collapseScopeHierarchy(
      buildScopesFromEndpoints(options.orgMode, options.enabledTools, READ_ONLY_OFF)
    )
      .map(canonicalScope)
      .filter((scope) => !permitted.has(scope))
  );
}

/**
 * Resolve the scopes the /authorize redirect may request, given what the client asked
 * for.
 *
 * A client-supplied `scope` is still honoured, which is deliberate upstream behaviour:
 * an operator who has constrained nothing lets the client choose. What a client may NOT
 * do is name a scope that exists only to serve a tool the operator blocked. Forwarding
 * the list verbatim let a caller ask for Mail.Send while every send tool was blocked,
 * and the bearer token the client receives is usable directly against Graph, outside
 * the tool surface the blocklist guards (#24).
 *
 * So the filter is narrow on purpose: it subtracts exactly the blocked-tool scopes and
 * leaves every other client request alone.
 */
function resolveAuthorizeScopes(
  options: AllowedScopeOptions = {},
  clientScope?: string | null
): string[] {
  // An explicit --allowed-scopes already pins the surface, and that path never consulted
  // the client's request. Unchanged.
  if (parseAllowedScopes(options.allowedScopes) !== undefined) {
    return resolveAuthScopes(options);
  }

  const derived = buildScopesFromEndpoints(
    options.orgMode,
    options.enabledTools,
    options.readOnly,
    // Do not request scopes for tools the operator blocked (#24).
    options.blockedTools
  );

  const requested = parseAllowedScopes(clientScope ?? undefined);
  if (!requested || requested.length === 0) {
    return derived;
  }

  // Compare canonically. The client controls the spelling, and Entra treats
  // `mail.send`, `Mail.Send` and `https://graph.microsoft.com/Mail.Send` as the same
  // permission, so an exact match against bare catalogue names is trivially bypassable.
  const prohibited = blockedToolScopes(options);
  if (prohibited.size === 0) {
    // Nothing is blocked, so there is nothing to subtract and upstream's behaviour of
    // honouring the client's request applies unchanged.
    return requested;
  }

  // Catalogue spelling still matters for SCOPE_HIERARCHY, which is keyed on exact
  // names and expresses relationships the suffix rules cannot derive. It is a
  // supplement to the canonical expansion below, never the only line of defence:
  // a scope absent from the catalogue must still be checked.
  const catalogueSpelling = new Map<string, string>(
    buildScopesFromEndpoints(options.orgMode, options.enabledTools, options.readOnly).map(
      (scope) => [canonicalScope(scope), scope]
    )
  );

  const granted = requested.filter((scope) => {
    const canonical = canonicalScope(scope);

    // `.default` is catalogue-independent: it asks for every statically consented app
    // permission, which necessarily includes the ones the operator blocked. There is no
    // narrower reading of it, so it cannot be honoured while a blocklist is in force.
    if (canonical === '.default') return false;

    // Reject when the scope IS prohibited or IMPLIES something prohibited, so a broader
    // scope cannot smuggle a blocked one in. Union of the case-insensitive suffix
    // expansion (works for any scope, catalogued or not) and the catalogue's own
    // hierarchy map (catches relationships the suffix rules do not encode).
    const implied = canonicalImpliedScopes(canonical);
    const catalogued = catalogueSpelling.get(canonical);
    if (catalogued) {
      for (const s of collapseScopeHierarchy([catalogued])) implied.add(canonicalScope(s));
    }

    for (const s of implied) {
      if (prohibited.has(s)) return false;
    }
    return true;
  });

  if (granted.length !== requested.length) {
    const dropped = requested.filter((scope) => !granted.includes(scope));
    logger.warn(`Ignoring client-requested scope(s) for blocked tools: ${dropped.join(', ')}`);
  }

  // Falling back keeps sign-in working when a client asks only for prohibited scopes,
  // rather than requesting an empty scope list.
  return granted.length > 0 ? granted : derived;
}

function buildScopeDiagnostics(
  toolScopes: string[],
  allowedScopesInput: string[]
): ScopeDiagnostics {
  const toolPermissions = [...toolScopes].sort((a, b) => a.localeCompare(b));
  const coveredAllowedScopes = new Set(collapseScopeHierarchy(allowedScopesInput));
  const missingAllowedScopesForTools = toolPermissions.filter(
    (scope) => !coveredAllowedScopes.has(scope)
  );

  return {
    permissions: toolPermissions.filter((scope) => coveredAllowedScopes.has(scope)),
    toolPermissions,
    effectivePermissions: toolPermissions.filter((scope) => coveredAllowedScopes.has(scope)),
    allowedScopes: [...allowedScopesInput].sort((a, b) => a.localeCompare(b)),
    disabledTools: [],
    missingAllowedScopesForTools,
    extraAllowedScopesNotUsedByTools: [...allowedScopesInput]
      .sort((a, b) => a.localeCompare(b))
      .filter((scope) => !isScopeUsedByTools(scope, toolPermissions)),
  };
}

interface LoginTestResult {
  success: boolean;
  message: string;
  userData?: {
    displayName: string;
    userPrincipalName: string;
  };
}

interface ExpectedAccountOptions {
  expectedUsername?: string;
  expectedHomeAccountId?: string;
}

interface AuthManagerCreateOptions {
  storage?: TokenCacheStorage;
}

/**
 * Summarises a silent-acquire failure for logging. MSAL throws AuthError subclasses
 * (e.g. InteractionRequiredAuthError) whose errorCode, subError and correlationId pin
 * the cause, such as invalid_grant from the token endpoint or interaction_required.
 * The log formatter only emits `message`, so the codes are folded into the string here.
 */
export function describeAuthError(error: unknown): string {
  if (error instanceof AuthError) {
    const suberror = error.subError ? ` / ${error.subError}` : '';
    return `${error.errorCode}${suberror} (correlationId: ${error.correlationId || 'none'}): ${error.errorMessage}`;
  }
  return (error as Error).message;
}

/** Home tenant id shared by all personal Microsoft accounts (MSA). */
const MSA_HOME_TENANT_ID = '9188040d-6c67-4c5b-b112-36a304b66dad';

/**
 * Builds a remediation hint when a personal Microsoft account's refresh token is
 * rejected on the default 'common' authority. As of June 2026 the token endpoint
 * returns invalid_grant for MSA refresh tokens issued via /common, while the
 * same login via /consumers refreshes fine - so the fix is a config change plus
 * one re-login, which a generic "token may have expired" message does not
 * convey. Returns null when the failure does not match that signature.
 */
export function consumersAuthorityHint(
  error: unknown,
  account: AccountInfo | null | undefined,
  authority: string | undefined
): string | null {
  if (
    error instanceof AuthError &&
    error.errorCode === 'invalid_grant' &&
    account?.tenantId === MSA_HOME_TENANT_ID &&
    (!authority || /\/common\/?$/i.test(authority))
  ) {
    return (
      `This looks like a known issue (June 2026) where Microsoft rejects refresh tokens ` +
      `issued to personal accounts via the default 'common' authority. If this server is ` +
      `used only with personal accounts, set MS365_MCP_TENANT_ID=consumers and re-login ` +
      `with: --login`
    );
  }
  return null;
}

class AuthManager {
  private config: Configuration;
  private scopes: string[];
  private msalApp: PublicClientApplication;
  private accessToken: string | null;
  private tokenExpiry: number | null;
  private oauthToken: string | null;
  private isOAuthMode: boolean;
  private selectedAccountId: string | null;
  private useInteractiveAuth: boolean;
  private expectedUsername: string | null;
  private expectedHomeAccountId: string | null;
  private storage: TokenCacheStorage;

  constructor(
    config: Configuration,
    scopes: string[] = [],
    expectedAccount?: ExpectedAccountOptions,
    storage?: TokenCacheStorage
  ) {
    logger.info(`And scopes are ${scopes.join(', ')}`, scopes);
    this.scopes = scopes;
    this.storage = storage ?? new DefaultTokenCacheStorage();
    // Register a cache plugin so MSAL reloads the newest persisted cache before every access
    // and persists rotations, keeping concurrent stdio processes coherent (issue #545).
    this.config = {
      ...config,
      cache: {
        ...config.cache,
        cachePlugin: buildDiskCoherencyCachePlugin(this.storage),
      },
    };
    this.msalApp = new PublicClientApplication(this.config);
    this.accessToken = null;
    this.tokenExpiry = null;
    this.selectedAccountId = null;
    this.useInteractiveAuth = false;
    this.expectedUsername = this.normalizeExpectedUsername(expectedAccount?.expectedUsername);
    this.expectedHomeAccountId = this.normalizeExpectedHomeAccountId(
      expectedAccount?.expectedHomeAccountId
    );

    const oauthTokenFromEnv = process.env.MS365_MCP_OAUTH_TOKEN;
    this.oauthToken = oauthTokenFromEnv ?? null;
    this.isOAuthMode = oauthTokenFromEnv != null;
  }

  /**
   * Creates an AuthManager instance with secrets loaded from the configured provider.
   * Uses Key Vault if MS365_MCP_KEYVAULT_URL is set, otherwise environment variables.
   */
  static async create(
    scopes: string[] = [],
    expectedAccount?: ExpectedAccountOptions,
    options: AuthManagerCreateOptions = {}
  ): Promise<AuthManager> {
    const secrets = await getSecrets();
    const config = createMsalConfig(secrets);
    const storage =
      options.storage ??
      (await createTokenCacheStorage({ allowCommandStorage: false, logProvider: true }));
    return new AuthManager(config, scopes, expectedAccount, storage);
  }

  async loadTokenCache(): Promise<void> {
    try {
      const cacheRaw = await this.storage.load('token-cache');
      if (cacheRaw) {
        this.msalApp.getTokenCache().deserialize(unwrapCache(cacheRaw).data);
      }

      // Load selected account
      await this.loadSelectedAccount();
    } catch (error) {
      logger.error(`Error loading token cache: ${(error as Error).message}`);
      if (this.storage.failClosed) {
        throw error;
      }
    }
  }

  private async loadSelectedAccount(): Promise<void> {
    try {
      const selectedAccountRaw = await this.storage.load('selected-account');
      if (selectedAccountRaw) {
        const parsed = JSON.parse(unwrapCache(selectedAccountRaw).data);
        this.selectedAccountId = parsed.accountId;
        logger.info(`Loaded selected account: ${this.selectedAccountId}`);
      }
    } catch (error) {
      logger.error(`Error loading selected account: ${(error as Error).message}`);
      if (this.storage.failClosed) {
        throw error;
      }
    }
  }

  private async saveSelectedAccount(): Promise<void> {
    try {
      const stamped = wrapCache(JSON.stringify({ accountId: this.selectedAccountId }));
      await this.storage.save('selected-account', stamped);
    } catch (error) {
      logger.error(`Error saving selected account: ${(error as Error).message}`);
      if (this.storage.failClosed) {
        throw error;
      }
    }
  }

  private normalizeExpectedUsername(value?: string): string | null {
    if (value === undefined) {
      return null;
    }
    const trimmed = value.trim();
    if (trimmed === '') {
      throw new Error('Expected Microsoft account username was provided but is empty.');
    }
    return trimmed.toLowerCase();
  }

  private normalizeExpectedHomeAccountId(value?: string): string | null {
    if (value === undefined) {
      return null;
    }
    const trimmed = value.trim();
    if (trimmed === '') {
      throw new Error('Expected Microsoft account homeAccountId was provided but is empty.');
    }
    return trimmed;
  }

  hasExpectedAccount(): boolean {
    return this.expectedUsername !== null || this.expectedHomeAccountId !== null;
  }

  private expectedAccountLabel(): string {
    const parts: string[] = [];
    if (this.expectedUsername) {
      parts.push(`username ${this.expectedUsername}`);
    }
    if (this.expectedHomeAccountId) {
      parts.push(`homeAccountId ${this.expectedHomeAccountId}`);
    }
    return parts.join(' and ');
  }

  private describeAccount(account: AccountInfo | null | undefined): string {
    return account?.username || account?.name || 'unknown';
  }

  private describeCachedAccounts(accounts: AccountInfo[]): string {
    if (accounts.length === 0) {
      return 'none';
    }
    return accounts.map((account) => this.describeAccount(account)).join(', ');
  }

  private accountMatchesExpected(account: AccountInfo | null | undefined): boolean {
    if (!this.hasExpectedAccount() || !account) {
      return !this.hasExpectedAccount();
    }
    if (this.expectedUsername && account.username?.toLowerCase() !== this.expectedUsername) {
      return false;
    }
    if (this.expectedHomeAccountId && account.homeAccountId !== this.expectedHomeAccountId) {
      return false;
    }
    return true;
  }

  private buildExpectedAccountMissingError(accounts: AccountInfo[]): Error {
    return new Error(
      `Expected Microsoft account '${this.expectedAccountLabel()}' not found in token cache. ` +
        `Cached accounts: ${this.describeCachedAccounts(accounts)}. ` +
        'Run --login after configuring the expected account, or use --select-account to recover.'
    );
  }

  private resolveExpectedAccountFromAccounts(accounts: AccountInfo[]): AccountInfo {
    if (!this.hasExpectedAccount()) {
      throw new Error('No expected Microsoft account is configured.');
    }

    const usernameMatch = this.expectedUsername
      ? accounts.find((account) => account.username?.toLowerCase() === this.expectedUsername)
      : undefined;
    const homeAccountIdMatch = this.expectedHomeAccountId
      ? accounts.find((account) => account.homeAccountId === this.expectedHomeAccountId)
      : undefined;

    if (this.expectedUsername && this.expectedHomeAccountId) {
      if (!usernameMatch || !homeAccountIdMatch) {
        throw this.buildExpectedAccountMissingError(accounts);
      }
      if (usernameMatch.homeAccountId !== homeAccountIdMatch.homeAccountId) {
        throw new Error(
          `Expected Microsoft account pins conflict: username ${this.expectedUsername} matched ` +
            `${this.describeAccount(usernameMatch)}, but homeAccountId ${this.expectedHomeAccountId} matched ` +
            `${this.describeAccount(homeAccountIdMatch)}.`
        );
      }
      return usernameMatch;
    }

    const expectedAccount = usernameMatch ?? homeAccountIdMatch;
    if (!expectedAccount) {
      throw this.buildExpectedAccountMissingError(accounts);
    }
    return expectedAccount;
  }

  async assertExpectedAccountAvailable(): Promise<void> {
    if (!this.hasExpectedAccount()) {
      return;
    }
    const accounts = await this.msalApp.getTokenCache().getAllAccounts();
    this.resolveExpectedAccountFromAccounts(accounts);
  }

  private async rejectUnexpectedLoginAccount(
    account: AccountInfo | null | undefined
  ): Promise<void> {
    if (!this.hasExpectedAccount()) {
      return;
    }

    if (this.accountMatchesExpected(account)) {
      return;
    }

    this.accessToken = null;
    this.tokenExpiry = null;

    if (account) {
      // The cache plugin (afterCacheAccess) persists during the acquire call, so a mismatched
      // account's tokens are already on disk by the time we get here. removeAccount triggers the
      // plugin again to persist the removal - but if it fails we must NOT claim the login was not
      // persisted, because the rejected account's tokens remain in the shared cache. Surface that
      // loudly and actionably instead of swallowing it (issue #545 hardening).
      try {
        await this.msalApp.getTokenCache().removeAccount(account);
      } catch (error) {
        logger.error(`Failed to remove unexpected account from cache: ${(error as Error).message}`);
        throw new Error(
          `Authenticated Microsoft account '${this.describeAccount(account)}' does not match expected ` +
            `Microsoft account '${this.expectedAccountLabel()}', and it could not be removed from the ` +
            `token cache (${(error as Error).message}). Its tokens may remain persisted - run --logout ` +
            `to clear the cache, then re-login.`
        );
      }
      throw new Error(
        `Authenticated Microsoft account '${this.describeAccount(account)}' does not match expected Microsoft account '${this.expectedAccountLabel()}'. Login was not persisted.`
      );
    }

    throw new Error(
      `Microsoft login did not return an account. Expected Microsoft account '${this.expectedAccountLabel()}'. Login was not persisted.`
    );
  }

  async setOAuthToken(token: string): Promise<void> {
    this.oauthToken = token;
    this.isOAuthMode = true;
  }

  async getToken(forceRefresh = false): Promise<string | null> {
    if (this.isOAuthMode && this.oauthToken) {
      return this.oauthToken;
    }

    if (this.accessToken && this.tokenExpiry && this.tokenExpiry > Date.now() && !forceRefresh) {
      return this.accessToken;
    }

    const currentAccount = await this.getCurrentAccount();

    if (currentAccount) {
      const silentRequest = {
        account: currentAccount,
        scopes: this.scopes,
      };

      try {
        const response = await this.msalApp.acquireTokenSilent(silentRequest);
        this.accessToken = response.accessToken;
        this.tokenExpiry = response.expiresOn ? new Date(response.expiresOn).getTime() : null;
        // Persistence is owned by the cache plugin (afterCacheAccess): when MSAL rotates the
        // refresh token it reloads-then-saves under the coherency protocol. A manual save here
        // would serialize the in-memory cache without the reload-before-write step and could
        // clobber a newer rotation a sibling process wrote in the meantime (issue #545).
        return this.accessToken;
      } catch (error) {
        const hint = consumersAuthorityHint(error, currentAccount, this.config.auth.authority);
        logger.error(
          `Silent token acquisition failed: ${describeAuthError(error)}${hint ? ` ${hint}` : ''}`
        );
        throw new Error(
          hint ? `Silent token acquisition failed. ${hint}` : 'Silent token acquisition failed'
        );
      }
    }

    throw new Error('No valid token found');
  }

  async getCurrentAccount(): Promise<AccountInfo | null> {
    const accounts = await this.msalApp.getTokenCache().getAllAccounts();

    if (this.hasExpectedAccount()) {
      return this.resolveExpectedAccountFromAccounts(accounts);
    }

    if (accounts.length === 0) {
      return null;
    }

    // If a specific account is selected, find it
    if (this.selectedAccountId) {
      const selectedAccount = accounts.find(
        (account: AccountInfo) => account.homeAccountId === this.selectedAccountId
      );
      if (selectedAccount) {
        return selectedAccount;
      }
      logger.warn(
        `Selected account ${this.selectedAccountId} not found, falling back to first account`
      );
    }

    // Fall back to first account (backward compatibility)
    return accounts[0];
  }

  async acquireTokenByDeviceCode(hack?: (message: string) => void): Promise<string | null> {
    const deviceCodeRequest = {
      scopes: this.scopes,
      deviceCodeCallback: (response: { message: string }) => {
        const text = ['\n', response.message, '\n'].join('');
        if (hack) {
          hack(text + 'After login run the "verify login" command');
        } else {
          console.log(text);
        }
        logger.info('Device code login initiated');
      },
    };

    try {
      logger.info('Requesting device code...');
      logger.info(`Requesting scopes: ${this.scopes.join(', ')}`);
      const response = await this.msalApp.acquireTokenByDeviceCode(deviceCodeRequest);
      logger.info(`Granted scopes: ${response?.scopes?.join(', ') || 'none'}`);
      logger.info('Device code login successful');
      this.accessToken = response?.accessToken || null;
      this.tokenExpiry = response?.expiresOn ? new Date(response.expiresOn).getTime() : null;
      await this.rejectUnexpectedLoginAccount(response?.account);

      // Set the newly authenticated account as selected if no account is currently selected
      if (!this.selectedAccountId && response?.account) {
        this.selectedAccountId = response.account.homeAccountId;
        await this.saveSelectedAccount();
        logger.info(`Auto-selected new account: ${response.account.username}`);
      }

      // MSAL persisted the new tokens via the cache plugin (afterCacheAccess) during the
      // acquire call; no manual save needed (issue #545).
      return this.accessToken;
    } catch (error) {
      logger.error(`Error in device code flow: ${(error as Error).message}`);
      throw error;
    }
  }

  setUseInteractiveAuth(value: boolean): void {
    this.useInteractiveAuth = value;
  }

  getUseInteractiveAuth(): boolean {
    return this.useInteractiveAuth;
  }

  async acquireTokenInteractive(hack?: (message: string) => void): Promise<string | null> {
    const open = (await import('open')).default;

    const interactiveRequest = {
      scopes: this.scopes,
      openBrowser: async (url: string) => {
        const message = 'Opening browser for Microsoft sign-in...';
        if (hack) {
          hack(message);
        }
        logger.info(message);
        await open(url);
      },
      successTemplate:
        '<h1>Authentication successful!</h1><p>You can close this window and return to your application.</p>',
      errorTemplate: '<h1>Authentication failed</h1><p>Something went wrong. Please try again.</p>',
    };

    try {
      logger.info('Requesting interactive browser login...');
      logger.info(`Requesting scopes: ${this.scopes.join(', ')}`);
      const response = await this.msalApp.acquireTokenInteractive(interactiveRequest);
      logger.info(`Granted scopes: ${response?.scopes?.join(', ') || 'none'}`);
      logger.info('Interactive browser login successful');
      this.accessToken = response?.accessToken || null;
      this.tokenExpiry = response?.expiresOn ? new Date(response.expiresOn).getTime() : null;
      await this.rejectUnexpectedLoginAccount(response?.account);

      // Set the newly authenticated account as selected if no account is currently selected
      if (!this.selectedAccountId && response?.account) {
        this.selectedAccountId = response.account.homeAccountId;
        await this.saveSelectedAccount();
        logger.info(`Auto-selected new account: ${response.account.username}`);
      }

      // MSAL persisted the new tokens via the cache plugin (afterCacheAccess) during the
      // acquire call; no manual save needed (issue #545).
      return this.accessToken;
    } catch (error) {
      logger.error(`Error in interactive browser flow: ${(error as Error).message}`);
      throw error;
    }
  }

  async testLogin(): Promise<LoginTestResult> {
    try {
      logger.info('Testing login...');
      const token = await this.getToken();

      if (!token) {
        logger.error('Login test failed - no token received');
        return {
          success: false,
          message: 'Login failed - no token received',
        };
      }

      logger.info('Token retrieved successfully, testing Graph API access...');

      try {
        const secrets = await getSecrets();
        const cloudEndpoints = getCloudEndpoints(secrets.cloudType);
        const response = await fetch(`${cloudEndpoints.graphApi}/v1.0/me`, {
          headers: {
            Authorization: `Bearer ${token}`,
          },
        });

        if (response.ok) {
          const userData = await response.json();
          logger.info('Graph API user data fetch successful');
          return {
            success: true,
            message: 'Login successful',
            userData: {
              displayName: userData.displayName,
              userPrincipalName: userData.userPrincipalName,
            },
          };
        } else {
          const errorText = await response.text();
          logger.error(`Graph API user data fetch failed: ${response.status} - ${errorText}`);
          return {
            success: false,
            message: `Login successful but Graph API access failed: ${response.status}`,
          };
        }
      } catch (graphError) {
        logger.error(`Error fetching user data: ${(graphError as Error).message}`);
        return {
          success: false,
          message: `Login successful but Graph API access failed: ${(graphError as Error).message}`,
        };
      }
    } catch (error) {
      logger.error(`Login test failed: ${(error as Error).message}`);
      return {
        success: false,
        message: `Login failed: ${(error as Error).message}`,
      };
    }
  }

  async logout(): Promise<boolean> {
    try {
      const accounts = await this.msalApp.getTokenCache().getAllAccounts();
      for (const account of accounts) {
        await this.msalApp.getTokenCache().removeAccount(account);
      }
      this.accessToken = null;
      this.tokenExpiry = null;
      this.selectedAccountId = null;

      await this.storage.delete('token-cache');
      await this.storage.delete('selected-account');

      return true;
    } catch (error) {
      logger.error(`Error during logout: ${(error as Error).message}`);
      throw error;
    }
  }

  // Multi-account support methods
  async listAccounts(): Promise<AccountInfo[]> {
    return await this.msalApp.getTokenCache().getAllAccounts();
  }

  async selectAccount(identifier: string): Promise<boolean> {
    const account = await this.resolveAccount(identifier);
    if (this.hasExpectedAccount() && !this.accountMatchesExpected(account)) {
      throw new Error(
        `Account '${identifier}' does not match expected Microsoft account '${this.expectedAccountLabel()}'.`
      );
    }

    this.selectedAccountId = account.homeAccountId;
    await this.saveSelectedAccount();

    // Clear cached tokens to force refresh with new account
    this.accessToken = null;
    this.tokenExpiry = null;

    logger.info(`Selected account: ${account.username} (${account.homeAccountId})`);
    return true;
  }

  async removeAccount(identifier: string): Promise<boolean> {
    const account = await this.resolveAccount(identifier);

    try {
      await this.msalApp.getTokenCache().removeAccount(account);

      // If this was the selected account, clear the selection
      if (this.selectedAccountId === account.homeAccountId) {
        this.selectedAccountId = null;
        await this.saveSelectedAccount();
        this.accessToken = null;
        this.tokenExpiry = null;
      }

      logger.info(`Removed account: ${account.username} (${account.homeAccountId})`);
      return true;
    } catch (error) {
      logger.error(`Failed to remove account ${identifier}: ${(error as Error).message}`);
      return false;
    }
  }

  getSelectedAccountId(): string | null {
    return this.selectedAccountId;
  }

  /**
   * Returns true if auth is in OAuth/HTTP mode (token supplied via env or setOAuthToken).
   * In this mode, account resolution should be skipped — the request context drives token selection.
   */
  isOAuthModeEnabled(): boolean {
    return this.isOAuthMode;
  }

  /**
   * Resolves an account by identifier (email or homeAccountId).
   * Resolution: username match (case-insensitive) → homeAccountId match → throw.
   */
  async resolveAccount(identifier: string): Promise<AccountInfo> {
    const accounts = await this.msalApp.getTokenCache().getAllAccounts();

    if (accounts.length === 0) {
      throw new Error('No accounts found. Please login first.');
    }

    const lowerIdentifier = identifier.toLowerCase();

    // Try username (email) match first
    let account =
      accounts.find((a: AccountInfo) => a.username?.toLowerCase() === lowerIdentifier) ?? null;

    // Fall back to homeAccountId match
    if (!account) {
      account = accounts.find((a: AccountInfo) => a.homeAccountId === identifier) ?? null;
    }

    if (!account) {
      const availableAccounts = accounts
        .map((a: AccountInfo) => a.username || a.name || 'unknown')
        .join(', ');
      throw new Error(
        `Account '${identifier}' not found. Available accounts: ${availableAccounts}`
      );
    }

    return account;
  }

  /**
   * Returns true if the MSAL cache contains more than one account.
   * Used to decide whether to inject the `account` parameter into tool schemas.
   */
  async isMultiAccount(): Promise<boolean> {
    if (this.hasExpectedAccount()) {
      return false;
    }
    const accounts = await this.msalApp.getTokenCache().getAllAccounts();
    return accounts.length > 1;
  }

  /**
   * Acquires a token for a specific account identified by username (email) or homeAccountId,
   * WITHOUT changing the persisted selectedAccountId.
   *
   * Resolution order:
   *  1. Exact match on username (case-insensitive)
   *  2. Exact match on homeAccountId
   *  3. If identifier is empty/undefined AND only 1 account exists → auto-select
   *  4. If identifier is empty/undefined AND multiple accounts → use selectedAccountId or throw
   *
   * @returns The access token string.
   */
  async getTokenForAccount(identifier?: string): Promise<string> {
    if (this.isOAuthMode && this.oauthToken) {
      // Refuse instead of silently returning the bearer's identity (discussion #467):
      // in OAuth mode the token comes from the connecting client and cannot be
      // switched to a cached MSAL account.
      if (identifier) {
        throw new Error(
          `Cannot switch to account '${identifier}': the server is in OAuth mode and always uses ` +
            `the identity of the supplied bearer token. Account switching requires stdio mode ` +
            `(or HTTP with --trust-proxy-auth).`
        );
      }
      return this.oauthToken;
    }

    let targetAccount: AccountInfo | null = null;

    if (this.hasExpectedAccount()) {
      const accounts = await this.msalApp.getTokenCache().getAllAccounts();
      targetAccount = this.resolveExpectedAccountFromAccounts(accounts);
      if (identifier) {
        const requestedAccount = await this.resolveAccount(identifier);
        if (requestedAccount.homeAccountId !== targetAccount.homeAccountId) {
          throw new Error(
            `Account '${identifier}' does not match expected Microsoft account '${this.expectedAccountLabel()}'.`
          );
        }
      }
    } else if (identifier) {
      // resolveAccount handles empty-cache check internally
      targetAccount = await this.resolveAccount(identifier);
    } else {
      const accounts = await this.msalApp.getTokenCache().getAllAccounts();

      if (accounts.length === 0) {
        throw new Error('No accounts found. Please login first.');
      }
      // No identifier provided
      if (accounts.length === 1) {
        targetAccount = accounts[0];
      } else {
        // Multiple accounts: resolve by explicit selectedAccountId only — never fall back to accounts[0].
        // getCurrentAccount() has backward-compat fallback to first account which is unsafe for multi-account routing.
        if (this.selectedAccountId) {
          targetAccount =
            accounts.find((a: AccountInfo) => a.homeAccountId === this.selectedAccountId) ?? null;
        }
        if (!targetAccount) {
          const availableAccounts = accounts
            .map((a: AccountInfo) => a.username || a.name || 'unknown')
            .join(', ');
          throw new Error(
            `Multiple accounts configured but no 'account' parameter provided and no default selected. ` +
              `Available accounts: ${availableAccounts}. ` +
              `Pass account="<email>" in your tool call or use select-account to set a default.`
          );
        }
      }
    }

    const silentRequest = {
      account: targetAccount,
      scopes: this.scopes,
    };

    try {
      const response = await this.msalApp.acquireTokenSilent(silentRequest);
      // Persistence is owned by the cache plugin (afterCacheAccess); see getToken (issue #545).
      return response.accessToken;
    } catch (error) {
      const hint = consumersAuthorityHint(error, targetAccount, this.config.auth.authority);
      logger.error(
        `Silent token acquisition failed: ${describeAuthError(error)}${hint ? ` ${hint}` : ''}`
      );
      throw new Error(
        `Failed to acquire token for account '${targetAccount.username || targetAccount.name || 'unknown'}'. ` +
          (hint ?? 'The token may have expired. Please re-login with: --login')
      );
    }
  }
}

export default AuthManager;
export {
  type AuthManagerCreateOptions,
  type ExpectedAccountOptions,
  buildAllowedScopeDiagnostics,
  buildScopesFromEndpoints,
  buildScopeDiagnostics,
  collapseScopeHierarchy,
  getEndpointRequiredScopes,
  getEndpointScopeGroups,
  getMissingAllowedScopes,
  getMissingAllowedScopesForGroups,
  getTokenCachePath,
  getSelectedAccountPath,
  parseAllowedScopes,
  resolveAuthScopes,
  resolveAuthorizeScopes,
  wrapCache,
  unwrapCache,
  pickNewest,
};
