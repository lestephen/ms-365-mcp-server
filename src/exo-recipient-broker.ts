/**
 * Internal read-only Exchange Online recipient-permission broker.
 *
 * Why this exists (EnviroKinetics ms365-mcp#17): "Send As" is an Exchange
 * recipient ACL, assigned to a user or group independently of Microsoft 365
 * group membership. Microsoft Graph does NOT expose it; the documented read
 * interface is the Exchange Online cmdlet `Get-EXORecipientPermission`
 * (https://learn.microsoft.com/en-us/powershell/module/exchangepowershell/get-exorecipientpermission).
 * So the `get-shared-draft-capability` operation cannot prove Send As from the
 * Graph/OBO token alone; it must read the Exchange ACL through this broker.
 *
 * Design: the broker logic here is PURE and transport-agnostic. All live
 * Exchange access is confined to an injected {@link ExoTransport}. This keeps
 * the SendAs-matching, recipient-type mapping, and staleness rules fully unit
 * testable with synthetic snapshots and no network, and confines every
 * live-boundary detail (REST/PowerShell surface, cmdlet field names, app-only
 * certificate auth, RBAC role) to the transport implementation.
 *
 * READ-ONLY BY CONSTRUCTION: the transport contract exposes a single read
 * method. There is no mutation entry point.
 */

import { CAPABILITY_FRESHNESS_WINDOW_MS } from './shared-draft-capability.js';

/** Recipient-type strings Exchange returns that we treat as a group identity. */
const GROUP_RECIPIENT_TYPES = new Set([
  'groupmailbox',
  'unifiedgroup',
  'mailuniversalsecuritygroup',
  'mailuniversaldistributiongroup',
  'dynamicdistributiongroup',
  'universalsecuritygroup',
  'universaldistributiongroup',
]);

/**
 * Default window within which the Exchange read is considered current. Shared
 * with the proof validity policy so freshness and validUntil cannot drift apart.
 */
const DEFAULT_FRESHNESS_WINDOW_MS = CAPABILITY_FRESHNESS_WINDOW_MS;

/**
 * Maximum tolerated clock skew for a FUTURE-dated Exchange observation. An
 * observedAt further ahead of the post-read clock than this is treated as
 * invalid (not current), so a bad or spoofed timestamp cannot satisfy
 * freshness.
 */
const MAX_FUTURE_SKEW_MS = 5_000; // 5 seconds

/**
 * The proof's `sharedIdentity.recipientType` is a closed enum of exactly
 * `'group' | 'shared_identity'` (see the consumer contract
 * `eki.doc-control-shared-draft-capability/v1`). Exchange returns a richer
 * `RecipientType` / `RecipientTypeDetails`; this maps it onto the two allowed
 * values.
 *
 * LIVE-BOUNDARY ASSUMPTION: the exact `RecipientTypeDetails` string that
 * doccontrol@envirokinetics.com resolves to (SharedMailbox vs GroupMailbox vs a
 * mail-enabled security group) must be confirmed against the live directory.
 * Both branches are handled here; a shared mailbox maps to 'shared_identity'
 * and any group form maps to 'group'.
 */
export function mapRecipientType(recipientType: string | undefined): 'group' | 'shared_identity' {
  const normalized = (recipientType ?? '').toLowerCase().replace(/[^a-z]/g, '');
  return GROUP_RECIPIENT_TYPES.has(normalized) ? 'group' : 'shared_identity';
}

/**
 * One recipient-permission row, already annotated by the transport with the
 * trustee's resolved Entra object id and primary address when Exchange could
 * resolve them. Matching on the object id (not the display Trustee string) is
 * the strong path; `trusteePrimaryAddress` is a fallback the transport may fill
 * for diagnostics only.
 */
export interface ExoAnnotatedPermission {
  /** Entra object id of the trustee, lowercased, when resolvable. */
  trusteeObjectId?: string;
  /** Primary SMTP address of the trustee, when resolvable. */
  trusteePrimaryAddress?: string;
  /**
   * Whether the trustee resolves to a single user or a group. Set by the
   * transport from the trustee's Exchange RecipientType. Undefined when the
   * transport could not classify it; an unclassified Deny is treated as
   * potentially applicable (conservative). Only a Deny whose trustee is a
   * DIFFERENT resolved user (`user`) is provably unrelated to the signed-in
   * user and can be excluded.
   */
  trusteeType?: 'user' | 'group';
  /** Exchange `AccessControlType`, e.g. 'Allow' or 'Deny'. */
  accessControlType: string;
  /** Exchange `AccessRights`, e.g. ['SendAs']. */
  accessRights: string[];
}

/** The resolved shared recipient the ACL was read for. */
export interface ExoResolvedRecipient {
  /** Entra object id (ExternalDirectoryObjectId), lowercased. */
  objectId?: string;
  /**
   * Primary SMTP address as Exchange reports it. Optional because a malformed
   * transport response may omit it; a missing address is a clean negative, not
   * a crash.
   */
  primaryAddress?: string;
  /** Raw Exchange RecipientTypeDetails / RecipientType string. */
  recipientType?: string;
}

/**
 * A point-in-time snapshot of the Send As ACL for one shared identity. The
 * transport stamps `observedAt` at the moment it read Exchange; the broker uses
 * it to reject stale data.
 */
export interface ExoRecipientPermissionSnapshot {
  observedAt: Date;
  recipient?: ExoResolvedRecipient;
  permissions: ExoAnnotatedPermission[];
}

/**
 * Read-only Exchange Online access contract. The single method resolves the
 * shared recipient and returns its Send As ACL as a snapshot. Implementations
 * MUST NOT mutate Exchange state.
 */
export interface ExoTransport {
  readSendAsAcl(sharedPrimaryAddress: string): Promise<ExoRecipientPermissionSnapshot>;
}

export interface SendAsGrantQuery {
  /** Entra object id of the signed-in (trustee) user. */
  signedInUserObjectId: string;
  /** Primary address of the signed-in user (diagnostics / fallback match). */
  signedInUserPrimaryAddress?: string;
  /** The approved shared identity whose ACL is read. */
  sharedPrimaryAddress: string;
  /**
   * Clock read AFTER the Exchange read returns, for freshness evaluation.
   * Injectable for tests; defaults to `new Date()`. Callers MUST NOT pass a
   * pre-read timestamp: freshness must account for the time spent awaiting
   * Exchange and must reject future-dated observations.
   */
  clock?: () => Date;
  /** Override the staleness window; defaults to 2 minutes. */
  freshnessWindowMs?: number;
}

export interface SendAsGrant {
  /** Lowercased Entra object id of the shared recipient. */
  recipientId?: string;
  /** Primary address of the shared recipient as Exchange reported it. */
  recipientPrimaryAddress?: string;
  recipientType: 'group' | 'shared_identity';
  /** Always the queried right; the proof's permissions.accessRight is 'SendAs'. */
  accessRight: 'SendAs';
  /** Lowercased trustee object id (the signed-in user). */
  trusteeObjectId: string;
  /** True only for an explicit, current, Allow + SendAs grant to this trustee. */
  granted: boolean;
  /** ISO-8601 timestamp of the underlying Exchange read. */
  sourceObservedAt: string;
  /** Set when the grant was denied specifically because the read was stale. */
  stale: boolean;
  /** Set when the shared recipient could not be resolved or did not match. */
  recipientResolved: boolean;
  /**
   * Set when a SendAs Deny row applies to the signed-in user or cannot be
   * excluded (a group or unclassified trustee), forcing a conservative
   * negative even though a direct user Allow exists.
   */
  unexcludableDeny: boolean;
}

function lower(value: string | undefined): string | undefined {
  return value === undefined ? undefined : value.toLowerCase();
}

/**
 * Evaluate the Send As grant for the signed-in user on the shared identity from
 * a live Exchange read. Pure over the injected transport. Returns a normalized
 * grant; `granted` is true only when Exchange currently records an explicit
 * `Allow` + `SendAs` entry for exactly this trustee on exactly this recipient,
 * within the freshness window. Group membership, ownership, and historical
 * state are never treated as equivalent (that is the whole point of #17).
 */
export async function readSendAsGrant(
  transport: ExoTransport,
  query: SendAsGrantQuery
): Promise<SendAsGrant> {
  const freshnessWindowMs = query.freshnessWindowMs ?? DEFAULT_FRESHNESS_WINDOW_MS;
  const trusteeObjectId = query.signedInUserObjectId.toLowerCase();
  const sharedTarget = query.sharedPrimaryAddress.toLowerCase();

  const snapshot = await transport.readSendAsAcl(query.sharedPrimaryAddress);
  // Capture the freshness clock AFTER the read returns, so the time spent
  // awaiting Exchange counts against the freshness window.
  const now = (query.clock ?? (() => new Date()))();
  const observedAtIso = snapshot.observedAt.toISOString();

  const recipient = snapshot.recipient;
  const recipientObjectId = lower(recipient?.objectId);
  // A recipient is resolved only when Exchange returns BOTH a directory object
  // id AND the exact requested primary address. A missing object id is a
  // negative result (never a ready proof with an unresolved recipient).
  const recipientResolved =
    !!recipient &&
    recipientObjectId !== undefined &&
    recipient.primaryAddress?.toLowerCase() === sharedTarget;
  const recipientType = mapRecipientType(recipient?.recipientType);

  const base: SendAsGrant = {
    recipientId: recipientObjectId,
    recipientPrimaryAddress: recipient?.primaryAddress,
    recipientType,
    accessRight: 'SendAs',
    trusteeObjectId,
    granted: false,
    sourceObservedAt: observedAtIso,
    stale: false,
    recipientResolved,
    unexcludableDeny: false,
  };

  // Wrong / unresolved shared identity: never a positive.
  if (!recipientResolved) {
    return base;
  }

  // Stale or invalidly future-dated ACL: not current proof. Age is measured
  // against the post-read clock; an observation dated further into the future
  // than the tolerated skew is rejected the same as an over-age one.
  const ageMs = now.getTime() - snapshot.observedAt.getTime();
  if (ageMs > freshnessWindowMs || ageMs < -MAX_FUTURE_SKEW_MS) {
    return { ...base, stale: true };
  }

  // Evaluate the SendAs rows. Granting requires a direct Allow to the signed-in
  // user's own object id. Blocking is deliberately conservative and READ-ONLY
  // (we never resolve group membership): a Deny to the user's own id blocks, and
  // so does ANY other SendAs Deny we cannot prove is unrelated to the user. Only
  // a Deny to a DIFFERENT resolved USER is provably unrelated and excludable; a
  // Deny to a group, or to an unresolved/unclassified trustee, is treated as
  // potentially applicable (the user might belong to that group), forcing a
  // negative even behind a direct Allow.
  let sendAsAllowed = false;
  let unexcludableDeny = false;
  for (const perm of snapshot.permissions) {
    const rightMatches = perm.accessRights.some((right) => right.toLowerCase() === 'sendas');
    if (!rightMatches) continue;
    const control = perm.accessControlType.toLowerCase();
    const permTrustee = lower(perm.trusteeObjectId);
    const isOurUser = permTrustee !== undefined && permTrustee === trusteeObjectId;

    if (control === 'allow') {
      if (isOurUser) sendAsAllowed = true;
      continue; // group/other Allow never grants us (no membership inference)
    }
    if (control !== 'deny') continue;

    if (isOurUser) {
      unexcludableDeny = true; // direct user Deny always wins
      continue;
    }
    // A Deny to some other trustee. Provably unrelated only if it is a resolved,
    // different USER; otherwise it might apply to us via group membership.
    const provablyUnrelated =
      perm.trusteeType === 'user' && permTrustee !== undefined && permTrustee !== trusteeObjectId;
    if (!provablyUnrelated) unexcludableDeny = true;
  }

  return { ...base, unexcludableDeny, granted: sendAsAllowed && !unexcludableDeny };
}

/**
 * The live transport is NOT provisioned in this build. Standing it up requires,
 * outside this repository:
 *  - an Entra app registration with the `Office 365 Exchange Online ->
 *    Exchange.ManageAsApp` application permission and admin consent;
 *  - a client certificate for that app; and
 *  - an Exchange RBAC role restricted to recipient reads (e.g. "View-Only
 *    Recipients") assigned to the app's service principal.
 * With those, the implementation runs `Get-EXORecipient` (to resolve the shared
 * identity and each SendAs trustee) and `Get-EXORecipientPermission -Identity
 * <sharedAddress>` over the Exchange Online REST admin endpoint, stamps
 * `observedAt = new Date()`, and annotates each row with the trustee object id.
 *
 * Until that provisioning and owner authorization exist, no transport is
 * configured and `get-shared-draft-capability` reports that the Exchange broker
 * is unavailable rather than guessing.
 */
export interface ExoBrokerConfig {
  transport: ExoTransport;
  /**
   * The Entra tenant that owns the doccontrol mailbox and that this transport is
   * provisioned for. The probe requires this to equal the approved Document
   * Control tenant AND the signed-in token's tenant, so a transport wired for a
   * different tenant cannot satisfy the check.
   */
  tenantId: string;
}

let configuredBroker: ExoBrokerConfig | undefined;

export function configureExoBroker(config: ExoBrokerConfig | undefined): void {
  configuredBroker = config;
}

export function getConfiguredExoBroker(): ExoBrokerConfig | undefined {
  return configuredBroker;
}

// Exposed for tests only.
export const __testing = {
  DEFAULT_FRESHNESS_WINDOW_MS,
  reset: () => {
    configuredBroker = undefined;
  },
};
