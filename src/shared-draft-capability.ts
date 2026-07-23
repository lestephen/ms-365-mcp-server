import { createHash } from 'node:crypto';

/**
 * Pure builder for the `eki.doc-control-shared-draft-capability/v1` proof
 * consumed by the EKI Document Control skill (EnviroKinetics/ms365-mcp#17).
 *
 * The consumer validates a CLOSED record and recomputes `proofSha256` as the
 * SHA-256 of the proof with `proofSha256` removed, canonicalized as
 * `json.dumps(value, ensure_ascii=False, separators=(",",":"), sort_keys=True)`
 * (see eki/scripts .../doc-control-capability-evidence.py). This module MUST
 * produce a byte-identical canonicalization so the digest matches across the
 * TypeScript producer and the Python consumer.
 *
 * The builder is pure and does no I/O: it takes already-resolved facts and
 * emits the proof. Readiness is DERIVED from those facts, never assumed.
 */

export const CAPABILITY_SCHEMA = 'eki.doc-control-shared-draft-capability/v1';
export const CAPABILITY_MODE = 'read_only_no_mailbox_mutation';
export const CONNECTOR_PROVIDER = 'eki-ms365-mcp';
export const CONNECTOR_OPERATION = 'get-shared-draft-capability';
export const REQUIRED_DRAFT_SCOPE = 'Mail.ReadWrite';

/**
 * The draft-composition operations Document Control requires. Each flag is the
 * structural fact "this operation is registered on the running connector
 * profile"; the builder does not invent them.
 */
export interface OperationCapabilities {
  createDraft: boolean;
  createReplyDraft: boolean;
  readDraft: boolean;
  readDraftMime: boolean;
  createAttachmentUploadSession: boolean;
  listAttachments: boolean;
  getDownloadUrl: boolean;
}

export interface CapabilityBuildInput {
  proofId: string;
  tenantId: string;
  sessionBindingSha256: string;
  signedInUser: { objectId: string; primaryAddress: string };
  sharedIdentity: {
    recipientId: string;
    primaryAddress: string;
    recipientType: 'group' | 'shared_identity';
  };
  /** Delegated scopes from the session `scp` claim. */
  delegatedScopes: string[];
  /** Whether Exchange currently grants SendAs to the signed-in user. */
  sendAsGranted: boolean;
  /** Structural registration facts for the draft operations. */
  operations: OperationCapabilities;
  /**
   * Whether this connector profile can transmit mail: derived structurally
   * from the running config (any send-mail tool registered, or a send scope
   * held by the session). Never hardcoded by the caller.
   */
  sendOperationExposed: boolean;
  observedAt: Date;
  validUntil: Date;
}

export interface SharedDraftCapabilityProof {
  schema: string;
  proofId: string;
  mode: string;
  ready: boolean;
  connector: {
    provider: string;
    operation: string;
    tenantId: string;
    sessionBindingSha256: string;
  };
  signedInUser: { objectId: string; primaryAddress: string };
  sharedIdentity: {
    recipientId: string;
    primaryAddress: string;
    recipientType: 'group' | 'shared_identity';
  };
  permissions: {
    delegatedScopes: string[];
    accessRight: 'SendAs';
    trusteeObjectId: string;
    granted: boolean;
  };
  operations: OperationCapabilities & { sendOperationExposed: boolean };
  customerMutationPerformed: false;
  emailSendPermitted: false;
  observedAt: string;
  validUntil: string;
  proofSha256: string;
}

/**
 * Recursively sort object keys so serialization is deterministic and matches
 * Python's `sort_keys=True`. Keys here are ASCII, so JS UTF-16 ordering and
 * Python code-point ordering agree.
 */
function sortValue(value: unknown): unknown {
  if (Array.isArray(value)) {
    return value.map(sortValue);
  }
  if (value !== null && typeof value === 'object') {
    const source = value as Record<string, unknown>;
    const sorted: Record<string, unknown> = {};
    for (const key of Object.keys(source).sort()) {
      sorted[key] = sortValue(source[key]);
    }
    return sorted;
  }
  return value;
}

/**
 * Canonical JSON matching the consumer's
 * `json.dumps(..., ensure_ascii=False, separators=(",",":"), sort_keys=True)`.
 * `JSON.stringify` already uses comma/colon separators with no spaces and emits
 * non-ASCII literally, so the only transform needed is recursive key sorting.
 */
export function canonicalJson(value: unknown): string {
  return JSON.stringify(sortValue(value));
}

export function canonicalSha256(value: unknown): string {
  return createHash('sha256').update(canonicalJson(value), 'utf8').digest('hex');
}

function lowerSafe(value: string): string {
  return value.toLowerCase();
}

/**
 * Build the closed proof and stamp its `proofSha256`. Readiness is true only
 * when: the session holds `Mail.ReadWrite`, Exchange grants SendAs, every
 * required draft operation is registered, and no send operation is exposed.
 */
export function buildSharedDraftCapabilityProof(
  input: CapabilityBuildInput
): SharedDraftCapabilityProof {
  const delegatedScopes = Array.from(new Set(input.delegatedScopes)).sort();
  const hasDraftScope = delegatedScopes.includes(REQUIRED_DRAFT_SCOPE);
  const operationsComplete =
    input.operations.createDraft &&
    input.operations.createReplyDraft &&
    input.operations.readDraft &&
    input.operations.readDraftMime &&
    input.operations.createAttachmentUploadSession &&
    input.operations.listAttachments &&
    input.operations.getDownloadUrl;

  const ready =
    input.sendAsGranted && hasDraftScope && operationsComplete && !input.sendOperationExposed;

  const trusteeObjectId = lowerSafe(input.signedInUser.objectId);

  const unsigned: Omit<SharedDraftCapabilityProof, 'proofSha256'> = {
    schema: CAPABILITY_SCHEMA,
    proofId: lowerSafe(input.proofId),
    mode: CAPABILITY_MODE,
    ready,
    connector: {
      provider: CONNECTOR_PROVIDER,
      operation: CONNECTOR_OPERATION,
      tenantId: lowerSafe(input.tenantId),
      sessionBindingSha256: input.sessionBindingSha256,
    },
    signedInUser: {
      objectId: trusteeObjectId,
      primaryAddress: input.signedInUser.primaryAddress.toLowerCase(),
    },
    sharedIdentity: {
      recipientId: lowerSafe(input.sharedIdentity.recipientId),
      primaryAddress: input.sharedIdentity.primaryAddress.toLowerCase(),
      recipientType: input.sharedIdentity.recipientType,
    },
    permissions: {
      delegatedScopes,
      accessRight: 'SendAs',
      trusteeObjectId,
      granted: input.sendAsGranted,
    },
    operations: {
      createDraft: input.operations.createDraft,
      createReplyDraft: input.operations.createReplyDraft,
      readDraft: input.operations.readDraft,
      readDraftMime: input.operations.readDraftMime,
      createAttachmentUploadSession: input.operations.createAttachmentUploadSession,
      listAttachments: input.operations.listAttachments,
      getDownloadUrl: input.operations.getDownloadUrl,
      sendOperationExposed: input.sendOperationExposed,
    },
    customerMutationPerformed: false,
    emailSendPermitted: false,
    observedAt: input.observedAt.toISOString(),
    validUntil: input.validUntil.toISOString(),
  };

  return { ...unsigned, proofSha256: canonicalSha256(unsigned) };
}
