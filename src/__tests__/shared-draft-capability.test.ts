import { describe, it, expect } from 'vitest';
import { createHash } from 'node:crypto';
import {
  buildSharedDraftCapabilityProof,
  canonicalJson,
  canonicalSha256,
  type CapabilityBuildInput,
  type OperationCapabilities,
} from '../shared-draft-capability.js';

const ALL_OPERATIONS: OperationCapabilities = {
  createDraft: true,
  createReplyDraft: true,
  readDraft: true,
  readDraftMime: true,
  createAttachmentUploadSession: true,
  listAttachments: true,
  getDownloadUrl: true,
};

function baseInput(overrides: Partial<CapabilityBuildInput> = {}): CapabilityBuildInput {
  return {
    proofId: 'sdc-44444444-4444-4444-8444-444444444444',
    tenantId: '11111111-1111-4111-8111-111111111111',
    sessionBindingSha256: 'a'.repeat(64),
    signedInUser: {
      objectId: '22222222-2222-4222-8222-222222222222',
      primaryAddress: 'engineer@envirokinetics.com',
    },
    sharedIdentity: {
      recipientId: '33333333-3333-4333-8333-333333333333',
      primaryAddress: 'doccontrol@envirokinetics.com',
      recipientType: 'shared_identity',
    },
    delegatedScopes: ['User.Read', 'Mail.ReadWrite', 'offline_access'],
    sendAsGranted: true,
    operations: { ...ALL_OPERATIONS },
    sendOperationExposed: false,
    observedAt: new Date('2026-07-23T04:00:00.000Z'),
    validUntil: new Date('2026-07-23T04:10:00.000Z'),
    ...overrides,
  };
}

describe('canonicalJson', () => {
  it('sorts keys recursively with no whitespace, matching Python json.dumps sort_keys', () => {
    expect(canonicalJson({ b: 1, a: { d: 2, c: 3 } })).toBe('{"a":{"c":3,"d":2},"b":1}');
  });

  it('preserves array order (arrays are sequences, not sorted)', () => {
    expect(canonicalJson({ x: [3, 1, 2] })).toBe('{"x":[3,1,2]}');
  });
});

describe('buildSharedDraftCapabilityProof', () => {
  it('produces a proofSha256 equal to the SHA-256 of the proof minus proofSha256', () => {
    const proof = buildSharedDraftCapabilityProof(baseInput());
    const { proofSha256, ...unsigned } = proof;
    const recomputed = createHash('sha256').update(canonicalJson(unsigned), 'utf8').digest('hex');
    expect(proofSha256).toBe(recomputed);
    expect(canonicalSha256(unsigned)).toBe(proofSha256);
  });

  it('matches the cross-language golden digest computed by the Python consumer', () => {
    // GOLDEN VECTOR: this hex is the output of the consumer's
    // canonical_sha256() (json.dumps(..., ensure_ascii=False,
    // separators=(",",":"), sort_keys=True) then sha256) over the exact unsigned
    // proof that baseInput() produces, verified end-to-end against
    // doc-control-capability-evidence.py validate_proof (ready: true). If this
    // assertion breaks, the TS producer and Python consumer have diverged on
    // canonicalization and the proof will be rejected in production.
    const GOLDEN = '53a71c7d80b9631377bd5bce2e12bcfd274bac683ba3099571173617378c1b01';
    const proof = buildSharedDraftCapabilityProof(baseInput());
    expect(proof.proofSha256).toBe(GOLDEN);
  });

  it('lowercases identifiers and casefolds addresses so they pass the consumer safe-id/email rules', () => {
    const proof = buildSharedDraftCapabilityProof(
      baseInput({
        proofId: 'SDC-ABCDEF',
        tenantId: 'AAAAAAAA-1111-4111-8111-111111111111',
        signedInUser: {
          objectId: 'BBBBBBBB-2222-4222-8222-222222222222',
          primaryAddress: 'Engineer@Envirokinetics.com',
        },
        sharedIdentity: {
          recipientId: 'CCCCCCCC-3333-4333-8333-333333333333',
          primaryAddress: 'DocControl@EnviroKinetics.com',
          recipientType: 'group',
        },
      })
    );
    expect(proof.proofId).toBe('sdc-abcdef');
    expect(proof.connector.tenantId).toBe('aaaaaaaa-1111-4111-8111-111111111111');
    expect(proof.signedInUser.objectId).toBe('bbbbbbbb-2222-4222-8222-222222222222');
    expect(proof.signedInUser.primaryAddress).toBe('engineer@envirokinetics.com');
    expect(proof.sharedIdentity.recipientId).toBe('cccccccc-3333-4333-8333-333333333333');
    expect(proof.sharedIdentity.primaryAddress).toBe('doccontrol@envirokinetics.com');
    expect(proof.permissions.trusteeObjectId).toBe(proof.signedInUser.objectId);
  });

  it('emits delegatedScopes as a sorted, de-duplicated set', () => {
    const proof = buildSharedDraftCapabilityProof(
      baseInput({ delegatedScopes: ['User.Read', 'Mail.ReadWrite', 'Mail.ReadWrite', 'User.Read'] })
    );
    expect(proof.permissions.delegatedScopes).toEqual(['Mail.ReadWrite', 'User.Read']);
  });

  it('holds the read-only invariants constant regardless of inputs', () => {
    const proof = buildSharedDraftCapabilityProof(baseInput());
    expect(proof.schema).toBe('eki.doc-control-shared-draft-capability/v1');
    expect(proof.mode).toBe('read_only_no_mailbox_mutation');
    expect(proof.connector.provider).toBe('eki-ms365-mcp');
    expect(proof.connector.operation).toBe('get-shared-draft-capability');
    expect(proof.customerMutationPerformed).toBe(false);
    expect(proof.emailSendPermitted).toBe(false);
    expect(proof.permissions.accessRight).toBe('SendAs');
  });

  describe('readiness derivation', () => {
    it('is ready for the exact positive case', () => {
      expect(buildSharedDraftCapabilityProof(baseInput()).ready).toBe(true);
    });

    it('is not ready when SendAs is not granted', () => {
      expect(buildSharedDraftCapabilityProof(baseInput({ sendAsGranted: false })).ready).toBe(
        false
      );
    });

    it('is not ready when Mail.ReadWrite is absent from the session scopes', () => {
      const proof = buildSharedDraftCapabilityProof(
        baseInput({ delegatedScopes: ['User.Read', 'offline_access'] })
      );
      expect(proof.ready).toBe(false);
    });

    it('is not ready when a required draft operation is not registered', () => {
      const proof = buildSharedDraftCapabilityProof(
        baseInput({ operations: { ...ALL_OPERATIONS, getDownloadUrl: false } })
      );
      expect(proof.ready).toBe(false);
    });

    it('is not ready when a send operation is exposed', () => {
      expect(buildSharedDraftCapabilityProof(baseInput({ sendOperationExposed: true })).ready).toBe(
        false
      );
    });
  });
});
