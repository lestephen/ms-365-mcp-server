import { describe, it, expect, vi } from 'vitest';
import {
  readSendAsGrant,
  mapRecipientType,
  type ExoTransport,
  type ExoRecipientPermissionSnapshot,
  type ExoAnnotatedPermission,
} from '../exo-recipient-broker.js';

const SHARED = 'doccontrol@envirokinetics.com';
const USER_OID = '22222222-2222-4222-8222-222222222222';
const RECIP_OID = '33333333-3333-4333-8333-333333333333';

function snapshot(
  overrides: Partial<ExoRecipientPermissionSnapshot> = {}
): ExoRecipientPermissionSnapshot {
  return {
    observedAt: new Date('2026-07-23T04:00:00.000Z'),
    recipient: {
      objectId: RECIP_OID,
      primaryAddress: SHARED,
      recipientType: 'SharedMailbox',
    },
    permissions: [],
    ...overrides,
  };
}

function transportOf(snap: ExoRecipientPermissionSnapshot): ExoTransport {
  return { readSendAsAcl: vi.fn(async () => snap) };
}

const sendAsAllow: ExoAnnotatedPermission = {
  trusteeObjectId: USER_OID,
  trusteePrimaryAddress: 'engineer@envirokinetics.com',
  accessControlType: 'Allow',
  accessRights: ['SendAs'],
};

const now = new Date('2026-07-23T04:00:30.000Z');

describe('mapRecipientType', () => {
  it('maps group forms to "group"', () => {
    for (const t of ['GroupMailbox', 'MailUniversalSecurityGroup', 'UnifiedGroup']) {
      expect(mapRecipientType(t)).toBe('group');
    }
  });
  it('maps mailbox forms to "shared_identity"', () => {
    for (const t of ['SharedMailbox', 'UserMailbox', undefined]) {
      expect(mapRecipientType(t)).toBe('shared_identity');
    }
  });
});

describe('readSendAsGrant', () => {
  it('grants for an explicit Allow + SendAs entry matching the trustee (positive)', async () => {
    const grant = await readSendAsGrant(transportOf(snapshot({ permissions: [sendAsAllow] })), {
      signedInUserObjectId: USER_OID,
      sharedPrimaryAddress: SHARED,
      now,
    });
    expect(grant.granted).toBe(true);
    expect(grant.accessRight).toBe('SendAs');
    expect(grant.trusteeObjectId).toBe(USER_OID);
    expect(grant.recipientId).toBe(RECIP_OID);
    expect(grant.recipientType).toBe('shared_identity');
    expect(grant.stale).toBe(false);
    expect(grant.recipientResolved).toBe(true);
  });

  it('denies a group member who holds no personal SendAs entry (negative)', async () => {
    // ACL grants SendAs to a different trustee (e.g. the group), not this user.
    const otherTrustee: ExoAnnotatedPermission = {
      ...sendAsAllow,
      trusteeObjectId: '99999999-9999-4999-8999-999999999999',
    };
    const grant = await readSendAsGrant(transportOf(snapshot({ permissions: [otherTrustee] })), {
      signedInUserObjectId: USER_OID,
      sharedPrimaryAddress: SHARED,
      now,
    });
    expect(grant.granted).toBe(false);
  });

  it('denies when a foreign-tenant user object id matches no ACL trustee', async () => {
    const grant = await readSendAsGrant(transportOf(snapshot({ permissions: [sendAsAllow] })), {
      signedInUserObjectId: 'ffffffff-0000-4000-8000-000000000000',
      sharedPrimaryAddress: SHARED,
      now,
    });
    expect(grant.granted).toBe(false);
  });

  it('denies when an explicit Deny + SendAs exists for the trustee even alongside an Allow', async () => {
    const deny: ExoAnnotatedPermission = { ...sendAsAllow, accessControlType: 'Deny' };
    const grant = await readSendAsGrant(
      transportOf(snapshot({ permissions: [sendAsAllow, deny] })),
      { signedInUserObjectId: USER_OID, sharedPrimaryAddress: SHARED, now }
    );
    expect(grant.granted).toBe(false);
  });

  it('denies when the trustee has a different right (e.g. FullAccess) but not SendAs', async () => {
    const fullAccess: ExoAnnotatedPermission = { ...sendAsAllow, accessRights: ['FullAccess'] };
    const grant = await readSendAsGrant(transportOf(snapshot({ permissions: [fullAccess] })), {
      signedInUserObjectId: USER_OID,
      sharedPrimaryAddress: SHARED,
      now,
    });
    expect(grant.granted).toBe(false);
  });

  it('denies and flags stale when the read is older than the freshness window', async () => {
    const staleNow = new Date('2026-07-23T04:10:00.000Z'); // 10 min after observedAt
    const grant = await readSendAsGrant(transportOf(snapshot({ permissions: [sendAsAllow] })), {
      signedInUserObjectId: USER_OID,
      sharedPrimaryAddress: SHARED,
      now: staleNow,
      freshnessWindowMs: 120_000,
    });
    expect(grant.granted).toBe(false);
    expect(grant.stale).toBe(true);
  });

  it('denies when the shared recipient does not resolve to the requested address', async () => {
    const grant = await readSendAsGrant(
      transportOf(
        snapshot({
          recipient: { objectId: RECIP_OID, primaryAddress: 'someoneelse@envirokinetics.com' },
          permissions: [sendAsAllow],
        })
      ),
      { signedInUserObjectId: USER_OID, sharedPrimaryAddress: SHARED, now }
    );
    expect(grant.granted).toBe(false);
    expect(grant.recipientResolved).toBe(false);
  });

  it('denies when the shared recipient is unresolved entirely', async () => {
    const grant = await readSendAsGrant(
      transportOf(snapshot({ recipient: undefined, permissions: [sendAsAllow] })),
      { signedInUserObjectId: USER_OID, sharedPrimaryAddress: SHARED, now }
    );
    expect(grant.granted).toBe(false);
    expect(grant.recipientResolved).toBe(false);
  });

  it('matches trustee object id case-insensitively', async () => {
    const upper: ExoAnnotatedPermission = {
      ...sendAsAllow,
      trusteeObjectId: USER_OID.toUpperCase(),
    };
    const grant = await readSendAsGrant(transportOf(snapshot({ permissions: [upper] })), {
      signedInUserObjectId: USER_OID,
      sharedPrimaryAddress: SHARED,
      now,
    });
    expect(grant.granted).toBe(true);
  });
});
