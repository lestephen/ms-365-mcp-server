import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';

vi.mock('../logger.js', () => ({
  default: { info: vi.fn(), warn: vi.fn(), error: vi.fn(), debug: vi.fn() },
}));
// Keep the real audit-log module (the tool imports getSessionClaims from it);
// only silence the sink.
vi.mock('../audit-log.js', async (importOriginal) => {
  const actual = await importOriginal<typeof import('../audit-log.js')>();
  return { ...actual, auditLog: vi.fn() };
});

import { UTILITY_TOOLS } from '../graph-tools.js';
import { requestContext } from '../request-context.js';
import { auditLog } from '../audit-log.js';
import {
  configureExoBroker,
  __testing as brokerTesting,
  type ExoTransport,
  type ExoRecipientPermissionSnapshot,
} from '../exo-recipient-broker.js';
import { canonicalSha256 } from '../shared-draft-capability.js';

const APPROVED_TENANT_ENV = 'MS365_MCP_DOCCONTROL_TENANT_ID';
const SHARED = 'doccontrol@envirokinetics.com';
const USER_OID = '22222222-2222-4222-8222-222222222222';
const RECIP_OID = '33333333-3333-4333-8333-333333333333';
const TENANT = '11111111-1111-4111-8111-111111111111';

const DRAFT_TOOLS = [
  'create-draft-email',
  'create-reply-draft',
  'get-mail-message',
  'get-mail-message-mime',
  'create-mail-attachment-upload-session',
  'list-mail-attachments',
  'get-download-url',
];

const tool = UTILITY_TOOLS.find((t) => t.name === 'get-shared-draft-capability')!;

function makeToken(claims: Record<string, unknown>): string {
  const header = Buffer.from(JSON.stringify({ alg: 'none', typ: 'JWT' })).toString('base64url');
  const payload = Buffer.from(JSON.stringify(claims)).toString('base64url');
  return `${header}.${payload}.`;
}

function positiveSnapshot(): ExoRecipientPermissionSnapshot {
  return {
    observedAt: new Date(),
    recipient: { objectId: RECIP_OID, primaryAddress: SHARED, recipientType: 'SharedMailbox' },
    permissions: [
      { trusteeObjectId: USER_OID, accessControlType: 'Allow', accessRights: ['SendAs'] },
    ],
  };
}

const DEFAULT_ME = { id: USER_OID, userPrincipalName: 'engineer@envirokinetics.com' };

/**
 * A graph client that serves an authenticated read-only GET /me and refuses any
 * other endpoint. This is how the probe resolves the signed-in identity; a
 * write verb or a second endpoint would throw and be caught as a failure.
 */
function meGraphClient(me: Record<string, unknown> = DEFAULT_ME): {
  makeRequest: ReturnType<typeof vi.fn>;
} {
  return {
    makeRequest: vi.fn(async (endpoint: string) => {
      if (endpoint === '/me') return me;
      throw new Error(`unexpected Graph call: ${endpoint}`);
    }),
  };
}

/** A graph client whose GET /me rejects, as Microsoft would for a forged token. */
function rejectingGraphClient(): { makeRequest: ReturnType<typeof vi.fn> } {
  return {
    makeRequest: vi.fn(async () => {
      throw new Error('Microsoft Graph API error: 401 Unauthorized');
    }),
  };
}

function ctxWith(
  registered: string[],
  graphClient: { makeRequest: ReturnType<typeof vi.fn> } = meGraphClient()
) {
  return {
    graphClient: graphClient as never,
    authManager: undefined,
    multiAccount: false,
    accountNames: [],
    registeredToolNames: new Set(registered),
  };
}

async function invoke(
  token: string | undefined,
  params: Record<string, unknown>,
  ctx: ReturnType<typeof ctxWith>
) {
  const run = () => tool.execute(params, ctx);
  const result = token ? await requestContext.run({ accessToken: token }, run) : await run();
  const text = result.content?.[0]?.text as string;
  return { result, json: text ? JSON.parse(text) : undefined };
}

const scopedToken = (scopes: string, extra: Record<string, unknown> = {}) =>
  makeToken({
    oid: USER_OID,
    tid: TENANT,
    upn: 'engineer@envirokinetics.com',
    sid: 'session-abc',
    scp: scopes,
    ...extra,
  });

let transport: ExoTransport & { readSendAsAcl: ReturnType<typeof vi.fn> };

beforeEach(() => {
  transport = { readSendAsAcl: vi.fn(async () => positiveSnapshot()) };
  configureExoBroker({ transport, tenantId: TENANT });
  process.env[APPROVED_TENANT_ENV] = TENANT;
});

afterEach(() => {
  brokerTesting.reset();
  delete process.env[APPROVED_TENANT_ENV];
  vi.clearAllMocks();
});

describe('get-shared-draft-capability tool', () => {
  it('emits a ready proof for the positive case using only a read-only GET /me (read-only by construction)', async () => {
    const graph = meGraphClient();
    const { result, json } = await invoke(
      scopedToken('Mail.ReadWrite User.Read'),
      {},
      ctxWith(DRAFT_TOOLS, graph)
    );
    expect(result.isError).toBeFalsy();
    expect(json.ready).toBe(true);
    expect(json.operations.sendOperationExposed).toBe(false);
    expect(json.permissions.granted).toBe(true);
    // Identity is bound from the authenticated /me response, not raw claims.
    expect(json.signedInUser.objectId).toBe(USER_OID);
    expect(json.signedInUser.primaryAddress).toBe('engineer@envirokinetics.com');
    expect(json.sharedIdentity.primaryAddress).toBe(SHARED);
    // read-only: the only Graph call is GET /me (single arg => no method override),
    // and the transport was only read.
    expect(graph.makeRequest).toHaveBeenCalledTimes(1);
    expect(graph.makeRequest).toHaveBeenCalledWith('/me');
    expect(graph.makeRequest.mock.calls[0][1]).toBeUndefined();
    expect(transport.readSendAsAcl).toHaveBeenCalledTimes(1);
    expect(transport.readSendAsAcl).toHaveBeenCalledWith(SHARED);
    // digest self-consistency
    const { proofSha256, ...unsigned } = json;
    expect(canonicalSha256(unsigned)).toBe(proofSha256);
  });

  it('fails closed when the signed-in identity cannot be authenticated (forged/invalid token)', async () => {
    const graph = rejectingGraphClient();
    const { result } = await invoke(scopedToken('Mail.ReadWrite'), {}, ctxWith(DRAFT_TOOLS, graph));
    expect(result.isError).toBe(true);
    // No proof is minted and the ACL is never read when identity is unproven.
    expect(transport.readSendAsAcl).not.toHaveBeenCalled();
  });

  it('binds the session without leaking the bearer token', async () => {
    const { json } = await invoke(scopedToken('Mail.ReadWrite'), {}, ctxWith(DRAFT_TOOLS));
    expect(json.connector.sessionBindingSha256).toMatch(/^[0-9a-f]{64}$/);
    // The binding is a hash, not the raw token material.
    expect(json.connector.sessionBindingSha256).not.toContain('session-abc');
  });

  it('is not ready when the session lacks Mail.ReadWrite', async () => {
    const { json } = await invoke(scopedToken('User.Read'), {}, ctxWith(DRAFT_TOOLS));
    expect(json.ready).toBe(false);
    expect(json.permissions.delegatedScopes).not.toContain('Mail.ReadWrite');
  });

  it('is not ready when Exchange does not grant SendAs to this trustee (group member)', async () => {
    transport.readSendAsAcl = vi.fn(async () => ({
      ...positiveSnapshot(),
      permissions: [
        {
          trusteeObjectId: '99999999-9999-4999-8999-999999999999',
          accessControlType: 'Allow',
          accessRights: ['SendAs'],
        },
      ],
    }));
    const { json } = await invoke(scopedToken('Mail.ReadWrite'), {}, ctxWith(DRAFT_TOOLS));
    expect(json.ready).toBe(false);
    expect(json.permissions.granted).toBe(false);
  });

  it('exposes a send operation (not ready) when a send-mail tool is registered', async () => {
    const { json } = await invoke(
      scopedToken('Mail.ReadWrite'),
      {},
      ctxWith([...DRAFT_TOOLS, 'send-mail'])
    );
    expect(json.operations.sendOperationExposed).toBe(true);
    expect(json.ready).toBe(false);
  });

  it('exposes a send operation (not ready) when the session holds a send scope', async () => {
    const { json } = await invoke(
      scopedToken('Mail.ReadWrite Mail.Send'),
      {},
      ctxWith(DRAFT_TOOLS)
    );
    expect(json.operations.sendOperationExposed).toBe(true);
    expect(json.ready).toBe(false);
  });

  it('detects a send scope case-insensitively and across the broadened send set', async () => {
    for (const scope of ['MAIL.SEND', 'SMTP.Send', 'full_access_as_user']) {
      transport.readSendAsAcl = vi.fn(async () => positiveSnapshot());
      const { json } = await invoke(
        scopedToken(`Mail.ReadWrite ${scope}`),
        {},
        ctxWith(DRAFT_TOOLS)
      );
      expect(json.operations.sendOperationExposed).toBe(true);
      expect(json.ready).toBe(false);
    }
  });

  it('is not ready for an explicit Deny SendAs ACL (through the full tool path)', async () => {
    transport.readSendAsAcl = vi.fn(async () => ({
      ...positiveSnapshot(),
      permissions: [
        { trusteeObjectId: USER_OID, accessControlType: 'Allow', accessRights: ['SendAs'] },
        { trusteeObjectId: USER_OID, accessControlType: 'Deny', accessRights: ['SendAs'] },
      ],
    }));
    const { json } = await invoke(scopedToken('Mail.ReadWrite'), {}, ctxWith(DRAFT_TOOLS));
    expect(json.permissions.granted).toBe(false);
    expect(json.ready).toBe(false);
  });

  it('is not ready for a stale ACL read (through the full tool path)', async () => {
    transport.readSendAsAcl = vi.fn(async () => ({
      ...positiveSnapshot(),
      observedAt: new Date(Date.now() - 10 * 60 * 1000), // 10 minutes old
    }));
    const { json } = await invoke(scopedToken('Mail.ReadWrite'), {}, ctxWith(DRAFT_TOOLS));
    expect(json.permissions.granted).toBe(false);
    expect(json.ready).toBe(false);
  });

  it('is not ready when Exchange cannot resolve the recipient object id (through the full tool path)', async () => {
    transport.readSendAsAcl = vi.fn(async () => ({
      ...positiveSnapshot(),
      recipient: { primaryAddress: SHARED, recipientType: 'SharedMailbox' }, // no objectId
    }));
    const { json } = await invoke(scopedToken('Mail.ReadWrite'), {}, ctxWith(DRAFT_TOOLS));
    expect(json.ready).toBe(false);
    expect(json.permissions.granted).toBe(false);
  });

  it('is not ready when a required draft operation is not registered', async () => {
    const missing = DRAFT_TOOLS.filter((t) => t !== 'get-download-url');
    const { json } = await invoke(scopedToken('Mail.ReadWrite'), {}, ctxWith(missing));
    expect(json.operations.getDownloadUrl).toBe(false);
    expect(json.ready).toBe(false);
  });

  it('rejects a request for any identity other than the approved shared identity', async () => {
    const { result } = await invoke(
      scopedToken('Mail.ReadWrite'),
      { sharedIdentity: 'someoneelse@envirokinetics.com' },
      ctxWith(DRAFT_TOOLS)
    );
    expect(result.isError).toBe(true);
    expect(transport.readSendAsAcl).not.toHaveBeenCalled();
  });

  it('errors clearly when the Exchange broker is not configured', async () => {
    configureExoBroker(undefined);
    const { result } = await invoke(scopedToken('Mail.ReadWrite'), {}, ctxWith(DRAFT_TOOLS));
    expect(result.isError).toBe(true);
    expect(result.content[0].text).toContain('broker is not configured');
  });

  it('errors when there is no authenticated delegated session', async () => {
    const { result } = await invoke(undefined, {}, ctxWith(DRAFT_TOOLS));
    expect(result.isError).toBe(true);
    expect(transport.readSendAsAcl).not.toHaveBeenCalled();
  });

  describe('tenant boundary (f1)', () => {
    it('fails closed when the signed-in tenant differs from the approved tenant', async () => {
      const foreignToken = scopedToken('Mail.ReadWrite', {
        tid: '99999999-9999-4999-8999-999999999999',
      });
      const { result } = await invoke(foreignToken, {}, ctxWith(DRAFT_TOOLS));
      expect(result.isError).toBe(true);
      // The ACL is never read once the tenant boundary is breached.
      expect(transport.readSendAsAcl).not.toHaveBeenCalled();
    });

    it('fails closed when the approved tenant is not configured', async () => {
      delete process.env[APPROVED_TENANT_ENV];
      const { result } = await invoke(scopedToken('Mail.ReadWrite'), {}, ctxWith(DRAFT_TOOLS));
      expect(result.isError).toBe(true);
      expect(transport.readSendAsAcl).not.toHaveBeenCalled();
    });

    it('fails closed when the configured broker serves a different tenant', async () => {
      configureExoBroker({ transport, tenantId: '99999999-9999-4999-8999-999999999999' });
      const { result } = await invoke(scopedToken('Mail.ReadWrite'), {}, ctxWith(DRAFT_TOOLS));
      expect(result.isError).toBe(true);
      expect(transport.readSendAsAcl).not.toHaveBeenCalled();
    });
  });

  describe('audited single-exit path (f2)', () => {
    it('audits a successful probe with status success', async () => {
      await invoke(scopedToken('Mail.ReadWrite'), {}, ctxWith(DRAFT_TOOLS));
      const events = vi.mocked(auditLog).mock.calls.map((c) => c[0]);
      expect(events).toHaveLength(1);
      expect(events[0]).toMatchObject({
        tool: 'get-shared-draft-capability',
        status: 'success',
      });
      expect(events[0].error_type).toBeUndefined();
    });

    it('audits a denied exit with status denied and an error class', async () => {
      const { result } = await invoke(
        scopedToken('Mail.ReadWrite'),
        { sharedIdentity: 'someoneelse@envirokinetics.com' },
        ctxWith(DRAFT_TOOLS)
      );
      expect(result.isError).toBe(true);
      const events = vi.mocked(auditLog).mock.calls.map((c) => c[0]);
      expect(events).toHaveLength(1);
      expect(events[0]).toMatchObject({ status: 'denied', error_type: 'unsupported_identity' });
    });

    it('audits an infrastructure failure with status error', async () => {
      const { result } = await invoke(
        scopedToken('Mail.ReadWrite'),
        {},
        ctxWith(DRAFT_TOOLS, rejectingGraphClient())
      );
      expect(result.isError).toBe(true);
      const events = vi.mocked(auditLog).mock.calls.map((c) => c[0]);
      expect(events).toHaveLength(1);
      expect(events[0]).toMatchObject({ status: 'error', error_type: 'graph_identity_rejected' });
    });

    it('never places token material in the audit event', async () => {
      await invoke(scopedToken('Mail.ReadWrite'), {}, ctxWith(DRAFT_TOOLS));
      const event = vi.mocked(auditLog).mock.calls[0][0];
      const serialized = JSON.stringify(event);
      expect(serialized).not.toContain('session-abc'); // sid
      expect(serialized).not.toContain('eyJ'); // no JWT segment
    });
  });
});
