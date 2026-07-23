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
import {
  configureExoTransport,
  __testing as brokerTesting,
  type ExoTransport,
  type ExoRecipientPermissionSnapshot,
} from '../exo-recipient-broker.js';
import { canonicalSha256 } from '../shared-draft-capability.js';

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

/** A graph client that fails loudly if the tool ever tries to call Graph. */
function noGraphClient(): { makeRequest: ReturnType<typeof vi.fn> } {
  return {
    makeRequest: vi.fn(() => {
      throw new Error('get-shared-draft-capability must not call Microsoft Graph');
    }),
  };
}

function ctxWith(
  registered: string[],
  graphClient: { makeRequest: ReturnType<typeof vi.fn> } = noGraphClient()
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
  configureExoTransport(transport);
});

afterEach(() => {
  brokerTesting.reset();
  vi.clearAllMocks();
});

describe('get-shared-draft-capability tool', () => {
  it('emits a ready proof for the positive case and never calls Graph (read-only by construction)', async () => {
    const graph = noGraphClient();
    const { result, json } = await invoke(
      scopedToken('Mail.ReadWrite User.Read'),
      {},
      ctxWith(DRAFT_TOOLS, graph)
    );
    expect(result.isError).toBeFalsy();
    expect(json.ready).toBe(true);
    expect(json.operations.sendOperationExposed).toBe(false);
    expect(json.permissions.granted).toBe(true);
    expect(json.signedInUser.objectId).toBe(USER_OID);
    expect(json.sharedIdentity.primaryAddress).toBe(SHARED);
    // read-only: no Graph call, and the transport was only read.
    expect(graph.makeRequest).not.toHaveBeenCalled();
    expect(transport.readSendAsAcl).toHaveBeenCalledTimes(1);
    expect(transport.readSendAsAcl).toHaveBeenCalledWith(SHARED);
    // digest self-consistency
    const { proofSha256, ...unsigned } = json;
    expect(canonicalSha256(unsigned)).toBe(proofSha256);
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
    configureExoTransport(undefined);
    const { result } = await invoke(scopedToken('Mail.ReadWrite'), {}, ctxWith(DRAFT_TOOLS));
    expect(result.isError).toBe(true);
    expect(result.content[0].text).toContain('broker is not configured');
  });

  it('errors when there is no authenticated delegated session', async () => {
    const { result } = await invoke(undefined, {}, ctxWith(DRAFT_TOOLS));
    expect(result.isError).toBe(true);
    expect(transport.readSendAsAcl).not.toHaveBeenCalled();
  });
});
