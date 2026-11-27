/**
 * Consumer contract: golden fixtures of the tool results and input schemas that
 * downstream consumer code (EnviroKinetics/eki helpers and hooks) parses.
 *
 * Every fixture comes from the real server: MicrosoftGraphServer builds the MCP server
 * with the production deployment's flags, a real GraphClient talks to a mocked `fetch`
 * that answers with Microsoft Graph sample payloads, and an in-memory MCP client calls
 * the tools over the wire. A change that alters what a consumer sees therefore shows up
 * as a diff under contracts/consumer/.
 *
 * Check mode (the default) fails when the regenerated output differs from the committed
 * files. `UPDATE_CONTRACTS=1 npx vitest run test/consumer-contract.test.ts` rewrites them.
 */
import { existsSync, mkdirSync, readdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { Client } from '@modelcontextprotocol/client';
import { InMemoryTransport, type McpServer } from '@modelcontextprotocol/server';
import { afterAll, afterEach, beforeAll, beforeEach, describe, expect, it, vi } from 'vitest';
import type AuthManager from '../src/auth.js';
import GraphClient from '../src/graph-client.js';
import {
  configureAttachmentMinting,
  resetAttachmentMinting,
} from '../src/lib/attachment-minting.js';
import { AttachmentTicketStore } from '../src/lib/attachment-tickets.js';
import { loadAttachmentUrlConfig } from '../src/lib/attachment-url-config.js';
import { __resetSharedBreakerForTests } from '../src/lib/graph-resilience.js';
import MicrosoftGraphServer from '../src/server.js';
import { getCombinedPresetPattern } from '../src/tool-categories.js';

// Production sets no MS365_MCP_* variable that changes a tool result, and some are read
// when a module loads. Clear any inherited from the caller's shell before the server
// modules import, so a local setting (MS365_MCP_BLOB_OUTPUT, for example) cannot change
// the fixtures. Vitest gives each test file its own worker environment.
vi.hoisted(() => {
  for (const key of Object.keys(process.env)) {
    if (key.startsWith('MS365_MCP_')) delete process.env[key];
  }
});

// Ticket ids are 32 bytes from randomBytes. Replaced with a counter-seeded digest so the
// minted URLs (and their signatures) are byte-stable; the counter resets for every case.
const ticketRandom = vi.hoisted(() => ({ counter: 0 }));
vi.mock('node:crypto', async (importOriginal) => {
  const actual = await importOriginal<typeof import('node:crypto')>();
  const randomBytes = ((size: number, callback?: unknown) => {
    if (callback !== undefined) {
      return (actual.randomBytes as (n: number, cb: unknown) => void)(size, callback);
    }
    const chunks: Buffer[] = [];
    let produced = 0;
    while (produced < size) {
      const chunk = actual
        .createHash('sha256')
        .update(`consumer-contract-${ticketRandom.counter++}`)
        .digest();
      chunks.push(chunk);
      produced += chunk.length;
    }
    return Buffer.concat(chunks).subarray(0, size);
  }) as typeof actual.randomBytes;
  return { ...actual, randomBytes, default: { ...actual, randomBytes } };
});

vi.mock('../src/logger.js', () => ({
  default: { info: vi.fn(), warn: vi.fn(), error: vi.fn(), debug: vi.fn() },
}));

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const CONTRACT_DIR = path.join(__dirname, '..', 'contracts', 'consumer');
const UPDATE = process.env.UPDATE_CONTRACTS === '1';

const NOW = '2030-01-01T00:00:00.000Z';
const NOW_EPOCH = Math.floor(Date.parse(NOW) / 1000);

// Production deployment flags, copied by hand from EnviroKinetics/ms365-mcp
// infrastructure/k8s/deployment.yaml. Nothing checks the copy: update it, and regenerate,
// whenever that file's server flags change.
const ATTACHMENT_URL_BASE = 'https://mcp.corp.envirokinetics.com';
const ATTACHMENT_URL_KEY = 'consumer-contract-fixed-hmac-key-not-a-secret';
const BLOCKED_TOOLS =
  '^((send|reply|reply-all|forward)-[a-z-]*(mail|draft)[a-z-]*|forward-calendar-event|share-drive-item|reply-to-group-thread|upload-file-content|upload-my-profile-photo|list-adhoc-call-transcripts|get-adhoc-call-transcript-content)$';
const EAGER_TOOLS = [
  'copilot-retrieve',
  'create-draft-email',
  'create-drive-item-share-link',
  'create-forward-draft',
  'create-group',
  'create-mail-attachment-upload-session',
  'create-onedrive-folder',
  'create-planner-bucket',
  'create-planner-task',
  'create-reply-all-draft',
  'create-reply-draft',
  'create-shared-mailbox-draft',
  'create-todo-task',
  'create-todo-task-list',
  'create-upload-session',
  'delete-mail-message',
  'download-bytes',
  'get-calendar-view',
  'get-current-user',
  'get-download-url',
  'get-drive-item',
  'get-drive-root-item',
  'get-mail-message',
  'get-mail-message-mime',
  'get-planner-plan',
  'get-planner-task',
  'get-planner-task-details',
  'get-sharepoint-site-by-path',
  'graph-batch',
  'list-channel-messages',
  'list-chat-messages',
  'list-drives',
  'list-folder-files',
  'list-groups',
  'list-mail-attachments',
  'list-mail-folder-messages',
  'list-mail-folders',
  'list-mail-messages',
  'list-my-memberships',
  'list-outlook-contacts',
  'list-plan-buckets',
  'list-plan-tasks',
  'list-planner-tasks',
  'list-sharepoint-site-drives',
  'list-sharepoint-site-list-items',
  'list-todo-task-lists',
  'search-onedrive-files',
  'search-query',
  'search-sharepoint-sites',
  'update-excel-range',
  'update-group',
  'update-mail-message',
  'update-planner-task',
  'update-planner-task-details',
];

const PRODUCTION_OPTIONS = {
  http: '3000',
  orgMode: true,
  enabledTools: getCombinedPresetPattern(['all'], { attachmentUrls: true }),
  discovery: true,
  eagerTools: EAGER_TOOLS,
  blockedTools: BLOCKED_TOOLS,
  enableAttachmentUrls: true,
  enableMcpApps: true,
};

// Not eager in production, so absent from the production tools/list; their schemas come
// from the same flags without --discovery. Every other schema tool must be eager.
const NON_EAGER_SCHEMA_TOOLS = new Set(['get-upload-url', 'add-mail-attachment']);

const SCHEMA_TOOLS = [
  'create-draft-email',
  'create-reply-draft',
  'create-reply-all-draft',
  'create-forward-draft',
  'create-shared-mailbox-draft',
  'update-mail-message',
  'add-mail-attachment',
  'create-mail-attachment-upload-session',
  'get-download-url',
  'get-upload-url',
  'execute-tool',
];

// Sample identifiers, shaped like the ones Graph returns. Outlook ids end in "=".
const MESSAGE_ID =
  'AAMkAGVmMDEzMTM4LTZmYWUtNDdkNC1hMDZiLTU1OGY5OTZhYmY4OABGAAAAAAAiQ8W967B7TKBjgx9rVEURBwAiIsqMbYjsT5e-T7KzowPTAAAAAAEMAAAiIsqMbYjsT5e-T7KzowPTAAAYbvZDAAA=';
const ATTACHMENT_ID =
  'AAMkAGVmMDEzMTM4LTZmYWUtNDdkNC1hMDZiLTU1OGY5OTZhYmY4OABGAAAAAAAiQ8W967B7TKBjgx9rVEURBwAiIsqMbYjsT5e-T7KzowPTAAAYbvZDAAABEgAQAJ5Ewq0Bq1VGhZ2n6wQfxDE=';
const MOVED_MESSAGE_ID =
  'AAMkAGVmMDEzMTM4LTZmYWUtNDdkNC1hMDZiLTU1OGY5OTZhYmY4OABGAAAAAAAiQ8W967B7TKBjgx9rVEURBwAiIsqMbYjsT5e-T7KzowPTAAAAAAEKAAAiIsqMbYjsT5e-T7KzowPTAAAYbvZEAAA=';
const DRAFT_ID =
  'AAMkAGVmMDEzMTM4LTZmYWUtNDdkNC1hMDZiLTU1OGY5OTZhYmY4OABGAAAAAAAiQ8W967B7TKBjgx9rVEURBwAiIsqMbYjsT5e-T7KzowPTAAAAAAEPAAAiIsqMbYjsT5e-T7KzowPTAAAYbvZFAAA=';
const DRIVE_ID = 'b!-RIj2DuyvEyV1T4NlOaMHk8XkS_I8MdFlUCq1BlcjgmhRfAj3-Z8RY2VpuvV_tpd';
const DRIVE_ITEM_ID = '01BYE5RZ6QN3ZWBTUFOFD3GSPGOHDJD36K';
const DRIVE_FOLDER_ID = '01BYE5RZ4CRXQ6WJOY7FDLFBYAHWQMGP6S';
const USER_ID = '48d31887-5fad-4d73-a9f5-3c356e68a038';
const GRAPH = 'https://graph.microsoft.com/v1.0';

const MAIL_ATTACHMENT_TARGET = `/me/messages/${MESSAGE_ID}/attachments/${ATTACHMENT_ID}/$value`;
const DRIVE_ITEM_TARGET = `/drives/${DRIVE_ID}/items/${DRIVE_ITEM_ID}/content`;

/** A mocked Graph answer, keyed by `METHOD url-after-/v1.0`. */
interface GraphRoute {
  status: number;
  body?: unknown;
  /** Raw body text, for non-JSON answers. */
  text?: string;
  headers?: Record<string, string>;
}

const STATUS_TEXT: Record<number, string> = {
  200: 'OK',
  201: 'Created',
  404: 'Not Found',
};

function graphErrorBody(code: string, message: string) {
  return {
    error: {
      code,
      message,
      innerError: {
        date: '2030-01-01T00:00:00',
        'request-id': '5f7a3c2e-8b1d-4e6a-9c0f-2d4b6a8e1c3f',
        'client-request-id': '5f7a3c2e-8b1d-4e6a-9c0f-2d4b6a8e1c3f',
      },
    },
  };
}

const NOT_FOUND = graphErrorBody(
  'ErrorItemNotFound',
  'The specified object was not found in the store.'
);

const ATTACHMENT_META = {
  '@odata.context': `${GRAPH}/$metadata#users('${USER_ID}')/messages('${MESSAGE_ID}')/attachments(name,contentType,size)/$entity`,
  '@odata.type': '#microsoft.graph.fileAttachment',
  '@odata.mediaContentType': 'application/pdf',
  id: ATTACHMENT_ID,
  name: 'Quotation Q-1042 Rev B.pdf',
  contentType: 'application/pdf',
  size: 284117,
};

const DRIVE_ITEM = {
  '@odata.context': `${GRAPH}/$metadata#drives('${encodeURIComponent(DRIVE_ID)}')/items/$entity`,
  '@microsoft.graph.downloadUrl':
    'https://contoso.sharepoint.com/sites/Engineering/_layouts/15/download.aspx?UniqueId=2d9c8a1e-6b3f-4c7d-9e0a-1f2b3c4d5e6f&Translate=false&tempauth=v1.eyJzaXRlaWQiOiJleGFtcGxlIn0.c2lnbmF0dXJl&ApiVersion=2.0',
  createdDateTime: '2029-11-14T16:22:41Z',
  eTag: '"{2D9C8A1E-6B3F-4C7D-9E0A-1F2B3C4D5E6F},3"',
  id: DRIVE_ITEM_ID,
  lastModifiedDateTime: '2029-12-18T09:05:12Z',
  name: 'General Arrangement Rev C.pdf',
  webUrl:
    'https://contoso.sharepoint.com/sites/Engineering/Shared%20Documents/General%20Arrangement%20Rev%20C.pdf',
  cTag: '"c:{2D9C8A1E-6B3F-4C7D-9E0A-1F2B3C4D5E6F},3"',
  size: 1843302,
  createdBy: {
    user: { email: 'adele.vance@contoso.com', id: USER_ID, displayName: 'Adele Vance' },
  },
  lastModifiedBy: {
    user: { email: 'adele.vance@contoso.com', id: USER_ID, displayName: 'Adele Vance' },
  },
  parentReference: {
    driveType: 'documentLibrary',
    driveId: DRIVE_ID,
    id: DRIVE_FOLDER_ID,
    name: 'Shared Documents',
    path: `/drives/${DRIVE_ID}/root:`,
    siteId: '2c8e4a1b-7d3f-4e9a-b6c0-5f1d2e3a4b5c',
  },
  file: {
    mimeType: 'application/pdf',
    hashes: { quickXorHash: 'm1o3SPmJhSmsS7+0n0yhGIs6v+Q=' },
  },
  fileSystemInfo: {
    createdDateTime: '2029-11-14T16:22:41Z',
    lastModifiedDateTime: '2029-12-18T09:05:12Z',
  },
  shared: { scope: 'users' },
};

const PROBE_PATH = `/me/messages/${MESSAGE_ID}/attachments/${ATTACHMENT_ID}?$select=name,contentType,size`;
const DRIVE_ITEM_PATH = `/drives/${DRIVE_ID}/items/${DRIVE_ITEM_ID}`;

function messageResource(id: string, parentFolderId: string, extra: Record<string, unknown>) {
  return {
    '@odata.context': `${GRAPH}/$metadata#users('${USER_ID}')/messages/$entity`,
    '@odata.etag': 'W/"CQAAABYAAAAiIsqMbYjsT5e/T7KzowPTAAAYbt3T"',
    id,
    createdDateTime: '2029-12-31T23:58:10Z',
    lastModifiedDateTime: '2030-01-01T00:00:00Z',
    changeKey: 'CQAAABYAAAAiIsqMbYjsT5e/T7KzowPTAAAYbt3T',
    categories: [],
    receivedDateTime: '2029-12-31T23:58:10Z',
    sentDateTime: '2029-12-31T23:58:08Z',
    hasAttachments: false,
    internetMessageId:
      '<SN6PR04MB4110E2C1F5D6B7A8C9D0E1F2A3B4C5D6@SN6PR04MB4110.namprd04.prod.outlook.com>',
    subject: 'Q-1042 quotation follow-up',
    bodyPreview: 'Hi Adele, please find the revised quotation attached.',
    importance: 'normal',
    parentFolderId,
    conversationId:
      'AAQkAGVmMDEzMTM4LTZmYWUtNDdkNC1hMDZiLTU1OGY5OTZhYmY4OAAQAMvKcH0lsUNKl0n9sEh7Lhs=',
    conversationIndex: 'AQHaXc7+y8pwfSWxQ0qXSf2wSHsuGw==',
    isDeliveryReceiptRequested: false,
    isReadReceiptRequested: false,
    isRead: true,
    isDraft: false,
    webLink: `https://outlook.office365.com/owa/?ItemID=${encodeURIComponent(id)}&exvsurl=1&viewmodel=ReadMessageItem`,
    inferenceClassification: 'focused',
    body: {
      contentType: 'html',
      content:
        '<html><body><p>Hi Adele, please find the revised quotation attached.</p></body></html>',
    },
    sender: { emailAddress: { name: 'Megan Bowen', address: 'megan.bowen@contoso.com' } },
    from: { emailAddress: { name: 'Megan Bowen', address: 'megan.bowen@contoso.com' } },
    toRecipients: [{ emailAddress: { name: 'Adele Vance', address: 'adele.vance@contoso.com' } }],
    ccRecipients: [],
    bccRecipients: [],
    replyTo: [],
    flag: { flagStatus: 'notFlagged' },
    ...extra,
  };
}

const MIME_TEXT = [
  'Received: from SN6PR04MB4110.namprd04.prod.outlook.com (2603:10b6:805:a2::14)',
  ' by SN6PR04MB4110.namprd04.prod.outlook.com with HTTPS; Mon, 31 Dec 2029 23:58:10 +0000',
  'From: Megan Bowen <megan.bowen@contoso.com>',
  'To: Adele Vance <adele.vance@contoso.com>',
  'Subject: Q-1042 quotation follow-up',
  'Date: Mon, 31 Dec 2029 23:58:08 +0000',
  'Message-ID: <SN6PR04MB4110E2C1F5D6B7A8C9D0E1F2A3B4C5D6@SN6PR04MB4110.namprd04.prod.outlook.com>',
  'Content-Language: en-US',
  'MIME-Version: 1.0',
  'Content-Type: multipart/alternative; boundary="_000_SN6PR04MB4110_"',
  '',
  '--_000_SN6PR04MB4110_',
  'Content-Type: text/plain; charset="us-ascii"',
  '',
  'Hi Adele, please find the revised quotation attached.',
  '',
  '--_000_SN6PR04MB4110_',
  'Content-Type: text/html; charset="us-ascii"',
  '',
  '<html><body><p>Hi Adele, please find the revised quotation attached.</p></body></html>',
  '',
  '--_000_SN6PR04MB4110_--',
  '',
].join('\r\n');

interface ContractCase {
  id: string;
  tool: string;
  route: 'direct' | 'execute-tool';
  input: Record<string, unknown>;
  graph: Record<string, GraphRoute>;
  notes: string;
}

const CASES: ContractCase[] = [
  {
    id: 'get-download-url.mail-attachment',
    tool: 'get-download-url',
    route: 'direct',
    input: { target: MAIL_ATTACHMENT_TARGET },
    graph: { [`GET ${PROBE_PATH}`]: { status: 200, body: ATTACHMENT_META } },
    notes:
      'Ticket result: downloadUrl (signed t/dgk/dgx/dgs on the attachment URL base), name/size/contentType from the probe, expiresAt, singleUse, note. Consumers must tolerate added keys.',
  },
  {
    id: 'get-download-url.drive-item',
    tool: 'get-download-url',
    route: 'direct',
    input: { target: DRIVE_ITEM_TARGET },
    graph: { [`GET ${DRIVE_ITEM_PATH}`]: { status: 200, body: DRIVE_ITEM } },
    notes:
      'With --enable-attachment-urls a drive file gets the same server-minted ticket shape as a mail attachment, not the Graph downloadUrl.',
  },
  {
    id: 'get-download-url.multi',
    tool: 'get-download-url',
    route: 'direct',
    input: { targets: [MAIL_ATTACHMENT_TARGET, DRIVE_ITEM_TARGET] },
    graph: {
      [`GET ${PROBE_PATH}`]: { status: 200, body: ATTACHMENT_META },
      [`GET ${DRIVE_ITEM_PATH}`]: { status: 200, body: DRIVE_ITEM },
    },
    notes:
      'targets form: {results:[...]} in input order, each entry the single-target result plus target; _meta.result_count counts minted links.',
  },
  {
    id: 'get-download-url.error-unknown-attachment',
    tool: 'get-download-url',
    route: 'direct',
    input: { target: MAIL_ATTACHMENT_TARGET },
    graph: { [`GET ${PROBE_PATH}`]: { status: 404, body: NOT_FOUND } },
    notes:
      'Probe 404 refuses to mint: isError, {error:"No download URL was made: Microsoft Graph API error: 404 ..."} with an id hint, _meta http_status/error_code.',
  },
  {
    id: 'get-upload-url.mail-draft',
    tool: 'get-upload-url',
    route: 'execute-tool',
    input: {
      target: `/me/messages/${DRAFT_ID}/attachments`,
      name: 'Quotation Q-1042 Rev B.pdf',
      contentType: 'application/pdf',
    },
    graph: {},
    notes:
      'get-upload-url is not an eager tool in production, so it is reached through execute-tool. Result: uploadUrl, method PUT, expiresAt, singleUse, maxBytes, name, note.',
  },
  {
    id: 'get-upload-url.drive-folder',
    tool: 'get-upload-url',
    route: 'execute-tool',
    input: {
      target: `/drives/${DRIVE_ID}/items/${DRIVE_FOLDER_ID}`,
      name: 'Performance Datasheet Rev A.xlsx',
      contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      conflictBehavior: 'rename',
    },
    graph: {},
    notes: 'Drive folder target adds conflictBehavior to the upload ticket result.',
  },
  {
    id: 'get-upload-url.multi',
    tool: 'get-upload-url',
    route: 'execute-tool',
    input: {
      files: [
        {
          target: `/me/messages/${DRAFT_ID}/attachments`,
          name: 'Quotation Q-1042 Rev B.pdf',
          contentType: 'application/pdf',
        },
        {
          target: `/drives/${DRIVE_ID}/items/${DRIVE_FOLDER_ID}`,
          name: 'Performance Datasheet Rev A.xlsx',
          contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
          conflictBehavior: 'rename',
        },
      ],
    },
    graph: {},
    notes: 'files form: {results:[...]} in input order, each entry carrying target and name.',
  },
  {
    id: 'create-mail-attachment-upload-session',
    tool: 'create-mail-attachment-upload-session',
    route: 'direct',
    input: {
      messageId: DRAFT_ID,
      body: {
        AttachmentItem: {
          attachmentType: 'file',
          name: 'Quotation Q-1042 Rev B.pdf',
          // Graph refuses upload sessions below 3 MB, so the success case declares 4 MB.
          size: 4194304,
        },
      },
    },
    graph: {
      [`POST /me/messages/${DRAFT_ID}/attachments/createUploadSession`]: {
        status: 201,
        body: {
          '@odata.context': `${GRAPH}/$metadata#microsoft.graph.uploadSession`,
          uploadUrl: `https://outlook.office.com/api/v2.0/Users('${USER_ID}@contoso.com')/Messages('${DRAFT_ID}')/AttachmentSessions('AAMkAGVmMDEzMTM4LTZmYWUtNDdkNC1hMDZiLTU1OGY5OTZhYmY4OAAuAAAAAAAiQ8W967B7TKBjgx9rVEURAQAiIsqMbYjsT5e-T7KzowPTAAAYbvZFAAA=')?authtoken=eyJhbGciOiJSUzI1NiIsImtpZCI6IkZBRDY1NDI2MkM2QUYyOTYxQUExRThDQUI3OEZGMUIyNzBFNzA3RTkiLCJ0eXAiOiJKV1QiLCJ4NXQiOiItdFpVSml4cThwWWFvZWpLdDRfeHNuRG5CLWsifQ.eyJvcmlnaW4iOiJodHRwczovL291dGxvb2sub2ZmaWNlLmNvbSJ9.c2lnbmF0dXJl`,
          expirationDateTime: '2030-01-01T06:00:00.0000000Z',
          nextExpectedRanges: ['0-'],
        },
      },
    },
    notes:
      'Graph uploadSession passed through: uploadUrl, expirationDateTime, nextExpectedRanges; the server drops @odata.context.',
  },
  {
    id: 'create-upload-session',
    tool: 'create-upload-session',
    route: 'direct',
    input: {
      driveId: DRIVE_ID,
      driveItemId: `${DRIVE_FOLDER_ID}:/General Arrangement Rev C.pdf:`,
      body: { item: { '@microsoft.graph.conflictBehavior': 'rename' } },
    },
    graph: {
      [`POST /drives/${DRIVE_ID}/items/${DRIVE_FOLDER_ID}%3A%2FGeneral%20Arrangement%20Rev%20C.pdf%3A/createUploadSession`]:
        {
          status: 200,
          body: {
            '@odata.context': `${GRAPH}/$metadata#microsoft.graph.uploadSession`,
            expirationDateTime: '2030-01-02T00:00:00.000Z',
            nextExpectedRanges: ['0-'],
            uploadUrl:
              "https://contoso.sharepoint.com/sites/Engineering/_api/v2.0/drives/b!-RIj2DuyvEyV1T4NlOaMHk8XkS_I8MdFlUCq1BlcjgmhRfAj3-Z8RY2VpuvV_tpd/items/01BYE5RZ4CRXQ6WJOY7FDLFBYAHWQMGP6S:/General%20Arrangement%20Rev%20C.pdf:/uploadSession?guid='1f2e3d4c-5b6a-4978-8695-a4b3c2d1e0f9'&overwrite=False&rename=True&dc=0&tempauth=v1.eyJzaXRlaWQiOiJleGFtcGxlIn0.c2lnbmF0dXJl",
          },
        },
    },
    notes: 'Drive uploadSession passed through: uploadUrl, expirationDateTime, nextExpectedRanges.',
  },
  {
    id: 'get-mail-message-mime',
    tool: 'get-mail-message-mime',
    route: 'direct',
    input: { messageId: MESSAGE_ID },
    graph: {
      [`GET /me/messages/${MESSAGE_ID}/$value`]: {
        status: 200,
        text: MIME_TEXT,
        headers: { 'content-type': 'text/plain' },
      },
    },
    notes:
      'Graph answers text/plain, so the MIME comes back as text JSON {message:"OK!", rawResponse:"<RFC 5322>"}, not as a blob resource item.',
  },
  {
    id: 'graph-error.404',
    tool: 'get-mail-message',
    route: 'direct',
    input: { messageId: MESSAGE_ID },
    graph: {
      [`GET /me/messages/${MESSAGE_ID}`]: { status: 404, body: NOT_FOUND },
    },
    notes:
      'Graph HTTP error on an ordinary Graph tool: isError, {error:"Microsoft Graph API error: 404 Not Found - <Graph JSON>"}, _meta http_status/error_code.',
  },
  {
    id: 'delete-mail-message',
    tool: 'delete-mail-message',
    route: 'direct',
    input: { messageId: MESSAGE_ID },
    graph: {
      [`POST /me/messages/${MESSAGE_ID}/move`]: {
        status: 201,
        body: messageResource(
          MOVED_MESSAGE_ID,
          'AAMkAGVmMDEzMTM4LTZmYWUtNDdkNC1hMDZiLTU1OGY5OTZhYmY4OAAuAAAAAAAiQ8W967B7TKBjgx9rVEURAQAiIsqMbYjsT5e-T7KzowPTAAAAAAEKAAA=',
          {}
        ),
      },
    },
    notes:
      'Move to deleteditems (POST /move with a fixed body): returns the moved message with its NEW id, not an empty success.',
  },
  {
    id: 'execute-tool.get-download-url.mail-attachment',
    tool: 'get-download-url',
    route: 'execute-tool',
    input: { target: MAIL_ATTACHMENT_TARGET },
    graph: { [`GET ${PROBE_PATH}`]: { status: 200, body: ATTACHMENT_META } },
    notes: 'Same ticket shape as the direct call when dispatched through execute-tool.',
  },
  {
    id: 'mcp-envelope.create-draft-email',
    tool: 'create-draft-email',
    route: 'direct',
    input: {
      body: {
        subject: 'Q-1042 quotation follow-up',
        body: {
          contentType: 'html',
          content: '<p>Hi Adele, please find the revised quotation attached.</p>',
        },
        toRecipients: [{ emailAddress: { address: 'adele.vance@contoso.com' } }],
      },
    },
    graph: {
      'POST /me/messages': {
        status: 201,
        body: messageResource(
          DRAFT_ID,
          'AAMkAGVmMDEzMTM4LTZmYWUtNDdkNC1hMDZiLTU1OGY5OTZhYmY4OAAuAAAAAAAiQ8W967B7TKBjgx9rVEURAQAiIsqMbYjsT5e-T7KzowPTAAAAAAEPAAA=',
          {
            isDraft: true,
            isRead: true,
            receivedDateTime: '2030-01-01T00:00:00Z',
            sentDateTime: '2030-01-01T00:00:00Z',
            sender: null,
            from: null,
            body: {
              contentType: 'html',
              content:
                '<html><head><meta http-equiv="Content-Type" content="text/html; charset=utf-8"></head><body><p>Hi Adele, please find the revised quotation attached.</p></body></html>',
            },
            toRecipients: [
              {
                emailAddress: {
                  name: 'adele.vance@contoso.com',
                  address: 'adele.vance@contoso.com',
                },
              },
            ],
          }
        ),
      },
    },
    notes:
      'Full tools/call wire envelope with --enable-mcp-apps on (create-draft-email has a mail-draft view): one pretty-printed JSON text item, no structuredContent, _meta http_status/response_bytes.',
  },
  {
    id: 'blocked.execute-tool.send-mail',
    tool: 'send-mail',
    route: 'execute-tool',
    input: {
      body: {
        message: {
          subject: 'Q-1042 quotation follow-up',
          body: { contentType: 'Text', content: 'Hi Adele, please find the revised quotation.' },
          toRecipients: [{ emailAddress: { address: 'adele.vance@contoso.com' } }],
        },
      },
    },
    graph: {},
    notes:
      'A blocked tool through execute-tool is indistinguishable from an unknown one: isError, {error:"Tool not found: send-mail", tip}. Graph is never called.',
  },
];

function stableJson(value: unknown): string {
  return JSON.stringify(value, null, 2) + '\n';
}

function decodeText(result: { content?: Array<{ type: string; text?: string }> }) {
  const first = result.content?.[0];
  if (first?.type !== 'text' || typeof first.text !== 'string') return undefined;
  try {
    return JSON.parse(first.text) as unknown;
  } catch {
    return undefined;
  }
}

/** A real Response for a mocked route. */
function respond(route: GraphRoute): Response {
  const text = route.text ?? (route.body === undefined ? '' : JSON.stringify(route.body));
  const headers = {
    ...(route.text === undefined
      ? { 'content-type': 'application/json; odata.metadata=minimal' }
      : {}),
    'request-id': '5f7a3c2e-8b1d-4e6a-9c0f-2d4b6a8e1c3f',
    ...route.headers,
  };
  return new Response(route.status === 204 ? null : text, {
    status: route.status,
    statusText: STATUS_TEXT[route.status] ?? '',
    headers,
  });
}

function mockAuthManager(): AuthManager {
  return {
    isOAuthModeEnabled: () => false,
    isMultiAccount: async () => false,
    isAppMode: () => false,
    getToken: async () => 'graph-access-token',
    getTokenForAccount: async () => 'graph-access-token',
    listAccounts: async () => [],
  } as unknown as AuthManager;
}

function buildMcpServer(options: Record<string, unknown>): McpServer {
  const authManager = mockAuthManager();
  const graphClient = new GraphClient(
    authManager,
    { clientId: '00000000-0000-0000-0000-000000000000', tenantId: 'common', cloudType: 'global' },
    'json',
    { orgMode: true }
  );
  const server = new MicrosoftGraphServer(authManager, options);
  Object.assign(server, { graphClient, version: 'contract' });
  return (server as unknown as { createMcpServer: () => McpServer }).createMcpServer();
}

async function connect(server: McpServer): Promise<Client> {
  const [serverSide, clientSide] = InMemoryTransport.createLinkedPair();
  await server.connect(serverSide);
  const client = new Client({ name: 'consumer-contract', version: '1' });
  await client.connect(clientSide);
  return client;
}

/** Every file the generator produces, keyed by path relative to CONTRACT_DIR. */
const generated = new Map<string, string>();

function checkOrWrite(relative: string, content: string): void {
  generated.set(relative, content);
  const file = path.join(CONTRACT_DIR, relative);
  if (UPDATE) {
    mkdirSync(path.dirname(file), { recursive: true });
    writeFileSync(file, content);
    return;
  }
  expect(existsSync(file), `${relative} is missing; run with UPDATE_CONTRACTS=1`).toBe(true);
  expect(readFileSync(file, 'utf8'), `${relative} drifted; run with UPDATE_CONTRACTS=1`).toBe(
    content
  );
}

describe('consumer contract', () => {
  let production: Client;
  let fullRegistration: Client;
  let fetchCalls: string[];
  let routes: Record<string, GraphRoute>;

  beforeAll(async () => {
    production = await connect(buildMcpServer(PRODUCTION_OPTIONS));
    // In production get-upload-url and add-mail-attachment are not eager, so they are not
    // in tools/list. Their schemas come from the same flags without --discovery.
    fullRegistration = await connect(
      buildMcpServer({ ...PRODUCTION_OPTIONS, discovery: false, eagerTools: undefined })
    );
  });

  afterAll(async () => {
    await production?.close();
    await fullRegistration?.close();
  });

  beforeEach(() => {
    vi.useFakeTimers({ toFake: ['Date'] });
    vi.setSystemTime(new Date(NOW));
    ticketRandom.counter = 0;
    __resetSharedBreakerForTests();
    const config = loadAttachmentUrlConfig(true, {
      MS365_MCP_ATTACHMENT_URL_BASE: ATTACHMENT_URL_BASE,
      MS365_MCP_ATTACHMENT_URL_KEY: ATTACHMENT_URL_KEY,
    })!;
    configureAttachmentMinting({ store: new AttachmentTicketStore(config.ttlSeconds), config });
    fetchCalls = [];
    routes = {};
    vi.stubGlobal(
      'fetch',
      vi.fn(async (url: string | URL, init?: Parameters<typeof fetch>[1]) => {
        const href = String(url);
        expect(href.startsWith(GRAPH), `unexpected fetch ${href}`).toBe(true);
        const key = `${(init?.method ?? 'GET').toUpperCase()} ${href.slice(GRAPH.length)}`;
        fetchCalls.push(key);
        const route = routes[key];
        expect(route, `no mocked Graph answer for ${key}`).toBeDefined();
        return respond(route);
      })
    );
  });

  afterEach(() => {
    vi.unstubAllGlobals();
    vi.useRealTimers();
    resetAttachmentMinting();
  });

  it.each(CASES.map((c) => [c.id, c] as const))('%s', async (_id, contractCase) => {
    routes = contractCase.graph;
    const call =
      contractCase.route === 'direct'
        ? { name: contractCase.tool, arguments: contractCase.input }
        : {
            name: 'execute-tool',
            arguments: { tool_name: contractCase.tool, parameters: contractCase.input },
          };
    const result = (await production.callTool(call)) as {
      content?: Array<{ type: string; text?: string }>;
    };
    // Every mocked Graph answer was used, and nothing else was asked for.
    expect([...fetchCalls].sort()).toEqual(Object.keys(contractCase.graph).sort());

    const decoded = decodeText(result);
    checkOrWrite(
      `results/${contractCase.id}.json`,
      stableJson({
        case: contractCase.id,
        tool: contractCase.tool,
        route: contractCase.route,
        input: contractCase.input,
        call,
        result,
        ...(decoded !== undefined && { decoded }),
        clock: { now: NOW, nowEpoch: NOW_EPOCH },
        notes: contractCase.notes,
      })
    );
  });

  it('records the advertised input schemas', async () => {
    const productionTools = (await production.listTools()).tools;
    const fullTools = (await fullRegistration.listTools()).tools;
    for (const name of SCHEMA_TOOLS) {
      const nonEager = NON_EAGER_SCHEMA_TOOLS.has(name);
      if (nonEager) {
        expect(
          productionTools.find((t) => t.name === name),
          `${name} is now eager in production; move it out of NON_EAGER_SCHEMA_TOOLS`
        ).toBeUndefined();
      }
      const tool = (nonEager ? fullTools : productionTools).find((t) => t.name === name);
      expect(
        tool,
        nonEager ? `${name} is not registered` : `${name} is not in the production tools/list`
      ).toBeDefined();
      checkOrWrite(`schemas/${name}.json`, stableJson(tool!.inputSchema));
    }
  });

  it('writes the manifest and leaves no stale files', () => {
    checkOrWrite(
      'manifest.json',
      stableJson({
        contractVersion: 1,
        generatedBy: 'test/consumer-contract.test.ts',
        cases: CASES.map((c) => c.id),
        schemas: SCHEMA_TOOLS,
      })
    );
    // Only meaningful after every case ran; a filtered run (-t) skips the sweep.
    if (generated.size !== CASES.length + SCHEMA_TOOLS.length + 1) return;
    for (const sub of ['results', 'schemas']) {
      const dir = path.join(CONTRACT_DIR, sub);
      for (const file of existsSync(dir) ? readdirSync(dir) : []) {
        const relative = `${sub}/${file}`;
        if (generated.has(relative)) continue;
        if (UPDATE) rmSync(path.join(dir, file));
        else expect.fail(`${relative} is not produced by any case; run with UPDATE_CONTRACTS=1`);
      }
    }
  });
});
