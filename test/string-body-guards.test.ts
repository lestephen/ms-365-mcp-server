import { afterEach, describe, expect, it, vi } from 'vitest';
import { Client } from '@modelcontextprotocol/client';
import { InMemoryTransport, McpServer } from '@modelcontextprotocol/server';
import { registerDiscoveryTools, registerGraphTools } from '../src/graph-tools.js';
import type GraphClient from '../src/graph-client.js';

/**
 * EnviroKinetics/ms365-mcp#92. execute-tool does not run a tool's zod schema, so a
 * request body passed as a JSON-encoded string reached executeGraphTool intact and was
 * sent to Graph verbatim as application/json, where Graph decodes it. The payload guards
 * walked objects only, so a string body slipped past both: a graph-batch carrying
 * POST /me/sendMail (a model-callable send through the blocklist) and an attachment
 * carrying base64 contentBytes. These drive the real catalogue through the real MCP
 * dispatch with the production blocklist, and assert Graph is never called.
 */
const PRODUCTION_BLOCKLIST =
  '^((send|reply|reply-all|forward)-[a-z-]*(mail|draft)[a-z-]*|forward-calendar-event|share-drive-item|reply-to-group-thread|upload-file-content|upload-my-profile-photo|list-adhoc-call-transcripts|get-adhoc-call-transcript-content)$';

const clients: Client[] = [];

afterEach(async () => {
  await Promise.all(clients.splice(0).map((client) => client.close()));
});

async function connect(mode: 'discovery' | 'direct' = 'discovery') {
  const graphRequest = vi.fn().mockResolvedValue({ content: [{ type: 'text', text: '{}' }] });
  const graph = { graphRequest } as unknown as GraphClient;
  const server = new McpServer({ name: 'string-body-test', version: '1' });
  const options = { orgMode: true, blockedTools: PRODUCTION_BLOCKLIST };
  if (mode === 'direct') registerGraphTools(server, graph, options);
  else registerDiscoveryTools(server, graph, options);
  const [a, b] = InMemoryTransport.createLinkedPair();
  await server.connect(a);
  const client = new Client({ name: 'test-client', version: '1' });
  await client.connect(b);
  clients.push(client);
  return { client, graphRequest };
}

async function executeTool(client: Client, toolName: string, parameters: unknown) {
  const result = await client.callTool({
    name: 'execute-tool',
    arguments: { tool_name: toolName, parameters },
  });
  const text = (result.content as Array<{ text: string }>)[0].text;
  return { result, payload: JSON.parse(text) as Record<string, unknown> };
}

const sendMailBatch = {
  requests: [
    {
      id: '1',
      method: 'POST',
      url: '/me/sendMail',
      headers: { 'Content-Type': 'application/json' },
      body: { message: { subject: 'x', toRecipients: [] } },
    },
  ],
};

const fileAttachment = {
  '@odata.type': '#microsoft.graph.fileAttachment',
  name: 'report.pdf',
  contentBytes: 'JVBERi0xLjQK',
};

describe('a JSON-string request body is guarded like an object body (#92)', () => {
  it('refuses a graph-batch whose string body carries a blocked send', async () => {
    const { client, graphRequest } = await connect();
    const { result, payload } = await executeTool(client, 'graph-batch', {
      body: JSON.stringify(sendMailBatch),
    });
    expect(result.isError).toBe(true);
    expect(JSON.stringify(payload)).toMatch(/send-mail|blocked/i);
    expect(graphRequest).not.toHaveBeenCalled();
  });

  it('still refuses the same batch as an object body', async () => {
    const { client, graphRequest } = await connect();
    const { result } = await executeTool(client, 'graph-batch', { body: sendMailBatch });
    expect(result.isError).toBe(true);
    expect(graphRequest).not.toHaveBeenCalled();
  });

  it('refuses inline contentBytes inside a string attachment body', async () => {
    const { client, graphRequest } = await connect();
    const { result, payload } = await executeTool(client, 'add-mail-attachment', {
      'message-id': 'AAMkAGI2',
      body: JSON.stringify(fileAttachment),
    });
    expect(result.isError).toBe(true);
    expect(payload.error).toBe('inline_bytes_refused');
    expect(graphRequest).not.toHaveBeenCalled();
  });

  it('refuses inline contentBytes in a graph-batch subrequest whose own body is a string', async () => {
    const batch = {
      requests: [
        {
          id: '1',
          method: 'POST',
          url: '/me/messages/AAMkAGI2/attachments',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify(fileAttachment),
        },
      ],
    };
    for (const body of [batch, JSON.stringify(batch)]) {
      const { client, graphRequest } = await connect();
      const { result, payload } = await executeTool(client, 'graph-batch', { body });
      expect(result.isError).toBe(true);
      expect(payload.error).toBe('inline_bytes_refused');
      expect(graphRequest).not.toHaveBeenCalled();
    }
  });

  it.each([
    [
      'a base64 drive upload',
      'PUT',
      '/me/drive/root:/notes.txt:/content',
      'text/plain',
      'aGVsbG8=',
    ],
    [
      'a base64 body with a non-JSON content type',
      'POST',
      '/me/onenote/sections/1-abc/pages',
      'text/html',
      'PGh0bWw+PC9odG1sPg==',
    ],
    [
      'a base64 attachment under a JSON content type',
      'POST',
      '/me/messages/AAMkAGI2/attachments',
      'application/json',
      Buffer.from(JSON.stringify(fileAttachment)).toString('base64'),
    ],
  ])(
    'refuses a graph-batch subrequest carrying %s (#93)',
    async (_label, method, url, type, body) => {
      const batch = {
        requests: [
          { id: '1', method: 'GET', url: '/me' },
          { id: '2', method, url, headers: { 'Content-Type': type }, body },
        ],
      };
      for (const outer of [batch, JSON.stringify(batch)]) {
        const { client, graphRequest } = await connect();
        const { result, payload } = await executeTool(client, 'graph-batch', { body: outer });
        expect(result.isError).toBe(true);
        expect(payload.error).toBe('opaque_subrequest_body_refused');
        expect(payload.subrequests).toEqual(['2']);
        expect(graphRequest).not.toHaveBeenCalled();
      }
    }
  );

  it('still forwards object and JSON-string subrequest bodies as written', async () => {
    const event = { subject: 'Sync', start: { dateTime: '2026-12-01T10:00:00', timeZone: 'UTC' } };
    const batch = {
      requests: [
        {
          id: '1',
          method: 'POST',
          url: '/me/events',
          headers: { 'Content-Type': 'application/json' },
          body: event,
        },
        {
          id: '2',
          method: 'POST',
          url: '/me/events',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify(event),
        },
      ],
    };
    const { client, graphRequest } = await connect();
    const { result } = await executeTool(client, 'graph-batch', { body: batch });
    expect(result.isError).toBeFalsy();
    expect(graphRequest).toHaveBeenCalledTimes(1);
    const sent = JSON.parse(graphRequest.mock.calls[0][1].body as string);
    expect(sent.requests[0].body).toEqual(event);
    expect(sent.requests[1].body).toBe(JSON.stringify(event));
  });

  it('forwards a top-level non-JSON string body unchanged, as upstream does', async () => {
    const { client, graphRequest } = await connect();
    const { result } = await executeTool(client, 'create-draft-email', { body: 'not json' });
    expect(result.isError).toBeFalsy();
    expect(graphRequest).toHaveBeenCalledTimes(1);
    expect(graphRequest.mock.calls[0][1].body).toBe('not json');
  });

  it('sends a valid string body exactly as the equivalent object body', async () => {
    const draft = { subject: 'Hello', body: { contentType: 'Text', content: 'Hi' } };
    const asObject = await connect();
    await executeTool(asObject.client, 'create-draft-email', { body: draft });
    const asString = await connect();
    const { result } = await executeTool(asString.client, 'create-draft-email', {
      body: JSON.stringify(draft),
    });
    expect(result.isError).toBeFalsy();
    expect(asString.graphRequest).toHaveBeenCalledTimes(1);
    expect(asString.graphRequest.mock.calls[0]).toEqual(asObject.graphRequest.mock.calls[0]);
  });

  it('leaves a reference attachment in a string body allowed', async () => {
    const { client, graphRequest } = await connect();
    const { result } = await executeTool(client, 'add-mail-attachment', {
      'message-id': 'AAMkAGI2',
      body: JSON.stringify({
        '@odata.type': '#microsoft.graph.referenceAttachment',
        name: 'report.pdf',
        sourceUrl: 'https://contoso.sharepoint.com/report.pdf',
      }),
    });
    expect(result.isError).toBeFalsy();
    expect(graphRequest).toHaveBeenCalledTimes(1);
  });

  it('refuses a graph-batch whose subrequest keys are capitalized (#96)', async () => {
    const { client, graphRequest } = await connect();
    const { result } = await executeTool(client, 'graph-batch', {
      body: { requests: [{ id: '1', Method: 'POST', Url: '/me/sendMail', Body: { message: {} } }] },
    });
    expect(result.isError).toBe(true);
    expect(graphRequest).not.toHaveBeenCalled();
  });

  it('refuses a send reached through a mail folder path (#95)', async () => {
    const { client, graphRequest } = await connect();
    const { result } = await executeTool(client, 'graph-batch', {
      body: {
        requests: [{ id: '1', method: 'POST', url: '/me/mailFolders/drafts/messages/AAMk/send' }],
      },
    });
    expect(result.isError).toBe(true);
    expect(graphRequest).not.toHaveBeenCalled();
  });

  it('sends an allowed mixed-case batch to Graph with canonical keys', async () => {
    const { client, graphRequest } = await connect();
    const { result } = await executeTool(client, 'graph-batch', {
      body: { requests: [{ ID: '1', Method: 'GET', Url: '/me' }] },
    });
    expect(result.isError).toBeFalsy();
    const sent = JSON.parse((graphRequest.mock.calls[0][1] as { body: string }).body);
    expect(sent.requests[0]).toEqual({ id: '1', method: 'GET', url: '/me' });
  });
});
