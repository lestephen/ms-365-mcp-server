import { describe, expect, it } from 'vitest';
import { buildBlockedOperationMatchers, findBlockedSubrequests } from '../src/lib/batch-guard.js';

/**
 * EnviroKinetics/ms365-mcp#24, narrowed to the route that actually matters.
 *
 * --blocked-tools matches tool NAMES. graph-batch is POST /$batch with arbitrary
 * {method, url, body} subrequests, so `POST /me/sendMail` inside a batch reached Graph
 * even with every send tool unreachable by name, and graph-batch declares no scopes of
 * its own so --allowed-scopes was no backstop either. That left the drafts-only policy
 * enforceable against an honest model but not against an injected one, which matters
 * because these skills read customer mail.
 *
 * Rather than build general method/path authorization for every request, this checks
 * batch subrequests against the operations the blocklist already prohibits. endpoints.json
 * carries method and pathPattern per tool, so the blocked names give the matchers directly.
 */
const BLOCKED = '^(send-mail|send-shared-mailbox-mail|reply-mail-message|forward-mail-message)$';

describe('blocked operation matchers', () => {
  it('derives matchers from the blocked tool names', () => {
    const matchers = buildBlockedOperationMatchers(BLOCKED);
    expect(matchers.length).toBeGreaterThanOrEqual(4);
    const names = matchers.map((m) => m.toolName);
    expect(names).toContain('send-mail');
    expect(names).toContain('send-shared-mailbox-mail');
  });

  it('is empty when nothing is blocked, so the guard is inert by default', () => {
    expect(buildBlockedOperationMatchers(undefined)).toEqual([]);
    expect(buildBlockedOperationMatchers('')).toEqual([]);
  });
});

describe('findBlockedSubrequests', () => {
  const matchers = buildBlockedOperationMatchers(BLOCKED);
  const check = (requests: unknown) => findBlockedSubrequests({ requests }, matchers);

  it('catches the bypass: POST /me/sendMail inside a batch', () => {
    const hits = check([{ id: '1', method: 'POST', url: '/me/sendMail', body: {} }]);
    expect(hits).toHaveLength(1);
    expect(hits[0].toolName).toBe('send-mail');
    expect(hits[0].id).toBe('1');
  });

  it('catches a shared-mailbox send, where the path carries an id segment', () => {
    const hits = check([{ id: '1', method: 'POST', url: '/users/ab12-cd34/sendMail', body: {} }]);
    expect(hits.map((h) => h.toolName)).toContain('send-shared-mailbox-mail');
  });

  it('catches a reply on a message id', () => {
    const hits = check([{ id: '9', method: 'POST', url: '/me/messages/AAMkAGI1/reply', body: {} }]);
    expect(hits.map((h) => h.toolName)).toContain('reply-mail-message');
  });

  it('expands placeholders embedded in Excel function-style path segments', () => {
    const excelMatchers = buildBlockedOperationMatchers('^delete-excel-range$');
    const hits = findBlockedSubrequests(
      {
        requests: [
          {
            id: 'excel-1',
            method: 'POST',
            url: "/drives/d/items/i/workbook/worksheets/w/range(address='A1:B2')/delete",
          },
        ],
      },
      excelMatchers
    );

    expect(hits).toHaveLength(1);
    expect(hits[0]).toMatchObject({ id: 'excel-1', toolName: 'delete-excel-range' });
  });

  it('ignores the query string when matching', () => {
    const hits = check([{ id: '1', method: 'POST', url: '/me/sendMail?foo=bar' }]);
    expect(hits).toHaveLength(1);
  });

  it('matches the method case-insensitively but does not match a different method', () => {
    expect(check([{ id: '1', method: 'post', url: '/me/sendMail' }])).toHaveLength(1);
    // GET on the same path is not the blocked operation.
    expect(check([{ id: '1', method: 'GET', url: '/me/sendMail' }])).toHaveLength(0);
  });

  it('still matches when a client sends an absolute Graph URL', () => {
    const hits = check([
      { id: '1', method: 'POST', url: 'https://graph.microsoft.com/v1.0/me/sendMail' },
    ]);
    expect(hits).toHaveLength(1);
  });

  it('allows operations that are not blocked', () => {
    const hits = check([
      { id: '1', method: 'GET', url: '/me/messages?$top=5' },
      { id: '2', method: 'POST', url: '/me/messages', body: { subject: 'draft' } },
      { id: '3', method: 'DELETE', url: '/me/messages/AAMk/attachments/xyz' },
    ]);
    expect(hits).toEqual([]);
  });

  it('does not confuse a longer path that merely starts the same way', () => {
    // /me/messages/{id}/reply is blocked; a deeper path under it is a different operation.
    const hits = check([
      { id: '1', method: 'POST', url: '/me/messages/AAMk/reply/something/else' },
    ]);
    expect(hits).toEqual([]);
  });

  it('reports every offending subrequest, not just the first', () => {
    const hits = check([
      { id: '1', method: 'POST', url: '/me/sendMail' },
      { id: '2', method: 'GET', url: '/me/messages' },
      { id: '3', method: 'POST', url: '/me/messages/AAMk/forward' },
    ]);
    expect(hits.map((h) => h.id).sort()).toEqual(['1', '3']);
  });

  it('tolerates a malformed batch body without throwing', () => {
    for (const body of [undefined, null, {}, { requests: null }, { requests: [null, 'x', {}] }]) {
      expect(() => findBlockedSubrequests(body, matchers)).not.toThrow();
    }
  });

  it('is inert when no matchers are configured', () => {
    expect(
      findBlockedSubrequests({ requests: [{ method: 'POST', url: '/me/sendMail' }] }, [])
    ).toEqual([]);
  });
});
