import { describe, expect, it } from 'vitest';
import { describeInlineBytes, findInlineByteFields } from '../src/lib/inline-bytes-guard.js';

/**
 * EnviroKinetics bans inline base64 in Microsoft 365 writes at any size.
 *
 * The point of doing it on the PAYLOAD rather than the tool name: add-mail-attachment
 * carries two shapes on one name. A fileAttachment has contentBytes and is exactly what
 * the ban is about; a referenceAttachment has a sourceUrl and no bytes, and is the
 * documented fallback when an upload-session PUT is unavailable. Name-blocking killed
 * both.
 */
describe('inline byte detection', () => {
  it('finds contentBytes on a fileAttachment', () => {
    const hits = findInlineByteFields({
      '@odata.type': '#microsoft.graph.fileAttachment',
      name: 'report.pdf',
      contentBytes: 'JVBERi0xLjQK',
    });

    expect(hits).toHaveLength(1);
    expect(hits[0].path).toBe('contentBytes');
    expect(hits[0].encodedLength).toBe(12);
  });

  it('allows a reference attachment, which is the whole reason this is payload-keyed', () => {
    expect(
      findInlineByteFields({
        '@odata.type': '#microsoft.graph.referenceAttachment',
        name: 'report.pdf',
        sourceUrl: 'https://envirokinetics.sharepoint.com/x.pdf',
        providerType: 'oneDriveBusiness',
        permission: 'view',
      })
    ).toEqual([]);
  });

  it('allows an item attachment and metadata-only writes', () => {
    expect(findInlineByteFields({ '@odata.type': '#microsoft.graph.itemAttachment' })).toEqual([]);
    expect(findInlineByteFields({ subject: 'hello', body: { content: 'hi' } })).toEqual([]);
  });

  it('treats an empty contentBytes as nothing to refuse', () => {
    expect(findInlineByteFields({ contentBytes: '' })).toEqual([]);
  });

  it('catches bytes smuggled inside a graph-batch subrequest', () => {
    const hits = findInlineByteFields({
      requests: [
        { id: '1', method: 'GET', url: '/me/messages' },
        {
          id: '2',
          method: 'POST',
          url: '/me/messages/AAA/attachments',
          body: { '@odata.type': '#microsoft.graph.fileAttachment', contentBytes: 'QUJD' },
        },
      ],
    });

    expect(hits).toHaveLength(1);
    expect(hits[0].path).toBe('requests[1].body.contentBytes');
  });

  it('is case-insensitive, since Graph accepts either casing', () => {
    expect(findInlineByteFields({ ContentBytes: 'QUJD' })).toHaveLength(1);
    expect(findInlineByteFields({ contentbytes: 'QUJD' })).toHaveLength(1);
  });

  it('is depth-bounded so a hostile shape cannot exhaust the stack', () => {
    // A guard that throws fails open, which is worse than one that misses a pathological
    // nesting no real Graph body has.
    let deep: Record<string, unknown> = { contentBytes: 'QUJD' };
    for (let i = 0; i < 200; i++) deep = { nested: deep };

    expect(() => findInlineByteFields(deep)).not.toThrow();
  });

  it('reports where the bytes were without echoing them', () => {
    const summary = describeInlineBytes(findInlineByteFields({ contentBytes: 'SECRETBASE64' }));

    expect(summary).toBe('contentBytes (12 chars)');
    expect(summary).not.toContain('SECRETBASE64');
  });
});
