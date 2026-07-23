import { describe, it, expect } from 'vitest';
import { readFileSync } from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { UTILITY_TOOLS, SEND_MAIL_TOOL_NAMES } from '../graph-tools.js';

/**
 * Synchronization guard (round-3 f3): keep SEND_MAIL_TOOL_NAMES in step with the
 * registered tool surface so a future send-capable mail tool cannot silently
 * escape deriveSendOperationExposed (which would let the restricted profile
 * report sendOperationExposed=false while a send tool is exposed).
 *
 * Every registered tool whose name looks like a send/reply/forward operation is
 * classified as exactly one of: a draft-compose tool (create-*), a non-mail
 * messaging surface (Teams chat/channel, group threads, calendar events,
 * activity notifications), or a mail send tool that MUST be in
 * SEND_MAIL_TOOL_NAMES. A newly added mail send tool matches none of the
 * exclusions and therefore fails this test until it is classified.
 */

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const endpoints = JSON.parse(
  readFileSync(path.join(__dirname, '..', 'endpoints.json'), 'utf8')
) as Array<{ toolName: string }>;

const ALL_TOOL_NAMES = [...endpoints.map((e) => e.toolName), ...UTILITY_TOOLS.map((u) => u.name)];

const SEND_ISH = /(send|reply|forward)/i;
// Names that send/reply/forward on a NON-mail surface. A new mail send tool
// would not match any of these, so it falls through to the assertion below.
const NON_MAIL_SURFACE = /(chat|channel|group-thread|calendar-event|activity)/;

describe('send-operation synchronization guard (f3)', () => {
  it('classifies every send/reply/forward tool and lists none unaccounted for', () => {
    // Sanity: the set is not accidentally emptied.
    expect(SEND_MAIL_TOOL_NAMES.size).toBeGreaterThan(0);

    const escaped = ALL_TOOL_NAMES.filter((name) => {
      if (!SEND_ISH.test(name)) return false; // not a send/reply/forward tool
      if (name.startsWith('create-')) return false; // create-*-draft composes, does not send
      if (NON_MAIL_SURFACE.test(name)) return false; // non-mail messaging surface
      return !SEND_MAIL_TOOL_NAMES.has(name); // a mail send tool must be in the set
    });

    expect(escaped).toEqual([]);
  });

  it('every SEND_MAIL_TOOL_NAMES entry is a registered tool (no stale entries)', () => {
    const registered = new Set(ALL_TOOL_NAMES);
    const stale = [...SEND_MAIL_TOOL_NAMES].filter((name) => !registered.has(name));
    expect(stale).toEqual([]);
  });
});
