import { describe, expect, it } from 'vitest';
import {
  buildScopesFromEndpoints,
  resolveAuthScopes,
  resolveAuthorizeScopes,
} from '../src/auth.js';

/**
 * EnviroKinetics/ms365-mcp#24, the cheaper half.
 *
 * `--blocked-tools` removes the tools but the token's scopes are derived from the
 * endpoint catalogue, so blocking send-mail still left Mail.Send on the token.
 * `graph-batch` is a passthrough that declares no scopes of its own, so a blocked
 * operation stayed fully authorized: `POST /me/sendMail` in a `$batch` subrequest
 * would succeed. Subtracting blocked tools from scope derivation does not close the
 * passthrough (that is the other half of #24) but it does mean the token cannot
 * perform the operation the operator prohibited.
 */
describe('blocked tools do not contribute scopes', () => {
  const SENDS = [
    'send-mail',
    'send-draft-message',
    'send-shared-mailbox-mail',
    'reply-mail-message',
    'reply-all-mail-message',
    'forward-mail-message',
  ];
  const BLOCKED = `^(${SENDS.join('|')})$`;

  it('drops Mail.Send when every send tool is blocked', () => {
    const withSends = buildScopesFromEndpoints(true, undefined, false);
    const withoutSends = buildScopesFromEndpoints(true, undefined, false, BLOCKED);

    expect(withSends).toContain('Mail.Send');
    expect(withoutSends).not.toContain('Mail.Send');
  });

  it('keeps the scopes the remaining tools still need', () => {
    const scopes = buildScopesFromEndpoints(true, undefined, false, BLOCKED);
    // Drafting still has to work: create-draft-email needs mailbox write access.
    expect(scopes).toContain('Mail.ReadWrite');
    // Unrelated surfaces must be untouched.
    expect(scopes.some((s) => s.startsWith('Files.'))).toBe(true);
  });

  it('is a no-op when nothing is blocked', () => {
    const a = buildScopesFromEndpoints(true, undefined, false);
    const b = buildScopesFromEndpoints(true, undefined, false, undefined);
    expect(b).toEqual(a);
  });

  it('applies on the allowed-scopes path too', () => {
    // resolveAuthScopes is the branch taken when --allowed-scopes is set, and it
    // derives tool scopes separately, so it needs the same subtraction.
    const withSends = resolveAuthScopes({
      orgMode: true,
      allowedScopes: 'Mail.Send Mail.ReadWrite',
    });
    const withoutSends = resolveAuthScopes({
      orgMode: true,
      allowedScopes: 'Mail.Send Mail.ReadWrite',
      blockedTools: BLOCKED,
    });

    expect(withSends).toContain('Mail.Send');
    expect(withoutSends).not.toContain('Mail.Send');
  });

  it('ignores an invalid blocklist pattern rather than dropping every scope', () => {
    // Failing closed here would request no scopes at all and brick sign-in. The
    // guardrail that must fail closed is registration, which compileBlockedToolsRegex
    // already makes fatal at startup, so by the time scopes are built the pattern has
    // been validated.
    const scopes = buildScopesFromEndpoints(true, undefined, false, '([bad');
    expect(scopes.length).toBeGreaterThan(0);
    expect(scopes).toContain('Mail.Send');
  });
});

/**
 * The authorize route is the third branch, and it was the hole.
 *
 * Scope derivation subtracts blocked tools, but /authorize only reached that
 * derivation when the client sent no `scope` parameter. A client that asked for
 * `scope=Mail.Send` had its list forwarded to Entra verbatim, so the drafts-only
 * policy was enforced on the MCP tool surface while the bearer token the client
 * walked away with could still send mail directly against Graph.
 */
describe('authorize-route scopes cannot exceed what the tool surface permits', () => {
  const BLOCKED =
    '^(send-mail|send-draft-message|send-shared-mailbox-mail|reply-mail-message|reply-all-mail-message|forward-mail-message)$';

  it('drops a client-requested scope the blocklist prohibits', () => {
    const scopes = resolveAuthorizeScopes(
      { orgMode: true, blockedTools: BLOCKED },
      'Mail.Send Mail.ReadWrite'
    );

    expect(scopes).not.toContain('Mail.Send');
    expect(scopes).toContain('Mail.ReadWrite');
  });

  it('leaves a client scope the blocklist does not prohibit alone', () => {
    // The filter subtracts blocked-tool scopes only. Upstream deliberately lets a client
    // choose its own scopes when the operator has set no --allowed-scopes, and
    // test/allowed-scopes.test.ts pins that; over-rejecting here would break it.
    const scopes = resolveAuthorizeScopes(
      { orgMode: true, blockedTools: BLOCKED },
      'Calendars.Read Mail.Read'
    );

    expect(scopes).toContain('Calendars.Read');
    expect(scopes).toContain('Mail.Read');
    expect(scopes).not.toContain('Mail.Send');
  });

  it('falls back to the derived scopes when the client requests none', () => {
    const derived = buildScopesFromEndpoints(true, undefined, false, BLOCKED);

    expect(resolveAuthorizeScopes({ orgMode: true, blockedTools: BLOCKED })).toEqual(derived);
    expect(resolveAuthorizeScopes({ orgMode: true, blockedTools: BLOCKED }, '')).toEqual(derived);
  });

  it('does not request an empty scope list when every requested scope is refused', () => {
    // Requesting nothing would brick sign-in, so a fully-refused request falls back to
    // what this configuration permits rather than to [].
    const scopes = resolveAuthorizeScopes({ orgMode: true, blockedTools: BLOCKED }, 'Mail.Send');

    expect(scopes).not.toContain('Mail.Send');
    expect(scopes.length).toBeGreaterThan(0);
  });

  it('constrains the client on the explicit allowed-scopes path too', () => {
    const scopes = resolveAuthorizeScopes(
      { orgMode: true, allowedScopes: 'Mail.ReadWrite', blockedTools: BLOCKED },
      'Mail.Send Mail.ReadWrite'
    );

    expect(scopes).not.toContain('Mail.Send');
    expect(scopes).toContain('Mail.ReadWrite');
  });

  /**
   * A first version of this filter compared the client's raw string against bare
   * catalogue names with an exact, case-sensitive Set.has(). Entra matches permission
   * names case-insensitively and accepts the fully qualified resource form, so every
   * spelling below reached the token while looking blocked.
   */
  describe('spellings of a blocked scope that must not slip through', () => {
    // Normalise the way Entra parses a scope: the permission is whatever follows the
    // LAST '/', so this sees through a resource URI, a doubled URI, a double slash and
    // the resource app-id GUID form alike. Deliberately not the same expression the
    // implementation uses; a helper that shares the implementation's blind spot cannot
    // detect a bypass, which is exactly what happened on the first pass here.
    const sends = (clientScope: string) =>
      resolveAuthorizeScopes({ orgMode: true, blockedTools: BLOCKED }, clientScope).map((s) =>
        s.trim().toLowerCase().split('/').pop()
      );

    it.each([
      ['lowercased', 'mail.send'],
      ['uppercased', 'MAIL.SEND'],
      ['fully qualified', 'https://graph.microsoft.com/Mail.Send'],
      ['fully qualified and lowercased', 'https://graph.microsoft.com/mail.send'],
      ['sovereign cloud host', 'https://graph.microsoft.us/Mail.Send'],
    ])('refuses a %s blocked scope', (_label, clientScope) => {
      expect(sends(clientScope)).not.toContain('mail.send');
    });

    it('refuses .default, which would grant every consented permission', () => {
      // .default is catalogue-independent: it yields everything statically consented on
      // the app registration, which includes every blocked scope at once.
      expect(sends('.default')).not.toContain('.default');
      expect(sends('https://graph.microsoft.com/.default')).not.toContain('.default');
    });

    it('still allows a fully qualified scope that is not blocked', () => {
      const scopes = resolveAuthorizeScopes(
        { orgMode: true, blockedTools: BLOCKED },
        'https://graph.microsoft.com/Calendars.Read'
      );

      expect(scopes).toContain('https://graph.microsoft.com/Calendars.Read');
    });

    // Round two of review found the first canonicalisation still had four holes. Each
    // of these was verified GRANTED against the previous implementation.
    it.each([
      ['double slash after the host', 'https://graph.microsoft.com//Mail.Send'],
      ['resource app-id GUID form', '00000003-0000-0000-c000-000000000000/Mail.Send'],
      ['doubled resource URI', 'https://graph.microsoft.com/https://graph.microsoft.com/Mail.Send'],
    ])('refuses the %s', (_label, clientScope) => {
      expect(sends(clientScope)).not.toContain('mail.send');
    });

    it('refuses .default behind a resource app-id GUID', () => {
      expect(sends('00000003-0000-0000-c000-000000000000/.default')).not.toContain('.default');
    });

    it('leaves .default alone when the operator blocked nothing', () => {
      // With no blocklist there is nothing to protect, and upstream lets the client
      // choose. Refusing here would change behaviour for every unblocked deployment.
      expect(resolveAuthorizeScopes({ orgMode: true }, '.default')).toContain('.default');
    });
  });

  it.each([
    ['a blocked scope', 'Mail.Send/'],
    ['.default', '.default/'],
  ])('refuses %s with a trailing slash', (_label, clientScope) => {
    // Canonicalising on the last '/' turns a trailing slash into an empty permission
    // name, which matched nothing prohibited and was granted verbatim.
    expect(
      resolveAuthorizeScopes({ orgMode: true, blockedTools: BLOCKED }, clientScope)
    ).not.toContain(clientScope);
  });

  /**
   * The prohibited set is a set difference, and both sides have to be expanded and
   * derived from a comparable baseline or the difference silently comes out empty.
   */
  describe('the prohibited set is derived soundly', () => {
    it.each([['Mail.Read'], ['Files.Read'], ['Calendars.Read']])(
      'refuses %s when every tool is blocked',
      (clientScope) => {
        // buildScopesFromEndpoints emits the collapsed higher form (Mail.ReadWrite) and
        // never the bare lower one, so an unexpanded unblocked side never marked
        // Mail.Read prohibited even with the whole catalogue blocked.
        expect(
          resolveAuthorizeScopes({ orgMode: true, blockedTools: '.*' }, clientScope)
        ).not.toContain(clientScope);
      }
    );

    it('does not let --read-only weaken the blocklist', () => {
      // Read-only strips write scopes from BOTH sides of the difference, cancelling the
      // delta: Mail.Send was refused with the blocklist alone and granted once
      // --read-only was added. A tightening flag must never loosen the filter.
      const withBlocklist = { orgMode: true, blockedTools: BLOCKED };

      expect(resolveAuthorizeScopes(withBlocklist, 'Mail.Send')).not.toContain('Mail.Send');
      expect(
        resolveAuthorizeScopes({ ...withBlocklist, readOnly: true }, 'Mail.Send')
      ).not.toContain('Mail.Send');
    });
  });

  /**
   * Case normalisation must not depend on the scope being in our endpoint catalogue.
   * Files.ReadWrite.All is a real delegated permission that endpoints.json does not
   * carry, so re-casing via a catalogue lookup misses it: the correctly-cased form was
   * refused while the lowercased form was granted, because the hierarchy suffix rules
   * are case-sensitive and never fired.
   */
  describe('super-scopes outside the endpoint catalogue', () => {
    const everything = { orgMode: true, blockedTools: '.*' };

    it.each([
      ['canonical casing', 'Files.ReadWrite.All'],
      ['lowercased', 'files.readwrite.all'],
      ['uppercased', 'FILES.READWRITE.ALL'],
    ])('refuses a %s super-scope of a prohibited permission', (_label, clientScope) => {
      expect(resolveAuthorizeScopes(everything, clientScope)).not.toContain(clientScope);
    });
  });
});
