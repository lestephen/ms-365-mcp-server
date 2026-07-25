import { describe, expect, it } from 'vitest';
import { buildScopesFromEndpoints, resolveAuthScopes } from '../src/auth.js';

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
