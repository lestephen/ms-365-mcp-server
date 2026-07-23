# get-shared-draft-capability: read-only shared-draft readiness

This documents the `get-shared-draft-capability` MCP operation and its internal
Exchange Online recipient-permission broker, added for
[EnviroKinetics/ms365-mcp#17](https://github.com/EnviroKinetics/ms365-mcp/issues/17).

## Problem

The autonomous `eki-doc-control` workflow must prove that the signed-in session
can prepare a final draft for the shared identity `doccontrol@envirokinetics.com`
BEFORE it performs earlier document-control mutations. That proof must be
read-only, current, bound to the authenticated session, and must not create,
update, delete, or send a mailbox item.

Two facts cannot be proven by the existing surface:

1. **Send As authorization.** "Send As" is an Exchange recipient ACL, assigned
   to a user or group independently of Microsoft 365 group membership. Microsoft
   Graph does not expose it; the documented read interface is the Exchange
   Online cmdlet `Get-EXORecipientPermission`. Group membership, ownership, and
   historical drafts are NOT equivalent evidence.
2. **Current draft-write authorization** for the same connector session,
   without invoking a write operation.

## Shape

`get-shared-draft-capability` is a hand-written utility tool (the same class as
`get-download-url`), not a generated Graph endpoint, because it composes token
claim inspection with an Exchange ACL read.

It returns the closed `eki.doc-control-shared-draft-capability/v1` proof consumed
by the EKI Document Control skill. Readiness is DERIVED, never assumed:

- **Mail.ReadWrite** is read from the `scp` claim of the connecting client's
  bearer token (`getSessionClaims` in `audit-log.ts`). No write call is made.
- **Send As** is read live through the `ExoTransport` broker
  (`exo-recipient-broker.ts`), requiring an explicit `Allow` + `SendAs` entry
  whose trustee resolves to the signed-in user's object id, on the exact shared
  recipient, within a freshness window. A `Deny` blocks it; a stale read denies.
- **operations** flags (createDraft, createReplyDraft, readDraft, readDraftMime,
  createAttachmentUploadSession, listAttachments, getDownloadUrl) are derived
  STRUCTURALLY from the set of tools actually registered on the running profile.
- **sendOperationExposed** is derived STRUCTURALLY: true if any send-mail tool is
  registered, or the session holds `Mail.Send` / `Mail.Send.Shared`. Both must be
  absent for a truthful `false`, which is why Document Control needs the
  restricted connector profile (no send scopes, no send tools). It is never
  hardcoded.

The proof's `proofSha256` is computed with a canonicalization byte-identical to
the Python consumer (`json.dumps(..., ensure_ascii=False, separators=(",",":"),
sort_keys=True)` then SHA-256), covered by a cross-language golden-vector test.

## sessionBindingSha256 (design point for owner confirmation)

`connector.sessionBindingSha256` is the SHA-256 of `tid|oid|sid` (auth session
id when present, else `tid|oid|jti`). It binds the proof to the signed-in
identity and auth session while deliberately excluding the raw bearer token and
its signature bytes. The consumer only requires a SHA-256, so the binding input
is a producer-side choice. If a stronger, per-request non-replayable binding is
wanted (e.g. mixing in a server nonce echoed by the consumer request), that is a
follow-up for the repository owner to confirm.

## Read-only by construction

The tool makes no Graph write call and the `ExoTransport` contract exposes only a
read method. The restricted Document Control profile is **send-denied**, not
globally `--read-only`: draft-create and attachment tools are POST writes to the
user's own mailbox and must stay registered for the proof's operation flags to be
true. Read-only-ness is a property of the probe operation itself, enforced and
tested; the profile removes only send scopes and send tools.

## Live boundary (not provisioned here)

The live `ExoTransport` is intentionally unconfigured in this build. Standing it
up requires, outside this repository:

- an Entra app with the `Office 365 Exchange Online -> Exchange.ManageAsApp`
  application permission and admin consent;
- a client certificate for that app; and
- an Exchange RBAC role restricted to recipient reads (e.g. "View-Only
  Recipients") assigned to the app's service principal.

With those, the transport runs `Get-EXORecipient` (to resolve the shared
identity and each SendAs trustee) and `Get-EXORecipientPermission` over the
Exchange Online REST admin endpoint, stamping `observedAt` at read time. Until
that provisioning and owner authorization exist, the tool reports that the
broker is unavailable rather than inferring Send As from Graph or group
membership. The exact Exchange REST/PowerShell response field names and the
`RecipientTypeDetails` value for `doccontrol@envirokinetics.com`
(SharedMailbox vs GroupMailbox vs a mail-enabled group) must be confirmed against
the live directory; both recipient-type branches are already handled.
