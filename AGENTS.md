# Agent guidance for the EKI fork of ms-365-mcp-server

This is EnviroKinetics' tracking fork of `Softeria/ms-365-mcp-server`. `main` is the EKI
deployment branch: upstream merged in, plus a carried set of EKI changes. Releases are
immutable tags `v<upstream base>-eki.<n>` that identify the exact reviewed server commit on
`main`, the one the `EnviroKinetics/ms365-mcp` harness pins as its submodule gitlink. EKI tags are placed by hand; do not rely on upstream's semantic-release
workflow for them.
Upstream is merged with merge commits, not rebased or squashed, so upstream history stays an
ancestor.

## Working rules

- Keep the carry small and upstream-shaped. Generic fixes should also be offered upstream;
  EKI-only policy stays here.
- When merging upstream, adopt upstream's version of anything it now does as well as or
  better than the carry, and say so in the PR.
- Before pushing, run `npm run verify` (generate, lint, format check, typecheck, build,
  test). CI also fails when `npm run generate` changes tracked files, so commit what it
  regenerates.

## Review calibration

One internal deployment, users signed in with Entra, reached both directly and through a
Pomerium broker. In a PR that merges upstream, review the EKI resolutions and the carry,
plus any upstream change that touches a control listed below; do not review the rest of
upstream's own code.

Must never regress (treat a weakening as a blocker):

- The `--blocked-tools` policy on every path: registration, `execute-tool` discovery,
  `graph-batch` subrequests, and OAuth scope grants. The policy itself (which tools are
  blocked) lives in the harness's `deployment.yaml`: agents prepare mail drafts and a human
  sends them; mail send, reply, and forward, plus a few named share and forward tools, are
  blocked. Other outbound tools, such as Teams posting, are allowed by that policy.
- The inline file-bytes guard (`src/lib/inline-bytes-guard.ts`), which refuses inline
  `contentBytes` in request bodies, with only its narrow reviewed exception for
  create-custom-emoji. Tools whose whole body is the bytes (upload-file-content,
  upload-my-profile-photo) are outside that guard and stay name-blocked in the harness.
- Binary target canonicalization and bounded inline downloads.
- Path-parameter encoding (`src/lib/unencoded-path-params.ts`).
- The Prometheus metrics on their dedicated listener.

Fine to defer or accept as-is:

- Style and naming in code that upstream owns.
- Speculative hardening that needs a compromised host or tenant admin.
- Refactors of working carry without a concrete defect.

Proportionality: prefer a documented residual or a follow-up issue over growing a PR.
"Real" is not the bar; "worth fixing now" is.
