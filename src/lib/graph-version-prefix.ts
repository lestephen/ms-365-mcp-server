/**
 * One definition of "a Microsoft Graph API version segment", shared by every guard that
 * has to canonicalize a Graph path before matching it.
 *
 * This exists because the definition was duplicated and the copies drifted.
 * `batch-guard` stripped the prefix only for ABSOLUTE subrequest URLs while
 * `message-signoff` stripped it unconditionally, so `POST /v1.0/me/sendMail` inside a
 * `$batch` slipped past the blocklist that `POST https://graph.microsoft.com/v1.0/me/sendMail`
 * and `POST /me/sendMail` both hit (EnviroKinetics/ms365-mcp#24).
 *
 * Graph exposes no top-level resource named `v1.0` or `beta`, so stripping it
 * unconditionally cannot shadow a real path, and erring toward stripping is the
 * fail-closed direction for a blocklist: it can only make MORE things match.
 */
export const GRAPH_VERSION_SEGMENT = /^(?:v1\.0|beta)$/i;

/** Drop a leading `v1.0`/`beta` segment from an already-split path. */
export function stripGraphVersionSegment(segments: string[]): string[] {
  return GRAPH_VERSION_SEGMENT.test(segments[0] ?? '') ? segments.slice(1) : segments;
}

/** Drop a leading `/v1.0` or `/beta` from a path string. */
export function stripGraphVersionPath(path: string): string {
  return path.replace(/^\/(?:v1\.0|beta)(?=\/)/i, '');
}
