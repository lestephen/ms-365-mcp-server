/**
 * Refuse request bodies that carry file bytes inline.
 *
 * EnviroKinetics bans inline base64 in Microsoft 365 writes at any size. The reason is
 * not efficiency: the blob lands in the agent transcript, where it has caused
 * corruption, truncation, and automated misuse misclassification that can terminate a
 * Claude thread. The supported path is an upload session, where the caller PUTs from
 * disk and the bytes never enter the model.
 *
 * Why this exists as well as `--blocked-tools`:
 *
 * `add-mail-attachment` has TWO payload shapes on one tool name. A `fileAttachment`
 * carries `contentBytes` and is exactly what the ban is about. A `referenceAttachment`
 * carries a `sourceUrl` and no bytes at all, and is the documented fallback when an
 * upload-session PUT is unavailable. Blocking the NAME killed both, which over-blocks
 * beyond the policy and removes a compliant route. Blocking the PAYLOAD is what the
 * policy actually says, and it is the same correction already made for batch
 * subrequests: check the operation, not the label.
 *
 * Scans recursively, so a `contentBytes` smuggled inside a `graph-batch` subrequest is
 * caught by the same pass. Callers hand it the NORMALIZED body for the same reason the
 * batch guard does: a passthrough client may flatten the request shape.
 *
 * LIMITATION: keyed on Graph's `contentBytes` field. Tools whose entire body is the
 * base64 (`requestFormat: 'binary'`, i.e. upload-file-content and
 * upload-my-profile-photo) have no non-inline shape at all, so those stay name-blocked
 * in the deployment manifest; there is nothing for a payload check to distinguish.
 */

/** A place in the request body where inline bytes were found. */
export interface InlineBytesHit {
  /** Dotted path to the offending field, e.g. `requests[0].body.contentBytes`. */
  path: string;
  /** Length of the encoded string, which is what would have entered the transcript. */
  encodedLength: number;
}

const FIELD = 'contentBytes';
const MAX_DEPTH = 12;

/**
 * Every inline-bytes field in `body`. Empty when there are none, which is the answer for
 * a reference attachment, an item attachment, or any metadata-only write.
 */
export function findInlineByteFields(body: unknown): InlineBytesHit[] {
  const hits: InlineBytesHit[] = [];

  const walk = (node: unknown, path: string, depth: number): void => {
    // Depth-bounded: the body is caller-controlled and a deeply nested object would
    // otherwise let a hostile shape exhaust the stack inside the guard, which fails open.
    if (depth > MAX_DEPTH || node === null || typeof node !== 'object') return;

    if (Array.isArray(node)) {
      node.forEach((item, index) => walk(item, `${path}[${index}]`, depth + 1));
      return;
    }

    for (const [key, value] of Object.entries(node as Record<string, unknown>)) {
      const here = path ? `${path}.${key}` : key;
      // Case-insensitive: Graph accepts the property either way, and a guard that only
      // matched the documented casing would be trivially sidestepped.
      if (key.toLowerCase() === FIELD.toLowerCase()) {
        if (typeof value === 'string' && value.length > 0) {
          hits.push({ path: here, encodedLength: value.length });
        }
        continue;
      }
      walk(value, here, depth + 1);
    }
  };

  walk(body, '', 0);
  return hits;
}

/** Operator-facing summary naming where the bytes were, without echoing them. */
export function describeInlineBytes(hits: InlineBytesHit[]): string {
  return hits.map((h) => `${h.path} (${h.encodedLength} chars)`).join(', ');
}
