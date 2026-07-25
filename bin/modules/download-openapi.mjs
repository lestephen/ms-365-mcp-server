import fs from 'fs';
import path from 'path';
import { createHash } from 'crypto';
import { fileURLToPath } from 'url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const DEFAULT_REPO_ROOT = path.join(__dirname, '..', '..');
const PIN_FILE = 'openapi-pin.json';

/**
 * The Graph OpenAPI specs are pinned to an immutable upstream commit.
 *
 * Fetching refs/heads/master meant CI built against whatever Microsoft had published
 * that minute, while a developer checkout silently reused whatever copy was already on
 * disk. The same commit could produce different clients, and nothing recorded which spec
 * was used. The download also wrote `response.text()` straight out with no integrity
 * check, so a truncated body became a malformed spec that only surfaced later as a
 * generated client missing schemas. See EnviroKinetics/ms365-mcp#27.
 *
 * Refresh deliberately with `npm run generate -- --refresh-spec`, which fetches master,
 * reports what it got, and tells you to update the pin and review the client diff.
 */
export function readSpecPin(repoRoot = DEFAULT_REPO_ROOT) {
  const file = path.join(repoRoot, PIN_FILE);
  const pin = JSON.parse(fs.readFileSync(file, 'utf8'));
  if (!pin.repo || !pin.ref || !pin.specs) {
    throw new Error(`${PIN_FILE} is missing repo, ref or specs`);
  }
  return pin;
}

/** Raw URL for a pinned spec version. Commit-pinned, so it cannot move under us. */
export function specUrl(pin, version, ref = pin.ref) {
  const spec = pin.specs[version];
  if (!spec) throw new Error(`No pinned spec for version ${version}`);
  return `https://raw.githubusercontent.com/${pin.repo}/${ref}/${spec.path}`;
}

/**
 * Reject a body that does not match the pin. Size is checked first, because a truncated
 * read is the likelier failure and its message is the more useful one.
 */
export function verifyDownload(version, buffer, expected) {
  if (expected?.bytes !== undefined && buffer.length !== expected.bytes) {
    throw new Error(
      `Downloaded ${version} spec is ${buffer.length} bytes, expected ${expected.bytes}. ` +
        'A short read means a truncated or interrupted fetch; nothing has been written. ' +
        'Retry, or refresh the pin with --refresh-spec if upstream genuinely changed.'
    );
  }
  if (expected?.sha256) {
    const actual = createHash('sha256').update(buffer).digest('hex');
    if (actual !== expected.sha256) {
      throw new Error(
        `Downloaded ${version} spec sha256 ${actual} does not match the pinned ` +
          `${expected.sha256}. Refresh the pin with --refresh-spec if upstream changed, ` +
          'and review the generated client diff before committing.'
      );
    }
  }
}

async function fetchSpec(url) {
  const response = await fetch(url);
  if (!response.ok) {
    throw new Error(`Failed to download: ${response.status} ${response.statusText}`);
  }
  const buffer = Buffer.from(await response.arrayBuffer());
  // content-length describes the ENCODED body. fetch decompresses transparently, so on a
  // gzipped response (which is what raw.githubusercontent.com sends for these specs, ~2 MB
  // compressed against ~38 MB decoded) comparing it to the decoded length is a guaranteed
  // false positive. Only compare when the body arrived unencoded, where a mismatch really
  // does mean a short read. The pinned byte count and digest are the authoritative check.
  const encoding = response.headers.get('content-encoding');
  const declared = response.headers.get('content-length');
  if (!encoding && declared && Number(declared) !== buffer.length) {
    throw new Error(`Truncated download: got ${buffer.length} bytes, server declared ${declared}.`);
  }
  return buffer;
}

/**
 * Download one pinned spec version into `targetFile`, verifying before writing.
 * Returns true when a download happened, false when an existing file was reused.
 */
export async function downloadGraphOpenAPI(
  targetDir,
  targetFile,
  version,
  { repoRoot = DEFAULT_REPO_ROOT, refreshSpec = false, forceDownload = false } = {}
) {
  if (!fs.existsSync(targetDir)) {
    console.log(`Creating directory: ${targetDir}`);
    fs.mkdirSync(targetDir, { recursive: true });
  }

  const pin = readSpecPin(repoRoot);
  const expected = pin.specs[version];

  if (fs.existsSync(targetFile) && !forceDownload && !refreshSpec) {
    // An existing file still has to match the pin. Otherwise a stale local copy quietly
    // becomes the build input, which is the developer half of #27.
    verifyDownload(version, fs.readFileSync(targetFile), expected);
    console.log(`OpenAPI specification already exists and matches the pin: ${targetFile}`);
    return false;
  }

  const ref = refreshSpec ? 'refs/heads/master' : pin.ref;
  console.log(`Downloading ${version} OpenAPI specification from ${specUrl(pin, version, ref)}`);

  const buffer = await fetchSpec(specUrl(pin, version, ref));

  if (refreshSpec) {
    const sha256 = createHash('sha256').update(buffer).digest('hex');
    console.log(
      `   refreshed ${version}: ${buffer.length} bytes, sha256 ${sha256}\n` +
        `   update ${PIN_FILE} with these values and the new upstream ref before committing`
    );
  } else {
    verifyDownload(version, buffer, expected);
  }

  fs.writeFileSync(targetFile, buffer);
  console.log(`OpenAPI specification downloaded to ${targetFile}`);
  return true;
}
