import type { Transport } from '@modelcontextprotocol/sdk/shared/transport.js';
import type { JSONRPCMessage } from '@modelcontextprotocol/sdk/types.js';
import logger from '../logger.js';

/**
 * Make emitted tool schemas acceptable to strict OpenAI-compatible providers.
 *
 * The MCP TypeScript SDK converts our zod tool schemas with zod-to-json-schema
 * but does not pass `$refStrategy` (see
 * `@modelcontextprotocol/sdk/server/zod-json-schema-compat.js`), so the library
 * default of 'root' applies. Any zod schema instance the converter has already
 * seen becomes a `$ref` to a JSON Pointer at the root of that tool's schema, for
 * example:
 *
 *     "$ref": "#/properties/body/properties/from"
 *     "$ref": "#/properties/body/properties/allowedRoles/items/anyOf/0"
 *
 * Those pointers are valid JSON Schema, and Anthropic's API accepts them, but
 * several strict providers reject the whole request. Moonshot (Kimi) answers with:
 *
 *     tools.function.parameters is not a valid moonshot flavored json schema,
 *     details: <At path '...$ref': references must start with #/$defs/>
 *
 * So every reference is rewritten to point under `#/$defs/`. The referenced
 * subschema is copied into `$defs` and the inline original is left in place, so
 * the schema keeps validating exactly as before. Self-referential pointers (Graph
 * entities such as mailFolder nest their own type) become recursive `$defs`
 * references, which is legal JSON Schema and which Moonshot accepts.
 *
 * The real fix belongs in the SDK, which should expose `$refStrategy`. Until it
 * does, this runs on the way out so it cannot change how tool arguments are
 * validated on the way in.
 */

type JsonRecord = Record<string, unknown>;

function isRecord(value: unknown): value is JsonRecord {
  return typeof value === 'object' && value !== null && !Array.isArray(value);
}

function visit(node: unknown, fn: (node: JsonRecord) => void): void {
  if (!node || typeof node !== 'object') return;
  if (Array.isArray(node)) {
    for (const item of node) visit(item, fn);
    return;
  }
  fn(node as JsonRecord);
  for (const value of Object.values(node as JsonRecord)) visit(value, fn);
}

function collectRefs(schema: unknown): string[] {
  const refs = new Set<string>();
  visit(schema, (node) => {
    if (typeof node.$ref === 'string') refs.add(node.$ref);
  });
  return [...refs];
}

/**
 * Turn a JSON Pointer into a readable, stable `$defs` key. The `properties` and
 * `items` hops carry no meaning for a human reader, so they are dropped:
 * `#/properties/body/properties/from` becomes `body_from`.
 */
function defNameFor(pointer: string): string {
  const name = pointer
    .slice(2)
    .split('/')
    .filter((segment) => segment !== 'properties')
    .map((segment) =>
      segment
        .replace(/~1/g, '/')
        .replace(/~0/g, '~')
        .replace(/[^A-Za-z0-9_]/g, '_')
    )
    .join('_');
  return name || 'root';
}

function unescapeSegment(segment: string): string {
  return segment.replace(/~1/g, '/').replace(/~0/g, '~');
}

/**
 * Resolve a root-relative JSON Pointer, following any `$ref` met part way along
 * the path. Returns undefined for a pointer that cannot be resolved, including
 * one that loops, rather than throwing: a tool with one odd pointer should still
 * be listed.
 */
function resolvePointer(root: unknown, pointer: string, seen = new Set<string>()): unknown {
  if (seen.has(pointer)) return undefined;
  seen.add(pointer);

  let current: unknown = root;
  for (const rawSegment of pointer.slice(2).split('/')) {
    if (isRecord(current) && typeof current.$ref === 'string') {
      current = resolvePointer(root, current.$ref, seen);
    }
    if (!current || typeof current !== 'object') return undefined;

    const segment = unescapeSegment(rawSegment);
    if (Array.isArray(current)) {
      const index = Number(segment);
      if (!Number.isInteger(index) || index < 0 || index >= current.length) return undefined;
      current = current[index];
    } else {
      const record = current as JsonRecord;
      if (!(segment in record)) return undefined;
      current = record[segment];
    }
  }
  return current;
}

/**
 * Rewrite every root-relative `$ref` in a JSON Schema to point under `#/$defs/`.
 * Returns the input unchanged when there is nothing to rewrite, and never
 * mutates the input.
 */
export function hoistRefsUnderDefs(schema: unknown): unknown {
  if (!isRecord(schema)) return schema;

  const pointers = collectRefs(schema)
    .filter((ref) => ref.startsWith('#/') && !ref.startsWith('#/$defs/'))
    .sort();
  if (pointers.length === 0) return schema;

  const root = structuredClone(schema) as JsonRecord;
  const existingDefs = isRecord(root.$defs) ? root.$defs : undefined;

  const names = new Map<string, string>();
  const taken = new Set<string>(existingDefs ? Object.keys(existingDefs) : []);
  for (const pointer of pointers) {
    const base = defNameFor(pointer);
    let name = base;
    let suffix = 2;
    while (taken.has(name)) name = `${base}_${suffix++}`;
    taken.add(name);
    names.set(pointer, name);
  }

  // Resolve every target against the untouched clone before any rewriting, so a
  // pointer never resolves through a reference this pass has already replaced.
  const hoisted: JsonRecord = {};
  for (const pointer of pointers) {
    const target = resolvePointer(root, pointer);
    if (target === undefined) {
      // Should not happen: every pointer the converter emits addresses a subschema it
      // just wrote, and no unresolvable pointer occurs across the current tool set.
      // If the converter ever changes, degrade to a permissive definition rather than
      // dropping the tool, but say so loudly: the emitted schema is then weaker than
      // the zod schema that still validates the arguments, and a caller could be led
      // to send something the server rejects.
      logger.warn(
        `Tool schema: could not resolve $ref ${pointer}; emitting a permissive {} for it. ` +
          'The advertised schema is now looser than the validated one.'
      );
      hoisted[names.get(pointer)!] = {};
      continue;
    }
    hoisted[names.get(pointer)!] = structuredClone(target);
  }

  const rewrite = (node: unknown) =>
    visit(node, (candidate) => {
      if (typeof candidate.$ref === 'string' && names.has(candidate.$ref)) {
        candidate.$ref = `#/$defs/${names.get(candidate.$ref)}`;
      }
    });
  rewrite(root);
  for (const definition of Object.values(hoisted)) rewrite(definition);

  root.$defs = { ...(existingDefs ?? {}), ...hoisted };
  return root;
}

/**
 * Rewrite the schemas in a `tools/list` result in place. Any other message, and
 * any malformed entry, is left exactly as it was.
 */
export function applyStrictToolSchemas(message: unknown): void {
  if (!isRecord(message)) return;
  const result = message.result;
  if (!isRecord(result) || !Array.isArray(result.tools)) return;

  for (const tool of result.tools) {
    if (!isRecord(tool)) continue;
    if (isRecord(tool.inputSchema)) tool.inputSchema = hoistRefsUnderDefs(tool.inputSchema);
    if (isRecord(tool.outputSchema)) tool.outputSchema = hoistRefsUnderDefs(tool.outputSchema);
  }
}

/**
 * Decorate a transport so every outgoing `tools/list` result carries strict
 * schemas. Applied at the transport boundary because the SDK builds tool schemas
 * inside its own request handler and offers no conversion hook. Returns the same
 * instance so callers keep their reference.
 */
export function withStrictToolSchemas<T extends Transport>(transport: T): T {
  const send = transport.send.bind(transport);
  transport.send = (message: JSONRPCMessage, options?: Parameters<Transport['send']>[1]) => {
    applyStrictToolSchemas(message);
    return send(message, options);
  };
  return transport;
}
