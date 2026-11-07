import { describe, expect, it } from 'vitest';
import { readFileSync } from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import yaml from 'js-yaml';
// @ts-expect-error - bin modules are plain ESM JavaScript with no type declarations
import { WRITABLE_NAVIGATION_PROPERTIES } from '../bin/modules/writable-request-bodies.mjs';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const repoRoot = path.join(__dirname, '..');

/**
 * EnviroKinetics/ms365-mcp#25, the direction the other guard cannot cover.
 *
 * test/writable-body-invariants.ts asserts that a writable property does not DISAPPEAR
 * from a request body. It cannot catch the opposite: a navigation property Microsoft
 * newly publishes is dropped silently, because the allowlist has no entry for something
 * that did not exist when the allowlist was written. That makes a spec refresh a place
 * where capability can quietly vanish.
 *
 * So every navigation property reachable from a request body must be classified, and a
 * refresh that introduces an unclassified one fails here and asks for a decision.
 */

type Classification = { keep: Record<string, string[]>; drop: Record<string, string[]> };

const classification: Classification = JSON.parse(
  readFileSync(path.join(repoRoot, 'nav-property-classification.json'), 'utf8')
);

/** Every (entity, navProperty) pair reachable from a POST/PATCH/PUT request body. */
function reachableNavPairs(): Array<[string, string]> {
  const spec = yaml.load(
    readFileSync(path.join(repoRoot, 'openapi', 'openapi-trimmed.yaml'), 'utf8')
  ) as {
    paths?: Record<string, Record<string, { requestBody?: unknown }>>;
    components?: { schemas?: Record<string, unknown> };
  };
  const schemas = spec.components?.schemas ?? {};

  const props = (sc: unknown): Record<string, unknown> => {
    const out: Record<string, unknown> = {};
    if (sc && typeof sc === 'object') {
      const rec = sc as Record<string, unknown>;
      if (rec.properties && typeof rec.properties === 'object') Object.assign(out, rec.properties);
      for (const branch of (rec.allOf as unknown[]) ?? []) Object.assign(out, props(branch));
    }
    return out;
  };

  const refsIn = (node: unknown, out = new Set<string>()): Set<string> => {
    if (!node || typeof node !== 'object') return out;
    if (Array.isArray(node)) {
      node.forEach((n) => refsIn(n, out));
      return out;
    }
    for (const [k, v] of Object.entries(node as Record<string, unknown>)) {
      if (k === '$ref' && typeof v === 'string' && v.startsWith('#/components/schemas/')) {
        out.add(v.slice('#/components/schemas/'.length));
      } else refsIn(v, out);
    }
    return out;
  };

  const seeds = new Set<string>();
  for (const pathItem of Object.values(spec.paths ?? {})) {
    for (const [method, op] of Object.entries(pathItem ?? {})) {
      if (!['post', 'patch', 'put'].includes(method)) continue;
      const schema = (op as Record<string, any>)?.requestBody?.content?.['application/json']
        ?.schema;
      if (!schema) continue;
      for (const ref of refsIn(schema)) seeds.add(ref.replace(/\.writable$/, ''));
    }
  }

  const reached = new Set<string>();
  const stack = [...seeds];
  while (stack.length) {
    const name = stack.pop()!;
    if (!name || reached.has(name) || name.endsWith('.writable')) continue;
    reached.add(name);
    for (const ref of refsIn(schemas[name] ?? {})) if (!reached.has(ref)) stack.push(ref);
  }

  const pairs: Array<[string, string]> = [];
  for (const name of [...reached].sort()) {
    for (const [prop, def] of Object.entries(props(schemas[name]))) {
      if (def && typeof def === 'object' && (def as any)['x-ms-navigationProperty'] === true) {
        pairs.push([name, prop]);
      }
    }
  }
  return pairs;
}

const classified = (entity: string, prop: string) =>
  (classification.keep[entity] ?? []).includes(prop)
    ? 'keep'
    : (classification.drop[entity] ?? []).includes(prop)
      ? 'drop'
      : undefined;

describe('every reachable navigation property is classified', () => {
  const pairs = reachableNavPairs();

  it('finds a non-trivial set to classify, so a broken walk cannot pass vacuously', () => {
    expect(pairs.length).toBeGreaterThan(100);
  });

  it('leaves nothing unclassified', () => {
    const unclassified = pairs.filter(([e, p]) => !classified(e, p));
    expect(
      unclassified,
      'Unclassified navigation properties. A Graph spec refresh introduced these, and ' +
        'they are currently being dropped from request bodies with no decision recorded. ' +
        'For each: either allowlist it in WRITABLE_NAVIGATION_PROPERTIES and add it under ' +
        '"keep" in nav-property-classification.json, or record it under "drop".\n' +
        unclassified.map(([e, p]) => `  ${e}.${p}`).join('\n')
    ).toEqual([]);
  });

  it('does not carry stale entries for properties Graph no longer publishes', () => {
    const live = new Set(pairs.map(([e, p]) => `${e}.${p}`));
    const stale: string[] = [];
    for (const bucket of ['keep', 'drop'] as const) {
      for (const [entity, props] of Object.entries(classification[bucket])) {
        for (const prop of props)
          if (!live.has(`${entity}.${prop}`)) stale.push(`${bucket}: ${entity}.${prop}`);
      }
    }
    expect(stale, 'Classification entries no longer reachable; remove them').toEqual([]);
  });
});

describe('the classification and the allowlist agree', () => {
  it('everything marked keep is actually allowlisted', () => {
    const notAllowlisted: string[] = [];
    for (const [entity, props] of Object.entries(classification.keep)) {
      const allowed: string[] = WRITABLE_NAVIGATION_PROPERTIES[entity] ?? [];
      for (const prop of props)
        if (!allowed.includes(prop)) notAllowlisted.push(`${entity}.${prop}`);
    }
    expect(
      notAllowlisted,
      'Marked "keep" but not in WRITABLE_NAVIGATION_PROPERTIES, so it is still being dropped'
    ).toEqual([]);
  });

  it('nothing marked drop is allowlisted', () => {
    const contradictions: string[] = [];
    for (const [entity, props] of Object.entries(classification.drop)) {
      const allowed: string[] = WRITABLE_NAVIGATION_PROPERTIES[entity] ?? [];
      for (const prop of props)
        if (allowed.includes(prop)) contradictions.push(`${entity}.${prop}`);
    }
    expect(
      contradictions,
      'Marked "drop" but allowlisted, so the record contradicts the code'
    ).toEqual([]);
  });
});
