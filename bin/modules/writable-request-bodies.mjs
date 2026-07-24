/**
 * Restrict POST/PATCH/PUT request bodies to properties a client can actually send.
 *
 * Microsoft's Graph OpenAPI metadata types a request body as the full entity, so
 * every read-only navigation property comes along for the ride. The effect on the
 * emitted MCP tool schemas is severe: `body` accounts for about two thirds of the
 * whole tool payload, and the worst offenders are pure noise, for example
 * `create-team-channel.body.filesFolder` (28 KB, server-assigned),
 * `create-sharepoint-list.body.drive` (26 KB, read-only), and a `createdByUser`
 * that expands to the entire Graph user entity on five separate tools.
 *
 * Beyond the size, it is actively misleading: the model is told it may set a
 * channel's `messages` when creating the channel.
 *
 * Graph marks these precisely with `x-ms-navigationProperty: true`, and does not
 * mark genuine inputs such as `displayName`, so this pass needs no name guessing.
 * A small allowlist covers the navigation properties Graph does accept on write.
 *
 * Response schemas are deliberately left alone. Request bodies are repointed at a
 * separate `.writable` variant rather than the shared entity being mutated, so
 * nothing that reads a Graph reply can regress.
 */

const WRITABLE_VARIANT_SUFFIX = '.writable';

/**
 * Navigation properties Graph accepts in a request body, keyed by the schema that
 * declares them. Entity-qualified on purpose: a global list of bare names would
 * also strike same-named properties on unrelated types.
 *
 * Keep this list conservative. Wrongly dropping a writable property silently
 * removes a capability, so anything uncertain belongs here until proven otherwise.
 */
const EXTENDED_PROPERTIES = ['singleValueExtendedProperties', 'multiValueExtendedProperties'];

const WRITABLE_NAVIGATION_PROPERTIES = {
  // A message, event or post can be created with its attachments inline. Extended
  // properties are writable on every Outlook type that declares them, and are how
  // callers stash their own sync metadata on an item.
  'microsoft.graph.message': ['attachments', ...EXTENDED_PROPERTIES],
  'microsoft.graph.event': ['attachments', ...EXTENDED_PROPERTIES],
  'microsoft.graph.todoTask': ['attachments', 'checklistItems', 'linkedResources', 'extensions'],
  'microsoft.graph.contact': EXTENDED_PROPERTIES,
  'microsoft.graph.contactFolder': EXTENDED_PROPERTIES,
  'microsoft.graph.mailFolder': EXTENDED_PROPERTIES,
  'microsoft.graph.calendar': EXTENDED_PROPERTIES,
  // POST /sites/{id}/lists/{id}/items requires `fields`.
  'microsoft.graph.listItem': ['fields'],
  // A list can be created with its column definitions.
  'microsoft.graph.list': ['columns'],
  // POST /chats and a private channel both require `members`.
  'microsoft.graph.chat': ['members'],
  'microsoft.graph.channel': ['members'],
  // Inline images in a Teams message are sent as hostedContents referenced from the
  // HTML body, so this is writable on create. `replies` is not: replies are posted
  // to their own endpoint.
  'microsoft.graph.chatMessage': ['hostedContents'],
  // An open extension can be created inline with the item that carries it.
  'microsoft.graph.post': ['attachments', 'extensions', ...EXTENDED_PROPERTIES],
  'microsoft.graph.todoTaskList': ['extensions'],
  // POST /teams accepts members. No create-team tool exists today, but keeping this
  // here means the silent-loss bug does not return if one is added.
  'microsoft.graph.team': ['members'],
  // Excel formatting is applied by PATCHing these onto the range format, so they
  // are the entire point of format-excel-range. Graph models them as navigation
  // properties even though they are writable.
  'microsoft.graph.workbookRangeFormat': ['borders', 'fill', 'font', 'protection'],
  'microsoft.graph.workbookWorksheet': ['protection'],
  'microsoft.graph.workbookChartAreaFormat': ['fill', 'font'],
  'microsoft.graph.workbookChartAxisFormat': ['font'],
  'microsoft.graph.workbookChartAxisTitleFormat': ['font'],
  'microsoft.graph.workbookChartDataLabelFormat': ['fill', 'font'],
  'microsoft.graph.workbookChartLegendFormat': ['fill', 'font'],
  'microsoft.graph.workbookChartPointFormat': ['fill'],
  'microsoft.graph.workbookChartSeriesFormat': ['fill'],
  'microsoft.graph.workbookChartTitleFormat': ['fill', 'font'],
};

const BODY_METHODS = new Set(['post', 'patch', 'put']);

function isRecord(value) {
  return value !== null && typeof value === 'object' && !Array.isArray(value);
}

function schemaNameFromRef(ref) {
  if (typeof ref !== 'string' || !ref.startsWith('#/components/schemas/')) return undefined;
  return ref.slice('#/components/schemas/'.length);
}

function isWritableVariantName(name) {
  return typeof name === 'string' && name.endsWith(WRITABLE_VARIANT_SUFFIX);
}

/**
 * Walk every request-body schema in the document.
 */
function forEachRequestBodySchema(spec, visit) {
  for (const pathItem of Object.values(spec.paths ?? {})) {
    if (!isRecord(pathItem)) continue;
    for (const [method, operation] of Object.entries(pathItem)) {
      if (!BODY_METHODS.has(method.toLowerCase()) || !isRecord(operation)) continue;
      const content = operation.requestBody?.content;
      if (!isRecord(content)) continue;
      for (const mediaType of Object.values(content)) {
        if (isRecord(mediaType) && isRecord(mediaType.schema)) visit(mediaType.schema, operation);
      }
    }
  }
}

/** Every `#/components/schemas/...` name referenced anywhere inside a node. */
function referencedSchemaNames(node, out = new Set()) {
  if (!node || typeof node !== 'object') return out;
  if (Array.isArray(node)) {
    for (const item of node) referencedSchemaNames(item, out);
    return out;
  }
  for (const [key, value] of Object.entries(node)) {
    if (key === '$ref') {
      const name = schemaNameFromRef(value);
      if (name) out.add(name);
      continue;
    }
    referencedSchemaNames(value, out);
  }
  return out;
}

/**
 * Names of the properties this schema would lose, ignoring anything reached
 * through a ref. Walks allOf/anyOf/oneOf so inherited properties are counted.
 */
function droppableNavigationProperties(schema, declaringName) {
  const allowed = new Set(WRITABLE_NAVIGATION_PROPERTIES[declaringName] ?? []);
  const dropped = [];
  const walk = (node) => {
    if (!isRecord(node)) return;
    for (const [property, definition] of Object.entries(node.properties ?? {})) {
      if (!isRecord(definition)) continue;
      if (definition['x-ms-navigationProperty'] !== true) continue;
      if (allowed.has(property)) continue;
      dropped.push(property);
    }
    for (const keyword of ['allOf', 'anyOf', 'oneOf']) {
      for (const branch of node[keyword] ?? []) walk(branch);
    }
  };
  walk(schema);
  return dropped;
}

/**
 * Schemas whose writable form would differ from the original, either because they
 * drop a navigation property themselves or because something they reference does.
 * Computed to a fixpoint over the reverse reference graph, so cycles terminate.
 *
 * Knowing this up front means a schema that would be copied unchanged, for example
 * a plain complex type such as driveRecipient, keeps its original name instead of
 * gaining a pointless `.writable` twin.
 */
function findAffectedSchemas(schemas) {
  const referencedBy = new Map();
  const affected = new Set();
  const queue = [];

  for (const [name, schema] of Object.entries(schemas)) {
    if (isWritableVariantName(name)) continue;
    for (const target of referencedSchemaNames(schema)) {
      if (!referencedBy.has(target)) referencedBy.set(target, new Set());
      referencedBy.get(target).add(name);
    }
    if (droppableNavigationProperties(schema, name).length > 0) {
      affected.add(name);
      queue.push(name);
    }
  }

  while (queue.length > 0) {
    const current = queue.pop();
    for (const parent of referencedBy.get(current) ?? []) {
      if (affected.has(parent)) continue;
      affected.add(parent);
      queue.push(parent);
    }
  }
  return affected;
}

export function restrictRequestBodiesToWritableProperties(spec) {
  const schemas = spec?.components?.schemas;
  const report = { variantsCreated: 0, removed: [] };
  if (!isRecord(schemas) || !isRecord(spec.paths)) return report;

  const affected = findAffectedSchemas(schemas);

  // name of the source schema -> name of its writable variant
  const variants = new Map();

  /**
   * Return the ref a request body should use for `name`. Schemas that would be
   * copied unchanged keep their original ref. The variant is registered before its
   * body is filled in, so a self-referential entity such as mailFolder.childFolders
   * terminates instead of recursing.
   */
  function writableVariantRef(name) {
    if (isWritableVariantName(name)) return `#/components/schemas/${name}`;
    if (variants.has(name)) return `#/components/schemas/${variants.get(name)}`;

    const source = schemas[name];
    // Nothing in this schema's subgraph changes, so reuse it as published.
    if (!isRecord(source) || !affected.has(name)) return `#/components/schemas/${name}`;

    const variantName = `${name}${WRITABLE_VARIANT_SUFFIX}`;
    variants.set(name, variantName);
    schemas[variantName] = {};
    report.variantsCreated += 1;

    const clone = structuredClone(source);
    stripNavigationProperties(clone, name);
    repointRefs(clone);
    schemas[variantName] = clone;
    return `#/components/schemas/${variantName}`;
  }

  /**
   * Remove navigation properties from a cloned schema, honouring the allowlist for
   * the schema that declares them. Descends through allOf/anyOf/oneOf so inherited
   * properties are covered too.
   */
  function stripNavigationProperties(node, declaringSchema) {
    if (!isRecord(node)) return;
    const allowed = new Set(WRITABLE_NAVIGATION_PROPERTIES[declaringSchema] ?? []);

    if (isRecord(node.properties)) {
      const deleted = new Set();
      for (const [property, definition] of Object.entries(node.properties)) {
        if (!isRecord(definition)) continue;
        if (definition['x-ms-navigationProperty'] !== true) continue;
        if (allowed.has(property)) continue;
        delete node.properties[property];
        deleted.add(property);
        report.removed.push({ schema: declaringSchema, property });
      }
      // Filter `required` by what this pass actually deleted, not by membership in
      // node.properties: Graph declares some required names in an allOf branch other
      // than the one carrying `properties`, and dropping those would weaken
      // validation for a body this pass did not otherwise touch.
      if (deleted.size > 0 && Array.isArray(node.required)) {
        node.required = node.required.filter((property) => !deleted.has(property));
      }
    }

    for (const keyword of ['allOf', 'anyOf', 'oneOf']) {
      for (const branch of node[keyword] ?? []) stripNavigationProperties(branch, declaringSchema);
    }
  }

  /**
   * Rewrite every ref inside a writable clone so it points at a writable variant.
   * Without this, a kept property such as listItem.fields would drag the unfiltered
   * entity graph back in.
   */
  function repointRefs(node) {
    if (!isRecord(node) && !Array.isArray(node)) return;
    if (Array.isArray(node)) {
      for (const item of node) repointRefs(item);
      return;
    }
    for (const [key, value] of Object.entries(node)) {
      if (key === '$ref') {
        const target = schemaNameFromRef(value);
        if (target && !isWritableVariantName(target)) node.$ref = writableVariantRef(target);
        continue;
      }
      repointRefs(value);
    }
  }

  forEachRequestBodySchema(spec, (schema) => {
    const direct = schemaNameFromRef(schema.$ref);
    if (direct) {
      // Already repointed by an earlier pass, so this stays idempotent.
      if (isWritableVariantName(direct)) return;
      schema.$ref = writableVariantRef(direct);
      return;
    }
    // Inline body schema. There is no declaring entity name, so the allowlist could
    // never match and stripping here would drop every navigation property including
    // the writable ones. Leave the inline schema alone and only follow its refs,
    // where a named entity does resolve.
    repointRefs(schema);
  });

  return report;
}

export { WRITABLE_NAVIGATION_PROPERTIES };
