import { describe, it, expect } from 'vitest';
// @ts-expect-error - bin modules are plain ESM JavaScript with no type declarations
import { restrictRequestBodiesToWritableProperties } from '../bin/modules/writable-request-bodies.mjs';

/**
 * Microsoft's Graph metadata types a POST/PATCH request body as the full entity,
 * including every read-only navigation property (the channel's filesFolder, the
 * list's drive, a createdByUser that expands to the whole user entity). None of
 * it is valid input, and it dominates the emitted tool schemas. These tests pin
 * the codegen pass that restricts request bodies to writable properties.
 */

function miniSpec() {
  return {
    paths: {
      '/me/mailFolders': {
        post: {
          operationId: 'create-mail-folder',
          requestBody: {
            content: {
              'application/json': {
                schema: { $ref: '#/components/schemas/microsoft.graph.mailFolder' },
              },
            },
          },
          responses: {
            '201': {
              content: {
                'application/json': {
                  schema: { $ref: '#/components/schemas/microsoft.graph.mailFolder' },
                },
              },
            },
          },
        },
      },
      '/sites/{id}/lists/{lid}/items': {
        post: {
          operationId: 'create-sharepoint-list-item',
          requestBody: {
            content: {
              'application/json': {
                schema: { $ref: '#/components/schemas/microsoft.graph.listItem' },
              },
            },
          },
        },
      },
      '/me/messages/{id}': {
        get: {
          operationId: 'get-mail-message',
          responses: {
            '200': {
              content: {
                'application/json': {
                  schema: { $ref: '#/components/schemas/microsoft.graph.message' },
                },
              },
            },
          },
        },
      },
    },
    components: {
      schemas: {
        'microsoft.graph.entity': {
          type: 'object',
          properties: { id: { type: 'string' } },
        },
        'microsoft.graph.mailFolder': {
          allOf: [
            { $ref: '#/components/schemas/microsoft.graph.entity' },
            {
              type: 'object',
              properties: {
                displayName: { type: 'string' },
                isHidden: { type: 'boolean' },
                // Read-only navigation properties, and a self-referential one.
                messages: {
                  type: 'array',
                  items: { $ref: '#/components/schemas/microsoft.graph.message' },
                  'x-ms-navigationProperty': true,
                },
                childFolders: {
                  type: 'array',
                  items: { $ref: '#/components/schemas/microsoft.graph.mailFolder' },
                  'x-ms-navigationProperty': true,
                },
              },
            },
          ],
        },
        'microsoft.graph.listItem': {
          type: 'object',
          properties: {
            name: { type: 'string' },
            // Navigation property, but genuinely required to create a list item.
            fields: {
              $ref: '#/components/schemas/microsoft.graph.fieldValueSet',
              'x-ms-navigationProperty': true,
            },
            // Navigation property that is purely server-assigned.
            createdByUser: {
              $ref: '#/components/schemas/microsoft.graph.user',
              'x-ms-navigationProperty': true,
            },
          },
        },
        'microsoft.graph.fieldValueSet': {
          type: 'object',
          properties: {
            Title: { type: 'string' },
            // Forces fieldValueSet itself to change, so a request body reaching it
            // through listItem.fields must be repointed too.
            owner: {
              $ref: '#/components/schemas/microsoft.graph.user',
              'x-ms-navigationProperty': true,
            },
          },
        },
        // A plain complex type with no navigation properties at all.
        'microsoft.graph.driveRecipient': {
          type: 'object',
          properties: { email: { type: 'string' } },
        },
        'microsoft.graph.user': {
          type: 'object',
          properties: { displayName: { type: 'string' }, employeeHireDate: { type: 'string' } },
        },
        'microsoft.graph.message': {
          type: 'object',
          properties: {
            subject: { type: 'string' },
            attachments: {
              type: 'array',
              items: { type: 'object' },
              'x-ms-navigationProperty': true,
            },
          },
        },
      },
    },
  };
}

function bodySchemaRef(spec: any, path: string, method: string): string {
  return spec.paths[path][method].requestBody.content['application/json'].schema.$ref;
}

function resolve(spec: any, ref: string) {
  return spec.components.schemas[ref.replace('#/components/schemas/', '')];
}

/** Merge properties across an allOf chain so assertions read naturally. */
function propsOf(spec: any, schema: any): Record<string, any> {
  const out: Record<string, any> = {};
  if (!schema || typeof schema !== 'object') return out;
  if (schema.$ref) return propsOf(spec, resolve(spec, schema.$ref));
  for (const branch of schema.allOf ?? []) Object.assign(out, propsOf(spec, branch));
  Object.assign(out, schema.properties ?? {});
  return out;
}

describe('restrictRequestBodiesToWritableProperties', () => {
  it('repoints a request body at a writable variant instead of the shared entity', () => {
    const spec = miniSpec();
    restrictRequestBodiesToWritableProperties(spec);

    const ref = bodySchemaRef(spec, '/me/mailFolders', 'post');
    expect(ref).not.toBe('#/components/schemas/microsoft.graph.mailFolder');
    expect(resolve(spec, ref)).toBeDefined();
  });

  it('drops read-only navigation properties from the request body', () => {
    const spec = miniSpec();
    restrictRequestBodiesToWritableProperties(spec);

    const body = propsOf(spec, resolve(spec, bodySchemaRef(spec, '/me/mailFolders', 'post')));
    expect(Object.keys(body).sort()).toEqual(['displayName', 'id', 'isHidden']);
    expect(body.messages).toBeUndefined();
    expect(body.childFolders).toBeUndefined();
  });

  it('keeps navigation properties that Graph genuinely accepts on write', () => {
    const spec = miniSpec();
    restrictRequestBodiesToWritableProperties(spec);

    const body = propsOf(
      spec,
      resolve(spec, bodySchemaRef(spec, '/sites/{id}/lists/{lid}/items', 'post'))
    );
    // `fields` is allowlisted for listItem: without it you cannot create an item.
    expect(body.fields).toBeDefined();
    // `createdByUser` is server-assigned, and is what expands to the whole user entity.
    expect(body.createdByUser).toBeUndefined();
  });

  it('leaves response schemas and the shared entity definitions untouched', () => {
    const spec = miniSpec();
    const before = JSON.parse(
      JSON.stringify(spec.components.schemas['microsoft.graph.mailFolder'])
    );
    restrictRequestBodiesToWritableProperties(spec);

    // The original entity keeps every navigation property, so anything reading a
    // Graph reply is unaffected.
    expect(spec.components.schemas['microsoft.graph.mailFolder']).toEqual(before);
    expect(
      spec.paths['/me/mailFolders'].post.responses['201'].content['application/json'].schema.$ref
    ).toBe('#/components/schemas/microsoft.graph.mailFolder');
    expect(
      spec.paths['/me/messages/{id}'].get.responses['200'].content['application/json'].schema.$ref
    ).toBe('#/components/schemas/microsoft.graph.message');
  });

  it('rewrites a nested ref when the referenced schema also changes', () => {
    const spec = miniSpec();
    restrictRequestBodiesToWritableProperties(spec);

    const listItemBodyRef = bodySchemaRef(spec, '/sites/{id}/lists/{lid}/items', 'post');
    const fieldsRef = propsOf(spec, resolve(spec, listItemBodyRef)).fields.$ref;
    // fieldValueSet drops its own navigation property, so a body reaching it via
    // the allowlisted `fields` must follow the writable variant.
    expect(fieldsRef).not.toBe('#/components/schemas/microsoft.graph.fieldValueSet');
    expect(propsOf(spec, resolve(spec, fieldsRef)).owner).toBeUndefined();
    expect(propsOf(spec, resolve(spec, fieldsRef)).Title).toBeDefined();
  });

  it('leaves a schema that would not change under its original name', () => {
    const spec = miniSpec();
    restrictRequestBodiesToWritableProperties(spec);

    // driveRecipient has no navigation properties, so cloning it to a `.writable`
    // twin would be pure churn and would break refs that name it directly.
    expect(spec.components.schemas['microsoft.graph.driveRecipient.writable']).toBeUndefined();
    expect(spec.components.schemas['microsoft.graph.driveRecipient']).toBeDefined();
  });

  it('terminates on self-referential entities', () => {
    const spec = miniSpec();
    // mailFolder.childFolders points back at mailFolder.
    expect(() => restrictRequestBodiesToWritableProperties(spec)).not.toThrow();
    const body = propsOf(spec, resolve(spec, bodySchemaRef(spec, '/me/mailFolders', 'post')));
    expect(body.childFolders).toBeUndefined();
  });

  it('is idempotent, so a second codegen pass changes nothing', () => {
    const spec = miniSpec();
    restrictRequestBodiesToWritableProperties(spec);
    const once = JSON.stringify(spec);
    restrictRequestBodiesToWritableProperties(spec);
    expect(JSON.stringify(spec)).toBe(once);
  });

  it('ignores GET and DELETE operations, which carry no request body', () => {
    const spec = miniSpec();
    expect(() => restrictRequestBodiesToWritableProperties(spec)).not.toThrow();
    expect(spec.paths['/me/messages/{id}'].get.requestBody).toBeUndefined();
  });

  it('reports what it removed so codegen output is auditable', () => {
    const spec = miniSpec();
    const report = restrictRequestBodiesToWritableProperties(spec);

    expect(report.variantsCreated).toBeGreaterThan(0);
    expect(report.removed).toEqual(
      expect.arrayContaining([
        expect.objectContaining({ schema: 'microsoft.graph.mailFolder', property: 'messages' }),
        expect.objectContaining({ schema: 'microsoft.graph.listItem', property: 'createdByUser' }),
      ])
    );
    // Allowlisted properties must not show up as removed.
    expect(report.removed.some((r: any) => r.property === 'fields')).toBe(false);
  });
});
