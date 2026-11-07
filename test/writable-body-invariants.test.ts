import { describe, it, expect } from 'vitest';
import { z } from 'zod';
import { api } from '../src/generated/client.js';

/**
 * Guards the request-body narrowing in bin/modules/writable-request-bodies.mjs.
 *
 * That pass strips properties Graph marks `x-ms-navigationProperty: true` from
 * request bodies, which removes about a third of the emitted tool-schema payload.
 * The risk is over-reach: silently dropping a property Graph does accept removes a
 * capability with no error anywhere. These tests pin both directions, so a change
 * to the allowlist has to be deliberate.
 *
 * The generated client is produced by `npm run generate` (run in CI before tests).
 */

function bodyShape(alias: string): Record<string, z.ZodTypeAny> {
  const endpoint = api.endpoints.find((e) => e.alias === alias);
  if (!endpoint) throw new Error(`No endpoint with alias ${alias}`);
  const body = endpoint.parameters?.find((p: { type: string }) => p.type === 'Body');
  if (!body) throw new Error(`Endpoint ${alias} has no Body parameter`);

  // Unwrap optional/default/nullable wrappers down to the object schema.
  let schema = body.schema as z.ZodTypeAny;
  for (let i = 0; i < 10; i += 1) {
    const inner = (schema as { _def?: { innerType?: z.ZodTypeAny } })._def?.innerType;
    if (!inner) break;
    schema = inner;
  }
  const shape = (schema as unknown as { shape?: Record<string, z.ZodTypeAny> }).shape;
  if (!shape) throw new Error(`Body of ${alias} did not resolve to an object schema`);
  return shape;
}

/**
 * Exhaustive manifest of the writable navigation properties this server advertises.
 *
 * Deliberately maintained independently of WRITABLE_NAVIGATION_PROPERTIES: a test
 * derived from the allowlist would lose its assertion whenever an allowlist entry
 * was deleted, which is the exact regression it needs to catch. Round 3 caught that
 * the earlier hand-picked samples left gaps, for example event's
 * multiValueExtendedProperties was preserved by the implementation but asserted
 * nowhere, so deleting it would have narrowed the schema silently.
 *
 * Covers the safety-critical direction only: a property that disappears from a body
 * is a silent capability loss and must fail here. A newly published Graph navigation
 * property that this pass then drops is not covered, and is the far less dangerous
 * direction, since the tool keeps working with the fields it already had.
 */
const EXT = ['singleValueExtendedProperties', 'multiValueExtendedProperties'];
const WRITABLE_MANIFEST: Array<[string, string[]]> = [
  // Outlook: attachments inline on create, plus caller-owned extended properties.
  ['create-draft-email', ['attachments', ...EXT]],
  ['create-shared-mailbox-draft', ['attachments', ...EXT]],
  ['update-mail-message', ['attachments', ...EXT]],
  ['create-calendar-event', ['attachments', ...EXT]],
  ['create-specific-calendar-event', ['attachments', ...EXT]],
  ['update-calendar-event', ['attachments', ...EXT]],
  ['update-specific-calendar-event', ['attachments', ...EXT]],
  ['create-mail-folder', EXT],
  ['create-mail-child-folder', EXT],
  ['update-mail-folder', EXT],
  ['create-calendar', EXT],
  ['update-calendar', EXT],
  ['create-contact-folder', EXT],
  ['update-contact-folder', EXT],
  ['create-outlook-contact', EXT],
  ['update-outlook-contact', EXT],
  // SharePoint.
  ['create-sharepoint-list', ['columns']],
  ['create-sharepoint-list-item', ['fields']],
  ['update-sharepoint-list-item', ['fields']],
  // Teams: members on create, and hostedContents for inline images.
  ['create-chat', ['members']],
  ['create-team-channel', ['members']],
  ['update-team-channel', ['members']],
  ['send-chat-message', ['attachments', 'hostedContents']],
  ['send-channel-message', ['attachments', 'hostedContents']],
  ['reply-to-chat-message', ['attachments', 'hostedContents']],
  ['reply-to-channel-message', ['attachments', 'hostedContents']],
  // To Do.
  ['create-todo-task', ['attachments', 'checklistItems', 'extensions', 'linkedResources']],
  ['update-todo-task', ['attachments', 'checklistItems', 'extensions', 'linkedResources']],
  ['create-todo-task-list', ['extensions']],
  ['update-todo-task-list', ['extensions']],
  // Excel formatting is applied by PATCHing these onto the range format.
  ['format-excel-range', ['borders', 'fill', 'font', 'protection']],
  // Graph takes only an @odata.bind reference here, so narrowing away the `message`
  // relationship left the body with nothing but a read-only id and no way for a model
  // to say which message to pin. endpoints.json now declares the real wire shape.
  ['pin-chat-message', ['message@odata.bind']],
];

// Reasoned spot checks kept alongside the manifest, because these carry the WHY.
const MUST_KEEP: Array<[string, string[], string]> = [
  [
    'format-excel-range',
    ['font', 'fill', 'borders', 'protection'],
    'these are the entire point of the tool: formatting is applied by PATCHing them',
  ],
  ['create-draft-email', ['attachments'], 'a draft can be created with attachments inline'],
  [
    'create-draft-email',
    ['singleValueExtendedProperties', 'multiValueExtendedProperties'],
    'callers stash their own sync metadata in extended properties',
  ],
  ['create-calendar-event', ['attachments', 'singleValueExtendedProperties'], 'same as mail'],
  ['create-calendar', ['singleValueExtendedProperties'], 'extended properties are writable'],
  ['create-contact-folder', ['singleValueExtendedProperties'], 'extended properties are writable'],
  ['create-sharepoint-list-item', ['fields'], 'POST list items requires fields'],
  ['create-sharepoint-list', ['columns'], 'a list can be created with its column definitions'],
  ['create-chat', ['members'], 'POST /chats requires members'],
  ['create-team-channel', ['members'], 'a private channel requires members'],
  [
    'create-todo-task',
    ['checklistItems', 'linkedResources', 'attachments'],
    'a task can carry these on create',
  ],
  // Graph accepts hostedContents when creating a chat or channel message, which is
  // how inline images are sent. Dropping it silently removed that capability.
  ['send-chat-message', ['hostedContents'], 'inline images are sent as hostedContents'],
  ['send-channel-message', ['hostedContents'], 'inline images are sent as hostedContents'],
  ['reply-to-chat-message', ['hostedContents'], 'inline images are sent as hostedContents'],
  ['reply-to-channel-message', ['hostedContents'], 'inline images are sent as hostedContents'],
  // An open extension can be created inline with the item that carries it.
  ['create-todo-task', ['extensions'], 'an open extension can be created with the task'],
  ['create-todo-task-list', ['extensions'], 'an open extension can be created with the list'],
];

// Read-only navigation properties that dominated the payload. `create-team-channel`
// alone carried 28 KB of filesFolder, and five tools carried a createdByUser that
// expanded to the entire Graph user entity.
const MUST_DROP: Array<[string, string[]]> = [
  ['create-team-channel', ['filesFolder', 'messages', 'sharedWithTeams', 'tabs']],
  ['create-sharepoint-list', ['drive', 'items', 'createdByUser', 'contentTypes']],
  ['create-sharepoint-list-item', ['analytics', 'createdByUser', 'versions']],
  ['create-mail-folder', ['messages', 'messageRules', 'childFolders']],
  ['create-onedrive-folder', ['createdByUser', 'lastModifiedByUser']],
  ['create-calendar', ['calendarView', 'events', 'calendarPermissions']],
  ['create-contact-folder', ['contacts']],
  ['create-onenote-notebook', ['sectionGroups', 'sections']],
  // plannerTask.details is a read-only relationship: it is written through the
  // plannerTaskDetails endpoint (see update-planner-task-details), not inline.
  ['create-planner-task', ['details']],
  ['create-planner-bucket', ['tasks']],
  // Replies are posted to their own endpoint, never inline with the parent message.
  ['send-chat-message', ['replies']],
  ['send-channel-message', ['replies']],
  // Apps and tabs are installed through their own endpoints, not the create body.
  ['create-chat', ['installedApps', 'tabs', 'pinnedMessages', 'lastMessagePreview']],
  ['create-team-channel', ['tabs']],
  // The nested message entity is not the input Graph accepts, so advertising it only
  // misled callers into filling an object that Graph rejects.
  ['pin-chat-message', ['message', 'id']],
];

describe('every writable navigation property in the manifest survives codegen', () => {
  for (const [alias, properties] of WRITABLE_MANIFEST) {
    for (const property of properties) {
      it(`${alias} keeps ${property}`, () => {
        expect(
          Object.keys(bodyShape(alias)),
          `${alias}.body.${property} vanished. Either an entry was removed from ` +
            'WRITABLE_NAVIGATION_PROPERTIES, or Graph stopped declaring it. Silently ' +
            'dropping a writable property removes a capability with no error anywhere.'
        ).toContain(property);
      });
    }
  }
});

describe('request bodies keep every writable property', () => {
  for (const [alias, properties, why] of MUST_KEEP) {
    for (const property of properties) {
      it(`${alias} keeps ${property} (${why})`, () => {
        expect(Object.keys(bodyShape(alias))).toContain(property);
      });
    }
  }
});

describe('request bodies drop read-only navigation properties', () => {
  for (const [alias, properties] of MUST_DROP) {
    for (const property of properties) {
      it(`${alias} drops ${property}`, () => {
        expect(Object.keys(bodyShape(alias))).not.toContain(property);
      });
    }
  }
});
