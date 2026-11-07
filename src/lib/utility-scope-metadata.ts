/** Fixed Graph permissions for utilities whose resource scope is known before invocation. */
export const UTILITY_SCOPE_METADATA = [
  {
    toolName: 'read-mail-text',
    method: 'GET',
    pathPattern: '/me/messages',
    scopes: ['Mail.Read'],
  },
];

// Utilities accepting arbitrary Graph targets deliberately have no inferred scope entry.
export function getUtilityScopeMetadata(toolName: string) {
  return UTILITY_SCOPE_METADATA.find((entry) => entry.toolName === toolName);
}
