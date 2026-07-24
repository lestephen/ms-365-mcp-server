import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import logger from '../logger.js';

/**
 * Operator-facing tool blocklist.
 *
 * `execute-tool` dispatches Graph calls by name, so a client-side deny rule keyed
 * on a tool name never matches once discovery mode is on: the call the client sees
 * and evaluates is `execute-tool`, and the name of the real operation sits inside
 * the arguments. A policy such as "draft mail but never send it" is therefore only
 * enforceable inside the server.
 *
 * The blocklist has to hold on every registration path, not just the Graph one.
 * Rather than threading a pattern through each registrar (`registerAuthTools`,
 * `registerGraphTools`, `registerDiscoveryTools`, the utility loop) and hoping no
 * future one is forgotten, `withToolBlocklist` wraps the server so a blocked name
 * cannot be registered at all, whoever asks. `buildToolsRegistry` still takes the
 * compiled regex directly, because the registry is what `execute-tool`,
 * `get-tool-schema` and `search-tools` read and is not a registration path.
 */

/**
 * Compile the --blocked-tools pattern.
 *
 * Unlike --enabled-tools, an unparseable pattern here is fatal. This is a guardrail
 * rather than a filter, so ignoring a typo would silently unblock exactly the tools
 * an operator asked to be unreachable. `parseArgs` performs the same check at
 * startup, so an invalid pattern is refused before the server listens.
 */
export function compileBlockedToolsRegex(pattern?: string): RegExp | undefined {
  if (!pattern) return undefined;
  try {
    const regex = new RegExp(pattern, 'i');
    logger.info(`Tool blocklist active with pattern: ${pattern}`);
    return regex;
  } catch (error) {
    throw new Error(
      `Invalid --blocked-tools regex ${JSON.stringify(pattern)}: ${(error as Error).message}. ` +
        'Refusing to start, because ignoring it would leave the blocked tools reachable.'
    );
  }
}

/**
 * Wrap a server so `tool()` silently skips any name matching the blocklist. Returns
 * the server unchanged when no pattern is set, so the default path is untouched.
 *
 * Blocking an auth tool such as `login` will make the server unusable. That is the
 * operator's call to make, and it fails visibly, so it is not second-guessed here.
 */
export function withToolBlocklist(server: McpServer, pattern?: string): McpServer {
  const blocked = compileBlockedToolsRegex(pattern);
  if (!blocked) return server;

  const guarded = Object.create(server) as McpServer;

  // Guard every registration entry point the SDK exposes, not just the one current
  // registrars happen to use. All call sites use tool() today, but this is the single
  // enforcement point for an operator policy, so it should not be one refactor away
  // from a bypass.
  for (const method of ['tool', 'registerTool'] as const) {
    const original = server[method];
    if (typeof original !== 'function') continue;
    const register = (original as (...args: unknown[]) => unknown).bind(server);
    (guarded as unknown as Record<string, unknown>)[method] = (
      name: string,
      ...rest: unknown[]
    ) => {
      if (typeof name === 'string' && blocked.test(name)) {
        logger.info(`Blocking tool ${name} - matches blocklist pattern`);
        // Returning undefined is safe: every caller either ignores the handle or
        // guards its own registration in a try/catch.
        return undefined;
      }
      return register(name, ...rest);
    };
  }

  return guarded;
}
