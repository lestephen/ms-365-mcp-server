/** Shared context for MCP `initialize.instructions` (hosts that forward it to the model). */
export type McpInstructionsContext = {
  orgMode: boolean;
  appMode?: boolean;
  readOnly: boolean;
  multiAccount: boolean;
  attachmentUrlsEnabled?: boolean;
};

function buildGeneralMcpInstructions(opts: McpInstructionsContext): string {
  const parts = [
    'Microsoft 365 MCP exposes Microsoft Graph through MCP tools. Use each tool name, description, and parameter schema as the source of truth.',
    'Microsoft Graph OData: do not combine $filter with $search on the same request. For lists, prefer modest $top (or top) and $select; avoid very large pages unless the user needs them.',
    'Mail and message $search uses KQL; the $search query parameter value must be double-quoted per Graph (see search-query-parameter in Microsoft Graph docs).',
    'When you need an organizational user or recipient address, resolve it with list-users (or another directory tool); do not invent SMTP addresses.',
    'Directory $search on collections such as /users or /groups requires ConsistencyLevel: eventual when the tool exposes that header.',
    'Teams chat and channel messages: prefer HTML contentType in the body; plain text is often mangled by Graph.',
    'Files / binary content: for large drive/SharePoint file content, prefer get-download-url to resolve a pre-authenticated URL for out-of-band download.',
    opts.attachmentUrlsEnabled
      ? 'Attachment URLs are enabled: for document reading and downloads, prefer get-download-url for mail attachments, meeting recordings, drive files and other authenticated byte resources, then hand the URL to an external converter or downloader. Use download-bytes only when the caller explicitly needs bytes. Server-minted URLs are single-use: mint a fresh URL for every fetch, including probes, conversion, continuation and retries; retain the same target, account and document selection. This is a preference, not enforcement: enabling attachment URLs does not scrub arbitrary response bodies; use existing tool filtering when a text-only tool surface is required.'
      : 'Use download-bytes for authenticated byte reads such as mail attachments, profile photos, Teams hosted content, and meeting recordings.',
    'In stdio mode (or HTTP with --http-local-file-tools), download-bytes-to-file writes authenticated bytes straight to a local absolute path instead of returning base64. These tools take relative Microsoft Graph paths, not absolute URLs. For uploads, prefer create-upload-session and send the bytes to the returned uploadUrl yourself, streaming them from disk; this server does not perform that PUT. To attach a file to a draft message or an event, or to put a file in a OneDrive/SharePoint folder, use get-upload-url when available and PUT the file from disk to the URL it returns (e.g. curl -T), or create-mail-attachment-upload-session for a draft. upload-file-content instead carries the entire file as a base64 tool argument, and a truncated argument decodes to a truncated file with no error, so do not build that string regardless of how small the file looks.',
  ];
  if (opts.appMode)
    parts.push(
      'App mode: there is no signed-in user, so /me paths do not work. Name the user in /users/{user-id} paths with a UPN or object id.'
    );
  if (opts.readOnly) parts.push('This server is read-only; writes to Microsoft 365 are disabled.');
  else
    parts.push(
      'Mail reply, reply-all and forward tools (including their drafts and shared mailboxes): Graph puts comment into an HTML body, so plain-text line breaks are usually lost; use <br> for line breaks or <p> for paragraphs.'
    );
  if (opts.multiAccount)
    parts.push('Multiple accounts: pass the account parameter when required (see list-accounts).');
  if (!opts.orgMode)
    parts.push('Work/school-only tools require starting the server with --org-mode.');
  return parts.join(' ');
}

const DISCOVERY_MODE_INSTRUCTIONS_ADDON =
  'DISCOVERY MODE ADD-ON: Graph is reached via search-tools → get-tool-schema → execute-tool (plus auth helpers). ' +
  'Workflow: (1) call search-tools with short natural-language keywords (BM25-ranked); ' +
  '(2) call get-tool-schema(tool_name) to see the parameters, required fields, and enum values; ' +
  '(3) call execute-tool with tool_name exactly as returned and parameters shaped per the schema. ' +
  'Skipping get-tool-schema is the leading cause of Graph 400 errors here. ' +
  'If search-tools returns no matches, retry with shorter or different keywords.';

const HYBRID_MODE_INSTRUCTIONS_ADDON =
  'HYBRID MODE ADD-ON: the most common Graph operations are already registered directly as named tools, ' +
  'and you can see them in your tool list. Prefer a named tool whenever one fits and call it straight away, ' +
  'with no search-tools or get-tool-schema step. ' +
  'Everything else Graph offers is still reachable, it is just not loaded up front: ' +
  'reach it via search-tools → get-tool-schema → execute-tool. ' +
  'Workflow for those: (1) call search-tools with short natural-language keywords (BM25-ranked); ' +
  '(2) call get-tool-schema(tool_name) to see the parameters, required fields, and enum values; ' +
  '(3) call execute-tool with tool_name exactly as returned and parameters shaped per the schema. ' +
  'Skipping get-tool-schema is the leading cause of Graph 400 errors here. ' +
  'If search-tools returns no matches, retry with shorter or different keywords. ' +
  'A tool missing from your tool list is not unavailable, so search for it before reporting that you cannot do something. ' +
  'Do NOT call a tool by name unless it appears in your tool list: search-tools and get-tool-schema describe tools that are ' +
  'not registered here, and calling one of those directly fails with "Tool not found". Both report the route per tool, as ' +
  'invoke_via in search results and as an invocation block in get-tool-schema; follow it.';

/**
 * Full MCP `initialize.instructions` string: general guidance for every mode, plus
 * a suffix describing how Graph is reached in this configuration.
 *
 * Three shapes: direct tools only (no suffix), discovery only, and hybrid (named
 * tools plus the discovery triad). The hybrid suffix has to say both paths exist,
 * otherwise the model routes everything through execute-tool and ignores the named
 * tools sitting in front of it.
 */
export function buildMcpServerInstructions(
  opts: McpInstructionsContext & { discovery: boolean; eagerTools?: string[] }
): string {
  const general = buildGeneralMcpInstructions(opts);
  if (!opts.discovery) return general;
  if (opts.eagerTools?.length) return `${general} ${HYBRID_MODE_INSTRUCTIONS_ADDON}`;
  return `${general} ${DISCOVERY_MODE_INSTRUCTIONS_ADDON}`;
}
