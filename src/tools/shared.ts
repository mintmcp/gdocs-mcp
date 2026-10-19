/**
 * Pieces every tool module shares: the error envelope, the Word-upload guard,
 * the tab-aware document fetch and the descriptions repeated across schemas.
 */

import { docsOnly as wrapDocsOnly } from "../lib/word.js";
import { resolveTab, type ResolvedTab } from "../lib/tabs.js";
import { type DocsBody } from "../lib/tableModel.js";
import { GoogleApiError, makeDriveRequest, makeDocsRequest } from "../lib/google.js";
import { type WriteControl } from "../lib/docs-structure.js";

// Tabs nest at most 3 levels deep, so three tabProperties selectors cover the
// whole tree without pulling any tab content
export const TAB_LIST_FIELDS = 'revisionId,tabs.tabProperties,tabs.childTabs.tabProperties,tabs.childTabs.childTabs.tabProperties';

export const TAB_ID_DESCRIPTION = 'ID of the tab to target in a multi-tab document: tabProperties.tabId, the value after ?tab= in the document URL (e.g. "t.abc123"), also listed in get_document\'s `tabs`. Omit to target the first tab.';

export const SUGGEST_DESCRIPTION = 'Make the change as a tracked suggestion (Google Docs "Suggesting" mode) instead of a direct edit, so a reviewer accepts or rejects it in the document. Suggestions show in get_document structure with their ids; review_suggestion accepts or rejects them. Defaults to false.';

export const SEGMENT_ID_DESCRIPTION = 'Target a header, footer or footnote instead of the body: the id get_document reports for it with include_structure=true (headers[].id, footers[].id, footnotes[].id). Indices inside a segment start at 0. Omit for the document body.';

export const REQUIRED_REVISION_DESCRIPTION = 'If provided, the write fails with a 400 if the document revision has changed since this revisionId. Use the `revisionId` returned by get_document(include_structure=true) for optimistic concurrency control.';

export const DOCUMENTS_SCOPE = "https://www.googleapis.com/auth/documents";
export const DRIVE_READONLY_SCOPE = "https://www.googleapis.com/auth/drive.readonly";
export const DRIVE_FILE_SCOPE = "https://www.googleapis.com/auth/drive.file";

export interface DriveFile {
  id: string;
  name: string;
  mimeType?: string;
  webViewLink?: string;
}

/**
 * True only when Drive judged the file itself unconvertible. Auth, throttling,
 * outages and requests that never completed are about the service, not the
 * content, and must not be reported to the caller as a bad document.
 */
export function isConversionRefusal(err: unknown): boolean {
  if (!(err instanceof GoogleApiError)) return false;
  const { status } = err;
  return status >= 400 && status < 500 &&
    status !== 401 && status !== 403 && status !== 404 &&
    status !== 408 && status !== 425 && status !== 429;
}

// The `code` for an error that isn't Google's. A failed fetch is a TypeError like
// a bug in our code, so the system code on its cause (ECONNRESET, ENOTFOUND,
// UND_ERR_CONNECT_TIMEOUT) comes first. A plain Error is a message we wrote,
// usually a rejected input, so it gets no code; any other class (TypeError,
// RangeError) points at a bug
function errorClass(err: Error): string | undefined {
  const cause = (err as { cause?: { code?: unknown } }).cause;
  if (typeof cause?.code === "string") return cause.code;
  return err.name === "Error" ? undefined : err.name;
}

/**
 * Format any thrown error into a structured MCP error response. Handlers MUST
 * wrap their bodies in try/catch and route caught errors through this so
 * downstream callers get machine-readable `{error, status, code, retryAfter}`
 * envelopes rather than plain strings.
 */
export function toolErrorResponse(err: unknown): { content: Array<{ type: 'text'; text: string }>; isError: true } {
  let payload: Record<string, unknown>;
  if (err instanceof GoogleApiError) {
    payload = {
      error: err.message,
      status: err.status,
      code: err.code,
      ...(err.retryAfter !== undefined ? { retryAfter: err.retryAfter } : {}),
      ...(err.details !== undefined ? { details: err.details } : {}),
      api: err.api,
    };
  } else if (err instanceof Error) {
    payload = { error: err.message, code: errorClass(err) };
  } else {
    payload = { error: String(err) };
  }
  return {
    content: [{ type: 'text', text: JSON.stringify(payload) }],
    isError: true,
  };
}

/** Write tools report what they did; in suggest mode that is a proposal, not an edit. */
export function doneMessage(message: string, suggest?: boolean): string {
  return suggest ? `${message} (as a suggestion, pending review)` : message;
}

/**
 * The error envelope for a step that failed after an earlier step already
 * changed the document: `partialSuccess` plus whatever ids the caller needs to
 * finish by hand instead of retrying and creating the thing twice.
 */
export function partialFailure(err: unknown, prefix: string, extra: Record<string, unknown>) {
  const base = toolErrorResponse(err);
  const parsed = JSON.parse(base.content[0].text);
  const payload = { ...parsed, error: `${prefix}: ${parsed.error}`, partialSuccess: true, ...extra };
  return { content: [{ type: 'text' as const, text: JSON.stringify(payload) }], isError: true as const };
}

/** The success envelope every tool returns: the output pretty-printed plus structured. */
export function toolResult<T>(output: T) {
  return {
    content: [{ type: 'text' as const, text: JSON.stringify(output, null, 2) }],
    structuredContent: output,
  };
}

/**
 * Refuse writes against Word uploads. Bound here so the wrapper stays pure.
 * Its Drive lookup and refusal run before the handler's own try, so their
 * throws go through toolErrorResponse like every other failure.
 */
export const docsOnly = <T,>(handler: (args: any, context: any) => Promise<T>) => {
  const guarded = wrapDocsOnly(handler, makeDriveRequest);
  return async (args: any, context: any) => {
    try {
      return await guarded(args, context);
    } catch (err) {
      return toolErrorResponse(err);
    }
  };
};

export async function fetchResolvedTab(
  documentId: string,
  accessToken: string,
  tabId?: string,
): Promise<{ doc: any; resolved: ResolvedTab; tabBody: DocsBody }> {
  const doc = await makeDocsRequest(`/${encodeURIComponent(documentId)}?includeTabsContent=true`, accessToken, { method: 'GET' });
  const resolved = resolveTab(doc, tabId);
  return { doc, resolved, tabBody: { revisionId: doc.revisionId, body: { content: resolved.body } } };
}

/** One batchUpdate call; `writeControl` is only sent when set. */
export function docsBatchUpdate(
  documentId: string,
  accessToken: string,
  requests: any[],
  writeControl?: WriteControl,
): Promise<any> {
  return makeDocsRequest(`/${encodeURIComponent(documentId)}:batchUpdate`, accessToken, {
    method: 'POST',
    body: JSON.stringify({ requests, ...(writeControl ? { writeControl } : {}) }),
  });
}
