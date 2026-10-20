import { z } from 'zod';
import { withGoogleAuth } from "../auth.js";
import { SCOPES } from "../scopes.js";
import { docsOnly as wrapDocsOnly } from "../lib/word.js";
import { attachLabelsMeta } from "../lib/driveLabels.js";
import { type ResolvedTab, resolveTab, summarizeTabs, type TabSummary } from "../lib/tabs.js";
import { GoogleApiError, makeDriveRequest, makeDocsRequest } from "../lib/google.js";
import { buildWriteControl, type WriteOptions } from "../lib/docRequests.js";

// Tabs nest at most 3 levels deep
const TAB_LIST_FIELDS = 'revisionId,tabs.tabProperties,tabs.childTabs.tabProperties,tabs.childTabs.childTabs.tabProperties';

export const documentIdField = z.string().describe('Google Doc ID');

export const tabIdField = z.string().min(1).optional().describe('ID of the tab to target in a multi-tab document: tabProperties.tabId, the value after ?tab= in the document URL (e.g. "t.abc123"), also listed in get_document\'s `tabs`. Omit to target the first tab.');

export const suggestField = z.boolean().optional().describe('Make the change as a tracked suggestion (Google Docs "Suggesting" mode) instead of a direct edit, so a reviewer accepts or rejects it in the document. Suggestions show in get_document structure with their ids; review_suggestion accepts or rejects them. Defaults to false.');

export const segmentIdField = z.string().min(1).optional().describe('Target a header, footer or footnote instead of the body: the id get_document reports for it with include_structure=true (headers[].id, footers[].id, footnotes[].id). Indices inside a segment start at 0. Omit for the document body.');

export const keysEnum = <T extends Record<string, unknown>>(table: T) =>
  z.enum(Object.keys(table) as [Extract<keyof T, string>, ...Array<Extract<keyof T, string>>]);

export const tabSummarySchema = z.object({
  tabId: z.string(),
  title: z.string().optional(),
  index: z.number().int().optional(),
  nestingLevel: z.number().int().optional(),
});

export const requiredRevisionField = z.string().optional().describe('If provided, the write fails with a 400 if the document revision has changed since this revisionId. Use the `revisionId` returned by get_document(include_structure=true) for optimistic concurrency control.');

export const editResultShape = { id: z.string(), message: z.string() };

export interface DriveFile {
  id: string;
  name: string;
  mimeType?: string;
  webViewLink?: string;
}

export function isConversionRefusal(err: unknown): boolean {
  if (!(err instanceof GoogleApiError)) return false;
  const { status } = err;
  return status >= 400 && status < 500 &&
    status !== 401 && status !== 403 && status !== 404 &&
    status !== 408 && status !== 425 && status !== 429;
}

// A failed fetch reports its cause's system code; a plain Error is our own message and gets none
function errorClass(err: Error): string | undefined {
  const cause = (err as { cause?: { code?: unknown } }).cause;
  if (typeof cause?.code === "string") return cause.code;
  return err.name === "Error" ? undefined : err.name;
}

export class PartialFailure extends Error {
  constructor(cause: unknown, prefix: string, readonly extra: Record<string, unknown>) {
    super(prefix, { cause });
  }
}

function errorPayload(err: unknown): Record<string, unknown> {
  if (err instanceof PartialFailure) {
    const inner = errorPayload(err.cause);
    return { ...inner, error: `${err.message}: ${inner.error}`, partialSuccess: true, ...err.extra };
  }
  if (err instanceof GoogleApiError) {
    return {
      error: err.message,
      status: err.status,
      code: err.code,
      ...(err.retryAfter !== undefined ? { retryAfter: err.retryAfter } : {}),
      ...(err.details !== undefined ? { details: err.details } : {}),
      api: err.api,
    };
  }
  if (err instanceof Error) return { error: err.message, code: errorClass(err) };
  return { error: String(err) };
}

function toolErrorResponse(err: unknown): { content: Array<{ type: 'text'; text: string }>; isError: true } {
  return {
    content: [{ type: 'text', text: JSON.stringify(errorPayload(err)) }],
    isError: true,
  };
}

export function doneMessage(message: string, suggest?: boolean): string {
  return suggest ? `${message} (as a suggestion, pending review)` : message;
}

export function toolResult<T>(output: T) {
  return {
    content: [{ type: 'text' as const, text: JSON.stringify(output, null, 2) }],
    structuredContent: output,
  };
}

export interface ToolContext {
  accessToken: string;
}

interface ToolSpec<S extends z.ZodRawShape> {
  description: string;
  schema: S;
  outputSchema?: z.ZodRawShape;
  readOnlyHint?: boolean;
  destructiveHint?: boolean;
  run: (args: z.infer<z.ZodObject<S>>, context: ToolContext) => Promise<unknown>;
}

export interface ToolDefinition {
  description: string;
  schema: z.ZodRawShape;
  outputSchema?: z.ZodRawShape;
  readOnlyHint?: boolean;
  destructiveHint?: boolean;
  scope: string;
  handler: (args: any) => Promise<any>;
}

export function defineTool<S extends z.ZodRawShape>(
  scope: string,
  { run, ...definition }: ToolSpec<S>,
  options: { docsOnly?: boolean; labels?: boolean } = {},
): ToolDefinition {
  const guarded = options.docsOnly ? wrapDocsOnly(run, makeDriveRequest) : run;
  const safe = async (args: z.infer<z.ZodObject<S>>, context: ToolContext) => {
    try {
      return await guarded(args, context);
    } catch (err) {
      return toolErrorResponse(err);
    }
  };
  const handler = withGoogleAuth(options.labels ? attachLabelsMeta((a: any) => a.document_id, safe) : safe);
  return { ...definition, scope, handler };
}

export const editTool = <S extends z.ZodRawShape>(spec: ToolSpec<S>): ToolDefinition =>
  defineTool(SCOPES.DOCUMENTS, spec, { docsOnly: true });

export const SUGGESTIONS_VIEW_MODES = ['DEFAULT_FOR_CURRENT_ACCESS', 'SUGGESTIONS_INLINE', 'PREVIEW_SUGGESTIONS_ACCEPTED', 'PREVIEW_WITHOUT_SUGGESTIONS'] as const;

export async function fetchResolvedTab(
  documentId: string,
  accessToken: string,
  tabId?: string,
  suggestionsViewMode?: typeof SUGGESTIONS_VIEW_MODES[number],
): Promise<{ doc: any; resolved: ResolvedTab }> {
  const viewMode = suggestionsViewMode ? `&suggestionsViewMode=${suggestionsViewMode}` : '';
  const doc = await makeDocsRequest(`/${encodeURIComponent(documentId)}?includeTabsContent=true${viewMode}`, accessToken, { method: 'GET' });
  return { doc, resolved: resolveTab(doc, tabId) };
}

export const withRefusalHint = (matches: (message: string) => boolean, hint: string, status = 400) => (err: unknown): never => {
  if (err instanceof GoogleApiError && err.status === status && matches(err.message)) {
    throw new GoogleApiError(`${err.message.replace(/(?<![.!?])$/, '.')} ${hint}`, err.status, err.api, { code: err.code, details: err.details });
  }
  throw err;
};

export async function fetchTabList(documentId: string, accessToken: string): Promise<{ tabs: TabSummary[]; revisionId?: string }> {
  const listing = await makeDocsRequest(`/${encodeURIComponent(documentId)}?fields=${TAB_LIST_FIELDS}`, accessToken, { method: 'GET' });
  return { tabs: summarizeTabs(listing?.tabs), revisionId: listing?.revisionId };
}

// Google can apply a batch yet drop its suggestions, reporting it only in commentUpdateState
export async function docsBatchUpdate(
  documentId: string,
  accessToken: string,
  requests: any[],
  options: WriteOptions = {},
): Promise<any> {
  const writeControl = buildWriteControl(options);
  const reply = await makeDocsRequest(`/${encodeURIComponent(documentId)}:batchUpdate`, accessToken, {
    method: 'POST',
    body: JSON.stringify({ requests, ...(writeControl ? { writeControl } : {}) }),
  });
  if (reply?.commentUpdateState === 'ALL_FAILED_UNKNOWN_REASON') {
    const unsaved = new Error('Google saved none of the comments or suggestions in this write (commentUpdateState ALL_FAILED_UNKNOWN_REASON)');
    if (!options.suggest) throw unsaved;
    throw new PartialFailure(
      unsaved,
      'The edit reached the document but its suggestion was not saved, so it may have landed as a direct edit. Read the document with get_document before retrying',
      { commentUpdateState: reply.commentUpdateState },
    );
  }
  return reply;
}
