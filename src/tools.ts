/**
 * Google Docs MCP Tools
 */

import { z } from 'zod';
import { withGoogleAuth as requirePermissionSecure } from "./auth.js";
import {
  buildTableInsertRequests,
  buildWriteControl,
  clearInheritedFormattingRequest,
  escapeDriveQueryName,
  extractHeadings,
  findLetterByContent,
  parseDocumentStructure,
  parseExportArtifacts,
  summarizeTabs,
  validateIndexRange,
  validateInsertIndex,
  type Reaction,
} from "./lib/docs-structure.js";
import {
  borderSchema,
  buildTableCellStyle,
  findTableAt,
  findTableStartingAt,
  hasMergedCells,
  nestingRefusal,
  normalizeTableData,
  styleFailureWarning,
  tableStartFor,
  MAX_CELL_CHARS,
  MAX_TABLE_COLUMNS,
  MAX_TABLE_ROWS,
  type DocsBody,
} from "./lib/tableModel.js";
import {
  docsOnly as wrapDocsOnly,
  fetchWordMagic,
  readWordUpload,
  toolResultWithNotice,
  NATIVE_DOC_MIME,
  WORD_READ_ONLY_NOTICE,
} from "./lib/word.js";
import {
  isWordMime,
  sniffWordFormat,
  DOCX_MIME,
  LEGACY_DOC_MIME,
  MAX_TEXT_CHARS as MAX_WORD_TEXT_CHARS,
} from "./lib/wordText.js";

const GOOGLE_DRIVE_API = 'https://www.googleapis.com/drive/v3';
const GOOGLE_DOCS_API = 'https://docs.googleapis.com/v1/documents';
const MAX_IMAGE_SIZE = 20 * 1024 * 1024; // 20MB

/**
 * Error thrown by Google API helpers. Carries enough structured detail
 * (status, Google's `error.status` enum, retry-after) for callers to surface
 * machine-readable error envelopes instead of opaque strings.
 */
class GoogleApiError extends Error {
  status: number;
  code?: string;
  retryAfter?: number;
  api: 'drive' | 'docs';
  // Google's `error.details[]` payload (e.g. `BadRequest.fieldViolations`,
  // `Help`, request-index hints). Surfaced so callers can pinpoint which
  // request in a multi-request batchUpdate actually failed.
  details?: unknown[];

  constructor(message: string, status: number, api: 'drive' | 'docs', opts: { code?: string; retryAfter?: number; details?: unknown[] } = {}) {
    super(message);
    this.name = 'GoogleApiError';
    this.status = status;
    this.api = api;
    this.code = opts.code;
    this.retryAfter = opts.retryAfter;
    this.details = opts.details;
  }
}

function safeJsonParse(text: string): any | null {
  try {
    return JSON.parse(text);
  } catch {
    return null;
  }
}

/**
 * Parse a non-OK Google API response into a `GoogleApiError`.
 *
 * Note: Google's Drive/Docs APIs return 404 for both "doesn't exist" and
 * "you don't have access" — we surface a disambiguated message so callers
 * don't make wrong assumptions.
 */
async function buildGoogleApiError(
  response: Response,
  api: 'drive' | 'docs'
): Promise<GoogleApiError> {
  const errorText = await response.text().catch(() => '');
  const errorJson = errorText ? safeJsonParse(errorText) : null;
  const googleMessage: string | undefined = errorJson?.error?.message;
  const googleCode: string | undefined = errorJson?.error?.status;
  const googleDetails: unknown[] | undefined = Array.isArray(errorJson?.error?.details) && errorJson.error.details.length > 0
    ? errorJson.error.details
    : undefined;

  let message: string;
  switch (response.status) {
    case 401:
      message = 'Authentication failed. Please re-authenticate.';
      break;
    case 403:
      message = googleMessage
        ? `Permission denied: ${googleMessage}`
        : `Permission denied. Make sure you have granted ${api === 'docs' ? 'Docs' : 'Drive'} access.`;
      break;
    case 404:
      message = api === 'docs'
        ? 'Document not found or you do not have permission to access it'
        : 'File or document not found or you do not have permission to access it';
      break;
    case 429:
      message = googleMessage || 'Rate limit exceeded. Retry after a short delay.';
      break;
    default:
      message = googleMessage || errorText || `Google ${api === 'docs' ? 'Docs' : 'Drive'} API error (${response.status})`;
  }

  let retryAfter: number | undefined;
  if (response.status === 429 || response.status === 503) {
    const header = response.headers.get('retry-after');
    if (header) {
      const parsed = parseInt(header, 10);
      if (!Number.isNaN(parsed) && parsed >= 0) retryAfter = parsed;
    }
  }

  return new GoogleApiError(message, response.status, api, { code: googleCode, retryAfter, details: googleDetails });
}

/**
 * Format any thrown error into a structured MCP error response. Handlers MUST
 * wrap their bodies in try/catch and route caught errors through this so
 * downstream callers get machine-readable `{error, status, code, retryAfter}`
 * envelopes rather than plain strings.
 */
function toolErrorResponse(err: unknown): { content: Array<{ type: 'text'; text: string }>; isError: true } {
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
    payload = { error: err.message };
  } else {
    payload = { error: String(err) };
  }
  return {
    content: [{ type: 'text', text: JSON.stringify(payload) }],
    isError: true,
  };
}

/**
 * Helper to make authenticated requests to Google Drive API
 */
async function makeDriveRequest(
  endpoint: string,
  accessToken: string,
  options: RequestInit = {}
): Promise<any> {
  const url = endpoint.startsWith('http') ? endpoint : `${GOOGLE_DRIVE_API}${endpoint}`;

  const response = await fetch(url, {
    ...options,
    headers: {
      'Authorization': `Bearer ${accessToken}`,
      'Accept': 'application/json',
      ...options.headers,
    },
  });

  if (!response.ok) {
    throw await buildGoogleApiError(response, 'drive');
  }

  return response.json();
}

/**
 * Helper to make authenticated requests to Google Docs API
 */
async function makeDocsRequest(
  endpoint: string,
  accessToken: string,
  options: RequestInit = {}
): Promise<any> {
  const url = endpoint.startsWith('http') ? endpoint : `${GOOGLE_DOCS_API}${endpoint}`;

  const response = await fetch(url, {
    ...options,
    headers: {
      'Authorization': `Bearer ${accessToken}`,
      'Content-Type': 'application/json',
      'Accept': 'application/json',
      ...options.headers,
    },
  });

  if (!response.ok) {
    throw await buildGoogleApiError(response, 'docs');
  }

  return response.json();
}

/** Refuse writes against Word uploads. Bound here so the wrapper stays pure. */
const docsOnly = <T,>(handler: (args: any, context: any) => Promise<T>) =>
  wrapDocsOnly(handler, makeDriveRequest);

interface InsertedTable {
  rows: number;
  columns: number;
  tableStartIndex: number;
  startIndexVerified: boolean;
  warning?: string;
}

/**
 * Insert a table and fill its cells, then optionally clear the formatting it
 * inherited from the insertion point.
 *
 * Once the insert batch returns the table exists, so nothing past that point
 * may throw: a retry would create a second table. Failures after it come back
 * as a warning on a successful result instead.
 */
async function insertTableAt(
  documentId: string,
  accessToken: string,
  tableData: string[][],
  index: number,
  clearInheritedFormatting: boolean,
): Promise<InsertedTable> {
  const columns = tableData[0].length;

  await makeDocsRequest(`/${encodeURIComponent(documentId)}:batchUpdate`, accessToken, {
    method: 'POST',
    body: JSON.stringify({
      requests: [
        { insertTable: { location: { index }, rows: tableData.length, columns } },
        ...buildTableInsertRequests(tableData, index),
      ],
    }),
  });

  const created = {
    rows: tableData.length,
    columns,
    tableStartIndex: tableStartFor(index),
    startIndexVerified: false,
  };

  if (!clearInheritedFormatting) {
    return created;
  }

  let table: { startIndex: number; endIndex: number } | undefined;
  try {
    const updated = await makeDocsRequest(
      `/${encodeURIComponent(documentId)}`, accessToken, { method: 'GET' },
    ) as DocsBody;
    table = findTableAt(updated, index);

    if (table) {
      await makeDocsRequest(`/${encodeURIComponent(documentId)}:batchUpdate`, accessToken, {
        method: 'POST',
        body: JSON.stringify({
          requests: [clearInheritedFormattingRequest(table.startIndex + 1, table.endIndex)],
        }),
      });
    }
  } catch (error: any) {
    // Google's error text is untrusted once the model reads it, so it stays in the log.
    console.error(
      `[gdocs-hosted] table style fail phase=${table ? 'style' : 'refetch'} ` +
      `status=${error?.status ?? 'none'} kind=${error?.name ?? 'unknown'}`,
    );
    return {
      ...created,
      startIndexVerified: Boolean(table),
      warning: styleFailureWarning(Boolean(table), error?.status),
    };
  }

  if (!table) {
    console.error(
      `[gdocs-hosted] table style fail phase=lookup msg=table not found at ${tableStartFor(index)}`,
    );
    return { ...created, warning: styleFailureWarning(false) };
  }

  return { ...created, startIndexVerified: true };
}

const tableOutput = (id: string, table: InsertedTable, message: string) => ({
  id,
  rows: table.rows,
  columns: table.columns,
  table_start_index: table.tableStartIndex,
  table_start_index_verified: table.startIndexVerified,
  message,
  ...(table.warning ? { warning: table.warning } : {}),
});

// Output schema fragments
const reportedBorderSchema = z.object({
  color: z.string().optional(),
  width: z.number().optional(),
  dash_style: z.string().optional(),
});

const reportedCellStyleSchema = z.object({
  background_color: z.string().optional(),
  border_top: reportedBorderSchema.optional(),
  border_right: reportedBorderSchema.optional(),
  border_bottom: reportedBorderSchema.optional(),
  border_left: reportedBorderSchema.optional(),
  content_alignment: z.string().optional(),
});

const structureElementSchema = z.object({
  type: z.string().optional(),
  startIndex: z.number().optional(),
  endIndex: z.number().optional(),
  text: z.string().optional(),
  inlineObjectId: z.string().optional(),
  headingLevel: z.number().int().min(1).max(6).optional(),
  rows: z.number().optional().describe('Table only: number of rows'),
  columns: z.number().optional().describe('Table only: number of columns'),
  cells: z.array(z.array(z.string())).optional().describe('Table only: cell text by row then column. Long cells are truncated and a cell holding a nested table reads as empty. Omitted for later tables once a document exhausts the structure budget, which counts both cell characters and cell count; `message` says when that happened. When has_merged_cells is true, a position covered by a merge reads as an empty string, so positions still line up with columns, but an empty cell may be a covered position rather than a genuinely empty one.'),
  has_merged_cells: z.boolean().optional().describe('Table only: the table has merged cells. Positions in `cells` still match columns, since a covered position reads as an empty string, but styling a range may affect cells outside it.'),
  styles: z.array(reportedCellStyleSchema).optional().describe('Table only, with include_table_styles=true: each distinct cell style listed once. Values are in the form update_table_style accepts, so they can be passed straight back.'),
  cell_styles: z.array(z.array(z.number())).optional().describe('Table only, with include_table_styles=true: an index into `styles` for each cell, positioned like `cells`. Cells sharing an index share a look, which is how you find the ranges to pass to update_table_style.'),
}).passthrough();

const tableToolOutputSchema = {
  id: z.string(),
  rows: z.number(),
  columns: z.number(),
  table_start_index: z.number(),
  table_start_index_verified: z.boolean().describe('True when the table was found at this index after insertion. False means the index is inferred from the API contract and was not confirmed — re-read the document before using it.'),
  message: z.string(),
  warning: z.string().optional().describe('Present only when the table was created but a follow-up step did not complete. Never re-insert the table in response to a warning.'),
};

const headingSchema = z.object({
  level: z.number().int().min(1).max(6),
  text: z.string(),
  startIndex: z.number().int(),
  endIndex: z.number().int(),
});

const tabSummarySchema = z.object({
  tabId: z.string(),
  title: z.string().optional(),
  index: z.number().int().optional(),
  nestingLevel: z.number().int().optional(),
});

const reactionSchema = z.object({
  author_name: z.string().optional(),
  emoji: z.string().optional(),
  timestamp: z.string().optional(),
}).passthrough();

const commentSchema = z.object({
  author_email: z.string().optional(),
  author_name: z.string().optional(),
  timestamp: z.string().optional(),
  content: z.string().optional(),
  reactions: z.array(reactionSchema).optional(),
}).passthrough();

const threadSchema = z.object({
  thread_id: z.string().optional().describe('Matches the `disco=` parameter in a Google Docs comment deep-link URL (`https://docs.google.com/document/d/<doc_id>/edit?disco=<thread_id>`) — use it to resolve such a link to the right thread.'),
  anchor_offset: z.object({
    start: z.number().optional(),
    end: z.number().optional(),
  }).passthrough().optional(),
  resolved: z.boolean().optional(),
  comments: z.array(commentSchema).optional(),
}).passthrough();

interface DriveCommentReply {
  id: string;
  content: string;
  author?: { displayName?: string; emailAddress?: string };
  createdTime: string;
}

interface DriveComment extends DriveCommentReply {
  quotedFileContent?: { value: string };
  resolved?: boolean;
  replies?: DriveCommentReply[];
}

interface ThreadOutput {
  // Half-open [start, end) span into `content` where the comment is
  // anchored. Absent when the comment has no quotedFileContent
  // (document-level comments) or whose quote can no longer be located in
  // the body (e.g., the doc was edited and the anchored span was deleted —
  // an "orphaned" thread). An absent `anchor_offset` is the explicit signal.
  anchor_offset?: { start: number; end: number };
  /** Drive's comment id, which is also the `disco=` deep-link parameter. */
  thread_id: string;
  resolved: boolean;
  comments: Array<{
    author_email?: string;
    author_name?: string;
    timestamp: string;
    content: string;
    reactions?: Reaction[];
  }>;
}

async function fetchAllComments(documentId: string, accessToken: string): Promise<DriveComment[]> {
  const all: DriveComment[] = [];
  let pageToken: string | undefined;
  do {
    const params = new URLSearchParams({
      fields: 'nextPageToken,comments(id,content,author(displayName,emailAddress),createdTime,quotedFileContent,resolved,replies(id,content,author(displayName,emailAddress),createdTime))',
      pageSize: '100',
    });
    if (pageToken) params.set('pageToken', pageToken);
    const result = await makeDriveRequest(
      `/files/${encodeURIComponent(documentId)}/comments?${params}`,
      accessToken
    ) as { comments?: DriveComment[]; nextPageToken?: string };
    all.push(...(result.comments || []));
    pageToken = result.nextPageToken;
  } while (pageToken);
  return all;
}


// Anchor each thread to its position in the document body using the comment's
// `quotedFileContent.value` (the snippet of doc text the comment was attached
// to). For threads with no quote (document-level comments) or quotes we can't
// locate (revision drift), assign a synthetic `t<N>` anchor that won't appear
// in the body. Threads are returned in reading order.
async function processCommentsForDocument(
  documentId: string,
  accessToken: string,
  plainText: string
): Promise<{ cleanedContent: string; threads: ThreadOutput[] }> {
  const { body: bodyOnly, letterContent, letterReactions, commentLetters } =
    parseExportArtifacts(plainText);

  const allComments = await fetchAllComments(documentId, accessToken);
  if (allComments.length === 0) {
    return { cleanedContent: bodyOnly, threads: [] };
  }

  const reactionsFor = (apiContent: string): Reaction[] | undefined => {
    const letter = findLetterByContent(letterContent, apiContent);
    const list = letter ? letterReactions.get(letter) : undefined;
    return list && list.length > 0 ? list : undefined;
  };

  // Locate each thread's quote in the pre-strip body.
  type Positioned = {
    cmt: DriveComment;
    quoteStart: number | null;
    quoteEnd: number | null;
    commentCount: number;
  };
  const positioned: Positioned[] = allComments.map((cmt) => {
    const quote = cmt.quotedFileContent?.value;
    const idx = quote ? bodyOnly.indexOf(quote) : -1;
    return {
      cmt,
      quoteStart: (quote && idx >= 0) ? idx : null,
      quoteEnd: (quote && idx >= 0) ? idx + quote.length : null,
      commentCount: 1 + (cmt.replies?.length || 0),
    };
  });

  // Group threads sharing a quote position so a single marker cluster
  // covering both can be stripped in one pass. Each group sums the marker
  // counts of its threads.
  type Group = { pos: number; expectedMarkers: number; entries: Positioned[] };
  const groups = new Map<number, Group>();
  for (const p of positioned) {
    if (p.quoteEnd === null) continue;
    let g = groups.get(p.quoteEnd);
    if (!g) {
      g = { pos: p.quoteEnd, expectedMarkers: 0, entries: [] };
      groups.set(p.quoteEnd, g);
    }
    g.expectedMarkers += p.commentCount;
    g.entries.push(p);
  }

  // Strip the export's `[<letter>]` cluster that sits immediately after each
  // thread's quote — those letters are export artifacts, not part of the
  // document's actual text. Process groups in ascending order, tracking a
  // running shift, so we can record each thread's final offset in the OUTPUT
  // body (the value the caller sees in `content`). Quote positions before a
  // strip are unaffected; positions after a strip shift left by `removedLen`.
  const ascendingPositions = [...groups.keys()].sort((a, b) => a - b);
  let shift = 0;
  let body = bodyOnly;
  const finalOffsets = new Map<Positioned, { start: number; end: number }>();

  for (const pos of ascendingPositions) {
    const group = groups.get(pos)!;
    const adjustedEnd = pos + shift;
    let end = adjustedEnd;
    let consumed = 0;
    while (consumed < group.expectedMarkers && end < body.length) {
      const m = body.slice(end).match(/^\[([a-z]+)\]/);
      if (!m || !commentLetters.has(m[1])) break;
      end += m[0].length;
      consumed++;
    }
    const removedLen = end - adjustedEnd;
    body = body.slice(0, adjustedEnd) + body.slice(end);

    // All entries in this group share the same anchor span — record final
    // offsets using the shift accumulated BEFORE this strip, since the strip
    // happens at end-of-quote and doesn't move the quote itself.
    for (const entry of group.entries) {
      finalOffsets.set(entry, {
        start: entry.quoteStart! + shift,
        end: adjustedEnd,
      });
    }
    shift -= removedLen;
  }

  // Emit threads in reading order. Anchored threads get sequential letter
  // labels (a, b, c, …); unanchored ones (no quote, or quote not findable in
  // current body — e.g., revision drift) get synthetic `t<N>` labels.
  const ordered = [...positioned].sort((a, b) => {
    if (a.quoteEnd === null && b.quoteEnd === null) return 0;
    if (a.quoteEnd === null) return 1;
    if (b.quoteEnd === null) return -1;
    return a.quoteEnd! - b.quoteEnd!;
  });

  const threads: ThreadOutput[] = ordered.map((p) => ({
    thread_id: p.cmt.id,
    anchor_offset: finalOffsets.get(p),
    resolved: !!p.cmt.resolved,
    comments: [
      {
        author_email: p.cmt.author?.emailAddress,
        author_name: p.cmt.author?.displayName,
        timestamp: p.cmt.createdTime,
        content: p.cmt.content,
        reactions: reactionsFor(p.cmt.content),
      },
      ...(p.cmt.replies || []).map((r) => ({
        author_email: r.author?.emailAddress,
        author_name: r.author?.displayName,
        timestamp: r.createdTime,
        content: r.content,
        reactions: reactionsFor(r.content),
      })),
    ],
  }));

  return { cleanedContent: body, threads };
}

/**
 * Google Docs Tools
 */
export class GoogleDocsTools {
  static getTools() {
    return {
      search_documents: {
        description: 'Search for Google Docs by name. Also finds Word (.doc/.docx) uploads, which get_document can read but the editing tools refuse; `mimeType` says which a result is. Returns matching documents with their IDs.',
        readOnlyHint: true,
        outputSchema: {
          documents: z.array(z.object({
            id: z.string(),
            name: z.string(),
            mimeType: z.string().optional().describe('A Word mimeType means the file is read-only here; call convert_to_google_doc to get an editable copy.'),
            createdTime: z.string().optional(),
            modifiedTime: z.string().optional(),
            webViewLink: z.string().optional(),
            owner: z.string().optional(),
          })),
          nextPageToken: z.string().nullable(),
        },
        schema: {
          name: z.string().describe('Search by document name (partial match)'),
          page_token: z.string().optional().describe('Token for fetching the next page of results'),
        },
        handler: requirePermissionSecure("https://www.googleapis.com/auth/drive.readonly", async ({ name, page_token }: any, context: any) => {
          try {
            const { accessToken } = context;

            // Drive query strings are single-quoted; escape backslash first
            // (so the next-step `'` escapes survive), then escape `'`, then
            // reject newlines outright — Drive's query parser does not accept
            // them and silently mismatched queries are worse than a clear error.
            if (name && /[\r\n]/.test(name)) {
              throw new Error('search name must not contain newline characters');
            }
            const escapedName = escapeDriveQueryName(name);
            // Word uploads are included because get_document can read them.
            let q = `(mimeType = '${NATIVE_DOC_MIME}' or mimeType = '${DOCX_MIME}' or mimeType = '${LEGACY_DOC_MIME}')`;
            if (escapedName) {
              q += ` and name contains '${escapedName}'`;
            }
            q += ` and trashed = false`;

            const params = new URLSearchParams({
              pageSize: '20',
              fields: 'nextPageToken,files(id,name,mimeType,createdTime,modifiedTime,webViewLink,owners)',
              supportsAllDrives: 'true',
              includeItemsFromAllDrives: 'true',
              q,
              ...(page_token && { pageToken: page_token }),
            });

            const result = await makeDriveRequest(`/files?${params}`, accessToken);

            const documents = (result.files || []).map((file: any) => ({
              id: file.id,
              name: file.name,
              mimeType: file.mimeType,
              createdTime: file.createdTime,
              modifiedTime: file.modifiedTime,
              webViewLink: file.webViewLink,
              owner: file.owners?.[0]?.emailAddress,
            }));

            const output = {
              documents,
              nextPageToken: result.nextPageToken || null,
            };
            return {
              content: [{ type: 'text', text: JSON.stringify(output, null, 2) }],
              structuredContent: output,
            };
          } catch (err) {
            return toolErrorResponse(err);
          }
        }),
      },

      get_document: {
        description: 'Read the contents of a Google Doc as plain text. Also reads Word (.doc and .docx) files uploaded to Drive, by parsing them directly — for those, `mimeType` is returned and `include_structure`/`include_comments` do not apply. Optionally include document structure with startIndex/endIndex for each element (needed for index-based editing tools like delete_content, insert_text, update_text_style, update_paragraph_style). Table elements in `structure` also carry `rows`, `columns` and `cells` (cell text by row then column), which is how you find the coordinates to pass to update_table_style; a table reporting has_merged_cells=true still lines its positions up with columns, since a merge-covered position reads as an empty string. Add include_table_styles=true to also see how each table is formatted: `styles` lists every distinct cell style once and `cell_styles` gives each cell an index into it, so cells sharing an index share a look. The values are in the form update_table_style accepts, so read a style and pass it straight back to copy a table\'s formatting. One caveat when replaying: a border reported without a `color` is written back as black, since Google needs a complete border, so check that field before copying a border across documents. When include_structure=true, the response also includes a `headings` array (level, text, startIndex, endIndex) to support "insert after the heading named X" flows without parsing the full structure. Set include_comments=true to also return comment thread metadata (author email, timestamp, replies, emoji reactions). `content` is returned verbatim with no inline markers; each thread carries `anchor_offset: { start, end }` — a half-open span into `content` indicating which text the comment was attached to. Threads with no findable position (document-level comments, or anchored text deleted by later edits) omit `anchor_offset`. If the document uses Tabs, a `tabs` summary is returned and the plain-text `content` concatenates all tabs (via Drive export); `structure` and `headings` indices refer to the FIRST tab only — index-based mutating tools in this server target the default tab. When `include_structure=true`, the response also includes the current `revisionId`; pass it as `required_revision_id` to a subsequent mutating tool to detect concurrent edits (the write will fail rather than silently overwrite). Use search_documents to find a document ID first.',
        readOnlyHint: true,
        outputSchema: {
          id: z.string(),
          title: z.string(),
          content: z.string(),
          webViewLink: z.string().optional(),
          revisionId: z.string().optional(),
          mimeType: z.string().optional().describe('Present for non-native files such as .docx uploads'),
          truncated: z.boolean().optional().describe('True when content was cut at the character cap'),
          message: z.string().optional().describe('Why content is partial, or which options do not apply'),
          structure: z.array(structureElementSchema).optional(),
          headings: z.array(headingSchema).optional(),
          tabs: z.array(tabSummarySchema).optional(),
          threads: z.array(threadSchema).optional(),
        },
        schema: {
          document_id: z.string().describe('Google Doc ID (from search_documents or a Google Docs URL)'),
          include_structure: z.boolean().optional().describe('Include document structure with startIndex/endIndex for each element. Required before using index-based tools.'),
          include_table_styles: z.boolean().optional().describe('Also report each table\'s cell styles, as a `styles` legend plus a `cell_styles` index grid. Requires include_structure=true. Use it to see how a table is currently formatted before matching or replicating it.'),
          include_comments: z.boolean().optional().describe('Include comment thread metadata (author email/name, timestamp, replies, resolved status, emoji reactions). `content` is returned verbatim — no inline markers are inserted. Each thread carries `anchor_offset: { start, end }` — a half-open span into `content` indicating which text the comment was attached to. Threads with no findable position (document-level comments, or anchor deleted by later edits) omit `anchor_offset`.'),
        },
        handler: requirePermissionSecure("https://www.googleapis.com/auth/drive.readonly", async ({ document_id, include_structure, include_comments, include_table_styles }: any, context: any) => {
          try {
            const { accessToken } = context;

            // Get file metadata for title and link
            const metadata = await makeDriveRequest(
              `/files/${encodeURIComponent(document_id)}?fields=name,mimeType,size,webViewLink&supportsAllDrives=true`,
              accessToken
            );

            const mimeType: string = metadata.mimeType || '';
            const webViewLink: string =
              metadata.webViewLink || `https://drive.google.com/file/d/${document_id}/view`;

            // Drive's export endpoint 403s on Word uploads, so these are parsed here.
            if (isWordMime(mimeType)) {
              const word = await readWordUpload({
                id: document_id,
                name: metadata.name,
                mimeType,
                size: metadata.size ? parseInt(metadata.size) : 0,
                webViewLink,
              }, accessToken, GOOGLE_DRIVE_API);

              const wordOutput: any = {
                id: document_id,
                title: metadata.name,
                webViewLink,
                content: word.text,
                mimeType,
              };
              const wordNotes: string[] = [];
              if (word.truncated) {
                wordOutput.truncated = true;
                wordNotes.push(
                  `The text was truncated at ${MAX_WORD_TEXT_CHARS} characters. Full file: ${webViewLink}`
                );
              }
              if (include_structure || include_comments) {
                wordNotes.push(
                  `'${metadata.name}' is a Word upload, not a Google Doc, ` +
                  `so document structure and comments are unavailable for it.`
                );
              }
              if (wordNotes.length) wordOutput.message = wordNotes.join(' ');

              return toolResultWithNotice(wordOutput, WORD_READ_ONLY_NOTICE);
            }

            if (mimeType && mimeType !== NATIVE_DOC_MIME) {
              throw new Error(
                `'${metadata.name}' has mimeType '${mimeType}', which this connector cannot read. ` +
                `This tool reads Google Docs and Word (.doc/.docx) files. ` +
                `Open it directly: ${webViewLink}`
              );
            }

            // Export document as plain text via Drive API.
            // Accept-Language pins the export footer (where comment reactions
            // surface as "X reacted with Y at Z") to the English template the
            // reaction parser expects.
            const exportUrl = `${GOOGLE_DRIVE_API}/files/${encodeURIComponent(document_id)}/export?mimeType=text/plain`;
            const response = await fetch(exportUrl, {
              headers: {
                'Authorization': `Bearer ${accessToken}`,
                'Accept-Language': 'en-US',
              },
            });

            if (!response.ok) {
              throw await buildGoogleApiError(response, 'drive');
            }

            let content = await response.text();

            const output: any = {
              id: document_id,
              title: metadata.name,
              webViewLink: metadata.webViewLink,
            };

            if (include_comments) {
              const { cleanedContent, threads } = await processCommentsForDocument(document_id, accessToken, content);
              content = cleanedContent;
              output.threads = threads;
            }

            output.content = content;

            // If structure requested, also fetch from Docs API.
            // `includeTabsContent=true` is required for tabbed documents —
            // without it, `tabs[*].documentTab.body.content` is omitted.
            // For tabbed documents, Docs omits the root `body.content` and
            // returns the per-tab body under `tabs[*].documentTab.body`; we
            // fall back to the first/default tab so structure and headings
            // line up with the indices our mutating tools target.
            if (include_structure) {
              const doc = await makeDocsRequest(`/${encodeURIComponent(document_id)}?includeTabsContent=true`, accessToken, { method: 'GET' });
              let body: any[] = doc.body?.content || [];
              if (body.length === 0 && Array.isArray(doc.tabs) && doc.tabs.length > 0) {
                body = doc.tabs[0]?.documentTab?.body?.content || [];
              }
              const parsed = parseDocumentStructure(body, include_table_styles === true);
              output.structure = parsed.elements;
              output.headings = extractHeadings(body);
              if (typeof doc.revisionId === 'string') {
                output.revisionId = doc.revisionId;
              }
              const tabs = summarizeTabs(doc.tabs);
              if (tabs.length > 0) {
                output.tabs = tabs;
              }
              const notes: string[] = [];
              if (parsed.tablesWithoutCells > 0) {
                notes.push(`${parsed.tablesWithoutCells} table(s) report dimensions without cell contents because the document exhausted the structure budget for table cells. Their text, if any, is in \`content\`.`);
              }
              if (parsed.tablesWithoutStyles > 0) {
                notes.push(`${parsed.tablesWithoutStyles} table(s) report no styles because the document exhausted the structure budget. Read those tables on their own to see their formatting.`);
              }
              if (notes.length) output.message = notes.join(' ');
            } else if (include_table_styles) {
              output.message = 'include_table_styles needs include_structure=true, since styles are reported on the table elements in `structure`. No structure was returned.';
            }

            return {
              content: [{ type: 'text', text: JSON.stringify(output, null, 2) }],
              structuredContent: output,
            };
          } catch (err) {
            return toolErrorResponse(err);
          }
        }),
      },

      create_document: {
        description: 'Create a new Google Doc with optional initial text content. Optionally place it in a specific folder (including shared drive folders).',
        outputSchema: {
          id: z.string(),
          title: z.string(),
          webViewLink: z.string(),
          message: z.string(),
        },
        schema: {
          title: z.string().describe('Title for the new document'),
          body: z.string().optional().describe('Optional initial plain text content'),
          parent_folder_id: z.string().optional().describe('ID of the folder to create the document in (supports shared drive folders)'),
        },
        handler: requirePermissionSecure("https://www.googleapis.com/auth/documents", async ({ title, body, parent_folder_id }: any, context: any) => {
          let file: { id: string; name: string } | undefined;
          try {
            const { accessToken } = context;

            // Create the document via Drive API so we can specify parent folder
            const fileMetadata: any = {
              name: title,
              mimeType: 'application/vnd.google-apps.document',
            };
            if (parent_folder_id) {
              fileMetadata.parents = [parent_folder_id];
            }

            file = await makeDriveRequest(
              `/files?supportsAllDrives=true`,
              accessToken,
              {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(fileMetadata),
              }
            ) as { id: string; name: string };

            // If body text provided, insert it via Docs API.
            // The Drive file has been created at this point — if this fails the
            // empty document still exists, so surface partialSuccess so the
            // caller can find it or clean up rather than re-creating.
            if (body) {
              try {
                await makeDocsRequest(`/${encodeURIComponent(file.id)}:batchUpdate`, accessToken, {
                  method: 'POST',
                  body: JSON.stringify({
                    requests: [{
                      insertText: {
                        location: { index: 1 },
                        text: body,
                      },
                    }],
                  }),
                });
              } catch (insertErr) {
                const base = toolErrorResponse(insertErr);
                const parsed = safeJsonParse(base.content[0].text) || {};
                const payload = {
                  ...parsed,
                  error: `Document was created but inserting body text failed: ${parsed.error ?? String(insertErr)}`,
                  partialSuccess: true,
                  fileId: file.id,
                  webViewLink: `https://docs.google.com/document/d/${file.id}`,
                };
                return {
                  content: [{ type: 'text', text: JSON.stringify(payload) }],
                  isError: true,
                };
              }
            }

            const webViewLink = `https://docs.google.com/document/d/${file.id}`;

            const output = {
              id: file.id,
              title: file.name,
              webViewLink,
              message: 'Document created successfully',
            };
            return {
              content: [{ type: 'text', text: JSON.stringify(output, null, 2) }],
              structuredContent: output,
            };
          } catch (err) {
            return toolErrorResponse(err);
          }
        }),
      },

      append_text: {
        description: 'Append plain text to the end of a Google Doc. Use get_document first to verify the document exists and see current content. Optionally pass `required_revision_id` (from get_document with include_structure=true) to fail the write if the document was edited concurrently.',
        outputSchema: {
          id: z.string(),
          message: z.string(),
        },
        schema: {
          document_id: z.string().describe('Google Doc ID'),
          text: z.string().describe('Plain text to append to the end of the document'),
          clear_inherited_formatting: z.boolean().optional().describe('Clear inherited formatting from preceding text on the newly appended content. Defaults to true.'),
          required_revision_id: z.string().optional().describe('If provided, the write fails with a 400 if the document revision has changed since this revisionId. Use the `revisionId` returned by get_document(include_structure=true) for optimistic concurrency control.'),
        },
        handler: requirePermissionSecure("https://www.googleapis.com/auth/documents", docsOnly(async ({ document_id, text, clear_inherited_formatting, required_revision_id }: any, context: any) => {
          try {
            const { accessToken } = context;

            // Get document to find the end index
            const doc = await makeDocsRequest(`/${encodeURIComponent(document_id)}`, accessToken, { method: 'GET' }) as { body: { content: Array<{ endIndex: number }> } };
            const endIndex = doc.body.content[doc.body.content.length - 1].endIndex - 1;

            // Insert text at the end, optionally clearing inherited formatting
            const requests: any[] = [{
              insertText: {
                location: { index: endIndex },
                text,
              },
            }];

            if (clear_inherited_formatting !== false) {
              requests.push(clearInheritedFormattingRequest(endIndex, endIndex + text.length));
            }

            const writeControl = buildWriteControl(required_revision_id);
            await makeDocsRequest(`/${encodeURIComponent(document_id)}:batchUpdate`, accessToken, {
              method: 'POST',
              body: JSON.stringify({ requests, ...(writeControl ? { writeControl } : {}) }),
            });

            const output = {
              id: document_id,
              message: 'Text appended successfully',
            };
            return {
              content: [{ type: 'text', text: JSON.stringify(output, null, 2) }],
              structuredContent: output,
            };
          } catch (err) {
            return toolErrorResponse(err);
          }
        })),
      },

      replace_text: {
        description: 'Replace all occurrences of a text string in a Google Doc. Use get_document first to see current content and verify the text to replace exists. Use empty new_text to delete occurrences. Optionally pass `required_revision_id` (from get_document with include_structure=true) to fail the write if the document was edited concurrently.',
        outputSchema: {
          id: z.string(),
          occurrencesChanged: z.number(),
          message: z.string(),
        },
        schema: {
          document_id: z.string().describe('Google Doc ID'),
          old_text: z.string().describe('Text to find (all occurrences will be replaced)'),
          new_text: z.string().describe('Replacement text (empty string to delete)'),
          match_case: z.boolean().optional().describe('Whether to match case (default true)'),
          required_revision_id: z.string().optional().describe('If provided, the write fails with a 400 if the document revision has changed since this revisionId. Use the `revisionId` returned by get_document(include_structure=true) for optimistic concurrency control.'),
        },
        handler: requirePermissionSecure("https://www.googleapis.com/auth/documents", docsOnly(async ({ document_id, old_text, new_text, match_case, required_revision_id }: any, context: any) => {
          try {
            const { accessToken } = context;

            // Docs API returns 400 on empty `old_text`; reject up-front with a
            // clear structured error.
            if (typeof old_text !== 'string' || old_text.length === 0) {
              throw new Error('old_text must be a non-empty string');
            }

            const writeControl = buildWriteControl(required_revision_id);
            const result = await makeDocsRequest(`/${encodeURIComponent(document_id)}:batchUpdate`, accessToken, {
              method: 'POST',
              body: JSON.stringify({
                requests: [{
                  replaceAllText: {
                    replaceText: new_text,
                    containsText: {
                      text: old_text,
                      matchCase: match_case !== false,
                    },
                  },
                }],
                ...(writeControl ? { writeControl } : {}),
              }),
            }) as { replies: Array<{ replaceAllText?: { occurrencesChanged: number } }> };

            const occurrencesChanged = result.replies?.[0]?.replaceAllText?.occurrencesChanged || 0;

            const output = {
              id: document_id,
              occurrencesChanged,
              message: occurrencesChanged > 0
                ? `Replaced ${occurrencesChanged} occurrence(s)`
                : 'No occurrences found',
            };
            return {
              content: [{ type: 'text', text: JSON.stringify(output, null, 2) }],
              structuredContent: output,
            };
          } catch (err) {
            return toolErrorResponse(err);
          }
        })),
      },

      delete_content: {
        description: 'Delete content in a Google Doc by index range. Use get_document with include_structure=true first to find the correct startIndex and endIndex. Optionally pass `required_revision_id` (from get_document with include_structure=true) to fail the write if the document was edited concurrently.',
        destructiveHint: true,
        outputSchema: {
          id: z.string(),
          message: z.string(),
        },
        schema: {
          document_id: z.string().describe('Google Doc ID'),
          startIndex: z.coerce.number().int().describe('Start index of content to delete (use get_document with include_structure to find indices)'),
          endIndex: z.coerce.number().int().describe('End index of content to delete (exclusive)'),
          required_revision_id: z.string().optional().describe('If provided, the write fails with a 400 if the document revision has changed since this revisionId. Use the `revisionId` returned by get_document(include_structure=true) for optimistic concurrency control.'),
        },
        handler: requirePermissionSecure("https://www.googleapis.com/auth/documents", docsOnly(async ({ document_id, startIndex, endIndex, required_revision_id }: any, context: any) => {
          try {
            const { accessToken } = context;

            validateIndexRange(startIndex, endIndex);

            const writeControl = buildWriteControl(required_revision_id);
            await makeDocsRequest(`/${encodeURIComponent(document_id)}:batchUpdate`, accessToken, {
              method: 'POST',
              body: JSON.stringify({
                requests: [{
                  deleteContentRange: {
                    range: {
                      startIndex,
                      endIndex,
                      segmentId: '',
                    },
                  },
                }],
                ...(writeControl ? { writeControl } : {}),
              }),
            });

            const output = {
              id: document_id,
              message: `Deleted content from index ${startIndex} to ${endIndex}`,
            };
            return {
              content: [{ type: 'text', text: JSON.stringify(output, null, 2) }],
              structuredContent: output,
            };
          } catch (err) {
            return toolErrorResponse(err);
          }
        })),
      },

      insert_text: {
        description: 'Insert text at a specific position in a Google Doc. Use get_document with include_structure=true to find the correct index. NOTE: Inserted text inherits formatting from surrounding text at the insertion point. Use clear_inherited_formatting=true to reset to plain formatting. Optionally pass `required_revision_id` (from get_document with include_structure=true) to fail the write if the document was edited concurrently.',
        outputSchema: {
          id: z.string(),
          message: z.string(),
        },
        schema: {
          document_id: z.string().describe('Google Doc ID'),
          text: z.string().describe('Text to insert'),
          index: z.coerce.number().int().describe('Position to insert at (use get_document with include_structure to find indices)'),
          clear_inherited_formatting: z.boolean().optional().describe('Clear inherited formatting from surrounding text on the newly inserted content. Defaults to false.'),
          required_revision_id: z.string().optional().describe('If provided, the write fails with a 400 if the document revision has changed since this revisionId. Use the `revisionId` returned by get_document(include_structure=true) for optimistic concurrency control.'),
        },
        handler: requirePermissionSecure("https://www.googleapis.com/auth/documents", docsOnly(async ({ document_id, text, index, clear_inherited_formatting, required_revision_id }: any, context: any) => {
          try {
            const { accessToken } = context;

            // Docs body content starts at index 1 (index 0 is the document
            // start sentinel); inserting at 0 always 400s.
            validateInsertIndex(index);

            const requests: any[] = [{
              insertText: {
                location: { index },
                text,
              },
            }];

            if (clear_inherited_formatting === true) {
              requests.push(clearInheritedFormattingRequest(index, index + text.length));
            }

            const writeControl = buildWriteControl(required_revision_id);
            await makeDocsRequest(`/${encodeURIComponent(document_id)}:batchUpdate`, accessToken, {
              method: 'POST',
              body: JSON.stringify({ requests, ...(writeControl ? { writeControl } : {}) }),
            });

            const output = {
              id: document_id,
              message: `Text inserted at index ${index}`,
            };
            return {
              content: [{ type: 'text', text: JSON.stringify(output, null, 2) }],
              structuredContent: output,
            };
          } catch (err) {
            return toolErrorResponse(err);
          }
        })),
      },

      update_text_style: {
        description: 'Apply formatting (bold, italic, underline, strikethrough, link) to a text range in a Google Doc. Use get_document with include_structure=true to find the correct indices. NOTE: If you also need to set a heading level via update_paragraph_style, do that first — heading changes reset character formatting. Optionally pass `required_revision_id` (from get_document with include_structure=true) to fail the write if the document was edited concurrently.',
        outputSchema: {
          id: z.string(),
          message: z.string(),
        },
        schema: {
          document_id: z.string().describe('Google Doc ID'),
          startIndex: z.coerce.number().int().describe('Start index of text range'),
          endIndex: z.coerce.number().int().describe('End index of text range (exclusive)'),
          bold: z.boolean().optional().describe('Set bold'),
          italic: z.boolean().optional().describe('Set italic'),
          underline: z.boolean().optional().describe('Set underline'),
          strikethrough: z.boolean().optional().describe('Set strikethrough'),
          link_url: z.string().optional().describe('Set hyperlink URL'),
          required_revision_id: z.string().optional().describe('If provided, the write fails with a 400 if the document revision has changed since this revisionId. Use the `revisionId` returned by get_document(include_structure=true) for optimistic concurrency control.'),
        },
        handler: requirePermissionSecure("https://www.googleapis.com/auth/documents", docsOnly(async ({ document_id, startIndex, endIndex, bold, italic, underline, strikethrough, link_url, required_revision_id }: any, context: any) => {
          try {
            const { accessToken } = context;

            validateIndexRange(startIndex, endIndex);

            // Build textStyle and fields dynamically from provided params
            const textStyle: any = {};
            const fields: string[] = [];

            if (bold !== undefined) {
              textStyle.bold = bold;
              fields.push('bold');
            }
            if (italic !== undefined) {
              textStyle.italic = italic;
              fields.push('italic');
            }
            if (underline !== undefined) {
              textStyle.underline = underline;
              fields.push('underline');
            }
            if (strikethrough !== undefined) {
              textStyle.strikethrough = strikethrough;
              fields.push('strikethrough');
            }
            if (link_url !== undefined) {
              // Empty string => remove the link, otherwise set it.
              if (link_url === '') {
                textStyle.link = null;
              } else {
                textStyle.link = { url: link_url };
              }
              fields.push('link');
            }

            if (fields.length === 0) {
              throw new Error('At least one style property must be provided (bold, italic, underline, strikethrough, link_url)');
            }

            const writeControl = buildWriteControl(required_revision_id);
            await makeDocsRequest(`/${encodeURIComponent(document_id)}:batchUpdate`, accessToken, {
              method: 'POST',
              body: JSON.stringify({
                requests: [{
                  updateTextStyle: {
                    range: { startIndex, endIndex },
                    textStyle,
                    fields: fields.join(','),
                  },
                }],
                ...(writeControl ? { writeControl } : {}),
              }),
            });

            const output = {
              id: document_id,
              message: `Applied text style (${fields.join(', ')}) to range ${startIndex}-${endIndex}`,
            };
            return {
              content: [{ type: 'text', text: JSON.stringify(output, null, 2) }],
              structuredContent: output,
            };
          } catch (err) {
            return toolErrorResponse(err);
          }
        })),
      },

      update_paragraph_style: {
        description: 'Apply paragraph formatting (heading level, alignment) to a range in a Google Doc. Use get_document with include_structure=true to find the correct indices. WARNING: Setting heading_level applies a named style that resets character-level formatting (bold, italic, etc.). If you also need to apply text styles, call update_paragraph_style first, then apply text styles after. Optionally pass `required_revision_id` (from get_document with include_structure=true) to fail the write if the document was edited concurrently.',
        outputSchema: {
          id: z.string(),
          message: z.string(),
        },
        schema: {
          document_id: z.string().describe('Google Doc ID'),
          startIndex: z.coerce.number().int().describe('Start index of paragraph range'),
          endIndex: z.coerce.number().int().describe('End index of paragraph range (exclusive)'),
          heading_level: z.coerce.number().int().min(0).max(6).optional().describe('Heading level: 0=normal text, 1-6=heading levels'),
          alignment: z.enum(['START', 'CENTER', 'END', 'JUSTIFIED']).optional().describe('Paragraph alignment'),
          required_revision_id: z.string().optional().describe('If provided, the write fails with a 400 if the document revision has changed since this revisionId. Use the `revisionId` returned by get_document(include_structure=true) for optimistic concurrency control.'),
        },
        handler: requirePermissionSecure("https://www.googleapis.com/auth/documents", docsOnly(async ({ document_id, startIndex, endIndex, heading_level, alignment, required_revision_id }: any, context: any) => {
          try {
            const { accessToken } = context;

            validateIndexRange(startIndex, endIndex);

            const paragraphStyle: any = {};
            const fields: string[] = [];

            if (heading_level !== undefined) {
              const headingMap: Record<number, string> = {
                0: 'NORMAL_TEXT',
                1: 'HEADING_1',
                2: 'HEADING_2',
                3: 'HEADING_3',
                4: 'HEADING_4',
                5: 'HEADING_5',
                6: 'HEADING_6',
              };
              const namedStyle = headingMap[heading_level];
              if (!namedStyle) {
                throw new Error('heading_level must be 0 (normal) or 1-6');
              }
              paragraphStyle.namedStyleType = namedStyle;
              fields.push('namedStyleType');
            }

            if (alignment !== undefined) {
              paragraphStyle.alignment = alignment;
              fields.push('alignment');
            }

            if (fields.length === 0) {
              throw new Error('At least one style property must be provided (heading_level, alignment)');
            }

            const writeControl = buildWriteControl(required_revision_id);
            await makeDocsRequest(`/${encodeURIComponent(document_id)}:batchUpdate`, accessToken, {
              method: 'POST',
              body: JSON.stringify({
                requests: [{
                  updateParagraphStyle: {
                    range: { startIndex, endIndex },
                    paragraphStyle,
                    fields: fields.join(','),
                  },
                }],
                ...(writeControl ? { writeControl } : {}),
              }),
            });

            const output = {
              id: document_id,
              message: `Applied paragraph style (${fields.join(', ')}) to range ${startIndex}-${endIndex}`,
            };
            return {
              content: [{ type: 'text', text: JSON.stringify(output, null, 2) }],
              structuredContent: output,
            };
          } catch (err) {
            return toolErrorResponse(err);
          }
        })),
      },

      append_table: {
        description: 'Insert a table with data at the end of a Google Doc. Supports ragged rows (will be padded with empty cells). Optionally pass `required_revision_id` (from get_document with include_structure=true) to fail the initial insert if the document was edited concurrently; the follow-up formatting clear is not revision-gated since it operates on the just-inserted table. If this call fails, check the document with get_document before retrying — a table that was already created would otherwise be inserted twice. On success the result reports table_start_index plus table_start_index_verified; if a warning comes back, the table already exists — never insert it again. Native Google Docs only: Word (.doc/.docx) uploads are read-only and this tool refuses them.',
        outputSchema: {
          id: z.string(),
          rows: z.number(),
          columns: z.number(),
          table_start_index: z.number(),
          table_start_index_verified: z.boolean().describe('True when the table was found at this index after insertion. False means the index is inferred from the API contract and was not confirmed — re-read the document before using it.'),
          message: z.string(),
          warning: z.string().optional(),
        },
        schema: {
          document_id: z.string().describe('Google Doc ID'),
          rows: z.array(z.array(z.string().max(MAX_CELL_CHARS)).max(MAX_TABLE_COLUMNS)).max(MAX_TABLE_ROWS).describe('Table data as array of rows, each row is array of cell strings'),
          clear_inherited_formatting: z.boolean().optional().describe('Clear inherited formatting on the newly appended table content. Defaults to true.'),
          required_revision_id: z.string().optional().describe('If provided, the initial table insert fails with a 400 if the document revision has changed since this revisionId. Use the `revisionId` returned by get_document(include_structure=true) for optimistic concurrency control.'),
        },
        handler: requirePermissionSecure("https://www.googleapis.com/auth/documents", docsOnly(async ({ document_id, rows, clear_inherited_formatting, required_revision_id }: any, context: any) => {
          try {
            const { accessToken } = context;

            const tableData = normalizeTableData(rows);
            const maxCols = tableData[0].length;

            // Get document to find the end index
            const doc = await makeDocsRequest(`/${encodeURIComponent(document_id)}`, accessToken, { method: 'GET' }) as { body: { content: Array<{ endIndex: number }> } };
            const endOfDoc = doc.body.content[doc.body.content.length - 1].endIndex;
            const insertIndex = Math.max(1, endOfDoc - 1);

            // Build requests: first insert empty table, then populate cells in reverse order
            const requests: any[] = [
              {
                insertTable: {
                  location: { index: insertIndex },
                  rows: tableData.length,
                  columns: maxCols,
                },
              },
              ...buildTableInsertRequests(tableData, insertIndex),
            ];

            const writeControl = buildWriteControl(required_revision_id);
            await makeDocsRequest(`/${encodeURIComponent(document_id)}:batchUpdate`, accessToken, {
              method: 'POST',
              body: JSON.stringify({ requests, ...(writeControl ? { writeControl } : {}) }),
            });

            // Clear inherited formatting on the newly inserted table — best-effort.
            // The table is already inserted at this point; if this second
            // batchUpdate fails, we still want to surface success and warn
            // about the formatting step rather than fail the whole operation.
            //
            // Locate OUR table by matching the known insertion point AND shape.
            // A naive `lastTable` lookup would strip formatting from a
            // different table if a concurrent edit appended one after our
            // insert; this also accounts for the rare case where another
            // client inserted at the same stale end position by additionally
            // matching rows/columns. If more than one table still satisfies
            // both filters, we refuse to guess and surface a warning instead
            // of silently formatting the wrong table.
            let formattingWarning: string | undefined;
            // The API contract puts the table one past the insertion index;
            // locating it below upgrades this from inferred to confirmed.
            let tableStartIndex = tableStartFor(insertIndex);
            let tableStartIndexVerified = false;
            if (clear_inherited_formatting !== false) {
              try {
                const updatedDoc = await makeDocsRequest(`/${encodeURIComponent(document_id)}`, accessToken, { method: 'GET' });
                const allTables = (updatedDoc.body?.content || []).filter((el: any) => el.table);
                const candidates = allTables.filter((t: any) =>
                  typeof t.startIndex === 'number' &&
                  typeof t.endIndex === 'number' &&
                  t.startIndex >= insertIndex - 1 &&
                  t.startIndex <= insertIndex + 1 &&
                  t.table?.rows === tableData.length &&
                  t.table?.columns === maxCols
                );
                if (candidates.length === 1) {
                  const ourTable = candidates[0];
                  tableStartIndex = ourTable.startIndex;
                  tableStartIndexVerified = true;
                  await makeDocsRequest(`/${encodeURIComponent(document_id)}:batchUpdate`, accessToken, {
                    method: 'POST',
                    body: JSON.stringify({
                      requests: [clearInheritedFormattingRequest(ourTable.startIndex + 1, ourTable.endIndex)],
                    }),
                  });
                } else if (candidates.length === 0) {
                  formattingWarning = 'Could not locate the inserted table to clear formatting (document may have been edited concurrently).';
                } else {
                  formattingWarning = `Skipped clearing formatting: ${candidates.length} tables match the insertion point and shape; refusing to guess which is ours.`;
                }
              } catch (formattingErr) {
                formattingWarning = formattingErr instanceof Error ? formattingErr.message : String(formattingErr);
              }
            }

            const output: {
              id: string;
              rows: number;
              columns: number;
              table_start_index: number;
              table_start_index_verified: boolean;
              message: string;
              warning?: string;
            } = {
              id: document_id,
              rows: tableData.length,
              columns: maxCols,
              table_start_index: tableStartIndex,
              table_start_index_verified: tableStartIndexVerified,
              message: `Table inserted with ${tableData.length} rows and ${maxCols} columns`,
            };
            if (formattingWarning) {
              output.warning = `Table inserted, but clearing inherited formatting failed: ${formattingWarning}. Do not insert the table again.`;
            }
            return {
              content: [{ type: 'text', text: JSON.stringify(output, null, 2) }],
              structuredContent: output,
            };
          } catch (err) {
            return toolErrorResponse(err);
          }
        })),
      },

      insert_table: {
        description: 'Insert a table with data at a specific index in a Google Doc — use this for any position other than the end of the document; index 1 is the top of the document body. To place a table between existing content, call get_document with include_structure=true and pass the startIndex of the paragraph the table should appear above. Google accepts most indices, so a wrong one usually misplaces the table instead of failing: an index inside a paragraph splits that paragraph, and an index inside a table cell nests the new table there (nested tables cannot be found afterwards, so their inherited formatting is left alone and table_start_index comes back unverified). If this call fails, check the document with get_document before retrying — a table that was already created would otherwise be inserted twice. On success the result reports table_start_index plus table_start_index_verified; if a warning comes back, the table already exists — never insert it again. Supports ragged rows (will be padded with empty cells). Native Google Docs only: Word (.doc/.docx) uploads are read-only and this tool refuses them.',
        outputSchema: tableToolOutputSchema,
        schema: {
          document_id: z.string().describe('Google Doc ID'),
          rows: z.array(z.array(z.string().max(MAX_CELL_CHARS)).max(MAX_TABLE_COLUMNS)).max(MAX_TABLE_ROWS).describe('Table data as array of rows, each row is array of cell strings. Pass empty rows (e.g. [[], [], []]) with columns set to create an empty table.'),
          index: z.number().int().min(1).describe('Document index to insert the table at. 1 is the top of the document body. Use get_document with include_structure=true to find indices.'),
          columns: z.number().int().min(1).max(MAX_TABLE_COLUMNS).optional().describe('Force the table width; short rows are padded with empty cells. Defaults to the widest row. Must not be smaller than the widest row.'),
          clear_inherited_formatting: z.boolean().optional().describe('Clear formatting inherited from the insertion point on the new table content. Defaults to true.'),
          allow_nested: z.boolean().optional().describe('Allow an index that falls inside an existing table, which nests the new table in that cell. Defaults to false, which refuses such an index and tells you which ones would work.'),
        },
        handler: requirePermissionSecure("https://www.googleapis.com/auth/documents", docsOnly(async ({ document_id, rows, index, columns, clear_inherited_formatting, allow_nested }: any, context: any) => {
          try {
            const { accessToken } = context;

            if (!Number.isInteger(index)) {
              throw new Error('index must be a whole number. Use get_document with include_structure=true to find valid indices.');
            }
            validateInsertIndex(index);

            const tableData = normalizeTableData(rows, columns);

            if (allow_nested !== true) {
              const doc = await makeDocsRequest(
                `/${encodeURIComponent(document_id)}`, accessToken, { method: 'GET' },
              ) as DocsBody;
              const refusal = nestingRefusal(doc, index);
              if (refusal) {
                throw new Error(refusal);
              }
            }

            const table = await insertTableAt(
              document_id,
              accessToken,
              tableData,
              index,
              clear_inherited_formatting !== false,
            );

            const output = tableOutput(
              document_id,
              table,
              `Table inserted with ${table.rows} rows and ${table.columns} columns` +
                (table.startIndexVerified ? `, starting at index ${table.tableStartIndex}` : ''),
            );
            return {
              content: [{ type: 'text', text: JSON.stringify(output, null, 2) }],
              structuredContent: output,
            };
          } catch (err) {
            return toolErrorResponse(err);
          }
        })),
      },

      update_table_style: {
        description: 'Style the cells of an existing table in a Google Doc: background colour, borders and vertical alignment. Pass table_start_index, which is the table\'s own start index: use the value insert_table or append_table returned, or the startIndex of a table element from get_document with include_structure=true (that element also reports rows, columns and cells so you can pick coordinates). This is NOT the index you pass to insert_table, which sits one below the table it creates. The cell range defaults to the whole table; narrow it with row_index, column_index, row_span and column_span, all zero-based. Colours are 6-digit hex such as #1A73E8. A `border` value applies to all four sides and any border_top/border_right/border_bottom/border_left replaces that side. Vertically adjacent cells share one edge, so sending `border` together with a per-side override loses that override wherever cells meet: the next row\'s top is written over the previous row\'s bottom, and only the outer edge of the table keeps it. To underline a header row inside a bordered grid, send the grid first and the override in a second call. content_alignment only shows on a row taller than its text, so on a single line table it changes nothing visible. This call creates nothing and applying the same style twice has the same effect as applying it once, so unlike insert_table a failure is safe to retry as-is. The write is pinned to the revision it validated against, so an edit by someone else between the two makes this fail rather than style the wrong cells. If the table has merged cells Google may style beyond the requested range, and a warning says so. Native Google Docs only: Word (.doc/.docx) uploads are read-only and this tool refuses them.',
        outputSchema: {
          id: z.string(),
          table_start_index: z.number(),
          rows: z.number(),
          columns: z.number(),
          cells_styled: z.number().describe('Size of the requested range. With merged cells Google may style more than this.'),
          applied: z.array(z.string()).describe('The style properties actually set'),
          message: z.string(),
          warning: z.string().optional().describe('Present when the result may not match the requested range exactly'),
        },
        schema: {
          document_id: z.string().describe('Google Doc ID'),
          table_start_index: z.number().int().min(1).describe('Start index of the table itself, from insert_table/append_table or a table element in get_document structure. Not the index passed to insert_table, which is one lower.'),
          row_index: z.number().int().min(0).optional().describe('First row of the range, zero-based. Defaults to 0.'),
          column_index: z.number().int().min(0).optional().describe('First column of the range, zero-based. Defaults to 0.'),
          row_span: z.number().int().min(1).optional().describe('How many rows to style. Defaults to the rest of the table.'),
          column_span: z.number().int().min(1).optional().describe('How many columns to style. Defaults to the rest of the table.'),
          background_color: z.string().optional().describe('Cell background as 6-digit hex, e.g. #EEEEEE'),
          border: borderSchema.optional().describe('Border applied to all four sides'),
          border_top: borderSchema.optional().describe('Overrides `border` for the top side'),
          border_right: borderSchema.optional().describe('Overrides `border` for the right side'),
          border_bottom: borderSchema.optional().describe('Overrides `border` for the bottom side'),
          border_left: borderSchema.optional().describe('Overrides `border` for the left side'),
          content_alignment: z.enum(['TOP', 'MIDDLE', 'BOTTOM']).optional().describe('Vertical alignment of cell content'),
        },
        handler: requirePermissionSecure("https://www.googleapis.com/auth/documents", docsOnly(async (args: any, context: any) => {
          try {
            const { document_id, table_start_index, row_index, column_index, row_span, column_span } = args;
            const { accessToken } = context;

            if (!Number.isInteger(table_start_index)) {
              throw new Error('table_start_index must be a whole number. Use get_document with include_structure=true and take the startIndex of a table element.');
            }

            const { tableCellStyle, fields } = buildTableCellStyle(args);
            if (fields.length === 0) {
              throw new Error('At least one style property must be provided (background_color, border, border_top, border_right, border_bottom, border_left, content_alignment).');
            }

            const doc = await makeDocsRequest(
              `/${encodeURIComponent(document_id)}`, accessToken, { method: 'GET' },
            ) as DocsBody;
            const element = findTableStartingAt(doc, table_start_index);
            if (!element) {
              const tables = (doc?.body?.content ?? []).filter((c) => c.table);
              // An index inside a table is a nested table, which the Docs API does not
              // expose as its own element, so no lookup can ever reach it.
              const enclosing = tables.find(
                (c) => (c.startIndex ?? 0) < table_start_index && table_start_index < c.endIndex
              );
              if (enclosing) {
                throw new Error(
                  `Index ${table_start_index} is inside the table at ${enclosing.startIndex}-${enclosing.endIndex}, ` +
                  `so it refers to a table nested in one of its cells. Nested tables cannot be styled. ` +
                  `Style the outer table at ${enclosing.startIndex} instead.`
                );
              }
              const starts = tables.map((c) => c.startIndex);
              const listed = starts.slice(0, 20).join(', ');
              const rest = starts.length > 20 ? ` and ${starts.length - 20} more` : '';
              throw new Error(
                `No table starts at index ${table_start_index}. ` +
                (starts.length
                  ? `Tables here start at ${listed}${rest}.`
                  : 'No tables were found.') +
                ' Only the first tab of the document is read.'
              );
            }

            const rows = element.table?.rows ?? 0;
            const columns = element.table?.columns ?? 0;
            const startRow = row_index ?? 0;
            const startColumn = column_index ?? 0;
            const size = `The table at index ${table_start_index} has ${rows} rows and ${columns} columns`;

            // Check the start before deriving a span from it, or the default span
            // goes negative and the range reads backwards.
            if (startRow >= rows || startColumn >= columns) {
              throw new Error(
                `${size}, so row_index ${startRow} and column_index ${startColumn} are out of range. ` +
                `Both are zero-based, so the last cell is row ${rows - 1}, column ${columns - 1}.`
              );
            }

            const rowSpan = row_span ?? rows - startRow;
            const columnSpan = column_span ?? columns - startColumn;

            if (startRow + rowSpan > rows || startColumn + columnSpan > columns) {
              throw new Error(
                `The requested range (rows ${startRow}-${startRow + rowSpan - 1}, columns ${startColumn}-${startColumn + columnSpan - 1}) ` +
                `does not fit. ${size}.`
              );
            }

            // The bounds above were checked against this revision, so pin the write to it
            // rather than let Google apply the range to a table a collaborator has resized.
            const body: any = {
              requests: [{
                updateTableCellStyle: {
                  tableRange: {
                    tableCellLocation: {
                      tableStartLocation: { index: table_start_index },
                      rowIndex: startRow,
                      columnIndex: startColumn,
                    },
                    rowSpan,
                    columnSpan,
                  },
                  tableCellStyle,
                  fields: fields.join(','),
                },
              }],
            };
            const writeControl = buildWriteControl(doc.revisionId);
            if (writeControl) {
              body.writeControl = writeControl;
            }

            await makeDocsRequest(`/${encodeURIComponent(document_id)}:batchUpdate`, accessToken, {
              method: 'POST',
              body: JSON.stringify(body),
            });

            const requested = rowSpan * columnSpan;
            const output: any = {
              id: document_id,
              table_start_index,
              rows,
              columns,
              cells_styled: requested,
              applied: fields,
              message: `Styled ${requested} cell(s) (${fields.join(', ')}) in the table at index ${table_start_index}`,
            };
            const warnings: string[] = [];
            if (!doc.revisionId) {
              warnings.push(
                'The document reported no revision, so this write could not be pinned to the revision it was validated against. ' +
                'A concurrent edit could have changed the table between the check and the write.'
              );
            }
            if (hasMergedCells(element)) {
              warnings.push(
                'This table contains merged cells, so Google may have styled cells outside the requested range. ' +
                'cells_styled counts the range that was asked for, not necessarily the cells affected. ' +
                'Read the document back if the exact extent matters.'
              );
            }
            if (warnings.length) {
              output.warning = warnings.join(' ');
            }
            return {
              content: [{ type: 'text', text: JSON.stringify(output, null, 2) }],
              structuredContent: output,
            };
          } catch (err) {
            return toolErrorResponse(err);
          }
        })),
      },

      convert_to_google_doc: {
        description:
          'Convert an uploaded Word (.doc or .docx) file into a NEW, editable native Google Doc. ' +
          'The original Word file is left untouched. Use this when the user wants to edit a file ' +
          'that search_documents or get_document reported with a Word mimeType. Drive performs the ' +
          'conversion, so formatting, tables and images are preserved — far better than re-typing ' +
          'the text into a new document.',
        outputSchema: {
          id: z.string().describe('ID of the new native Google Doc'),
          name: z.string(),
          webViewLink: z.string(),
          sourceId: z.string().describe('The Word file this was converted from, unchanged'),
          message: z.string(),
        },
        schema: {
          file_id: z.string().describe('Drive file ID of the .doc or .docx to convert'),
          name: z.string().optional().describe('Name for the new Doc (defaults to the original name)'),
        },
        handler: requirePermissionSecure("https://www.googleapis.com/auth/drive.file", async ({ file_id, name }: any, context: any) => {
          try {
            const { accessToken } = context;

            const meta = await makeDriveRequest(
              `/files/${encodeURIComponent(file_id)}?fields=name,mimeType,webViewLink&supportsAllDrives=true`,
              accessToken
            );
            const sourceMime: string = meta.mimeType || '';

            if (sourceMime === NATIVE_DOC_MIME) {
              throw new Error(
                `'${meta.name}' is already a native Google Doc and is editable as-is. Nothing to convert.`
              );
            }
            if (!isWordMime(sourceMime)) {
              throw new Error(
                `'${meta.name}' is not a Word file (${sourceMime}), so it cannot be converted to a Google Doc.`
              );
            }
            if (!sniffWordFormat(await fetchWordMagic(file_id, accessToken, GOOGLE_DRIVE_API))) {
              throw new Error(
                `'${meta.name}' is named like a Word file but its contents are not a Word ` +
                `document at all, so there is nothing to convert. Open it directly: ${meta.webViewLink}`
              );
            }

            // Drive owns the conversion, so its refusal arrives bare: give it the
            // file and a link like every other message this tool produces.
            let result: { id: string; name: string; webViewLink?: string };
            try {
              result = await makeDriveRequest(
                `/files/${encodeURIComponent(file_id)}/copy?supportsAllDrives=true`,
                accessToken,
                {
                  method: 'POST',
                  headers: { 'Content-Type': 'application/json' },
                  body: JSON.stringify({
                    mimeType: NATIVE_DOC_MIME,
                    ...(name ? { name } : {}),
                  }),
                }
              ) as { id: string; name: string; webViewLink?: string };
            } catch (err) {
              const reason = (err instanceof Error ? err.message : String(err)).replace(/\.\s*$/, '');
              throw new Error(
                `Drive could not convert '${meta.name}' into a Google Doc: ${reason}. ` +
                `It may be a different Office format stored under a Word name. ` +
                `Open it directly: ${meta.webViewLink}`
              );
            }

            const output = {
              id: result.id,
              name: result.name,
              webViewLink: result.webViewLink || `https://docs.google.com/document/d/${result.id}/edit`,
              sourceId: file_id,
              message:
                `Converted to a new native Google Doc, which is fully editable. ` +
                `The original '${meta.name}' is unchanged.`,
            };
            return {
              content: [{ type: 'text', text: JSON.stringify(output, null, 2) }],
              structuredContent: output,
            };
          } catch (err) {
            return toolErrorResponse(err);
          }
        }),
      },

      get_document_images: {
        description: 'Extract all inline images from a Google Doc. Returns each image as an inline image content block. Use get_document with include_structure=true first to see where images are positioned in the document. Google Docs only: the Docs API cannot open Word (.doc/.docx) uploads, so this tool refuses them.',
        readOnlyHint: true,
        schema: {
          document_id: z.string().describe('Google Doc ID (from search_documents or a Google Docs URL)'),
        },
        handler: requirePermissionSecure("https://www.googleapis.com/auth/drive.readonly", docsOnly(async ({ document_id }: any, context: any) => {
          try {
            const { accessToken } = context;

            const doc = await makeDocsRequest(`/${encodeURIComponent(document_id)}`, accessToken, { method: 'GET' });
            const inlineObjects = doc.inlineObjects || {};
            const objectIds = Object.keys(inlineObjects);

            if (objectIds.length === 0) {
              return {
                content: [{ type: 'text' as const, text: 'No images found in this document.' }],
              };
            }

            const content: any[] = [];

            for (const objectId of objectIds) {
              const obj = inlineObjects[objectId];
              const imageProps = obj?.inlineObjectProperties?.embeddedObject;
              const imageUrl = imageProps?.imageProperties?.sourceUri || imageProps?.imageProperties?.contentUri;
              const title = imageProps?.title || '';
              const description = imageProps?.description || '';

              if (!imageUrl) continue;

              try {
                const response = await fetch(imageUrl, {
                  headers: { 'Authorization': `Bearer ${accessToken}` },
                });

                if (!response.ok) {
                  content.push({ type: 'text' as const, text: `Failed to fetch image '${title || objectId}' (status ${response.status}).` });
                  continue;
                }

                const contentLength = parseInt(response.headers.get('content-length') || '0', 10);
                if (contentLength > MAX_IMAGE_SIZE) {
                  content.push({ type: 'text' as const, text: `Image '${title || objectId}' exceeds the 20MB limit, skipping.` });
                  continue;
                }

                const contentType = response.headers.get('content-type') || 'image/png';
                const arrayBuffer = await response.arrayBuffer();

                if (arrayBuffer.byteLength > MAX_IMAGE_SIZE) {
                  content.push({ type: 'text' as const, text: `Image '${title || objectId}' exceeds the 20MB limit, skipping.` });
                  continue;
                }
                const bytes = new Uint8Array(arrayBuffer);
                let binary = '';
                for (let i = 0; i < bytes.length; i++) {
                  binary += String.fromCharCode(bytes[i]);
                }
                const base64Data = btoa(binary);

                if (title || description) {
                  content.push({ type: 'text' as const, text: `Image: ${title || description}` });
                }
                content.push({ type: 'image' as const, data: base64Data, mimeType: contentType });
              } catch (fetchErr) {
                const msg = fetchErr instanceof Error ? fetchErr.message : String(fetchErr);
                content.push({ type: 'text' as const, text: `Failed to fetch image '${title || objectId}': ${msg}` });
              }
            }

            if (content.length === 0) {
              return {
                content: [{ type: 'text' as const, text: 'Found image references but could not fetch any images.' }],
              };
            }

            return { content };
          } catch (err) {
            return toolErrorResponse(err);
          }
        })),
      },

    };
  }
}
