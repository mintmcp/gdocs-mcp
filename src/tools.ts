/**
 * Google Docs MCP Tools
 */

import { z } from 'zod';
import { withGoogleAuth as requirePermissionSecure } from "./auth.js";

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

  constructor(message: string, status: number, api: 'drive' | 'docs', opts: { code?: string; retryAfter?: number } = {}) {
    super(message);
    this.name = 'GoogleApiError';
    this.status = status;
    this.api = api;
    this.code = opts.code;
    this.retryAfter = opts.retryAfter;
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

  return new GoogleApiError(message, response.status, api, { code: googleCode, retryAfter });
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

/**
 * Parse document body.content into structural elements with indices
 */
function parseDocumentStructure(content: any[]): Array<{ type: string; startIndex: number; endIndex: number; text?: string; inlineObjectId?: string }> {
  const elements: Array<{ type: string; startIndex: number; endIndex: number; text?: string; inlineObjectId?: string }> = [];

  for (const element of content) {
    if (element.paragraph) {
      const text = element.paragraph.elements
        ?.map((el: any) => el.textRun?.content || '')
        .join('') || '';
      elements.push({
        type: 'paragraph',
        startIndex: element.startIndex ?? 0,
        endIndex: element.endIndex,
        text,
      });

      // Extract inline images from paragraph elements
      for (const el of element.paragraph.elements || []) {
        if (el.inlineObjectElement) {
          elements.push({
            type: 'inlineImage',
            startIndex: el.startIndex,
            endIndex: el.endIndex,
            inlineObjectId: el.inlineObjectElement.inlineObjectId,
          });
        }
      }
    } else if (element.table) {
      elements.push({
        type: 'table',
        startIndex: element.startIndex,
        endIndex: element.endIndex,
      });
    } else if (element.sectionBreak) {
      elements.push({
        type: 'sectionBreak',
        startIndex: element.startIndex ?? 0,
        endIndex: element.endIndex,
      });
    } else if (element.tableOfContents) {
      elements.push({
        type: 'tableOfContents',
        startIndex: element.startIndex,
        endIndex: element.endIndex,
      });
    }
  }

  return elements;
}

/**
 * Build a Docs API `updateTextStyle` request that clears inherited character
 * formatting (bold, italic, underline, strikethrough, link) on the given range.
 * Centralised so all "insert + optionally reset formatting" tools agree on the
 * exact set of fields cleared.
 */
function clearInheritedFormattingRequest(startIndex: number, endIndex: number): any {
  return {
    updateTextStyle: {
      range: { startIndex, endIndex },
      textStyle: {},
      fields: 'bold,italic,underline,strikethrough,link',
    },
  };
}

/**
 * Build table insert requests in reverse order to avoid index shifts.
 *
 * Table structure in Google Docs:
 *   |||
 *   |X||X||X||X| |
 *   |X||X||X||X| |
 *   ...
 *
 * Row index formula: row_index = 3 + row * (2 * num_cols) + row
 * Cell index formula: insert_index = row_index + col * 2 + 1 + table_insert_index
 *
 * Insert in reverse order so earlier inserts don't shift later indices.
 */
function buildTableInsertRequests(
  tableData: string[][],
  tableInsertIndex: number
): any[] {
  const requests: any[] = [];
  const numRows = tableData.length;
  if (numRows === 0) return requests;
  const numCols = tableData[0]?.length ?? 0;
  if (numCols === 0) return requests;

  for (let row = numRows - 1; row >= 0; row--) {
    const rowIndex = 3 + row * 2 * numCols + row;
    for (let col = numCols - 1; col >= 0; col--) {
      const insertIndex = rowIndex + col * 2 + 1 + tableInsertIndex;
      const text = String(tableData[row][col]);
      if (text) {
        requests.push({
          insertText: {
            text,
            location: { index: insertIndex },
          },
        });
      }
    }
  }

  return requests;
}

// Output schema fragments
const structureElementSchema = z.object({
  type: z.string().optional(),
  startIndex: z.number().optional(),
  endIndex: z.number().optional(),
  text: z.string().optional(),
  inlineObjectId: z.string().optional(),
}).passthrough();

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

interface Reaction {
  author_name: string;
  emoji: string;
  timestamp: string;
}

interface ThreadOutput {
  // Half-open [start, end) span into `content` where the comment is
  // anchored. Absent when the comment has no quotedFileContent
  // (document-level comments) or whose quote can no longer be located in
  // the body (e.g., the doc was edited and the anchored span was deleted —
  // an "orphaned" thread). An absent `anchor_offset` is the explicit signal.
  anchor_offset?: { start: number; end: number };
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

function normalizeForMatch(text: string): string {
  return text.replace(/\s+/g, ' ').trim();
}

// Reactions only ever surface in the plain-text export footer (Drive Comments
// API has no reaction field). Format is the English template
// `<displayName> reacted with <emoji> at YYYY-MM-DD HH:MM (AM|PM)`. The HH is
// 24-hour despite the AM/PM suffix, so we drop the suffix when parsing.
const REACTION_RE = /^(.+?) reacted with (\S+) at (\d{4}-\d{2}-\d{2} \d{2}:\d{2})(?:\s*[AP]M)?\s*$/;

interface ParsedExport {
  // Body with the trailing letter-block removed but inline `[letter]` markers
  // preserved. Inline markers are rewritten downstream at known thread anchor
  // positions so legitimate `[a]` text in the doc body is left untouched.
  body: string;
  // letter (e.g. "a") → first non-empty line after the marker (= comment text)
  letterContent: Map<string, string>;
  // letter → reactions parsed from continuation lines under that marker
  letterReactions: Map<string, Reaction[]>;
  // Set of letters used as comment markers in the trailing block — used by
  // the rewrite step to decide which `[X]` runs are real anchors.
  commentLetters: Set<string>;
}

// Drive's plain-text export injects letter markers (`[a][b][c]…`) inline at
// every comment/reply anchor position and lists each letter's comment text
// (plus any reaction continuation lines) in a trailing block. The trailing
// block is unambiguous and safe to strip; the inline markers are not — a
// legitimate `[a]` typed into the doc body looks identical to a comment
// marker, so we leave them in place here and let the caller rewrite only
// at confirmed thread anchor positions.
function parseExportArtifacts(plainText: string): ParsedExport {
  const lines = plainText.split('\n');

  let blockStart = -1;
  for (let candidate = 0; candidate < lines.length; candidate++) {
    if (!/^\[[a-z]+\]/.test(lines[candidate])) continue;
    let inMarker = false;
    let valid = true;
    for (let i = candidate; i < lines.length; i++) {
      const line = lines[i];
      if (line.trim() === '') continue;
      if (/^\[[a-z]+\]/.test(line)) { inMarker = true; continue; }
      if (!inMarker) { valid = false; break; }
    }
    if (valid) { blockStart = candidate; break; }
  }

  const letterContent = new Map<string, string>();
  const letterReactions = new Map<string, Reaction[]>();
  const commentLetters = new Set<string>();

  if (blockStart >= 0) {
    let currentLetter: string | null = null;
    for (let i = blockStart; i < lines.length; i++) {
      const line = lines[i].replace(/\r$/, '');
      const m = line.match(/^\[([a-z]+)\](.*)$/);
      if (m) {
        currentLetter = m[1];
        commentLetters.add(currentLetter);
        letterContent.set(currentLetter, m[2].trim());
        letterReactions.set(currentLetter, []);
      } else if (currentLetter !== null && line.trim() !== '') {
        const rx = line.match(REACTION_RE);
        if (rx) {
          letterReactions.get(currentLetter)!.push({
            author_name: rx[1].trim(),
            emoji: rx[2],
            // Emit ISO-like local time (no TZ — Google doesn't surface one).
            // The AM/PM suffix is cosmetic; HH is already 24-hour.
            timestamp: rx[3].replace(' ', 'T'),
          });
        }
        // Non-reaction continuation lines (e.g., "N total reactions",
        // multi-line comment text) are intentionally dropped.
      }
    }
  }

  const body = blockStart >= 0
    ? lines.slice(0, blockStart).join('\n')
    : plainText;

  return {
    body: body.replace(/[\s\r\n]+$/, ''),
    letterContent,
    letterReactions,
    commentLetters,
  };
}

function findLetterByContent(
  letterContent: Map<string, string>,
  apiContent: string
): string | undefined {
  const target = normalizeForMatch(apiContent);
  if (!target) return undefined;
  for (const [letter, content] of letterContent) {
    if (normalizeForMatch(content) === target) return letter;
  }
  return undefined;
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
        description: 'Search for Google Docs by name. Returns matching documents with their IDs.',
        readOnlyHint: true,
        outputSchema: {
          documents: z.array(z.object({
            id: z.string(),
            name: z.string(),
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

            let q = `mimeType = 'application/vnd.google-apps.document'`;
            if (name) {
              q += ` and name contains '${name.replace(/'/g, "\\'")}'`;
            }
            q += ` and trashed = false`;

            const params = new URLSearchParams({
              pageSize: '20',
              fields: 'nextPageToken,files(id,name,createdTime,modifiedTime,webViewLink,owners)',
              supportsAllDrives: 'true',
              includeItemsFromAllDrives: 'true',
              q,
              ...(page_token && { pageToken: page_token }),
            });

            const result = await makeDriveRequest(`/files?${params}`, accessToken);

            const documents = (result.files || []).map((file: any) => ({
              id: file.id,
              name: file.name,
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
        description: 'Read the contents of a Google Doc as plain text. Optionally include document structure with startIndex/endIndex for each element (needed for index-based editing tools like delete_content, insert_text, update_text_style, update_paragraph_style). Set include_comments=true to also return comment thread metadata (author email, timestamp, replies, emoji reactions). `content` is returned verbatim with no inline markers; each thread carries `anchor_offset: { start, end }` — a half-open span into `content` indicating which text the comment was attached to. Threads with no findable position (document-level comments, or anchored text deleted by later edits) omit `anchor_offset`. Use search_documents to find a document ID first.',
        readOnlyHint: true,
        outputSchema: {
          id: z.string(),
          title: z.string(),
          content: z.string(),
          webViewLink: z.string().optional(),
          structure: z.array(structureElementSchema).optional(),
          threads: z.array(threadSchema).optional(),
        },
        schema: {
          document_id: z.string().describe('Google Doc ID (from search_documents or a Google Docs URL)'),
          include_structure: z.boolean().optional().describe('Include document structure with startIndex/endIndex for each element. Required before using index-based tools.'),
          include_comments: z.boolean().optional().describe('Include comment thread metadata (author email/name, timestamp, replies, resolved status, emoji reactions). `content` is returned verbatim — no inline markers are inserted. Each thread carries `anchor_offset: { start, end }` — a half-open span into `content` indicating which text the comment was attached to. Threads with no findable position (document-level comments, or anchor deleted by later edits) omit `anchor_offset`.'),
        },
        handler: requirePermissionSecure("https://www.googleapis.com/auth/drive.readonly", async ({ document_id, include_structure, include_comments }: any, context: any) => {
          try {
            const { accessToken } = context;

            // Get file metadata for title and link
            const metadata = await makeDriveRequest(
              `/files/${encodeURIComponent(document_id)}?fields=name,webViewLink&supportsAllDrives=true`,
              accessToken
            );

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

            // If structure requested, also fetch from Docs API
            if (include_structure) {
              const doc = await makeDocsRequest(`/${encodeURIComponent(document_id)}`, accessToken, { method: 'GET' });
              output.structure = parseDocumentStructure(doc.body?.content || []);
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
        description: 'Append plain text to the end of a Google Doc. Use get_document first to verify the document exists and see current content.',
        outputSchema: {
          id: z.string(),
          message: z.string(),
        },
        schema: {
          document_id: z.string().describe('Google Doc ID'),
          text: z.string().describe('Plain text to append to the end of the document'),
          clear_inherited_formatting: z.boolean().optional().describe('Clear inherited formatting from preceding text on the newly appended content. Defaults to true.'),
        },
        handler: requirePermissionSecure("https://www.googleapis.com/auth/documents", async ({ document_id, text, clear_inherited_formatting }: any, context: any) => {
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

            await makeDocsRequest(`/${encodeURIComponent(document_id)}:batchUpdate`, accessToken, {
              method: 'POST',
              body: JSON.stringify({ requests }),
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
        }),
      },

      replace_text: {
        description: 'Replace all occurrences of a text string in a Google Doc. Use get_document first to see current content and verify the text to replace exists. Use empty new_text to delete occurrences.',
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
        },
        handler: requirePermissionSecure("https://www.googleapis.com/auth/documents", async ({ document_id, old_text, new_text, match_case }: any, context: any) => {
          try {
            const { accessToken } = context;

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
        }),
      },

      delete_content: {
        description: 'Delete content in a Google Doc by index range. Use get_document with include_structure=true first to find the correct startIndex and endIndex.',
        destructiveHint: true,
        outputSchema: {
          id: z.string(),
          message: z.string(),
        },
        schema: {
          document_id: z.string().describe('Google Doc ID'),
          startIndex: z.coerce.number().int().describe('Start index of content to delete (use get_document with include_structure to find indices)'),
          endIndex: z.coerce.number().int().describe('End index of content to delete (exclusive)'),
        },
        handler: requirePermissionSecure("https://www.googleapis.com/auth/documents", async ({ document_id, startIndex, endIndex }: any, context: any) => {
          try {
            const { accessToken } = context;

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
        }),
      },

      insert_text: {
        description: 'Insert text at a specific position in a Google Doc. Use get_document with include_structure=true to find the correct index. NOTE: Inserted text inherits formatting from surrounding text at the insertion point. Use clear_inherited_formatting=true to reset to plain formatting.',
        outputSchema: {
          id: z.string(),
          message: z.string(),
        },
        schema: {
          document_id: z.string().describe('Google Doc ID'),
          text: z.string().describe('Text to insert'),
          index: z.coerce.number().int().describe('Position to insert at (use get_document with include_structure to find indices)'),
          clear_inherited_formatting: z.boolean().optional().describe('Clear inherited formatting from surrounding text on the newly inserted content. Defaults to false.'),
        },
        handler: requirePermissionSecure("https://www.googleapis.com/auth/documents", async ({ document_id, text, index, clear_inherited_formatting }: any, context: any) => {
          try {
            const { accessToken } = context;

            const requests: any[] = [{
              insertText: {
                location: { index },
                text,
              },
            }];

            if (clear_inherited_formatting === true) {
              requests.push(clearInheritedFormattingRequest(index, index + text.length));
            }

            await makeDocsRequest(`/${encodeURIComponent(document_id)}:batchUpdate`, accessToken, {
              method: 'POST',
              body: JSON.stringify({ requests }),
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
        }),
      },

      update_text_style: {
        description: 'Apply formatting (bold, italic, underline, strikethrough, link) to a text range in a Google Doc. Use get_document with include_structure=true to find the correct indices. NOTE: If you also need to set a heading level via update_paragraph_style, do that first — heading changes reset character formatting.',
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
        },
        handler: requirePermissionSecure("https://www.googleapis.com/auth/documents", async ({ document_id, startIndex, endIndex, bold, italic, underline, strikethrough, link_url }: any, context: any) => {
          try {
            const { accessToken } = context;

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
        }),
      },

      update_paragraph_style: {
        description: 'Apply paragraph formatting (heading level, alignment) to a range in a Google Doc. Use get_document with include_structure=true to find the correct indices. WARNING: Setting heading_level applies a named style that resets character-level formatting (bold, italic, etc.). If you also need to apply text styles, call update_paragraph_style first, then apply text styles after.',
        outputSchema: {
          id: z.string(),
          message: z.string(),
        },
        schema: {
          document_id: z.string().describe('Google Doc ID'),
          startIndex: z.coerce.number().int().describe('Start index of paragraph range'),
          endIndex: z.coerce.number().int().describe('End index of paragraph range (exclusive)'),
          heading_level: z.coerce.number().int().optional().describe('Heading level: 0=normal text, 1-6=heading levels'),
          alignment: z.enum(['START', 'CENTER', 'END', 'JUSTIFIED']).optional().describe('Paragraph alignment'),
        },
        handler: requirePermissionSecure("https://www.googleapis.com/auth/documents", async ({ document_id, startIndex, endIndex, heading_level, alignment }: any, context: any) => {
          try {
            const { accessToken } = context;

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
        }),
      },

      append_table: {
        description: 'Insert a table with data at the end of a Google Doc. Supports ragged rows (will be padded with empty cells).',
        outputSchema: {
          id: z.string(),
          rows: z.number(),
          columns: z.number(),
          message: z.string(),
          warning: z.string().optional(),
        },
        schema: {
          document_id: z.string().describe('Google Doc ID'),
          rows: z.array(z.array(z.string())).describe('Table data as array of rows, each row is array of cell strings'),
          clear_inherited_formatting: z.boolean().optional().describe('Clear inherited formatting on the newly appended table content. Defaults to true.'),
        },
        handler: requirePermissionSecure("https://www.googleapis.com/auth/documents", async ({ document_id, rows, clear_inherited_formatting }: any, context: any) => {
          try {
            const { accessToken } = context;

            if (!rows || rows.length === 0) {
              throw new Error('Table must have at least one row');
            }

            // Normalize ragged rows: find max columns and pad shorter rows
            const tableData: string[][] = rows.map((row: string[]) => [...row]);
            let maxCols = 0;
            for (const row of tableData) {
              maxCols = Math.max(maxCols, row.length);
            }
            if (maxCols === 0) {
              throw new Error('Table must have at least one column');
            }
            for (const row of tableData) {
              while (row.length < maxCols) {
                row.push('');
              }
            }

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

            await makeDocsRequest(`/${encodeURIComponent(document_id)}:batchUpdate`, accessToken, {
              method: 'POST',
              body: JSON.stringify({ requests }),
            });

            // Clear inherited formatting on the newly inserted table — best-effort.
            // The table is already inserted at this point; if this second
            // batchUpdate fails, we still want to surface success and warn
            // about the formatting step rather than fail the whole operation.
            let formattingWarning: string | undefined;
            if (clear_inherited_formatting !== false) {
              try {
                const updatedDoc = await makeDocsRequest(`/${encodeURIComponent(document_id)}`, accessToken, { method: 'GET' });
                const tables = updatedDoc.body.content.filter((el: any) => el.table);
                const lastTable = tables[tables.length - 1];
                if (lastTable) {
                  await makeDocsRequest(`/${encodeURIComponent(document_id)}:batchUpdate`, accessToken, {
                    method: 'POST',
                    body: JSON.stringify({
                      requests: [clearInheritedFormattingRequest(lastTable.startIndex + 1, lastTable.endIndex)],
                    }),
                  });
                }
              } catch (formattingErr) {
                formattingWarning = formattingErr instanceof Error ? formattingErr.message : String(formattingErr);
              }
            }

            const output: {
              id: string;
              rows: number;
              columns: number;
              message: string;
              warning?: string;
            } = {
              id: document_id,
              rows: tableData.length,
              columns: maxCols,
              message: `Table inserted with ${tableData.length} rows and ${maxCols} columns`,
            };
            if (formattingWarning) {
              output.warning = `Table inserted, but clearing inherited formatting failed: ${formattingWarning}`;
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
      get_document_images: {
        description: 'Extract all inline images from a Google Doc. Returns each image as an inline image content block. Use get_document with include_structure=true first to see where images are positioned in the document.',
        readOnlyHint: true,
        schema: {
          document_id: z.string().describe('Google Doc ID (from search_documents or a Google Docs URL)'),
        },
        handler: requirePermissionSecure("https://www.googleapis.com/auth/drive.readonly", async ({ document_id }: any, context: any) => {
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
        }),
      },

    };
  }
}
