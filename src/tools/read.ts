/** Read tools: search_documents, get_document and get_document_images */

import { z } from 'zod';
import { withGoogleAuth as requirePermissionSecure } from "../auth.js";
import {
  compactMarkdownExport,
  MAX_MARKDOWN_CHARS,
  escapeDriveQueryName,
  extractHeadings,
  findLetterByContent,
  isGoogleServedUrl,
  parseDocumentStructure,
  parseExportArtifacts,
  summarizeTabs,
  type Reaction,
} from "../lib/docs-structure.js";
import {
  readWordUpload,
  toolResultWithNotice,
  NATIVE_DOC_MIME,
  WORD_READ_ONLY_NOTICE,
} from "../lib/word.js";
import {
  isWordMime,
  DOCX_MIME,
  LEGACY_DOC_MIME,
  MAX_TEXT_CHARS as MAX_WORD_TEXT_CHARS,
} from "../lib/wordText.js";
import { collectInlineObjects, renderStructureText, resolveTab } from "../lib/tabs.js";
import { listHeaderFooters, segmentText } from "../lib/segments.js";
import { GOOGLE_DRIVE_API, buildGoogleApiError, makeDriveRequest, makeDocsRequest } from '../lib/google.js';
import { attachLabelsMeta } from '../lib/driveLabels.js';
import {
  DRIVE_READONLY_SCOPE,
  TAB_LIST_FIELDS,
  fetchResolvedTab,
  toolErrorResponse,
  toolResult,
} from './shared.js';

const MAX_IMAGE_SIZE = 20 * 1024 * 1024; // 20MB

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
  inlineObjectId: z.string().optional().describe('inlineImage only: the id replace_image takes'),
  footnoteId: z.string().optional().describe('footnoteReference only: the segment_id of the footnote text'),
  headingLevel: z.number().int().min(1).max(6).optional(),
  bullet: z.object({ nestingLevel: z.number().int() }).optional().describe('Paragraph only: present when the paragraph is a list item, with its nesting level (0 = top level).'),
  rows: z.number().optional().describe('Table only: number of rows'),
  columns: z.number().optional().describe('Table only: number of columns'),
  cells: z.array(z.array(z.string())).optional().describe('Table only: cell text by row then column. Long cells are truncated and a cell holding a nested table reads as empty. Omitted for later tables once a document exhausts the structure budget, which counts both cell characters and cell count; `message` says when that happened. When has_merged_cells is true, a position covered by a merge reads as an empty string, so positions still line up with columns, but an empty cell may be a covered position rather than a genuinely empty one.'),
  has_merged_cells: z.boolean().optional().describe('Table only: the table has merged cells. Positions in `cells` still match columns, since a covered position reads as an empty string, but styling a range may affect cells outside it.'),
  styles: z.array(reportedCellStyleSchema).optional().describe('Table only, with include_table_styles=true: each distinct cell style listed once. Values are in the form update_table_style accepts, so they can be passed straight back.'),
  cell_styles: z.array(z.array(z.number())).optional().describe('Table only, with include_table_styles=true: an index into `styles` for each cell, positioned like `cells`. Cells sharing an index share a look, which is how you find the ranges to pass to update_table_style.'),
  suggestions: z.array(z.object({ id: z.string(), kind: z.enum(['insertion', 'deletion', 'formatting']) })).optional().describe('Pending suggestions touching this element, by id. Pass the ids to review_suggestion to accept or reject them.'),
}).passthrough();

const headerFooterSchema = z.object({
  id: z.string().describe('Pass as segment_id to the editing tools'),
  type: z.enum(['DEFAULT', 'FIRST_PAGE', 'EVEN_PAGE', 'SECTION']),
  text: z.string(),
});

const SUGGESTIONS_VIEW_MODES = ['DEFAULT_FOR_CURRENT_ACCESS', 'SUGGESTIONS_INLINE', 'PREVIEW_SUGGESTIONS_ACCEPTED', 'PREVIEW_WITHOUT_SUGGESTIONS'] as const;

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

export const readTools = {
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
    handler: requirePermissionSecure(DRIVE_READONLY_SCOPE, async ({ name, page_token }: any, context: any) => {
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

        return toolResult({
          documents,
          nextPageToken: result.nextPageToken || null,
        });
      } catch (err) {
        return toolErrorResponse(err);
      }
    }),
  },

  get_document: {
    description: 'Read the contents of a Google Doc as plain text, or as markdown with format="markdown" (headings, bold, links, lists and tables preserved; best for understanding a formatted document). Also reads Word (.doc and .docx) files uploaded to Drive, by parsing them directly — for those, `mimeType` is returned and `include_structure`/`include_comments` do not apply. Optionally include document structure with startIndex/endIndex for each element (needed for index-based editing tools like delete_content, insert_text, update_text_style, update_paragraph_style). Table elements in `structure` also carry `rows`, `columns` and `cells` (cell text by row then column), which is how you find the coordinates to pass to update_table_style; a table reporting has_merged_cells=true still lines its positions up with columns, since a merge-covered position reads as an empty string. Add include_table_styles=true to also see how each table is formatted: `styles` lists every distinct cell style once and `cell_styles` gives each cell an index into it, so cells sharing an index share a look. The values are in the form update_table_style accepts, so read a style and pass it straight back to copy a table\'s formatting. One caveat when replaying: a border reported without a `color` is written back as black, since Google needs a complete border, so check that field before copying a border across documents. When include_structure=true, the response also includes a `headings` array (level, text, startIndex, endIndex) to support "insert after the heading named X" flows without parsing the full structure, plus `headers` and `footers` (id, type, text) for the tab being read; their ids are the segment_id the editing tools and set_header_footer work with. Set include_comments=true to also return comment thread metadata (author email, timestamp, replies, emoji reactions). `content` is returned verbatim with no inline markers; each thread carries `anchor_offset: { start, end }` — a half-open span into `content` indicating which text the comment was attached to. Threads with no findable position (document-level comments, or anchored text deleted by later edits) omit `anchor_offset`. If the document uses Tabs, a `tabs` summary (id, title, nesting) is always returned. Pass `tab_id` to read ONE tab: `content`, `structure` and `headings` then all describe that tab, rendered from the same source so the indices are exactly what the editing tools need; pass the same tab_id to those tools. Without `tab_id`, `content` concatenates all tabs (via Drive export) while `structure`/`headings` indices refer to the FIRST tab only, which is also the tab the editing tools target when not given a tab_id. With `tab_id`, tables in `content` render as tab-separated rows and comment reactions are unavailable; comment threads are document-level, so ones anchored in other tabs are still returned but without `anchor_offset`. The response also includes the current `revisionId`; pass it as `required_revision_id` to a subsequent mutating tool to detect concurrent edits (the write will fail rather than silently overwrite). Use search_documents to find a document ID first.',
    readOnlyHint: true,
    outputSchema: {
      id: z.string(),
      title: z.string(),
      content: z.string(),
      webViewLink: z.string().optional(),
      revisionId: z.string().optional(),
      mimeType: z.string().optional().describe('Present for non-native files such as .docx uploads'),
      format: z.literal('markdown').optional().describe('Present when content is markdown rather than plain text'),
      truncated: z.boolean().optional().describe('True when content was cut at the character cap'),
      message: z.string().optional().describe('Why content is partial, or which options do not apply'),
      structure: z.array(structureElementSchema).optional(),
      headings: z.array(headingSchema).optional(),
      headers: z.array(headerFooterSchema).optional().describe('With include_structure: the headers of the tab being read, with their ids and text'),
      footers: z.array(headerFooterSchema).optional().describe('With include_structure: the footers of the tab being read'),
      footnotes: z.array(z.object({ id: z.string().describe('Pass as segment_id to the editing tools'), text: z.string() })).optional().describe('With include_structure: the footnotes of the tab being read; the footnoteReference elements in `structure` say where each mark sits'),
      tabs: z.array(tabSummarySchema).optional(),
      tab: z.object({
        tabId: z.string(),
        title: z.string().optional(),
      }).optional().describe('Present when tab_id was passed: the tab this response describes'),
      threads: z.array(threadSchema).optional(),
    },
    schema: {
      document_id: z.string().describe('Google Doc ID (from search_documents or a Google Docs URL)'),
      tab_id: z.string().min(1).optional().describe('Read a single tab of a multi-tab document: tabProperties.tabId, the value after ?tab= in the document URL (e.g. "t.abc123"), also listed in this tool\'s `tabs` output. With tab_id, `content`, `structure` and `headings` all describe that tab and their indices are valid for the editing tools when passed the same tab_id. Without it, `content` concatenates all tabs while structure indices cover only the first tab.'),
      format: z.enum(['text', 'markdown']).optional().describe('"text" (default) returns plain text. "markdown" renders headings, bold, italic, links, lists and tables as markdown; embedded images are replaced with a placeholder (get_document_images returns them) and the result is capped at 200,000 characters. It reads the whole document, so it cannot be combined with tab_id, and comment anchors are plain-text offsets, so it cannot be combined with include_comments.'),
      include_structure: z.boolean().optional().describe('Include document structure with startIndex/endIndex for each element. Required before using index-based tools.'),
      include_table_styles: z.boolean().optional().describe('Also report each table\'s cell styles, as a `styles` legend plus a `cell_styles` index grid. Requires include_structure=true. Use it to see how a table is currently formatted before matching or replicating it.'),
      include_comments: z.boolean().optional().describe('Include comment thread metadata (author email/name, timestamp, replies, resolved status, emoji reactions). `content` is returned verbatim — no inline markers are inserted. Each thread carries `anchor_offset: { start, end }` — a half-open span into `content` indicating which text the comment was attached to. Threads with no findable position (document-level comments, or anchor deleted by later edits) omit `anchor_offset`.'),
      suggestions_view_mode: z.enum(SUGGESTIONS_VIEW_MODES).optional().describe('How pending suggestions (Google Docs "Suggesting" mode) appear in `structure`, `headings` and per-tab `content`: SUGGESTIONS_INLINE (the default for editors) shows suggested insertions and still-pending deletions as text, with their ids in each element\'s `suggestions`; PREVIEW_SUGGESTIONS_ACCEPTED shows the document as if every suggestion were accepted; PREVIEW_WITHOUT_SUGGESTIONS as if every one were rejected. The whole-document `content` export never includes pending suggestions, so on a document with suggestions read one tab with tab_id, or pass PREVIEW_WITHOUT_SUGGESTIONS, for indices that match it.'),
    },
    handler: requirePermissionSecure(DRIVE_READONLY_SCOPE, attachLabelsMeta((a: any) => a.document_id, async ({ document_id, tab_id, format, include_structure, include_comments, include_table_styles, suggestions_view_mode }: any, context: any) => {
      try {
        const { accessToken } = context;
        const asMarkdown = format === 'markdown';
        const markdownConflicts = [
          ...(tab_id ? ['tab_id (markdown is exported for the whole document)'] : []),
          ...(include_comments ? ['include_comments (comment anchors are plain-text offsets)'] : []),
        ];
        if (asMarkdown && markdownConflicts.length > 0) {
          throw new Error(`format "markdown" cannot be combined with ${markdownConflicts.join(' or ')}. Drop format or the other option.`);
        }
        if (suggestions_view_mode && !include_structure && !tab_id) {
          throw new Error('suggestions_view_mode only affects the Docs read behind `structure` and per-tab `content`; pass include_structure=true or tab_id with it, since the whole-document export never includes pending suggestions.');
        }

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
          if (include_structure || include_comments || tab_id || asMarkdown) {
            wordNotes.push(
              `'${metadata.name}' is a Word upload, not a Google Doc, ` +
              `so document structure, tabs, comments and markdown are unavailable for it.`
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

        const needsDocsContent = Boolean(include_structure || tab_id);
        const viewMode = suggestions_view_mode ? `&suggestionsViewMode=${suggestions_view_mode}` : '';
        let doc: any;
        let docsFetchNote: string | undefined;
        try {
          doc = await makeDocsRequest(
            needsDocsContent
              ? `/${encodeURIComponent(document_id)}?includeTabsContent=true${viewMode}`
              : `/${encodeURIComponent(document_id)}?fields=${TAB_LIST_FIELDS}`,
            accessToken, { method: 'GET' });
        } catch (docsErr) {
          // a plain text read should not fail because the tabs listing did
          if (needsDocsContent) throw docsErr;
          docsFetchNote = 'The tabs listing could not be fetched, so `tabs` and `revisionId` are missing. Retry if you need them.';
        }

        const resolved = needsDocsContent ? resolveTab(doc, tab_id) : undefined;
        const parsed = resolved ? parseDocumentStructure(resolved.body, include_table_styles === true) : undefined;

        let content: string;
        const markdownNotes: string[] = [];
        let markdownTruncated = false;
        if (parsed && tab_id) {
          content = renderStructureText(parsed.elements);
        } else {
          // Accept-Language pins the export footer to the English template
          // the comment-reaction parser expects
          const exportMime = encodeURIComponent(asMarkdown ? 'text/markdown' : 'text/plain');
          const exportUrl = `${GOOGLE_DRIVE_API}/files/${encodeURIComponent(document_id)}/export?mimeType=${exportMime}`;
          const response = await fetch(exportUrl, {
            headers: {
              'Authorization': `Bearer ${accessToken}`,
              'Accept-Language': 'en-US',
            },
          });

          if (!response.ok) {
            throw await buildGoogleApiError(response, 'drive');
          }

          content = await response.text();
          if (asMarkdown) {
            const compacted = compactMarkdownExport(content);
            content = compacted.content;
            markdownTruncated = compacted.truncated;
            if (compacted.imagesOmitted > 0) {
              markdownNotes.push(`${compacted.imagesOmitted} embedded image(s) were replaced with "embedded-image-omitted"; get_document_images returns them.`);
            }
            if (compacted.truncated) {
              markdownNotes.push(`The markdown was truncated at ${MAX_MARKDOWN_CHARS} characters. Use format "text" or read one tab with tab_id for the rest.`);
            }
          }
        }

        const output: any = {
          id: document_id,
          title: metadata.name,
          webViewLink: metadata.webViewLink,
        };
        if (tab_id && resolved) {
          output.tab = {
            tabId: resolved.tabId,
            ...(resolved.title !== undefined ? { title: resolved.title } : {}),
          };
        }

        if (include_comments) {
          const { cleanedContent, threads } = await processCommentsForDocument(document_id, accessToken, content);
          content = cleanedContent;
          output.threads = threads;
        }

        output.content = content;
        if (asMarkdown) output.format = 'markdown';
        if (markdownTruncated) output.truncated = true;

        const tabs = summarizeTabs(doc?.tabs);
        if (tabs.length > 0) output.tabs = tabs;
        if (typeof doc?.revisionId === 'string') output.revisionId = doc.revisionId;

        const notes: string[] = [];
        if (docsFetchNote) notes.push(docsFetchNote);
        notes.push(...markdownNotes);
        if (parsed && resolved && include_structure) {
          output.structure = parsed.elements;
          output.headings = extractHeadings(resolved.body);
          const headers = listHeaderFooters('header', resolved.documentStyle, resolved.headers);
          const footers = listHeaderFooters('footer', resolved.documentStyle, resolved.footers);
          const footnotes = Object.entries(resolved.footnotes).map(([id, note]) => ({ id, text: segmentText(note) }));
          if (headers.length) output.headers = headers;
          if (footers.length) output.footers = footers;
          if (footnotes.length) output.footnotes = footnotes;
          if (parsed.tablesWithoutCells > 0) {
            notes.push(`${parsed.tablesWithoutCells} table(s) report dimensions without cell contents because the document exhausted the structure budget for table cells. Their text, if any, is in \`content\`.`);
          }
          if (parsed.tablesWithoutStyles > 0) {
            notes.push(`${parsed.tablesWithoutStyles} table(s) report no styles because the document exhausted the structure budget. Read those tables on their own to see their formatting.`);
          }
        } else if (include_table_styles) {
          notes.push('include_table_styles needs include_structure=true, since styles are reported on the table elements in `structure`. No structure was returned.');
        }
        if (notes.length) output.message = notes.join(' ');

        return toolResult(output);
      } catch (err) {
        return toolErrorResponse(err);
      }
    })),
  },

  get_document_images: {
    description: 'Extract all inline images from a Google Doc, across every tab, or from one tab when tab_id is passed. Returns each image as an inline image content block. Use get_document with include_structure=true first to see where images are positioned in the document.',
    readOnlyHint: true,
    schema: {
      document_id: z.string().describe('Google Doc ID (from search_documents or a Google Docs URL)'),
      tab_id: z.string().min(1).optional().describe('Only extract images from this tab: tabProperties.tabId, the value after ?tab= in the document URL, also listed in get_document\'s `tabs`. Omit for all tabs.'),
    },
    handler: requirePermissionSecure(DRIVE_READONLY_SCOPE, attachLabelsMeta((a: any) => a.document_id, async ({ document_id, tab_id }: any, context: any) => {
      try {
        const { accessToken } = context;

        const { doc, resolved } = await fetchResolvedTab(document_id, accessToken, tab_id);
        const inlineObjects = tab_id ? resolved.inlineObjects : collectInlineObjects(doc);
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
          const imageUrl = imageProps?.imageProperties?.contentUri || imageProps?.imageProperties?.sourceUri;
          const title = imageProps?.title || '';
          const description = imageProps?.description || '';

          if (!imageUrl) continue;

          try {
            const response = await fetch(imageUrl, {
              headers: isGoogleServedUrl(imageUrl) ? { 'Authorization': `Bearer ${accessToken}` } : {},
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
