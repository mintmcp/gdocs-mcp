import { z } from 'zod';
import {
  escapeDriveQueryName,
  headingsOf,
  isGoogleServedUrl,
  parseDocumentStructure,
  renderBodyText,
} from "../lib/docs-structure.js";
import { anchorThreads, compactMarkdownExport, type DriveComment, MAX_MARKDOWN_CHARS } from "../lib/exportText.js";
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
import { collectInlineObjects, type ResolvedTab, type TabSummary } from "../lib/tabs.js";
import { listHeaderFooters, summarizeSegment } from "../lib/segments.js";
import { GOOGLE_DRIVE_API, buildGoogleApiError, makeDriveRequest } from '../lib/google.js';
import { SCOPES } from "../scopes.js";
import {
  defineTool,
  fetchResolvedTab,
  fetchTabList,
  SUGGESTIONS_VIEW_MODES,
  tabSummarySchema,
  toolResult,
} from './shared.js';

const MAX_IMAGE_SIZE = 20 * 1024 * 1024; // 20MB
// Each image can be 20MB in memory, so only a few download at once
const IMAGE_FETCH_CONCURRENCY = 4;

type ContentBlock = { type: 'text'; text: string } | { type: 'image'; data: string; mimeType: string };

async function mapWithLimit<T, R>(items: T[], limit: number, fn: (item: T) => Promise<R>): Promise<R[]> {
  const results: R[] = new Array(items.length);
  let next = 0;
  const worker = async () => {
    while (next < items.length) {
      const i = next++;
      results[i] = await fn(items[i]);
    }
  };
  await Promise.all(Array.from({ length: Math.min(limit, items.length) }, worker));
  return results;
}

async function imageBlocks(objectId: string, obj: any, accessToken: string): Promise<ContentBlock[]> {
  const imageProps = obj?.inlineObjectProperties?.embeddedObject;
  const imageUrl = imageProps?.imageProperties?.contentUri || imageProps?.imageProperties?.sourceUri;
  const title = imageProps?.title || '';
  const description = imageProps?.description || '';
  if (!imageUrl) return [];

  const name = title || objectId;
  try {
    const response = await fetch(imageUrl, {
      headers: isGoogleServedUrl(imageUrl) ? { 'Authorization': `Bearer ${accessToken}` } : {},
    });
    if (!response.ok) return [{ type: 'text', text: `Failed to fetch image '${name}' (status ${response.status}).` }];

    const tooLarge: ContentBlock = { type: 'text', text: `Image '${name}' exceeds the 20MB limit, skipping.` };
    if (parseInt(response.headers.get('content-length') || '0', 10) > MAX_IMAGE_SIZE) return [tooLarge];
    const bytes = await response.arrayBuffer();
    if (bytes.byteLength > MAX_IMAGE_SIZE) return [tooLarge];

    return [
      ...(title || description ? [{ type: 'text' as const, text: `Image: ${title || description}` }] : []),
      { type: 'image', data: Buffer.from(bytes).toString('base64'), mimeType: response.headers.get('content-type') || 'image/png' },
    ];
  } catch (fetchErr) {
    const msg = fetchErr instanceof Error ? fetchErr.message : String(fetchErr);
    return [{ type: 'text', text: `Failed to fetch image '${name}': ${msg}` }];
  }
}

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

const suggestionsSchema = z.array(z.object({ id: z.string(), kind: z.enum(['insertion', 'deletion', 'formatting']) }));

const segmentSchema = z.object({
  id: z.string().describe('Pass as segment_id to the editing tools'),
  text: z.string(),
  suggestions: suggestionsSchema.optional().describe('Pending suggestions inside it, by id, for review_suggestion'),
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
  suggestions: suggestionsSchema.optional().describe('Pending suggestions touching this element, by id. Pass the ids to review_suggestion to accept or reject them.'),
}).passthrough();

const headerFooterSchema = segmentSchema.extend({
  type: z.enum(['DEFAULT', 'FIRST_PAGE', 'EVEN_PAGE', 'SECTION']),
});

const headingSchema = z.object({
  level: z.number().int().min(1).max(6),
  text: z.string(),
  startIndex: z.number().int(),
  endIndex: z.number().int(),
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

const getDocumentSchema = {
    document_id: z.string().describe('Google Doc ID (from search_documents or a Google Docs URL)'),
    tab_id: z.string().min(1).optional().describe('Read a single tab of a multi-tab document: tabProperties.tabId, the value after ?tab= in the document URL (e.g. "t.abc123"), also listed in this tool\'s `tabs` output. With tab_id, `content`, `structure` and `headings` all describe that tab and their indices are valid for the editing tools when passed the same tab_id. Without it, `content` concatenates all tabs while structure indices cover only the first tab.'),
    format: z.enum(['text', 'markdown']).optional().describe('"text" (default) returns plain text. "markdown" renders headings, bold, italic, links, lists and tables as markdown; embedded images are replaced with a placeholder (get_document_images returns them) and the result is capped at 200,000 characters. It reads the whole document, so it cannot be combined with tab_id, and comment anchors are plain-text offsets, so it cannot be combined with include_comments.'),
    include_structure: z.boolean().optional().describe('Include document structure with startIndex/endIndex for each element. Required before using index-based tools.'),
    include_table_styles: z.boolean().optional().describe('Also report each table\'s cell styles, as a `styles` legend plus a `cell_styles` index grid. Requires include_structure=true. Use it to see how a table is currently formatted before matching or replicating it.'),
    include_comments: z.boolean().optional().describe('Include comment thread metadata (author email/name, timestamp, replies, resolved status, emoji reactions). `content` is returned verbatim — no inline markers are inserted. Each thread carries `anchor_offset: { start, end }` — a half-open span into `content` indicating which text the comment was attached to. Threads with no findable position (document-level comments, or anchor deleted by later edits) omit `anchor_offset`.'),
    suggestions_view_mode: z.enum(SUGGESTIONS_VIEW_MODES).optional().describe('How pending suggestions (Google Docs "Suggesting" mode) appear in `structure`, `headings` and per-tab `content`: SUGGESTIONS_INLINE (the default for editors) shows suggested insertions and still-pending deletions as text, with their ids in each element\'s `suggestions`; PREVIEW_SUGGESTIONS_ACCEPTED shows the document as if every suggestion were accepted; PREVIEW_WITHOUT_SUGGESTIONS as if every one were rejected. Only SUGGESTIONS_INLINE gives indices the editing tools can use: the preview modes shift every index after a pending suggestion, so read with them, then read again without them before an index-based edit. The whole-document `content` export never includes pending suggestions; read one tab with tab_id for text that matches the indices.'),
};

type GetDocumentArgs = z.infer<z.ZodObject<typeof getDocumentSchema>>;

interface ReadPart {
  fields: Record<string, unknown>;
  notes: string[];
}

function validateReadOptions({ format, tab_id, include_comments, include_structure, suggestions_view_mode }: GetDocumentArgs): void {
  const markdownConflicts = [
    ...(tab_id ? ['tab_id (markdown is exported for the whole document)'] : []),
    ...(include_comments ? ['include_comments (comment anchors are plain-text offsets)'] : []),
  ];
  if (format === 'markdown' && markdownConflicts.length > 0) {
    throw new Error(`format "markdown" cannot be combined with ${markdownConflicts.join(' or ')}. Drop format or the other option.`);
  }
  if (suggestions_view_mode && !include_structure && !tab_id) {
    throw new Error('suggestions_view_mode only affects the Docs read behind `structure` and per-tab `content`; pass include_structure=true or tab_id with it, since the whole-document export never includes pending suggestions.');
  }
}

async function readWordDocument(args: GetDocumentArgs, metadata: any, webViewLink: string, accessToken: string) {
  const word = await readWordUpload({
    id: args.document_id,
    name: metadata.name,
    mimeType: metadata.mimeType,
    size: metadata.size ? parseInt(metadata.size) : 0,
    webViewLink,
  }, accessToken, GOOGLE_DRIVE_API);

  const notes: string[] = [];
  if (word.truncated) {
    notes.push(`The text was truncated at ${MAX_WORD_TEXT_CHARS} characters. Full file: ${webViewLink}`);
  }
  if (args.include_structure || args.include_comments || args.tab_id || args.format === 'markdown' || args.suggestions_view_mode) {
    notes.push(
      `'${metadata.name}' is a Word upload, not a Google Doc, ` +
      `so document structure, tabs, comments, suggestions and markdown are unavailable for it.`
    );
  }
  return toolResultWithNotice({
    id: args.document_id,
    title: metadata.name,
    webViewLink,
    content: word.text,
    mimeType: metadata.mimeType,
    ...(word.truncated ? { truncated: true } : {}),
    ...(notes.length ? { message: notes.join(' ') } : {}),
  }, WORD_READ_ONLY_NOTICE);
}

interface TabListing {
  tabs: TabSummary[];
  revisionId?: string;
  note?: string;
}

async function readTabListing(documentId: string, accessToken: string): Promise<TabListing> {
  try {
    return await fetchTabList(documentId, accessToken);
  } catch {
    return { tabs: [], note: 'The tabs listing could not be fetched, so `tabs` and `revisionId` are missing. Retry if you need them.' };
  }
}

function tabContentPart(resolved: ResolvedTab): ReadPart & { content: string } {
  return {
    content: renderBodyText(resolved.body),
    fields: { tab: { tabId: resolved.tabId, ...(resolved.title !== undefined ? { title: resolved.title } : {}) } },
    notes: [],
  };
}

async function exportDocument(documentId: string, accessToken: string, asMarkdown: boolean): Promise<ReadPart & { content: string }> {
  // Accept-Language pins the footer to the English template the reaction parser reads
  const exportMime = encodeURIComponent(asMarkdown ? 'text/markdown' : 'text/plain');
  const response = await fetch(`${GOOGLE_DRIVE_API}/files/${encodeURIComponent(documentId)}/export?mimeType=${exportMime}`, {
    headers: {
      'Authorization': `Bearer ${accessToken}`,
      'Accept-Language': 'en-US',
    },
  });
  if (!response.ok) {
    throw await buildGoogleApiError(response, 'drive');
  }
  const text = await response.text();
  if (!asMarkdown) return { content: text, fields: {}, notes: [] };

  const compacted = compactMarkdownExport(text);
  const notes: string[] = [];
  if (compacted.imagesOmitted > 0) {
    notes.push(`${compacted.imagesOmitted} embedded image(s) were replaced with "embedded-image-omitted"; get_document_images returns them.`);
  }
  if (compacted.truncated) {
    notes.push(`The markdown was truncated at ${MAX_MARKDOWN_CHARS} characters. Use format "text" or read one tab with tab_id for the rest.`);
  }
  return {
    content: compacted.content,
    fields: { format: 'markdown', ...(compacted.truncated ? { truncated: true } : {}) },
    notes,
  };
}

function structurePart(
  resolved: ResolvedTab,
  { include_table_styles, suggestions_view_mode: suggestionsViewMode }: GetDocumentArgs,
): ReadPart {
  const parsed = parseDocumentStructure(resolved.body, include_table_styles === true);
  const fields: Record<string, unknown> = {
    structure: parsed.elements,
    headings: headingsOf(parsed.elements),
  };
  const headers = listHeaderFooters('header', resolved.documentStyle, resolved.headers);
  const footers = listHeaderFooters('footer', resolved.documentStyle, resolved.footers);
  const footnotes = Object.entries(resolved.footnotes).map(([id, note]) => summarizeSegment(id, note));
  if (headers.length) fields.headers = headers;
  if (footers.length) fields.footers = footers;
  if (footnotes.length) fields.footnotes = footnotes;

  const notes: string[] = [];
  if (parsed.tablesWithoutCells > 0) {
    notes.push(`${parsed.tablesWithoutCells} table(s) report dimensions without cell contents because the document exhausted the structure budget for table cells. Their text, if any, is in \`content\`.`);
  }
  if (parsed.tablesWithoutStyles > 0) {
    notes.push(`${parsed.tablesWithoutStyles} table(s) report no styles because the document exhausted the structure budget. Read those tables on their own to see their formatting.`);
  }
  if (suggestionsViewMode?.startsWith('PREVIEW_')) {
    notes.push(`The indices in \`structure\` describe the ${suggestionsViewMode} preview, not the document as it is: do not pass them to the editing tools. Read again without suggestions_view_mode before an index-based edit.`);
  }
  return { fields, notes };
}

export const readTools = {
  search_documents: defineTool(SCOPES.DRIVE_READONLY, {
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
    run: async ({ name, page_token }, { accessToken }) => {
      // Drive's query parser rejects newlines, so refuse them rather than search for something else
      if (name && /[\r\n]/.test(name)) {
        throw new Error('search name must not contain newline characters');
      }
      const escapedName = escapeDriveQueryName(name);
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
    },
  }),

  get_document: defineTool(SCOPES.DRIVE_READONLY, {
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
      footnotes: z.array(segmentSchema).optional().describe('With include_structure: the footnotes of the tab being read; the footnoteReference elements in `structure` say where each mark sits'),
      tabs: z.array(tabSummarySchema).optional(),
      tab: z.object({
        tabId: z.string(),
        title: z.string().optional(),
      }).optional().describe('Present when tab_id was passed: the tab this response describes'),
      threads: z.array(threadSchema).optional(),
    },
    schema: getDocumentSchema,
    run: async (args, { accessToken }) => {
      const { document_id, tab_id, include_structure, include_comments, include_table_styles } = args;
      const metadata = await makeDriveRequest(
        `/files/${encodeURIComponent(document_id)}?fields=name,mimeType,size,webViewLink&supportsAllDrives=true`,
        accessToken
      );
      const mimeType: string = metadata.mimeType || '';
      const webViewLink: string =
        metadata.webViewLink || `https://drive.google.com/file/d/${document_id}/view`;

      // Drive's export endpoint 403s on Word uploads, so these are parsed here.
      if (isWordMime(mimeType)) return readWordDocument(args, metadata, webViewLink, accessToken);
      if (mimeType && mimeType !== NATIVE_DOC_MIME) {
        throw new Error(
          `'${metadata.name}' has mimeType '${mimeType}', which this connector cannot read. ` +
          `This tool reads Google Docs and Word (.doc/.docx) files. ` +
          `Open it directly: ${webViewLink}`
        );
      }
      validateReadOptions(args);

      const docs = include_structure || tab_id
        ? fetchResolvedTab(document_id, accessToken, tab_id, args.suggestions_view_mode).then(({ resolved }) => resolved)
        : undefined;
      const tabListing: Promise<TabListing> = docs ?? readTabListing(document_id, accessToken);
      const [resolved, listing, read, comments] = await Promise.all([
        docs,
        tabListing,
        docs && tab_id ? docs.then(tabContentPart) : exportDocument(document_id, accessToken, args.format === 'markdown'),
        include_comments ? fetchAllComments(document_id, accessToken) : undefined,
      ]);
      const notes = [...(listing.note ? [listing.note] : []), ...read.notes];

      const output: Record<string, unknown> = {
        id: document_id,
        title: metadata.name,
        webViewLink: metadata.webViewLink,
        ...read.fields,
        content: read.content,
      };
      if (comments) {
        const { cleanedContent, threads } = anchorThreads(read.content, comments, !tab_id);
        output.content = cleanedContent;
        output.threads = threads;
      }
      if (listing.tabs.length > 0) output.tabs = listing.tabs;
      if (typeof listing.revisionId === 'string') output.revisionId = listing.revisionId;

      if (resolved && include_structure) {
        const structure = structurePart(resolved, args);
        Object.assign(output, structure.fields);
        notes.push(...structure.notes);
      } else if (include_table_styles) {
        notes.push('include_table_styles needs include_structure=true, since styles are reported on the table elements in `structure`. No structure was returned.');
      }
      if (notes.length) output.message = notes.join(' ');

      return toolResult(output);
    },
  }, { labels: true }),

  get_document_images: defineTool(SCOPES.DRIVE_READONLY, {
    description: 'Extract all inline images from a Google Doc, across every tab, or from one tab when tab_id is passed. Returns each image as an inline image content block. Use get_document with include_structure=true first to see where images are positioned in the document.',
    readOnlyHint: true,
    schema: {
      document_id: z.string().describe('Google Doc ID (from search_documents or a Google Docs URL)'),
      tab_id: z.string().min(1).optional().describe('Only extract images from this tab: tabProperties.tabId, the value after ?tab= in the document URL, also listed in get_document\'s `tabs`. Omit for all tabs.'),
    },
    run: async ({ document_id, tab_id }, { accessToken }) => {
      const { doc } = await fetchResolvedTab(document_id, accessToken, tab_id);
      const inlineObjects = collectInlineObjects(doc, tab_id);
      const objectIds = Object.keys(inlineObjects);

      if (objectIds.length === 0) {
        return {
          content: [{ type: 'text' as const, text: 'No images found in this document.' }],
        };
      }

      const perImage = await mapWithLimit(objectIds, IMAGE_FETCH_CONCURRENCY, (objectId) => imageBlocks(objectId, inlineObjects[objectId], accessToken));
      const content = perImage.flat();

      if (content.length === 0) {
        return {
          content: [{ type: 'text' as const, text: 'Found image references but could not fetch any images.' }],
        };
      }

      return { content };
    },
  }, { labels: true }),
};
