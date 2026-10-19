/**
 * Pure helpers for parsing Google Docs document structure and building Docs
 * API batchUpdate requests. Extracted from tools.ts so they can be unit-tested
 * in isolation without mocking the network or the MCP tool wrapper.
 *
 * Everything in this file is referentially transparent — no I/O, no global
 * state, no SDK dependencies. Tool handlers in tools.ts compose these helpers
 * with the Docs/Drive REST calls.
 */

import {
  describeTable,
  hexToRgbColor,
  newStructureBudget,
  type ReportedCellStyle,
} from './tableModel.js';

/**
 * Reactions parsed from Drive's plain-text export footer.
 */
export interface Reaction {
  author_name: string;
  emoji: string;
  timestamp: string;
}

/**
 * Tab summary returned to the caller when a document uses Tabs.
 */
export interface TabSummary {
  tabId: string;
  title?: string;
  index?: number;
  nestingLevel?: number;
}

/**
 * A heading paragraph extracted from `body.content`, with the indices needed
 * to drive index-based mutating tools (`insert_text`, `delete_content`, etc.).
 */
export interface Heading {
  level: number;
  text: string;
  startIndex: number;
  endIndex: number;
}

/**
 * A structural element flattened from `body.content`. Mirrors the small subset
 * of Google Docs element types this server cares about for index-based edits.
 * Table elements additionally carry the shape and styling fields assembled by
 * `describeTable`, so a caller can locate cells without a second read.
 */
export interface StructureElement {
  type: string;
  startIndex: number;
  endIndex: number;
  text?: string;
  inlineObjectId?: string;
  footnoteId?: string;
  headingLevel?: number;
  bullet?: { nestingLevel: number };
  rows?: number;
  columns?: number;
  cells?: string[][];
  has_merged_cells?: boolean;
  styles?: ReportedCellStyle[];
  cell_styles?: number[][];
  suggestions?: Suggestion[];
}

/** A pending suggestion touching an element, by id, for review_suggestion. */
export interface Suggestion {
  id: string;
  kind: 'insertion' | 'deletion' | 'formatting';
}

/**
 * Collect the suggestion ids on a paragraph's runs or on a table, each once,
 * in the order they appear. Insertions and deletions carry id lists; a
 * suggested style change is a map keyed by suggestion id under a
 * `suggested...Changes` field (text style, paragraph style, bullets, cells).
 */
export function collectSuggestions(nodes: any[]): Suggestion[] {
  const out: Suggestion[] = [];
  const seen = new Set<string>();
  const add = (id: unknown, kind: Suggestion['kind']) => {
    if (typeof id !== 'string' || seen.has(id)) return;
    seen.add(id);
    out.push({ id, kind });
  };
  for (const node of nodes) {
    if (!node || typeof node !== 'object') continue;
    for (const id of node.suggestedInsertionIds ?? []) add(id, 'insertion');
    for (const id of node.suggestedDeletionIds ?? []) add(id, 'deletion');
    for (const [field, value] of Object.entries(node)) {
      if (!/^suggested.*Changes$/.test(field) || !value || typeof value !== 'object') continue;
      for (const id of Object.keys(value)) add(id, 'formatting');
    }
  }
  return out;
}

/**
 * Structure plus a count of tables whose cells or styles were dropped, so
 * get_document can say so rather than silently returning a partial picture.
 */
export interface ParsedStructure {
  elements: StructureElement[];
  tablesWithoutCells: number;
  tablesWithoutStyles: number;
}

/**
 * Parsed artifacts from Drive's plain-text export. See `parseExportArtifacts`.
 */
export interface ParsedExport {
  body: string;
  letterContent: Map<string, string>;
  letterReactions: Map<string, Reaction[]>;
  commentLetters: Set<string>;
}

export interface WriteControl {
  requiredRevisionId?: string;
  writeMode?: 'SUGGEST';
}

/**
 * Build the `writeControl` object for a Docs batchUpdate. When
 * `requiredRevisionId` is provided, Google rejects the write if the document's
 * current revision doesn't match — surfacing optimistic-concurrency conflicts
 * as a 400 instead of silently overwriting a concurrent edit. With `suggest`
 * every request in the batch lands as a tracked suggestion instead of an edit.
 * Returns `undefined` when neither was requested so callers can omit it.
 */
export function buildWriteControl(
  requiredRevisionId: string | undefined,
  suggest = false,
): WriteControl | undefined {
  const pinned = Boolean(requiredRevisionId) && typeof requiredRevisionId === 'string';
  if (!pinned && !suggest) return undefined;
  return {
    ...(pinned ? { requiredRevisionId } : {}),
    ...(suggest ? { writeMode: 'SUGGEST' as const } : {}),
  };
}

/**
 * Flatten the nested `tabs` array (a tab can have `childTabs`) into a single
 * summary list. Returns an empty array for documents without tabs. Each entry
 * surfaces `tabId` + `title` (+ position metadata) so callers know multi-tab
 * documents exist; this server's index-based mutating tools target the body
 * of the default tab and do not yet support cross-tab editing.
 */
export function summarizeTabs(tabs: any[] | undefined): TabSummary[] {
  if (!tabs || tabs.length === 0) return [];
  const out: TabSummary[] = [];
  const walk = (list: any[], depth: number) => {
    for (const tab of list) {
      const props = tab?.tabProperties || {};
      if (props.tabId) {
        out.push({
          tabId: props.tabId,
          ...(props.title !== undefined ? { title: props.title } : {}),
          ...(props.index !== undefined ? { index: props.index } : {}),
          ...(props.nestingLevel !== undefined ? { nestingLevel: props.nestingLevel } : { nestingLevel: depth }),
        });
      }
      if (Array.isArray(tab?.childTabs) && tab.childTabs.length > 0) {
        walk(tab.childTabs, depth + 1);
      }
    }
  };
  walk(tabs, 0);
  return out;
}

/**
 * Map a Docs `namedStyleType` enum (e.g. `HEADING_3`) to its numeric heading
 * level (1-6). Returns `undefined` for non-heading styles like `NORMAL_TEXT`,
 * `TITLE`, `SUBTITLE` — callers should treat those as body paragraphs.
 */
export function namedStyleToHeadingLevel(namedStyleType: string | undefined): number | undefined {
  if (!namedStyleType) return undefined;
  const m = /^HEADING_([1-6])$/.exec(namedStyleType);
  return m ? parseInt(m[1], 10) : undefined;
}

/**
 * Extract heading paragraphs from `body.content` for ergonomic
 * "insert-after-heading" / "find-section" flows. Returns headings in document
 * order with the indices needed to drive `insert_text` / `delete_content`.
 * The `text` is trimmed of trailing newlines but preserves internal spacing.
 */
export function extractHeadings(content: any[]): Heading[] {
  const headings: Heading[] = [];
  for (const element of content) {
    if (!element.paragraph) continue;
    const level = namedStyleToHeadingLevel(element.paragraph.paragraphStyle?.namedStyleType);
    if (level === undefined) continue;
    const text = (element.paragraph.elements
      ?.map((el: any) => el.textRun?.content || '')
      .join('') || '').replace(/\n+$/, '');
    headings.push({
      level,
      text,
      startIndex: element.startIndex ?? 0,
      endIndex: element.endIndex,
    });
  }
  return headings;
}

/**
 * Parse document body.content into structural elements with indices.
 *
 * Recognises paragraphs (with embedded inline images), tables, section breaks,
 * and tables of contents. Unknown element types are skipped silently — Docs
 * occasionally introduces new structural elements and we'd rather emit a
 * partial structure than crash.
 *
 * Tables are described down to their cell text, and to their per-cell styling
 * when `includeStyles` is set. Both draw on one document-wide budget: a
 * document with many tables cannot drive unbounded work, and whatever the
 * budget refuses is counted so the caller can report it.
 */
export function parseDocumentStructure(
  content: any[],
  includeStyles = false,
): ParsedStructure {
  const elements: StructureElement[] = [];
  const budget = newStructureBudget();

  for (const element of content) {
    if (element.paragraph) {
      const text = element.paragraph.elements
        ?.map((el: any) => el.textRun?.content || '')
        .join('') || '';
      const headingLevel = namedStyleToHeadingLevel(element.paragraph.paragraphStyle?.namedStyleType);
      const bullet = element.paragraph.bullet;
      const suggestions = collectSuggestions([
        element.paragraph,
        ...(element.paragraph.elements || []).map((el: any) => el.textRun ?? el.inlineObjectElement ?? el.pageBreak ?? el.footnoteReference ?? {}),
      ]);
      elements.push({
        type: 'paragraph',
        startIndex: element.startIndex ?? 0,
        endIndex: element.endIndex,
        text,
        ...(headingLevel !== undefined ? { headingLevel } : {}),
        ...(bullet ? { bullet: { nestingLevel: bullet.nestingLevel ?? 0 } } : {}),
        ...(suggestions.length ? { suggestions } : {}),
      });

      // Inline images and footnote marks get elements of their own, so their ids are addressable
      for (const el of element.paragraph.elements || []) {
        if (el.inlineObjectElement) {
          elements.push({
            type: 'inlineImage',
            startIndex: el.startIndex,
            endIndex: el.endIndex,
            inlineObjectId: el.inlineObjectElement.inlineObjectId,
          });
        } else if (el.footnoteReference?.footnoteId) {
          elements.push({
            type: 'footnoteReference',
            startIndex: el.startIndex,
            endIndex: el.endIndex,
            footnoteId: el.footnoteReference.footnoteId,
          });
        }
      }
    } else if (element.table) {
      // Suggested rows, cells and cell styles sit on the rows and cells, not on the table
      const rows = element.table.tableRows ?? [];
      const suggestions = collectSuggestions([element.table, ...rows, ...rows.flatMap((r: any) => r.tableCells ?? [])]);
      elements.push({
        type: 'table',
        ...describeTable(element, budget, includeStyles),
        ...(suggestions.length ? { suggestions } : {}),
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

  return {
    elements,
    tablesWithoutCells: budget.tablesWithoutCells,
    tablesWithoutStyles: budget.tablesWithoutStyles,
  };
}

/**
 * Build a Docs API `updateTextStyle` request that clears inherited character
 * formatting (bold, italic, underline, strikethrough, font size, colour, link)
 * on the given range.
 * Centralised so all "insert + optionally reset formatting" tools agree on the
 * exact set of fields cleared.
 */
export function clearInheritedFormattingRequest(startIndex: number, endIndex: number, tabId?: string, segmentId?: string): any {
  return {
    updateTextStyle: {
      range: { startIndex, endIndex, ...(tabId ? { tabId } : {}), ...(segmentId ? { segmentId } : {}) },
      textStyle: {},
      fields: 'bold,italic,underline,strikethrough,fontSize,foregroundColor,link',
    },
  };
}

export type ListStyle = 'bullet' | 'numbered' | 'checkbox' | 'none';

const LIST_PRESETS: Record<Exclude<ListStyle, 'none'>, string> = {
  bullet: 'BULLET_DISC_CIRCLE_SQUARE',
  numbered: 'NUMBERED_DECIMAL_ALPHA_ROMAN',
  checkbox: 'BULLET_CHECKBOX',
};

/** Google's glyph presets, for callers who want a specific look. */
export const BULLET_PRESETS = [
  'BULLET_DISC_CIRCLE_SQUARE', 'BULLET_DIAMONDX_ARROW3D_SQUARE', 'BULLET_CHECKBOX', 'BULLET_ARROW_DIAMOND_DISC',
  'BULLET_STAR_CIRCLE_SQUARE', 'BULLET_ARROW3D_CIRCLE_SQUARE', 'BULLET_LEFTTRIANGLE_DIAMOND_DISC',
  'BULLET_DIAMONDX_HOLLOWDIAMOND_SQUARE', 'BULLET_DIAMOND_CIRCLE_SQUARE', 'NUMBERED_DECIMAL_ALPHA_ROMAN',
  'NUMBERED_DECIMAL_ALPHA_ROMAN_PARENS', 'NUMBERED_DECIMAL_NESTED', 'NUMBERED_UPPERALPHA_ALPHA_ROMAN',
  'NUMBERED_UPPERROMAN_UPPERALPHA_DECIMAL', 'NUMBERED_ZERODECIMAL_ALPHA_ROMAN',
] as const;

/**
 * Bullets are their own request type, not a paragraph style field, so they
 * ride alongside updateParagraphStyle in the same batch
 */
export function listStyleRequest(startIndex: number, endIndex: number, listStyle: ListStyle, tabId?: string, segmentId?: string, preset?: string): any {
  const range = { startIndex, endIndex, ...(tabId ? { tabId } : {}), ...(segmentId ? { segmentId } : {}) };
  if (listStyle === 'none') return { deleteParagraphBullets: { range } };
  return { createParagraphBullets: { range, bulletPreset: preset ?? LIST_PRESETS[listStyle] } };
}

export const NAMED_STYLES = ['NORMAL_TEXT', 'TITLE', 'SUBTITLE', 'HEADING_1', 'HEADING_2', 'HEADING_3', 'HEADING_4', 'HEADING_5', 'HEADING_6'] as const;

export interface ParagraphStyleArgs {
  heading_level?: number;
  named_style?: string;
  alignment?: string;
  line_spacing?: number;
  space_above_pt?: number;
  space_below_pt?: number;
  indent_start_pt?: number;
  indent_end_pt?: number;
  indent_first_line_pt?: number;
  keep_with_next?: boolean;
  page_break_before?: boolean;
}

const PARAGRAPH_DIMENSIONS: Array<[keyof ParagraphStyleArgs, string]> = [
  ['space_above_pt', 'spaceAbove'],
  ['space_below_pt', 'spaceBelow'],
  ['indent_start_pt', 'indentStart'],
  ['indent_end_pt', 'indentEnd'],
  ['indent_first_line_pt', 'indentFirstLine'],
];

/** The updateParagraphStyle payload for update_paragraph_style, with its field mask in argument order */
export function buildParagraphStyle(args: ParagraphStyleArgs): { paragraphStyle: Record<string, unknown>; fields: string[] } {
  const paragraphStyle: Record<string, unknown> = {};
  const fields: string[] = [];

  if (args.heading_level !== undefined && args.named_style !== undefined) {
    throw new Error('Pass one of heading_level or named_style, not both: both set the paragraph\'s named style.');
  }
  if (args.heading_level !== undefined) {
    if (!Number.isInteger(args.heading_level) || args.heading_level < 0 || args.heading_level > 6) {
      throw new Error('heading_level must be 0 (normal) or 1-6');
    }
    paragraphStyle.namedStyleType = args.heading_level === 0 ? 'NORMAL_TEXT' : `HEADING_${args.heading_level}`;
    fields.push('namedStyleType');
  }
  if (args.named_style !== undefined) {
    paragraphStyle.namedStyleType = args.named_style;
    fields.push('namedStyleType');
  }
  if (args.alignment !== undefined) {
    paragraphStyle.alignment = args.alignment;
    fields.push('alignment');
  }
  if (args.line_spacing !== undefined) {
    paragraphStyle.lineSpacing = args.line_spacing;
    fields.push('lineSpacing');
  }
  for (const [arg, field] of PARAGRAPH_DIMENSIONS) {
    const value = args[arg];
    if (value === undefined) continue;
    paragraphStyle[field] = { magnitude: value, unit: 'PT' };
    fields.push(field);
  }
  if (args.keep_with_next !== undefined) {
    paragraphStyle.keepWithNext = args.keep_with_next;
    fields.push('keepWithNext');
  }
  if (args.page_break_before !== undefined) {
    paragraphStyle.pageBreakBefore = args.page_break_before;
    fields.push('pageBreakBefore');
  }
  return { paragraphStyle, fields };
}

export interface TextStyleArgs {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strikethrough?: boolean;
  small_caps?: boolean;
  font_size?: number;
  font_family?: string;
  color?: string;
  background_color?: string;
  baseline_offset?: 'NONE' | 'SUPERSCRIPT' | 'SUBSCRIPT';
  link_url?: string;
}

/** The updateTextStyle payload for update_text_style, with its field mask in argument order */
export function buildTextStyle(args: TextStyleArgs): { textStyle: Record<string, unknown>; fields: string[] } {
  const textStyle: Record<string, unknown> = {};
  const fields: string[] = [];
  for (const key of ['bold', 'italic', 'underline', 'strikethrough'] as const) {
    if (args[key] !== undefined) {
      textStyle[key] = args[key];
      fields.push(key);
    }
  }
  if (args.font_size !== undefined) {
    textStyle.fontSize = { magnitude: args.font_size, unit: 'PT' };
    fields.push('fontSize');
  }
  if (args.font_family !== undefined) {
    textStyle.weightedFontFamily = { fontFamily: args.font_family };
    fields.push('weightedFontFamily');
  }
  if (args.color !== undefined) {
    textStyle.foregroundColor = { color: hexToRgbColor(args.color, 'color') };
    fields.push('foregroundColor');
  }
  if (args.background_color !== undefined) {
    // An empty colour object is how Google clears a highlight
    textStyle.backgroundColor = args.background_color === '' ? {} : { color: hexToRgbColor(args.background_color, 'background_color') };
    fields.push('backgroundColor');
  }
  if (args.baseline_offset !== undefined) {
    textStyle.baselineOffset = args.baseline_offset;
    fields.push('baselineOffset');
  }
  if (args.small_caps !== undefined) {
    textStyle.smallCaps = args.small_caps;
    fields.push('smallCaps');
  }
  if (args.link_url !== undefined) {
    textStyle.link = args.link_url === '' ? null : { url: args.link_url };
    fields.push('link');
  }
  if (fields.length === 0) {
    throw new Error('At least one style property must be provided (bold, italic, underline, strikethrough, small_caps, font_size, font_family, color, background_color, baseline_offset, link_url)');
  }
  return { textStyle, fields };
}

/**
 * A paragraph split off the end of a heading or list item inherits that
 * paragraph style; this resets it so appended text reads as body text
 */
export function clearInheritedParagraphStyleRequests(startIndex: number, endIndex: number, tabId?: string): any[] {
  const range = { startIndex, endIndex, ...(tabId ? { tabId } : {}) };
  return [
    { updateParagraphStyle: { range, paragraphStyle: { namedStyleType: 'NORMAL_TEXT' }, fields: 'namedStyleType' } },
    { deleteParagraphBullets: { range } },
  ];
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
 * Empty cells are skipped (no insertText request emitted) since they're
 * already empty after the table insert.
 */
export function buildTableInsertRequests(
  tableData: string[][],
  tableInsertIndex: number,
  tabId?: string,
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
            location: { index: insertIndex, ...(tabId ? { tabId } : {}) },
          },
        });
      }
    }
  }

  return requests;
}

/**
 * Escape a string for safe inclusion in a Drive query `name contains '...'`
 * clause. Drive query strings are single-quoted; escape backslash first so the
 * next-step `'` escapes survive, then escape `'`. Returns `undefined` for an
 * empty/undefined input so the caller can omit the clause entirely.
 *
 * Newline characters MUST be checked separately and rejected by the caller —
 * Drive's query parser rejects them, and a silent mismatch is worse than a
 * clear error.
 */
export function escapeDriveQueryName(name: string | undefined): string | undefined {
  if (!name) return undefined;
  return name.replace(/\\/g, '\\\\').replace(/'/g, "\\'");
}

export function isGoogleServedUrl(url: string): boolean {
  let parsed: URL;
  try {
    parsed = new URL(url);
  } catch {
    return false;
  }
  if (parsed.protocol !== 'https:') return false;
  const host = parsed.hostname;
  return host === 'googleusercontent.com' || host.endsWith('.googleusercontent.com') ||
    host === 'googleapis.com' || host.endsWith('.googleapis.com');
}

/**
 * Validate an index range for a Docs body mutation. Throws a descriptive error
 * if the range is invalid. Centralised so all mutating tools that take
 * (startIndex, endIndex) agree on the validation rules.
 */
export function validateIndexRange(startIndex: number, endIndex: number): void {
  if (!Number.isFinite(startIndex) || startIndex < 0) {
    throw new Error('startIndex must be >= 0');
  }
  if (!Number.isFinite(endIndex) || endIndex <= startIndex) {
    throw new Error('endIndex must be greater than startIndex');
  }
}

/**
 * Validate an insertion index. Docs body content starts at index 1 (index 0
 * is the document start sentinel); inserting at 0 always 400s. Header, footer
 * and footnote segments have no sentinel, so their text starts at index 0.
 */
export function validateInsertIndex(index: number, inSegment = false): void {
  const min = inSegment ? 0 : 1;
  if (!Number.isFinite(index) || index < min) {
    throw new Error(inSegment
      ? 'index must be >= 0 (header, footer and footnote text starts at index 0)'
      : 'index must be >= 1 (Docs body starts at index 1)');
  }
}

/**
 * Normalise whitespace for fuzzy text matching when correlating Drive comments
 * back to the export footer's letter markers.
 */
export function normalizeForMatch(text: string): string {
  return text.replace(/\s+/g, ' ').trim();
}

// Reactions only ever surface in the plain-text export footer (Drive Comments
// API has no reaction field). Format is the English template
// `<displayName> reacted with <emoji> at YYYY-MM-DD HH:MM (AM|PM)`. The HH is
// 24-hour despite the AM/PM suffix, so we drop the suffix when parsing.
const REACTION_RE = /^(.+?) reacted with (\S+) at (\d{4}-\d{2}-\d{2} \d{2}:\d{2})(?:\s*[AP]M)?\s*$/;

/**
 * Drive's plain-text export injects letter markers (`[a][b][c]…`) inline at
 * every comment/reply anchor position and lists each letter's comment text
 * (plus any reaction continuation lines) in a trailing block. The trailing
 * block is unambiguous and safe to strip; the inline markers are not — a
 * legitimate `[a]` typed into the doc body looks identical to a comment
 * marker, so we leave them in place here and let the caller rewrite only
 * at confirmed thread anchor positions.
 */
export function parseExportArtifacts(plainText: string): ParsedExport {
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

/**
 * Find the letter marker (e.g. "a") whose comment text matches `apiContent`
 * (whitespace-normalised). Returns `undefined` if no match.
 */
export function findLetterByContent(
  letterContent: Map<string, string>,
  apiContent: string,
): string | undefined {
  const target = normalizeForMatch(apiContent);
  if (!target) return undefined;
  for (const [letter, content] of letterContent) {
    if (normalizeForMatch(content) === target) return letter;
  }
  return undefined;
}

export const MAX_MARKDOWN_CHARS = 200_000;

const EMBEDDED_IMAGE_RE = /data:image\/[\w.+-]+;base64,[A-Za-z0-9+/=]+/g;

/**
 * Drive's markdown export inlines every image as a base64 data URI, which can
 * run to megabytes for a few screenshots; get_document_images serves those
 */
export function compactMarkdownExport(markdown: string, maxChars = MAX_MARKDOWN_CHARS): { content: string; imagesOmitted: number; truncated: boolean } {
  let imagesOmitted = 0;
  const stripped = markdown.replace(EMBEDDED_IMAGE_RE, () => {
    imagesOmitted++;
    return 'embedded-image-omitted';
  });
  const truncated = stripped.length > maxChars;
  return { content: truncated ? stripped.slice(0, maxChars) : stripped, imagesOmitted, truncated };
}
