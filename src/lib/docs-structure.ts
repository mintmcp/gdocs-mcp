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
  headingLevel?: number;
  rows?: number;
  columns?: number;
  cells?: string[][];
  has_merged_cells?: boolean;
  styles?: ReportedCellStyle[];
  cell_styles?: number[][];
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

/**
 * Build the `writeControl` object for a Docs batchUpdate. When
 * `requiredRevisionId` is provided, Google rejects the write if the document's
 * current revision doesn't match — surfacing optimistic-concurrency conflicts
 * as a 400 instead of silently overwriting a concurrent edit. Returns
 * `undefined` when no constraint was requested so callers can spread it
 * conditionally into the batchUpdate body.
 */
export function buildWriteControl(
  requiredRevisionId: string | undefined,
): { requiredRevisionId: string } | undefined {
  if (!requiredRevisionId || typeof requiredRevisionId !== 'string') return undefined;
  return { requiredRevisionId };
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
      elements.push({
        type: 'paragraph',
        startIndex: element.startIndex ?? 0,
        endIndex: element.endIndex,
        text,
        ...(headingLevel !== undefined ? { headingLevel } : {}),
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
      elements.push({ type: 'table', ...describeTable(element, budget, includeStyles) });
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
 * formatting (bold, italic, underline, strikethrough, link) on the given range.
 * Centralised so all "insert + optionally reset formatting" tools agree on the
 * exact set of fields cleared.
 */
export function clearInheritedFormattingRequest(startIndex: number, endIndex: number): any {
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
 * Empty cells are skipped (no insertText request emitted) since they're
 * already empty after the table insert.
 */
export function buildTableInsertRequests(
  tableData: string[][],
  tableInsertIndex: number,
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
 * is the document start sentinel); inserting at 0 always 400s.
 */
export function validateInsertIndex(index: number): void {
  if (!Number.isFinite(index) || index < 1) {
    throw new Error('index must be >= 1 (Docs body starts at index 1)');
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
