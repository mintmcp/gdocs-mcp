import {
  describeTable,
  newStructureBudget,
  type ReportedCellStyle,
} from './tableRead.js';

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

export interface Suggestion {
  id: string;
  kind: 'insertion' | 'deletion' | 'formatting';
}

function collectSuggestions(nodes: any[]): Suggestion[] {
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

function suggestionNodes(content: any[]): any[] {
  return content.flatMap((element: any) => {
    if (element.sectionBreak) return [element.sectionBreak];
    if (element.tableOfContents) return [element.tableOfContents, ...suggestionNodes(element.tableOfContents.content ?? [])];
    if (element.paragraph) {
      return [element.paragraph, ...(element.paragraph.elements ?? []).flatMap((el: any) => Object.values(el))];
    }
    if (element.table) {
      const rows = element.table.tableRows ?? [];
      const cells = rows.flatMap((r: any) => r.tableCells ?? []);
      return [element.table, ...rows, ...cells, ...suggestionNodes(cells.flatMap((c: any) => c.content ?? []))];
    }
    return [];
  });
}

export function contentSuggestions(content: any[]): Suggestion[] {
  return collectSuggestions(suggestionNodes(content));
}

export interface ParsedStructure {
  elements: StructureElement[];
  tablesWithoutCells: number;
  tablesWithoutStyles: number;
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

export function headingsOf(elements: StructureElement[]): Heading[] {
  return elements.flatMap((el) => el.headingLevel === undefined ? [] : [{
    level: el.headingLevel,
    text: (el.text ?? '').replace(/\n+$/, ''),
    startIndex: el.startIndex,
    endIndex: el.endIndex,
  }]);
}

const paragraphText = (paragraph: any): string =>
  (paragraph.elements ?? []).map((el: any) => el.textRun?.content || '').join('');

export function renderBodyText(content: any[]): string {
  return content.map((element: any) => {
    if (element.paragraph) return paragraphText(element.paragraph);
    if (element.tableOfContents) return renderBodyText(element.tableOfContents.content ?? []);
    if (element.table) {
      const rows = (element.table.tableRows ?? []).map((row: any) =>
        (row.tableCells ?? []).map((cell: any) => renderBodyText(cell.content ?? []).replace(/\n$/, '')).join('\t'));
      return rows.length ? rows.join('\n') + '\n' : '';
    }
    return '';
  }).join('');
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
      const text = paragraphText(element.paragraph);
      const headingLevel = namedStyleToHeadingLevel(element.paragraph.paragraphStyle?.namedStyleType);
      const bullet = element.paragraph.bullet;
      const suggestions = contentSuggestions([element]);
      elements.push({
        type: 'paragraph',
        startIndex: element.startIndex ?? 0,
        endIndex: element.endIndex,
        text,
        ...(headingLevel !== undefined ? { headingLevel } : {}),
        ...(bullet ? { bullet: { nestingLevel: bullet.nestingLevel ?? 0 } } : {}),
        ...(suggestions.length ? { suggestions } : {}),
      });

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
      const suggestions = contentSuggestions([element]);
      elements.push({
        type: 'table',
        ...describeTable(element, budget, includeStyles),
        ...(suggestions.length ? { suggestions } : {}),
      });
    } else if (element.sectionBreak || element.tableOfContents) {
      const suggestions = contentSuggestions([element]);
      elements.push({
        type: element.sectionBreak ? 'sectionBreak' : 'tableOfContents',
        startIndex: element.startIndex ?? 0,
        endIndex: element.endIndex,
        ...(suggestions.length ? { suggestions } : {}),
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
