/**
 * Google Docs table mechanics: reading tables out of a document and building the
 * requests that create or style them. Pure translation to and from Google's
 * representation, so nothing here performs IO or knows about MCP.
 */

import { z } from 'zod';

// Bound what we build for Google; an oversized payload would otherwise exhaust the isolate.
export const MAX_TABLE_ROWS = 1000;
export const MAX_TABLE_COLUMNS = 100;
export const MAX_TABLE_CELLS = 10_000;
export const MAX_CELL_CHARS = 10_000;
export const MAX_TABLE_TEXT_CHARS = 500_000;

// `content` already carries the full document, so table cells in the structure
// output are for locating coordinates, not for a second copy of the text.
export const MAX_STRUCTURE_CELL_CHARS = 200;
export const MAX_STRUCTURE_TABLE_CHARS = 100_000;
export const MAX_STRUCTURE_CELLS = 10_000;

export const MAX_BORDER_WIDTH_PT = 20;

export interface DocsTableElement {
  table?: {
    rows?: number;
    columns?: number;
    tableRows?: Array<{ tableCells?: Array<{ content?: any[]; tableCellStyle?: any }> }>;
  };
  startIndex?: number;
  endIndex: number;
}

export interface DocsBody {
  revisionId?: string;
  body?: { content?: DocsTableElement[] };
}

/** insertTable writes a newline before the table, so it starts one past the index. */
export const tableStartFor = (index: number) => index + 1;

/**
 * Takes a table's own start index, not the insertion index findTableAt takes.
 * insert_table's index is one below the table it creates; these must not be conflated.
 */
export function findTableStartingAt(doc: DocsBody, tableStartIndex: number): DocsTableElement | undefined {
  return (doc?.body?.content ?? []).find((c) => c.table && c.startIndex === tableStartIndex);
}

/** Exact start only: a looser match could style a table we did not create. */
export function findTableAt(doc: DocsBody, index: number): { startIndex: number; endIndex: number } | undefined {
  const el = findTableStartingAt(doc, tableStartFor(index));
  return el ? { startIndex: el.startIndex!, endIndex: el.endIndex } : undefined;
}

/** Google's tableRange is not guaranteed rectangular once cells are merged. */
export function hasMergedCells(element: DocsTableElement): boolean {
  return (element?.table?.tableRows ?? []).some((row) =>
    (row.tableCells ?? []).some(
      (c) => (c.tableCellStyle?.rowSpan ?? 1) > 1 || (c.tableCellStyle?.columnSpan ?? 1) > 1
    )
  );
}

/** A nested table contributes no paragraph, so it reads as empty rather than recursing. */
export function cellText(cellContent: any[]): string {
  let text = '';
  for (const el of cellContent || []) {
    for (const run of el.paragraph?.elements || []) {
      text += (run.textRun?.content || '').slice(0, MAX_STRUCTURE_CELL_CHARS + 1);
      if (text.length > MAX_STRUCTURE_CELL_CHARS) {
        return `${text.slice(0, MAX_STRUCTURE_CELL_CHARS)}...`;
      }
    }
  }
  return text.replace(/\n$/, '');
}

export interface CollectedCells {
  cells: string[][];
  size: number;
  count: number;
  complete: boolean;
}

/**
 * Read a table's cells, stopping the moment either budget would be crossed so a
 * document cannot drive unbounded work. `complete` is false when it stopped early.
 */
export function collectCells(
  element: DocsTableElement,
  textBudget: number,
  cellBudget: number
): CollectedCells {
  const cells: string[][] = [];
  let size = 0;
  let count = 0;

  for (const row of element.table?.tableRows ?? []) {
    const built: string[] = [];
    for (const c of row.tableCells ?? []) {
      const text = cellText(c.content ?? []);
      size += text.length;
      count += 1;
      if (size > textBudget || count > cellBudget) {
        return { cells, size, count, complete: false };
      }
      built.push(text);
    }
    cells.push(built);
  }

  return { cells, size, count, complete: true };
}

export interface TableStructure {
  startIndex: number;
  endIndex: number;
  rows?: number;
  columns?: number;
  cells?: string[][];
  has_merged_cells?: boolean;
  styles?: ReportedCellStyle[];
  cell_styles?: number[][];
}

/**
 * Both budgets span the whole document rather than each table, since per-table
 * caps would let many small tables rebuild the same unbounded walk. Once spent
 * they stay spent, so no later table is walked at all.
 */
export interface StructureBudget {
  text: number;
  cells: number;
  spent: boolean;
  tablesWithoutCells: number;
  tablesWithoutStyles: number;
}

export const newStructureBudget = (): StructureBudget => ({
  text: MAX_STRUCTURE_TABLE_CHARS,
  cells: MAX_STRUCTURE_CELLS,
  spent: false,
  tablesWithoutCells: 0,
  tablesWithoutStyles: 0,
});

/** Describes one table and charges what it reports to the document budget. */
export function describeTable(
  element: DocsTableElement,
  budget: StructureBudget,
  includeStyles: boolean
): TableStructure {
  const entry: TableStructure = {
    // Google omits startIndex on the first element of the body.
    startIndex: element.startIndex ?? 0,
    endIndex: element.endIndex,
    rows: element.table?.rows,
    columns: element.table?.columns,
  };

  const collected = budget.spent ? undefined : collectCells(element, budget.text, budget.cells);
  if (!collected?.complete) {
    budget.spent = true;
    budget.tablesWithoutCells += 1;
    if (includeStyles) budget.tablesWithoutStyles += 1;
    return entry;
  }

  budget.text -= collected.size;
  budget.cells -= collected.count;
  entry.cells = collected.cells;
  if (hasMergedCells(element)) {
    entry.has_merged_cells = true;
  }

  if (includeStyles) {
    // Styles come out of the same budget so one flag cannot quietly multiply
    // the size of the response.
    const styled = collectCellStyles(element);
    if (styled.size > budget.text) {
      budget.tablesWithoutStyles += 1;
      return entry;
    }
    budget.text -= styled.size;
    entry.styles = styled.styles;
    entry.cell_styles = styled.cellStyles;
  }

  return entry;
}

/** One source for which reported side maps to which Google field, used both directions. */
const BORDER_SIDES = [
  ['border_top', 'borderTop'],
  ['border_right', 'borderRight'],
  ['border_bottom', 'borderBottom'],
  ['border_left', 'borderLeft'],
] as const;

interface GoogleBorder {
  color?: { color?: { rgbColor?: { red?: number; green?: number; blue?: number } } };
  width?: { magnitude?: number; unit?: string };
  dashStyle?: string;
}

export interface GoogleCellStyle {
  backgroundColor?: { color?: { rgbColor?: { red?: number; green?: number; blue?: number } } };
  borderTop?: GoogleBorder;
  borderRight?: GoogleBorder;
  borderBottom?: GoogleBorder;
  borderLeft?: GoogleBorder;
  contentAlignment?: string;
  rowSpan?: number;
  columnSpan?: number;
}

export interface ReportedBorder {
  color?: string;
  width?: number;
  dash_style?: string;
}

export interface ReportedCellStyle {
  background_color?: string;
  border_top?: ReportedBorder;
  border_right?: ReportedBorder;
  border_bottom?: ReportedBorder;
  border_left?: ReportedBorder;
  content_alignment?: string;
}

/** Google omits a channel that is zero, so pure red arrives as { red: 1 }. */
export function rgbToHex(rgbColor: { red?: number; green?: number; blue?: number } | undefined): string {
  const channel = (v: number | undefined) =>
    Math.round(Math.min(1, Math.max(0, v ?? 0)) * 255).toString(16).padStart(2, '0');
  return `#${channel(rgbColor?.red)}${channel(rgbColor?.green)}${channel(rgbColor?.blue)}`.toUpperCase();
}

function describeBorder(border: GoogleBorder | undefined): ReportedBorder | undefined {
  if (!border) return undefined;
  const out: ReportedBorder = {};
  const rgb = border.color?.color?.rgbColor;
  if (rgb) out.color = rgbToHex(rgb);
  if (border.width?.magnitude !== undefined) out.width = border.width.magnitude;
  if (border.dashStyle) out.dash_style = border.dashStyle;
  return Object.keys(out).length ? out : undefined;
}

/** Reports only what update_table_style can write back, so a read value can be replayed. */
export function describeCellStyle(tableCellStyle: GoogleCellStyle | undefined): ReportedCellStyle {
  const out: ReportedCellStyle = {};
  const background = tableCellStyle?.backgroundColor?.color?.rgbColor;
  if (background) out.background_color = rgbToHex(background);

  for (const [side, googleSide] of BORDER_SIDES) {
    const border = describeBorder(tableCellStyle?.[googleSide]);
    if (border) out[side] = border;
  }

  if (tableCellStyle?.contentAlignment) out.content_alignment = tableCellStyle.contentAlignment;
  return out;
}

export interface CollectedStyles {
  styles: ReportedCellStyle[];
  cellStyles: number[][];
  size: number;
}

/**
 * Each distinct style is listed once and cells index into it. Real tables repeat
 * a handful of looks, and the index grid is what shows which ranges share one.
 */
export function collectCellStyles(element: DocsTableElement): CollectedStyles {
  const styles: ReportedCellStyle[] = [];
  const seen = new Map<string, number>();
  const cellStyles: number[][] = [];

  for (const row of element.table?.tableRows ?? []) {
    const built: number[] = [];
    for (const c of row.tableCells ?? []) {
      const style = describeCellStyle(c.tableCellStyle);
      const key = JSON.stringify(style);
      let index = seen.get(key);
      if (index === undefined) {
        index = styles.length;
        seen.set(key, index);
        styles.push(style);
      }
      built.push(index);
    }
    cellStyles.push(built);
  }

  // The index grid is part of the payload, so it is charged too.
  return {
    styles,
    cellStyles,
    size: JSON.stringify(styles).length + JSON.stringify(cellStyles).length,
  };
}

// buildTableInsertRequests lives in docs-structure.ts, which already owned the
// index arithmetic before tables grew a module of their own.

export function normalizeTableData(rows: string[][], columns?: number): string[][] {
  if (!rows || rows.length === 0) {
    throw new Error('Table must have at least one row');
  }

  if (rows.length > MAX_TABLE_ROWS) {
    throw new Error(`Table has ${rows.length} rows, above the ${MAX_TABLE_ROWS} row limit`);
  }

  if (columns !== undefined) {
    if (!Number.isInteger(columns) || columns < 1) {
      throw new Error('columns must be a whole number of at least 1');
    }
    if (columns > MAX_TABLE_COLUMNS) {
      throw new Error(`columns is ${columns}, above the ${MAX_TABLE_COLUMNS} column limit`);
    }
  }

  const widest = rows.reduce((max, row) => Math.max(max, (row || []).length), 0);
  if (widest > MAX_TABLE_COLUMNS) {
    throw new Error(`The widest row has ${widest} cells, above the ${MAX_TABLE_COLUMNS} column limit`);
  }
  if (columns !== undefined && columns < widest) {
    throw new Error(
      `columns is ${columns} but the widest row has ${widest} cells. ` +
      `Raise columns or shorten that row — cell data is never dropped.`
    );
  }

  const width = columns ?? widest;
  if (width === 0) {
    throw new Error('Table must have at least one column');
  }
  if (rows.length * width > MAX_TABLE_CELLS) {
    throw new Error(
      `Table has ${rows.length * width} cells (${rows.length} rows x ${width} columns), ` +
      `above the ${MAX_TABLE_CELLS} cell limit`
    );
  }

  const data = rows.map((row) =>
    Array.from({ length: width }, (_, i) => (i < (row || []).length ? String(row[i]) : '')));

  let totalChars = 0;
  data.forEach((row, r) => row.forEach((cell, c) => {
    if (cell.length > MAX_CELL_CHARS) {
      throw new Error(
        `Cell at row ${r + 1}, column ${c + 1} has ${cell.length} characters, ` +
        `above the ${MAX_CELL_CHARS} character limit`
      );
    }
    totalChars += cell.length;
    if (totalChars > MAX_TABLE_TEXT_CHARS) {
      throw new Error(
        `Table content exceeds the ${MAX_TABLE_TEXT_CHARS} character limit across all cells. ` +
        `Split it into several smaller tables.`
      );
    }
  }));

  return data;
}

export type BorderInput = { color?: string; width?: number; dash_style?: 'SOLID' | 'DOT' | 'DASH' };

export interface CellStyleArgs {
  background_color?: string;
  border?: BorderInput;
  border_top?: BorderInput;
  border_right?: BorderInput;
  border_bottom?: BorderInput;
  border_left?: BorderInput;
  content_alignment?: 'TOP' | 'MIDDLE' | 'BOTTOM';
}

export const borderSchema = z.object({
  color: z.string().optional().describe('6-digit hex, e.g. #1A73E8. Defaults to #000000.'),
  width: z.number().min(0).max(MAX_BORDER_WIDTH_PT).optional().describe('Width in points. Defaults to 1.'),
  dash_style: z.enum(['SOLID', 'DOT', 'DASH']).optional().describe('Defaults to SOLID.'),
});

export function hexToRgbColor(hex: string, label: string) {
  const match = /^#?([0-9a-fA-F]{6})$/.exec(String(hex));
  if (!match) {
    throw new Error(
      `${label} must be a 6-digit hex colour such as #1A73E8, received '${String(hex).slice(0, 20)}'.`
    );
  }
  const value = parseInt(match[1], 16);
  return {
    rgbColor: {
      red: ((value >> 16) & 255) / 255,
      green: ((value >> 8) & 255) / 255,
      blue: (value & 255) / 255,
    },
  };
}

/** Google rejects a partial border, so an unset property falls back to a default here. */
function buildBorder(side: BorderInput, label: string) {
  if (typeof side !== 'object' || Array.isArray(side)) {
    throw new Error(
      `${label} must be an object with optional color, width and dash_style, received '${String(side).slice(0, 20)}'.`
    );
  }
  const width = side.width ?? 1;
  if (!Number.isFinite(width) || width < 0 || width > MAX_BORDER_WIDTH_PT) {
    throw new Error(
      `${label}.width must be a number between 0 and ${MAX_BORDER_WIDTH_PT} points, received '${String(width).slice(0, 20)}'.`
    );
  }
  return {
    color: { color: hexToRgbColor(side.color ?? '#000000', `${label}.color`) },
    width: { magnitude: width, unit: 'PT' },
    dashStyle: side.dash_style ?? 'SOLID',
  };
}

export function buildTableCellStyle(args: CellStyleArgs): { tableCellStyle: any; fields: string[] } {
  const tableCellStyle: any = {};
  const fields: string[] = [];

  if (args.background_color !== undefined) {
    tableCellStyle.backgroundColor = { color: hexToRgbColor(args.background_color, 'background_color') };
    fields.push('backgroundColor');
  }
  if (args.content_alignment !== undefined) {
    tableCellStyle.contentAlignment = args.content_alignment;
    fields.push('contentAlignment');
  }

  for (const [side, googleSide] of BORDER_SIDES) {
    const override = args[side];
    const border = override ?? args.border;
    if (!border) continue;
    tableCellStyle[googleSide] = buildBorder(border, override ? side : 'border');
    fields.push(googleSide);
  }

  return { tableCellStyle, fields };
}

/**
 * Google nests a table rather than refusing when the index lands inside one, so the
 * caller only finds out by reading the document back.
 */
export function nestingRefusal(doc: DocsBody, index: number): string | undefined {
  const content = doc?.body?.content ?? [];
  const position = content.findIndex(
    (c) => c.table && (c.startIndex ?? 0) <= index && index < c.endIndex
  );
  if (position === -1) {
    return undefined;
  }

  const table = content[position];
  const before = content[position - 1]?.startIndex;
  return `index ${index} is inside the table at ${table.startIndex}-${table.endIndex}. ` +
    `Use ${table.endIndex} to place the new table after it` +
    (before !== undefined ? `, or ${before} to place it before it` : '') +
    `. Pass allow_nested=true to nest it inside that cell instead.`;
}

export function styleFailureWarning(located: boolean, status?: number): string {
  const cause = status === 401
    ? 'the credential was rejected before its inherited formatting could be cleared'
    : status === 403
      ? 'access was denied before its inherited formatting could be cleared — the document may be ' +
        'read-only for this account, or a quota was hit'
      : located
        ? 'clearing its inherited formatting failed'
        : 'it could not be located afterwards, so its inherited formatting was not cleared and ' +
          'table_start_index is unverified';

  const unverified = !located && status ? ', and table_start_index is unverified' : '';

  const recovery = located
    ? 'call get_document with include_structure=true to recover the table range, then apply ' +
      'update_text_style over it'
    : 'call get_document with include_structure=true to find the table';

  const first = status === 401 ? ' Re-authenticate first.' : status === 403 ? ' Check access first.' : '';

  return `The table was created, but ${cause}${unverified}. Do not insert it again — ${recovery}.${first}`;
}
