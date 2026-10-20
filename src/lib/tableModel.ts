import { buildStyle, docLocation, hexToRgbColor, pt, type StyleField } from './docRequests.js';
import { BORDER_SIDES, type CellRange, type DocsTableElement, type TableLayout, type TableShape } from './tableRead.js';

// Bound what we build for Google; an oversized payload would otherwise exhaust the isolate.
export const MAX_TABLE_ROWS = 1000;
export const MAX_TABLE_COLUMNS = 100;
export const MAX_TABLE_CELLS = 10_000;
export const MAX_CELL_CHARS = 10_000;
export const MAX_TABLE_TEXT_CHARS = 500_000;

export const MAX_BORDER_WIDTH_PT = 20;

/** insertTable writes a newline before the table, so it starts one past the index. */
export const tableStartFor = (index: number) => index + 1;

/**
 * Takes a table's own start index, not the insertion index findTableAt takes.
 * insert_table's index is one below the table it creates; these must not be conflated.
 */
function findTableStartingAt(content: DocsTableElement[], tableStartIndex: number): DocsTableElement | undefined {
  return content.find((c) => c.table && c.startIndex === tableStartIndex);
}

// Exact start and shape only: a looser match could style a table a collaborator put there
export function findTableAt(content: DocsTableElement[], index: number, shape: TableShape): { startIndex: number; endIndex: number } | undefined {
  const el = findTableStartingAt(content, tableStartFor(index));
  if (!el || el.table?.rows !== shape.rows || el.table?.columns !== shape.columns) return undefined;
  return { startIndex: el.startIndex!, endIndex: el.endIndex };
}

// Reverse order so no insert shifts a later one; cell (r, c) sits at index + 4 + r * (2 * cols + 1) + 2 * c
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
            location: docLocation(insertIndex, tabId),
          },
        });
      }
    }
  }

  return requests;
}

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
    width: pt(width),
    dashStyle: side.dash_style ?? 'SOLID',
  };
}

const CELL_STYLE_FIELDS: Array<StyleField<CellStyleArgs>> = [
  { arg: 'background_color', field: 'backgroundColor', encode: (hex) => ({ color: hexToRgbColor(hex, 'background_color') }) },
  { arg: 'content_alignment', field: 'contentAlignment' },
];

export function buildTableCellStyle(args: CellStyleArgs): { tableCellStyle: Record<string, unknown>; fields: string[] } {
  const { style: tableCellStyle, fields } = buildStyle(CELL_STYLE_FIELDS, args);
  for (const [side, googleSide] of BORDER_SIDES) {
    const override = args[side];
    const border = override ?? args.border;
    if (!border) continue;
    tableCellStyle[googleSide] = buildBorder(border, override ? side : 'border');
    fields.push(googleSide);
  }
  if (fields.length === 0) {
    const names = ['background_color', 'border', ...BORDER_SIDES.map(([side]) => side), 'content_alignment'];
    throw new Error(`At least one style property must be provided (${names.join(', ')}).`);
  }
  return { tableCellStyle, fields };
}

function enclosingTable(content: DocsTableElement[], index: number): DocsTableElement | undefined {
  return content.find((c) => c.table && (c.startIndex ?? 0) <= index && index < c.endIndex);
}

/**
 * Google nests a table rather than refusing when the index lands inside one, so the
 * caller only finds out by reading the document back.
 */
export function nestingRefusal(content: DocsTableElement[], index: number): string | undefined {
  const table = enclosingTable(content, index);
  if (!table) return undefined;

  const before = content[content.indexOf(table) - 1]?.startIndex;
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

export function locateTable(content: DocsTableElement[], tableStartIndex: number, tabId?: string): DocsTableElement {
  const element = findTableStartingAt(content, tableStartIndex);
  if (element) return element;

  const enclosing = enclosingTable(content, tableStartIndex);
  if (enclosing) {
    throw new Error(
      `Index ${tableStartIndex} is inside the table at ${enclosing.startIndex}-${enclosing.endIndex}, ` +
      `so it refers to a table nested in one of its cells. Nested tables cannot be changed. ` +
      `Use the outer table at ${enclosing.startIndex} instead.`
    );
  }
  const starts = content.filter((c) => c.table).map((c) => c.startIndex);
  const listed = starts.slice(0, 20).join(', ');
  const rest = starts.length > 20 ? ` and ${starts.length - 20} more` : '';
  throw new Error(
    `No table starts at index ${tableStartIndex}. ` +
    (starts.length ? `Tables here start at ${listed}${rest}.` : 'No tables were found.') +
    (tabId ? ` Only the tab '${tabId}' was read; pass a different tab_id for tables in other tabs.` : '')
  );
}

export function validateCellRange(
  shape: TableShape,
  tableStartIndex: number,
  args: { row_index?: number; column_index?: number; row_span?: number; column_span?: number },
): CellRange {
  const { rows, columns } = shape;
  const rowIndex = args.row_index ?? 0;
  const columnIndex = args.column_index ?? 0;
  const size = `The table at index ${tableStartIndex} has ${rows} rows and ${columns} columns`;

  if (rowIndex >= rows || columnIndex >= columns) {
    throw new Error(
      `${size}, so row_index ${rowIndex} and column_index ${columnIndex} are out of range. ` +
      `Both are zero-based, so the last cell is row ${rows - 1}, column ${columns - 1}.`
    );
  }

  const rowSpan = args.row_span ?? rows - rowIndex;
  const columnSpan = args.column_span ?? columns - columnIndex;

  if (rowIndex + rowSpan > rows || columnIndex + columnSpan > columns) {
    throw new Error(
      `The requested range (rows ${rowIndex}-${rowIndex + rowSpan - 1}, columns ${columnIndex}-${columnIndex + columnSpan - 1}) ` +
      `does not fit. ${size}.`
    );
  }

  return { rowIndex, columnIndex, rowSpan, columnSpan };
}

export const TABLE_ACTIONS = [
  'insert_row', 'insert_column', 'delete_row', 'delete_column',
  'merge_cells', 'unmerge_cells', 'set_column_width', 'pin_header_rows',
] as const;
export type TableAction = typeof TABLE_ACTIONS[number];

export const MIN_COLUMN_WIDTH_PT = 5;
export const MAX_COLUMN_WIDTH_PT = 2000;

export interface TableModificationArgs {
  action: TableAction;
  row_index?: number;
  column_index?: number;
  row_span?: number;
  column_span?: number;
  position?: 'before' | 'after';
  width_pt?: number;
  pinned_rows?: number;
  suggest?: boolean;
}

export interface TableModification {
  request: any;
  shape: TableShape;
  message: string;
}

function mergedExtent(merges: CellRange[], cell: CellRange, axis: 'row' | 'column'): { first: number; count: number } {
  const merge = merges.find((m) =>
    cell.rowIndex >= m.rowIndex && cell.rowIndex < m.rowIndex + m.rowSpan &&
    cell.columnIndex >= m.columnIndex && cell.columnIndex < m.columnIndex + m.columnSpan);
  if (axis === 'row') return merge ? { first: merge.rowIndex, count: merge.rowSpan } : { first: cell.rowIndex, count: 1 };
  return merge ? { first: merge.columnIndex, count: merge.columnSpan } : { first: cell.columnIndex, count: 1 };
}

export function buildTableModification(
  shape: TableLayout,
  tableStartLocation: { index: number; tabId?: string },
  args: TableModificationArgs,
): TableModification {
  const { action } = args;
  const tableStartIndex = tableStartLocation.index;
  const cellLocation = (range: CellRange) => ({
    tableStartLocation,
    rowIndex: range.rowIndex,
    columnIndex: range.columnIndex,
  });
  const after = args.position !== 'before';
  const extent = (cell: CellRange, axis: 'row' | 'column') => {
    const { first, count } = mergedExtent(shape.merges, cell, axis);
    const lines = count === 1 ? `${axis} ${first}` : `${axis}s ${first}-${first + count - 1}`;
    return { first, count, lines, merged: count === 1 ? '' : `, which the merged cell at row ${cell.rowIndex}, column ${cell.columnIndex} spans` };
  };

  switch (action) {
    case 'insert_row':
    case 'insert_column': {
      const axis = action === 'insert_row' ? 'row' : 'column';
      const cell = validateCellRange(shape, tableStartIndex, { ...args, row_span: 1, column_span: 1 });
      const { first, count, lines, merged } = extent(cell, axis);
      const side = axis === 'row' ? (after ? 'below' : 'above') : (after ? 'right of' : 'left of');
      return {
        request: axis === 'row'
          ? { insertTableRow: { tableCellLocation: cellLocation(cell), insertBelow: after } }
          : { insertTableColumn: { tableCellLocation: cellLocation(cell), insertRight: after } },
        shape: axis === 'row'
          ? { rows: shape.rows + 1, columns: shape.columns }
          : { rows: shape.rows, columns: shape.columns + 1 },
        message: `Inserted a ${axis} ${side} ${lines}${merged}; the new ${axis} is ${axis} ${after ? first + count : first}`,
      };
    }
    case 'delete_row':
    case 'delete_column': {
      const axis = action === 'delete_row' ? 'row' : 'column';
      const total = axis === 'row' ? shape.rows : shape.columns;
      if (total <= 1) {
        throw new Error(`This table has a single ${axis}; deleting the last ${axis} deletes the table. Use delete_content over the table's range for that.`);
      }
      const cell = validateCellRange(shape, tableStartIndex, { ...args, row_span: 1, column_span: 1 });
      const { count, lines, merged } = extent(cell, axis);
      if (count >= total) {
        throw new Error(`The cell at row ${cell.rowIndex}, column ${cell.columnIndex} is merged across ${lines}, which Google deletes together, and that is the whole table. Unmerge it first, or use delete_content over the table's range to delete the table.`);
      }
      return {
        request: axis === 'row'
          ? { deleteTableRow: { tableCellLocation: cellLocation(cell) } }
          : { deleteTableColumn: { tableCellLocation: cellLocation(cell) } },
        shape: axis === 'row'
          ? { rows: shape.rows - count, columns: shape.columns }
          : { rows: shape.rows, columns: shape.columns - count },
        message: `Deleted ${lines}${merged}`,
      };
    }
    case 'merge_cells':
    case 'unmerge_cells': {
      const spans = action === 'merge_cells' ? { row_span: args.row_span ?? 1, column_span: args.column_span ?? 1 } : {};
      const range = validateCellRange(shape, tableStartIndex, { ...args, ...spans });
      if (action === 'merge_cells' && range.rowSpan * range.columnSpan < 2) {
        throw new Error('merge_cells needs a range of more than one cell: set row_span and/or column_span above 1.');
      }
      const tableRange = { tableCellLocation: cellLocation(range), rowSpan: range.rowSpan, columnSpan: range.columnSpan };
      const cells = `${range.rowSpan}x${range.columnSpan} cells from row ${range.rowIndex}, column ${range.columnIndex}`;
      return {
        request: action === 'merge_cells' ? { mergeTableCells: { tableRange } } : { unmergeTableCells: { tableRange } },
        shape,
        message: action === 'merge_cells' ? `Merged ${cells}` : `Unmerged ${cells}`,
      };
    }
    case 'set_column_width': {
      if (args.suggest) {
        throw new Error('Google does not take column widths as a suggestion: call set_column_width without suggest.');
      }
      const width = args.width_pt;
      if (width === undefined || !Number.isFinite(width) || width < MIN_COLUMN_WIDTH_PT || width > MAX_COLUMN_WIDTH_PT) {
        throw new Error(`set_column_width needs width_pt between ${MIN_COLUMN_WIDTH_PT} and ${MAX_COLUMN_WIDTH_PT} points.`);
      }
      let columnIndices: number[] | undefined;
      if (args.column_index !== undefined) {
        validateCellRange(shape, tableStartIndex, { column_index: args.column_index, row_span: 1, column_span: 1 });
        columnIndices = [args.column_index];
      }
      return {
        request: {
          updateTableColumnProperties: {
            tableStartLocation,
            ...(columnIndices ? { columnIndices } : {}),
            tableColumnProperties: { widthType: 'FIXED_WIDTH', width: pt(width) },
            fields: 'widthType,width',
          },
        },
        shape,
        message: columnIndices ? `Set column ${columnIndices[0]} to ${width}pt wide` : `Set every column to ${width}pt wide`,
      };
    }
    case 'pin_header_rows': {
      const pinned = args.pinned_rows;
      if (pinned === undefined || !Number.isInteger(pinned) || pinned < 0) {
        throw new Error('pin_header_rows needs pinned_rows: how many top rows repeat on each page, 0 to unpin.');
      }
      if (pinned > shape.rows) {
        throw new Error(`pinned_rows is ${pinned}, but this table has only ${shape.rows} row(s).`);
      }
      return {
        request: { pinTableHeaderRows: { tableStartLocation, pinnedHeaderRowsCount: pinned } },
        shape,
        message: pinned === 0 ? 'Unpinned the header rows' : `Pinned the top ${pinned} row(s) as a repeating header`,
      };
    }
  }
}
