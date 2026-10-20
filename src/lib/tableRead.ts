// `content` carries the full text, so structure cells only need enough to locate them
export const MAX_STRUCTURE_CELL_CHARS = 200;
export const MAX_STRUCTURE_TABLE_CHARS = 100_000;
export const MAX_STRUCTURE_CELLS = 10_000;

export interface DocsTableElement {
  table?: {
    rows?: number;
    columns?: number;
    tableRows?: Array<{ tableCells?: Array<{ content?: any[]; tableCellStyle?: any }> }>;
  };
  startIndex?: number;
  endIndex: number;
}

export interface TableShape {
  rows: number;
  columns: number;
}

const tableShape = (element: DocsTableElement): TableShape => ({
  rows: element.table?.rows ?? 0,
  columns: element.table?.columns ?? 0,
});

export interface CellRange {
  rowIndex: number;
  columnIndex: number;
  rowSpan: number;
  columnSpan: number;
}

export interface TableLayout extends TableShape {
  merges: CellRange[];
}

export function tableLayout(element: DocsTableElement): TableLayout {
  const merges: CellRange[] = [];
  (element.table?.tableRows ?? []).forEach((row, rowIndex) => {
    (row.tableCells ?? []).forEach((cell, columnIndex) => {
      const rowSpan = cell.tableCellStyle?.rowSpan ?? 1;
      const columnSpan = cell.tableCellStyle?.columnSpan ?? 1;
      if (rowSpan > 1 || columnSpan > 1) merges.push({ rowIndex, columnIndex, rowSpan, columnSpan });
    });
  });
  return { ...tableShape(element), merges };
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
  if (tableLayout(element).merges.length > 0) {
    entry.has_merged_cells = true;
  }

  if (includeStyles) {
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

export const BORDER_SIDES = [
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

  return {
    styles,
    cellStyles,
    size: JSON.stringify(styles).length + JSON.stringify(cellStyles).length,
  };
}
