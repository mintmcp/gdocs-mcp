import { describe, it, expect } from 'vitest';
import { parseDocumentStructure } from './docs-structure.js';
import {
  MAX_STRUCTURE_CELL_CHARS,
  MAX_STRUCTURE_TABLE_CHARS,
  MAX_STRUCTURE_CELLS,
  MAX_BORDER_WIDTH_PT,
  hexToRgbColor,
  buildTableCellStyle,
  collectCells,
  collectCellStyles,
  describeCellStyle,
  rgbToHex,
  describeTable,
  newStructureBudget,
} from './tableModel.js';

describe('table structure', () => {
  const cell = (text: string) => ({
    content: [{ paragraph: { elements: [{ textRun: { content: `${text}\n` } }] } }],
  });
  const table = (grid: string[][], startIndex = 1) => ({
    startIndex,
    endIndex: startIndex + 100,
    table: {
      rows: grid.length,
      columns: grid[0]?.length ?? 0,
      tableRows: grid.map((row) => ({ tableCells: row.map(cell) })),
    },
  });

  it('reports dimensions and cell text for a table', () => {
    const { elements } = parseDocumentStructure([table([['Week', 'Status'], ['1', 'Done']])]);

    expect(elements[0]).toMatchObject({
      type: 'table',
      rows: 2,
      columns: 2,
      cells: [['Week', 'Status'], ['1', 'Done']],
    });
  });

  it('leaves a paragraph element untouched', () => {
    const { elements } = parseDocumentStructure([
      { startIndex: 0, endIndex: 5, paragraph: { elements: [{ textRun: { content: 'hi\n' } }] } },
    ]);

    expect(elements[0]).toEqual({ type: 'paragraph', startIndex: 0, endIndex: 5, text: 'hi\n' });
  });

  it('reads a cell containing a nested table as empty rather than recursing', () => {
    const nested = { table: { rows: 1, columns: 1, tableRows: [{ tableCells: [cell('deep')] }] } };
    const outer = {
      startIndex: 1,
      endIndex: 50,
      table: { rows: 1, columns: 1, tableRows: [{ tableCells: [{ content: [nested] }] }] },
    };

    const { elements } = parseDocumentStructure([outer]);

    expect(elements[0].cells).toEqual([['']]);
  });

  it('flags a merged table and keeps the covered position as an empty cell', () => {
    // Verified live: Google emits the covered position rather than dropping it,
    // for both column and row merges, so positions stay aligned with columns.
    const merged = {
      startIndex: 1,
      endIndex: 50,
      table: {
        rows: 2,
        columns: 3,
        tableRows: [
          { tableCells: [
            { ...cell('spans two'), tableCellStyle: { columnSpan: 2 } },
            { content: [] },
            cell('c'),
          ] },
          { tableCells: [cell('a'), cell('b'), cell('c')] },
        ],
      },
    };

    const { elements } = parseDocumentStructure([merged]);

    expect(elements[0].has_merged_cells).toBe(true);
    expect(elements[0].columns).toBe(3);
    expect(elements[0].cells![0]).toEqual(['spans two', '', 'c']);
  });

  it('leaves an ordinary table unflagged', () => {
    const { elements } = parseDocumentStructure([table([['a', 'b']])]);
    expect(elements[0].has_merged_cells).toBeUndefined();
  });

  it('truncates a long cell instead of copying the document twice', () => {
    const { elements } = parseDocumentStructure([table([['x'.repeat(5_000)]])]);

    expect(elements[0].cells![0][0].length).toBeLessThanOrEqual(MAX_STRUCTURE_CELL_CHARS + 3);
    expect(elements[0].cells![0][0].endsWith('...')).toBe(true);
  });

  it('omits cells once the document budget is spent but still reports dimensions', () => {
    // Two fifths of the budget each, so the third table crosses it.
    const per = MAX_STRUCTURE_TABLE_CHARS * 0.4;
    const wide = Array.from({ length: per / MAX_STRUCTURE_CELL_CHARS }, () => ['y'.repeat(MAX_STRUCTURE_CELL_CHARS)]);
    const { elements, tablesWithoutCells } = parseDocumentStructure([
      table(wide, 1), table(wide, 500), table(wide, 900),
    ]);

    expect(elements[0].cells).toBeDefined();
    expect(elements[1].cells).toBeDefined();
    expect(elements[2].cells).toBeUndefined();
    expect(elements[2]).toMatchObject({ rows: wide.length, columns: 1 });
    expect(tablesWithoutCells).toBe(1);
  });

  it('keeps omitting cells after the budget is spent, however small the next table is', () => {
    const sized = (chars: number) => Array.from(
      { length: chars / MAX_STRUCTURE_CELL_CHARS },
      () => ['z'.repeat(MAX_STRUCTURE_CELL_CHARS)]
    );

    const { elements, tablesWithoutCells } = parseDocumentStructure([
      table(sized(MAX_STRUCTURE_TABLE_CHARS * 0.9), 1),
      table(sized(MAX_STRUCTURE_TABLE_CHARS * 0.2), 500),
      table([['tiny']], 900),
    ]);

    expect(elements[0].cells).toBeDefined();
    expect(elements[1].cells).toBeUndefined();
    expect(elements[2].cells).toBeUndefined();
    expect(tablesWithoutCells).toBe(2);
  });

  it('stops walking a table with a huge number of empty cells', () => {
    const empties = Array.from({ length: 200 }, () => Array.from({ length: 200 }, () => ''));
    const { elements, tablesWithoutCells } = parseDocumentStructure([table(empties)]);

    expect(elements[0].cells).toBeUndefined();
    expect(elements[0]).toMatchObject({ rows: 200, columns: 200 });
    expect(tablesWithoutCells).toBe(1);
  });

  it('spends the cell budget across the document, not per table', () => {
    // Each table sits under the cap on its own; together they blow past it.
    const small = Array.from({ length: 2_000 }, () => ['']);
    const { elements } = parseDocumentStructure(
      Array.from({ length: 20 }, (_, i) => table(small, i * 100))
    );

    const built = elements.reduce(
      (n, e) => n + (e.cells ?? []).reduce((m, r) => m + r.length, 0),
      0
    );

    expect(built).toBeLessThanOrEqual(MAX_STRUCTURE_CELLS);
    expect(elements.some((e) => e.cells === undefined)).toBe(true);
  });

  it('omits cells for a single table that is larger than the whole budget', () => {
    const huge = Array.from({ length: 60 }, () => Array.from({ length: 60 }, () => 'y'.repeat(100)));
    const { elements, tablesWithoutCells } = parseDocumentStructure([table(huge)]);

    expect(elements[0].cells).toBeUndefined();
    expect(elements[0]).toMatchObject({ rows: 60, columns: 60 });
    expect(tablesWithoutCells).toBe(1);
  });
});

describe('cell style building', () => {
  it('converts hex to the API 0-1 float form', () => {
    expect(hexToRgbColor('#FF8000', 'background_color')).toEqual({
      rgbColor: { red: 1, green: 128 / 255, blue: 0 },
    });
  });

  it('accepts a hex value without the leading hash', () => {
    expect(hexToRgbColor('000000', 'background_color')).toEqual({
      rgbColor: { red: 0, green: 0, blue: 0 },
    });
  });

  it('names the parameter and the value it rejected', () => {
    expect(() => hexToRgbColor('red', 'background_color')).toThrow(/background_color/);
    expect(() => hexToRgbColor('red', 'background_color')).toThrow(/#RRGGBB|#1A73E8/);
  });

  it('never echoes an oversized colour value back in full', () => {
    expect(() => hexToRgbColor('#'.repeat(500), 'border.color')).toThrow(/border\.color/);
    try {
      hexToRgbColor('#'.repeat(500), 'border.color');
    } catch (e: any) {
      expect(e.message.length).toBeLessThan(200);
    }
  });

  it('applies one border to all four sides', () => {
    const { tableCellStyle, fields } = buildTableCellStyle({
      border: { color: '#000000', width: 1, dash_style: 'SOLID' },
    });

    expect([...fields].sort()).toEqual(['borderBottom', 'borderLeft', 'borderRight', 'borderTop']);
    expect(tableCellStyle.borderTop).toEqual({
      color: { color: { rgbColor: { red: 0, green: 0, blue: 0 } } },
      width: { magnitude: 1, unit: 'PT' },
      dashStyle: 'SOLID',
    });
    expect(tableCellStyle.borderLeft).toEqual(tableCellStyle.borderTop);
  });

  it('lets a per-side value replace the shared one outright', () => {
    const { tableCellStyle } = buildTableCellStyle({
      border: { color: '#000000', width: 1 },
      border_bottom: { color: '#FF0000', width: 3 },
    });

    expect(tableCellStyle.borderBottom.width.magnitude).toBe(3);
    expect(tableCellStyle.borderBottom.color.color.rgbColor.red).toBe(1);
    expect(tableCellStyle.borderTop.width.magnitude).toBe(1);
  });

  it('fills a partial border with defaults so the fields mask cannot blank a side', () => {
    const { tableCellStyle } = buildTableCellStyle({ border_top: { width: 2 } });

    expect(tableCellStyle.borderTop).toEqual({
      color: { color: { rgbColor: { red: 0, green: 0, blue: 0 } } },
      width: { magnitude: 2, unit: 'PT' },
      dashStyle: 'SOLID',
    });
  });

  it('sets only what was asked for', () => {
    const { tableCellStyle, fields } = buildTableCellStyle({
      background_color: '#EEEEEE',
      content_alignment: 'MIDDLE',
    });

    expect([...fields].sort()).toEqual(['backgroundColor', 'contentAlignment']);
    expect(tableCellStyle.borderTop).toBeUndefined();
  });

  it('refuses a border wider than the cap', () => {
    expect(() => buildTableCellStyle({ border: { width: 10_000 } }))
      .toThrow(new RegExp(String(MAX_BORDER_WIDTH_PT)));
  });

  it('refuses a negative border width', () => {
    expect(() => buildTableCellStyle({ border: { width: -1 } })).toThrow(/width/);
  });

  it('refuses a border width that is not a number', () => {
    expect(() => buildTableCellStyle({ border: { width: Number.NaN } })).toThrow(/width/);
  });

  it('refuses a border that is not an object rather than painting a default one', () => {
    expect(() => buildTableCellStyle({ border_top: true as any })).toThrow(/border_top must be an object/);
    expect(() => buildTableCellStyle({ border: 'thick' as any })).toThrow(/border must be an object/);
  });

  it('truncates an oversized width value in the error it reports', () => {
    try {
      buildTableCellStyle({ border: { width: 'w'.repeat(500) as any } });
      throw new Error('should have thrown');
    } catch (e: any) {
      expect(e.message).toMatch(/border\.width/);
      expect(e.message.length).toBeLessThan(200);
    }
  });

  it('returns no fields when nothing was given', () => {
    expect(buildTableCellStyle({}).fields).toEqual([]);
  });
});

describe('collectCells', () => {
  const cell = (text: string) => ({
    content: [{ paragraph: { elements: [{ textRun: { content: `${text}\n` } }] } }],
  });
  const table = (grid: string[][]) => ({
    endIndex: 100,
    table: {
      rows: grid.length,
      columns: grid[0]?.length ?? 0,
      tableRows: grid.map((row) => ({ tableCells: row.map(cell) })),
    },
  });

  it('reports what it read when both budgets hold', () => {
    const result = collectCells(table([['ab', 'c'], ['d', '']]), 1_000, 1_000);

    expect(result.complete).toBe(true);
    expect(result.cells).toEqual([['ab', 'c'], ['d', '']]);
    expect(result.size).toBe(4);
    expect(result.count).toBe(4);
  });

  it('stops on the character budget and says it is incomplete', () => {
    const result = collectCells(table([['aaaa', 'bbbb']]), 5, 1_000);

    expect(result.complete).toBe(false);
    expect(result.size).toBeGreaterThan(5);
  });

  it('stops on the cell budget even when the cells are empty', () => {
    const empty = Array.from({ length: 50 }, () => ['', '']);
    const result = collectCells(table(empty), 1_000, 10);

    expect(result.complete).toBe(false);
    expect(result.size).toBe(0);
    expect(result.count).toBe(11);
  });

  it('treats a table with no rows as complete and empty', () => {
    const result = collectCells({ endIndex: 5, table: { rows: 0, columns: 0 } }, 100, 100);

    expect(result).toEqual({ cells: [], size: 0, count: 0, complete: true });
  });
});

describe('reading cell styles', () => {
  const rgb = (red?: number, green?: number, blue?: number) => {
    const c: any = {};
    if (red !== undefined) c.red = red;
    if (green !== undefined) c.green = green;
    if (blue !== undefined) c.blue = blue;
    return { color: { rgbColor: c } };
  };

  it('converts a colour back to the hex it was written as', () => {
    expect(rgbToHex({ red: 1, green: 128 / 255, blue: 0 })).toBe('#FF8000');
  });

  it('treats an omitted channel as zero, which is how Google sends pure red', () => {
    expect(rgbToHex({ red: 1 })).toBe('#FF0000');
    expect(rgbToHex({})).toBe('#000000');
  });

  it('round trips every channel through hexToRgbColor', () => {
    for (const hex of ['#1A73E8', '#EEEEEE', '#000000', '#FFFFFF', '#FCE8B2']) {
      expect(rgbToHex(hexToRgbColor(hex, 'x').rgbColor)).toBe(hex);
    }
  });

  it('reports background, borders and alignment in the shape update_table_style accepts', () => {
    const style = describeCellStyle({
      backgroundColor: rgb(1, 1, 1),
      borderBottom: { color: rgb(1, 0, 0), width: { magnitude: 4, unit: 'PT' }, dashStyle: 'DASH' },
      contentAlignment: 'MIDDLE',
    });

    expect(style).toEqual({
      background_color: '#FFFFFF',
      border_bottom: { color: '#FF0000', width: 4, dash_style: 'DASH' },
      content_alignment: 'MIDDLE',
    });
  });

  it('omits what Google did not report rather than inventing defaults', () => {
    expect(describeCellStyle({ contentAlignment: 'TOP' })).toEqual({ content_alignment: 'TOP' });
    expect(describeCellStyle({})).toEqual({});
    expect(describeCellStyle(undefined)).toEqual({});
  });

  it('omits a background with no colour, which is how Google reports transparent', () => {
    expect(describeCellStyle({ backgroundColor: { color: {} } })).toEqual({});
  });

  it('gives identical cells one legend entry and keeps positions', () => {
    const header = { backgroundColor: rgb(0, 0, 1) };
    const plain = { backgroundColor: rgb(1, 1, 1) };
    const element = {
      endIndex: 50,
      table: {
        rows: 2,
        columns: 3,
        tableRows: [
          { tableCells: [{ tableCellStyle: header }, { tableCellStyle: header }, { tableCellStyle: header }] },
          { tableCells: [{ tableCellStyle: plain }, { tableCellStyle: header }, { tableCellStyle: plain }] },
        ],
      },
    };

    const { styles, cellStyles } = collectCellStyles(element);

    expect(styles).toHaveLength(2);
    expect(cellStyles).toEqual([[0, 0, 0], [1, 0, 1]]);
    expect(styles[0]).toEqual({ background_color: '#0000FF' });
    expect(styles[1]).toEqual({ background_color: '#FFFFFF' });
  });

  it('collapses a uniformly unstyled table to a single empty legend entry', () => {
    const element = {
      endIndex: 20,
      table: { rows: 2, columns: 2, tableRows: [{ tableCells: [{}, {}] }, { tableCells: [{}, {}] }] },
    };

    const { styles, cellStyles } = collectCellStyles(element);

    expect(styles).toEqual([{}]);
    expect(cellStyles).toEqual([[0, 0], [0, 0]]);
  });

  it('reports the character cost so the caller can charge it to a budget', () => {
    const element = {
      endIndex: 20,
      table: { rows: 1, columns: 1, tableRows: [{ tableCells: [{ tableCellStyle: { contentAlignment: 'TOP' } }] }] },
    };

    const { size } = collectCellStyles(element);

    expect(size).toBeGreaterThan(0);
    expect(size).toBe(
      JSON.stringify([{ content_alignment: 'TOP' }]).length + JSON.stringify([[0]]).length
    );
  });
});

describe('style budget accounting', () => {
  it('charges the index grid, not just the legend', () => {
    const cell = { tableCellStyle: {} };
    const wide = {
      endIndex: 100,
      table: {
        rows: 40,
        columns: 40,
        tableRows: Array.from({ length: 40 }, () => ({ tableCells: Array.from({ length: 40 }, () => cell) })),
      },
    };

    const { styles, size } = collectCellStyles(wide);

    // One legend entry for 1,600 identical cells, so anything near the legend
    // size alone would mean the grid was free.
    expect(styles).toEqual([{}]);
    expect(size).toBeGreaterThan(1_600);
  });
});

describe('describeTable and the document budget', () => {
  const cell = (text: string) => ({
    content: [{ paragraph: { elements: [{ textRun: { content: `${text}\n` } }] } }],
  });
  const table = (grid: string[][], startIndex = 1) => ({
    startIndex,
    endIndex: startIndex + 100,
    table: {
      rows: grid.length,
      columns: grid[0]?.length ?? 0,
      tableRows: grid.map((row) => ({ tableCells: row.map(cell) })),
    },
  });

  it('charges what it reports to the budget it was given', () => {
    const budget = newStructureBudget();
    const before = budget.text;

    const entry = describeTable(table([['ab', 'cd']]), budget, false);

    expect(entry.cells).toEqual([['ab', 'cd']]);
    expect(budget.text).toBe(before - 4);
    expect(budget.cells).toBe(MAX_STRUCTURE_CELLS - 2);
    expect(budget.spent).toBe(false);
  });

  it('marks the budget spent and counts the table once it will not fit', () => {
    const budget = newStructureBudget();
    budget.cells = 1;

    const entry = describeTable(table([['a', 'b']]), budget, false);

    expect(entry.cells).toBeUndefined();
    expect(entry.rows).toBe(1);
    expect(budget.spent).toBe(true);
    expect(budget.tablesWithoutCells).toBe(1);
  });

  it('counts a dropped table against styles too when styles were asked for', () => {
    const budget = newStructureBudget();
    budget.cells = 1;

    describeTable(table([['a', 'b']]), budget, true);

    expect(budget.tablesWithoutStyles).toBe(1);
  });

  it('keeps cells but drops styles when only the style budget is short', () => {
    const budget = newStructureBudget();
    budget.text = 6;

    const entry = describeTable(table([['ab', 'cd']]), budget, true);

    expect(entry.cells).toEqual([['ab', 'cd']]);
    expect(entry.styles).toBeUndefined();
    expect(budget.tablesWithoutStyles).toBe(1);
    expect(budget.spent).toBe(false);
  });
});
