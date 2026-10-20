import { afterEach, describe, expect, it, vi } from "vitest";
import { batchBodies, batchBody, callTool, errorText, paragraph, stubDoc, stubRefusal } from "../lib/mcpHarness.js";

const cell = (text: string, start: number) => ({ content: [paragraph(text, start)], tableCellStyle: { rowSpan: 1, columnSpan: 1 } });

const TABLE = {
  startIndex: 8,
  endIndex: 30,
  table: {
    rows: 2,
    columns: 3,
    tableRows: [
      { tableCells: [cell("a\n", 10), cell("b\n", 12), cell("c\n", 14)] },
      { tableCells: [cell("d\n", 18), cell("e\n", 20), cell("f\n", 22)] },
    ],
  },
};

const DOC = {
  revisionId: "rev-1",
  body: { content: [{ startIndex: 0, endIndex: 1, sectionBreak: {} }, paragraph("Title\n", 1), TABLE, paragraph("\n", 30)] },
};

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

const call = (args: Record<string, unknown>) => callTool("modify_table", { table_start_index: 8, ...args });

describe("modify_table rows and columns", () => {
  it("inserts a row below the reference cell, pinned to the revision it read", async () => {
    const calls = stubDoc(DOC);
    const result = await call({ action: "insert_row", row_index: 1 });
    expect(result.isError).toBeFalsy();
    const body = batchBody(calls);
    expect(body.requests).toEqual([{
      insertTableRow: {
        tableCellLocation: { tableStartLocation: { index: 8 }, rowIndex: 1, columnIndex: 0 },
        insertBelow: true,
      },
    }]);
    expect(body.writeControl).toEqual({ requiredRevisionId: "rev-1" });
    expect(result.structuredContent).toMatchObject({ id: "doc1", action: "insert_row", rows: 3, columns: 3 });
  });

  it("inserts above / left when asked and defaults the reference row and column to 0", async () => {
    const above = stubDoc(DOC);
    await call({ action: "insert_row", position: "before" });
    expect(batchBody(above).requests[0].insertTableRow).toMatchObject({ insertBelow: false, tableCellLocation: { rowIndex: 0, columnIndex: 0 } });
    const left = stubDoc(DOC);
    await call({ action: "insert_column", column_index: 2, position: "before", tab_id: undefined });
    expect(batchBody(left).requests[0].insertTableColumn).toEqual({
      tableCellLocation: { tableStartLocation: { index: 8 }, rowIndex: 0, columnIndex: 2 },
      insertRight: false,
    });
  });

  it("deletes a row or a column and reports the new shape", async () => {
    const row = stubDoc(DOC);
    const result = await call({ action: "delete_row", row_index: 0 });
    expect(batchBody(row).requests).toEqual([{ deleteTableRow: { tableCellLocation: { tableStartLocation: { index: 8 }, rowIndex: 0, columnIndex: 0 } } }]);
    expect(result.structuredContent).toMatchObject({ rows: 1, columns: 3 });
    const col = stubDoc(DOC);
    const colResult = await call({ action: "delete_column", column_index: 2 });
    expect(batchBody(col).requests[0]).toHaveProperty("deleteTableColumn");
    expect(colResult.structuredContent).toMatchObject({ rows: 2, columns: 2 });
  });

  it("refuses to delete the last row or column, since that deletes the table", async () => {
    const oneRow = { ...DOC, body: { content: [{ ...TABLE, table: { ...TABLE.table, rows: 1 } }] } };
    stubDoc(oneRow);
    expect(errorText(await call({ action: "delete_row", row_index: 0 }))).toMatch(/last row/);
  });

  it("counts every row a merged cell spans as deleted, and refuses when that is the whole table", async () => {
    const merged = (rowSpan: number, columnSpan: number) => {
      const rows = TABLE.table.tableRows.map((r) => ({ tableCells: r.tableCells.map((c) => ({ ...c })) }));
      rows[0].tableCells[1] = { ...rows[0].tableCells[1], tableCellStyle: { rowSpan, columnSpan } };
      return { ...DOC, body: { content: [DOC.body.content[0], DOC.body.content[1], { ...TABLE, table: { ...TABLE.table, tableRows: rows } }, DOC.body.content[3]] } };
    };

    stubDoc(merged(1, 2));
    const cols = await call({ action: "delete_column", column_index: 2 });
    expect(cols.structuredContent).toMatchObject({ rows: 2, columns: 1 });
    expect((cols.structuredContent as any).message).toMatch(/columns 1-2, which the merged cell/);

    stubDoc(merged(2, 1));
    expect(errorText(await call({ action: "delete_row", row_index: 1, column_index: 1 }))).toMatch(/merged across rows 0-1.*whole table/);

    stubDoc(merged(2, 1));
    expect((await call({ action: "delete_row", row_index: 1, column_index: 0 })).structuredContent).toMatchObject({ rows: 1 });
  });

  it("reports where an insert next to a merged cell lands: beside the whole merge", async () => {
    const rows = TABLE.table.tableRows.map((r) => ({ tableCells: r.tableCells.map((c) => ({ ...c })) }));
    rows[0].tableCells[1] = { ...rows[0].tableCells[1], tableCellStyle: { rowSpan: 2, columnSpan: 1 } };
    const doc = { ...DOC, body: { content: [DOC.body.content[0], DOC.body.content[1], { ...TABLE, table: { ...TABLE.table, tableRows: rows } }, DOC.body.content[3]] } };

    stubDoc(doc);
    const below = await call({ action: "insert_row", row_index: 0, column_index: 1 });
    expect((below.structuredContent as any).message).toMatch(/below rows 0-1, which the merged cell .*; the new row is row 2/);
    stubDoc(doc);
    const above = await call({ action: "insert_row", row_index: 1, column_index: 0, position: "before" });
    expect((above.structuredContent as any).message).toMatch(/above row 1; the new row is row 1/);
  });

  it("rejects a cell outside the table", async () => {
    stubDoc(DOC);
    expect(errorText(await call({ action: "insert_row", row_index: 2 }))).toMatch(/2 rows and 3 columns/);
  });

  it("explains when no table starts at the index", async () => {
    stubDoc(DOC);
    expect(errorText(await call({ action: "insert_row", table_start_index: 40 }))).toMatch(/No table starts at index 40.*start at 8/);
    stubDoc(DOC);
    expect(errorText(await call({ action: "insert_row", table_start_index: 9 }))).toMatch(/nested.*outer table at 8/);
  });
});

describe("modify_table merges, widths and pinned rows", () => {
  it("merges a rectangular range and unmerges it", async () => {
    const merged = stubDoc(DOC);
    const result = await call({ action: "merge_cells", row_index: 0, column_index: 0, row_span: 2, column_span: 2 });
    expect(result.isError).toBeFalsy();
    expect(batchBody(merged).requests).toEqual([{
      mergeTableCells: {
        tableRange: { tableCellLocation: { tableStartLocation: { index: 8 }, rowIndex: 0, columnIndex: 0 }, rowSpan: 2, columnSpan: 2 },
      },
    }]);
    const unmerged = stubDoc(DOC);
    await call({ action: "unmerge_cells", row_index: 0, column_index: 0, row_span: 2, column_span: 2 });
    expect(batchBody(unmerged).requests[0]).toHaveProperty("unmergeTableCells");
  });

  it("needs a span of more than one cell to merge, and the range must fit", async () => {
    stubDoc(DOC);
    expect(errorText(await call({ action: "merge_cells", row_index: 0, column_index: 0 }))).toMatch(/more than one cell/);
    stubDoc(DOC);
    expect(errorText(await call({ action: "merge_cells", row_index: 1, column_index: 2, row_span: 2, column_span: 1 }))).toMatch(/does not fit/);
  });

  it("sets a column width in points, for one column or all", async () => {
    const one = stubDoc(DOC);
    await call({ action: "set_column_width", column_index: 1, width_pt: 120 });
    expect(batchBody(one).requests).toEqual([{
      updateTableColumnProperties: {
        tableStartLocation: { index: 8 },
        columnIndices: [1],
        tableColumnProperties: { widthType: "FIXED_WIDTH", width: { magnitude: 120, unit: "PT" } },
        fields: "widthType,width",
      },
    }]);
    const all = stubDoc(DOC);
    await call({ action: "set_column_width", width_pt: 72 });
    expect(batchBody(all).requests[0].updateTableColumnProperties).not.toHaveProperty("columnIndices");
    stubDoc(DOC);
    expect(errorText(await call({ action: "set_column_width" }))).toMatch(/width_pt/);
    const suggested = stubDoc(DOC);
    expect(errorText(await call({ action: "set_column_width", width_pt: 72, suggest: true }))).toMatch(/without suggest/);
    expect(batchBodies(suggested)).toHaveLength(0);
  });

  it("pins header rows and unpins with 0", async () => {
    const pinned = stubDoc(DOC);
    await call({ action: "pin_header_rows", pinned_rows: 1 });
    expect(batchBody(pinned).requests).toEqual([{ pinTableHeaderRows: { tableStartLocation: { index: 8 }, pinnedHeaderRowsCount: 1 } }]);
    stubDoc(DOC);
    expect(errorText(await call({ action: "pin_header_rows", pinned_rows: 3 }))).toMatch(/only 2 row/);
  });

  it("carries the tab on every location and supports suggest", async () => {
    const tabbed = {
      revisionId: "rev-1",
      tabs: [{ tabProperties: { tabId: "t.0", index: 0 }, documentTab: DOC }],
    };
    const calls = stubDoc(tabbed);
    await call({ action: "insert_column", column_index: 0, tab_id: "t.0", suggest: true });
    const body = batchBody(calls);
    expect(body.requests[0].insertTableColumn.tableCellLocation.tableStartLocation).toEqual({ index: 8, tabId: "t.0" });
    expect(body.writeControl).toEqual({ requiredRevisionId: "rev-1", writeMode: "SUGGEST" });
  });
});

describe("modify_table refusals from Google", () => {
  it("names the pinned header rows when Google calls a merge non-rectangular", async () => {
    stubRefusal(DOC, "Invalid requests[0].mergeTableCells: A non-rectangular collection of cells cannot be merged.");
    expect(errorText(await call({ action: "merge_cells", row_index: 0, row_span: 2 }))).toMatch(/non-rectangular.*pinned header rows/);
  });
});
