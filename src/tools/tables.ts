import { z } from 'zod';
import {
  clearInheritedFormattingRequest,
  docLocation,
  endOfBodyIndex,
  validateInsertIndex,
  type WriteOptions,
} from "../lib/docRequests.js";
import {
  buildTableCellStyle,
  buildTableInsertRequests,
  buildTableModification,
  MAX_BORDER_WIDTH_PT,
  findTableAt,
  locateTable,
  nestingRefusal,
  normalizeTableData,
  styleFailureWarning,
  tableStartFor,
  validateCellRange,
  MAX_CELL_CHARS,
  MAX_COLUMN_WIDTH_PT,
  MAX_TABLE_COLUMNS,
  MAX_TABLE_ROWS,
  MIN_COLUMN_WIDTH_PT,
  TABLE_ACTIONS,
} from "../lib/tableModel.js";
import { tableLayout } from "../lib/tableRead.js";
import { log, errorFields } from "../lib/log.js";
import {
  docsBatchUpdate,
  documentIdField,
  doneMessage,
  editTool,
  fetchResolvedTab,
  requiredRevisionField,
  suggestField,
  tabIdField,
  toolResult,
  withRefusalHint,
} from './shared.js';

const borderSchema = z.object({
  color: z.string().optional().describe('6-digit hex, e.g. #1A73E8. Defaults to #000000.'),
  width: z.number().min(0).max(MAX_BORDER_WIDTH_PT).optional().describe('Width in points. Defaults to 1.'),
  dash_style: z.enum(['SOLID', 'DOT', 'DASH']).optional().describe('Defaults to SOLID.'),
});

interface InsertedTable {
  rows: number;
  columns: number;
  tableStartIndex: number;
  startIndexVerified: boolean;
  warning?: string;
}

interface TablePlacement {
  tableData: string[][];
  index: number;
  tabId?: string;
  clearInheritedFormatting: boolean;
  control: WriteOptions;
}

// Nothing past the insert may throw: a retry would create a second table, so failures become a warning
async function insertTableAt(
  documentId: string,
  accessToken: string,
  { tableData, index, tabId, clearInheritedFormatting, control }: TablePlacement,
): Promise<InsertedTable> {
  const columns = tableData[0].length;

  await docsBatchUpdate(documentId, accessToken, [
    { insertTable: { location: docLocation(index, tabId), rows: tableData.length, columns } },
    ...buildTableInsertRequests(tableData, index, tabId),
  ], control);

  const created = {
    rows: tableData.length,
    columns,
    tableStartIndex: tableStartFor(index),
    startIndexVerified: false,
  };

  if (!clearInheritedFormatting) {
    return created;
  }

  let table: { startIndex: number; endIndex: number } | undefined;
  try {
    const { resolved } = await fetchResolvedTab(documentId, accessToken, tabId);
    table = findTableAt(resolved.body, index, { rows: tableData.length, columns });

    if (table) {
      await docsBatchUpdate(documentId, accessToken, [
        clearInheritedFormattingRequest(table.startIndex + 1, table.endIndex, tabId),
      ], { read: resolved.revisionId, suggest: control.suggest });
    }
  } catch (error: any) {
    // Google's error text is untrusted once the model reads it, so it stays in the log.
    log("error", "table_style_failed", { phase: table ? "style" : "refetch", ...errorFields(error) });
    return {
      ...created,
      startIndexVerified: Boolean(table),
      warning: styleFailureWarning(Boolean(table), error?.status),
    };
  }

  if (!table) {
    log("error", "table_style_failed", { phase: "lookup", tableStartIndex: tableStartFor(index) });
    return { ...created, warning: styleFailureWarning(false) };
  }

  return { ...created, startIndexVerified: true };
}

// Google calls a merge across pinned header rows "non-rectangular" even when it is a rectangle
const explainMergeRefusal = withRefusalHint(
  (message) => /non-rectangular/i.test(message),
  'Google also refuses a merge that crosses the pinned header rows, and a table imported from markdown starts with its first row pinned: unpin with pin_header_rows (pinned_rows 0) first, or keep the merge on one side of them.',
);

const insertedTableResult = (id: string, table: InsertedTable, suggest?: boolean) => toolResult({
  id,
  rows: table.rows,
  columns: table.columns,
  table_start_index: table.tableStartIndex,
  table_start_index_verified: table.startIndexVerified,
  message: doneMessage(
    `Table inserted with ${table.rows} rows and ${table.columns} columns` +
      (table.startIndexVerified ? `, starting at index ${table.tableStartIndex}` : ''),
    suggest,
  ),
  ...(table.warning ? { warning: table.warning } : {}),
});

const tableToolOutputSchema = {
  id: z.string(),
  rows: z.number(),
  columns: z.number(),
  table_start_index: z.number(),
  table_start_index_verified: z.boolean().describe('True when the table was found at this index after insertion. False means the index is inferred from the API contract and was not confirmed — re-read the document before using it.'),
  message: z.string(),
  warning: z.string().optional().describe('Present only when the table was created but a follow-up step did not complete. Never re-insert the table in response to a warning.'),
};

export const tableTools = {
  append_table: editTool({
    description: 'Insert a table with data at the end of a Google Doc. Supports ragged rows (will be padded with empty cells). Optionally pass `required_revision_id` (from get_document with include_structure=true) to fail the initial insert if the document was edited concurrently; without it the table goes at the end of the revision just read, even if someone edits in between, and the follow-up formatting clear does the same with the revision read back after the insert. If this call fails, check the document with get_document before retrying — a table that was already created would otherwise be inserted twice. On success the result reports table_start_index plus table_start_index_verified; if a warning comes back, the table already exists — never insert it again. Native Google Docs only: Word (.doc/.docx) uploads are read-only and this tool refuses them.',
    outputSchema: tableToolOutputSchema,
    schema: {
      document_id: documentIdField,
      rows: z.array(z.array(z.string().max(MAX_CELL_CHARS)).max(MAX_TABLE_COLUMNS)).max(MAX_TABLE_ROWS).describe('Table data as array of rows, each row is array of cell strings'),
      tab_id: tabIdField,
      clear_inherited_formatting: z.boolean().optional().describe('Clear inherited formatting on the newly appended table content. Defaults to true.'),
      required_revision_id: z.string().optional().describe('If provided, the initial table insert fails with a 400 if the document revision has changed since this revisionId. Use the `revisionId` returned by get_document(include_structure=true) for optimistic concurrency control.'),
      suggest: suggestField,
    },
    run: async ({ document_id, rows, tab_id, clear_inherited_formatting, required_revision_id, suggest }, { accessToken }) => {
      const { resolved } = await fetchResolvedTab(document_id, accessToken, tab_id);
      const table = await insertTableAt(document_id, accessToken, {
        tableData: normalizeTableData(rows),
        index: endOfBodyIndex(resolved.body),
        tabId: resolved.tabId,
        clearInheritedFormatting: clear_inherited_formatting !== false,
        control: { required: required_revision_id, read: resolved.revisionId, suggest },
      });
      return insertedTableResult(document_id, table, suggest);
    },
  }),

  insert_table: editTool({
    description: 'Insert a table with data at a specific index in a Google Doc — use this for any position other than the end of the document; index 1 is the top of the document body. To place a table between existing content, call get_document with include_structure=true and pass the startIndex of the paragraph the table should appear above. Google accepts most indices, so a wrong one usually misplaces the table instead of failing: an index inside a paragraph splits that paragraph, and an index inside a table cell nests the new table there (nested tables cannot be found afterwards, so their inherited formatting is left alone and table_start_index comes back unverified). If this call fails, check the document with get_document before retrying — a table that was already created would otherwise be inserted twice. On success the result reports table_start_index plus table_start_index_verified; if a warning comes back, the table already exists — never insert it again. Supports ragged rows (will be padded with empty cells). Native Google Docs only: Word (.doc/.docx) uploads are read-only and this tool refuses them.',
    outputSchema: tableToolOutputSchema,
    schema: {
      document_id: documentIdField,
      rows: z.array(z.array(z.string().max(MAX_CELL_CHARS)).max(MAX_TABLE_COLUMNS)).max(MAX_TABLE_ROWS).describe('Table data as array of rows, each row is array of cell strings. Pass empty rows (e.g. [[], [], []]) with columns set to create an empty table.'),
      index: z.number().int().describe('Document index to insert the table at. 1 is the top of the document body. Use get_document with include_structure=true to find indices.'),
      tab_id: tabIdField,
      columns: z.number().int().min(1).max(MAX_TABLE_COLUMNS).optional().describe('Force the table width; short rows are padded with empty cells. Defaults to the widest row. Must not be smaller than the widest row.'),
      clear_inherited_formatting: z.boolean().optional().describe('Clear formatting inherited from the insertion point on the new table content. Defaults to true.'),
      allow_nested: z.boolean().optional().describe('Allow an index that falls inside an existing table, which nests the new table in that cell. Defaults to false, which refuses such an index and tells you which ones would work.'),
      required_revision_id: requiredRevisionField,
      suggest: suggestField,
    },
    run: async ({ document_id, rows, index, tab_id, columns, clear_inherited_formatting, allow_nested, required_revision_id, suggest }, { accessToken }) => {
      validateInsertIndex(index);

      const tableData = normalizeTableData(rows, columns);

      // Target the revision the nesting check read, so an edit in between cannot move the index
      let readRevisionId: string | undefined;
      if (allow_nested !== true) {
        const { resolved } = await fetchResolvedTab(document_id, accessToken, tab_id);
        const refusal = nestingRefusal(resolved.body, index);
        if (refusal) throw new Error(refusal);
        readRevisionId = resolved.revisionId;
      }

      const table = await insertTableAt(document_id, accessToken, {
        tableData,
        index,
        tabId: tab_id,
        clearInheritedFormatting: clear_inherited_formatting !== false,
        control: { required: required_revision_id, read: readRevisionId, suggest },
      });
      return insertedTableResult(document_id, table, suggest);
    },
  }),

  update_table_style: editTool({
    description: 'Style the cells of an existing table in a Google Doc: background colour, borders and vertical alignment. Pass table_start_index, which is the table\'s own start index: use the value insert_table or append_table returned, or the startIndex of a table element from get_document with include_structure=true (that element also reports rows, columns and cells so you can pick coordinates). This is NOT the index you pass to insert_table, which sits one below the table it creates. The cell range defaults to the whole table; narrow it with row_index, column_index, row_span and column_span, all zero-based. Colours are 6-digit hex such as #1A73E8. A `border` value applies to all four sides and any border_top/border_right/border_bottom/border_left replaces that side. Vertically adjacent cells share one edge, so sending `border` together with a per-side override loses that override wherever cells meet: the next row\'s top is written over the previous row\'s bottom, and only the outer edge of the table keeps it. To underline a header row inside a bordered grid, send the grid first and the override in a second call. content_alignment only shows on a row taller than its text, so on a single line table it changes nothing visible. This call creates nothing and applying the same style twice has the same effect as applying it once, so unlike insert_table a failure is safe to retry as-is. The write is pinned to the revision it validated against, so an edit by someone else between the two makes this fail rather than style the wrong cells. If the table has merged cells Google may style beyond the requested range, and a warning says so. Native Google Docs only: Word (.doc/.docx) uploads are read-only and this tool refuses them.',
    outputSchema: {
      id: z.string(),
      table_start_index: z.number(),
      rows: z.number(),
      columns: z.number(),
      cells_styled: z.number().describe('Size of the requested range. With merged cells Google may style more than this.'),
      applied: z.array(z.string()).describe('The style properties actually set'),
      message: z.string(),
      warning: z.string().optional().describe('Present when the result may not match the requested range exactly'),
    },
    schema: {
      document_id: documentIdField,
      table_start_index: z.number().int().min(1).describe('Start index of the table itself, from insert_table/append_table or a table element in get_document structure. Not the index passed to insert_table, which is one lower.'),
      tab_id: tabIdField,
      row_index: z.number().int().min(0).optional().describe('First row of the range, zero-based. Defaults to 0.'),
      column_index: z.number().int().min(0).optional().describe('First column of the range, zero-based. Defaults to 0.'),
      row_span: z.number().int().min(1).optional().describe('How many rows to style. Defaults to the rest of the table.'),
      column_span: z.number().int().min(1).optional().describe('How many columns to style. Defaults to the rest of the table.'),
      background_color: z.string().optional().describe('Cell background as 6-digit hex, e.g. #EEEEEE'),
      border: borderSchema.optional().describe('Border applied to all four sides'),
      border_top: borderSchema.optional().describe('Overrides `border` for the top side'),
      border_right: borderSchema.optional().describe('Overrides `border` for the right side'),
      border_bottom: borderSchema.optional().describe('Overrides `border` for the bottom side'),
      border_left: borderSchema.optional().describe('Overrides `border` for the left side'),
      content_alignment: z.enum(['TOP', 'MIDDLE', 'BOTTOM']).optional().describe('Vertical alignment of cell content'),
      suggest: suggestField,
    },
    run: async (args, { accessToken }) => {
      const { document_id, table_start_index, tab_id, row_index, column_index, row_span, column_span, suggest } = args;

      const { tableCellStyle, fields } = buildTableCellStyle(args);

      const { resolved } = await fetchResolvedTab(document_id, accessToken, tab_id);
      const element = locateTable(resolved.body, table_start_index, resolved.tabId);
      const { rows, columns, merges } = tableLayout(element);
      const { rowIndex: startRow, columnIndex: startColumn, rowSpan, columnSpan } =
        validateCellRange({ rows, columns }, table_start_index, { row_index, column_index, row_span, column_span });

      // Pin to the revision the bounds were checked against, in case a collaborator resized the table
      await docsBatchUpdate(document_id, accessToken, [{
        updateTableCellStyle: {
          tableRange: {
            tableCellLocation: {
              tableStartLocation: docLocation(table_start_index, resolved.tabId),
              rowIndex: startRow,
              columnIndex: startColumn,
            },
            rowSpan,
            columnSpan,
          },
          tableCellStyle,
          fields: fields.join(','),
        },
      }], { required: resolved.revisionId, suggest });

      const requested = rowSpan * columnSpan;
      return toolResult({
        id: document_id,
        table_start_index,
        rows,
        columns,
        cells_styled: requested,
        applied: fields,
        message: doneMessage(`Styled ${requested} cell(s) (${fields.join(', ')}) in the table at index ${table_start_index}`, suggest),
        ...(merges.length > 0 ? {
          warning: 'This table contains merged cells, so Google may have styled cells outside the requested range. ' +
            'cells_styled counts the range that was asked for, not necessarily the cells affected. ' +
            'Read the document back if the exact extent matters.',
        } : {}),
      });
    },
  }),

  modify_table: editTool({
    description: 'Change the structure of an existing table in a Google Doc: insert or delete a row or column, merge or unmerge cells, set column widths, or pin header rows that repeat on every page. Pass table_start_index, the table\'s own start index from get_document with include_structure=true or from insert_table/append_table. Rows and columns are zero-based. insert_row and insert_column go after the reference cell (below or right) by default, before it with position="before"; delete_row and delete_column remove the reference row or column; merge_cells and unmerge_cells take a rectangular range (row_index, column_index, row_span, column_span; merge spans default to 1, unmerge spans to the rest of the table) and merging keeps the text of every cell, joined into the top-left one; set_column_width takes width_pt for column_index or for every column; pin_header_rows takes pinned_rows (0 unpins). Each call makes one change, pinned to the revision it validated against, so a concurrent edit makes it fail rather than hit the wrong cells. Structural changes shift every index after the table: re-read with get_document before further index-based edits. Native Google Docs only.',
    outputSchema: {
      id: z.string(),
      action: z.enum(TABLE_ACTIONS),
      table_start_index: z.number(),
      rows: z.number().describe('Rows after the change'),
      columns: z.number().describe('Columns after the change'),
      message: z.string(),
    },
    schema: {
      document_id: documentIdField,
      table_start_index: z.number().int().min(1).describe('Start index of the table itself, from get_document structure or insert_table/append_table'),
      action: z.enum(TABLE_ACTIONS).describe('The change to make'),
      tab_id: tabIdField,
      row_index: z.number().int().min(0).optional().describe('Reference row, zero-based. Defaults to 0.'),
      column_index: z.number().int().min(0).optional().describe('Reference column, zero-based. Defaults to 0. For set_column_width, omit to set every column.'),
      position: z.enum(['before', 'after']).optional().describe('insert_row / insert_column: insert before (above / left of) or after (below / right of) the reference cell. Defaults to after.'),
      row_span: z.number().int().min(1).optional().describe('merge_cells / unmerge_cells: rows in the range. Defaults to the rest of the table.'),
      column_span: z.number().int().min(1).optional().describe('merge_cells / unmerge_cells: columns in the range. Defaults to the rest of the table.'),
      width_pt: z.number().min(MIN_COLUMN_WIDTH_PT).max(MAX_COLUMN_WIDTH_PT).optional().describe(`set_column_width: width in points (${MIN_COLUMN_WIDTH_PT}-${MAX_COLUMN_WIDTH_PT}); a letter page body is about 468pt wide`),
      pinned_rows: z.number().int().min(0).optional().describe('pin_header_rows: how many top rows repeat on each page; 0 unpins'),
      suggest: suggestField,
    },
    run: async (args, { accessToken }) => {
      const { document_id, table_start_index, tab_id, suggest } = args;
      const { resolved } = await fetchResolvedTab(document_id, accessToken, tab_id);
      const element = locateTable(resolved.body, table_start_index, resolved.tabId);
      const change = buildTableModification(tableLayout(element), docLocation(table_start_index, resolved.tabId), args);

      await docsBatchUpdate(document_id, accessToken, [change.request], { required: resolved.revisionId, suggest })
        .catch(explainMergeRefusal);

      return toolResult({
        id: document_id,
        action: args.action,
        table_start_index,
        rows: change.shape.rows,
        columns: change.shape.columns,
        message: doneMessage(`${change.message} in the table at index ${table_start_index}`, suggest),
      });
    },
  }),
};
