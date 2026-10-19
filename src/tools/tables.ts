/** Table tools: append_table, insert_table and update_table_style */

import { z } from 'zod';
import { withGoogleAuth as requirePermissionSecure } from "../auth.js";
import {
  buildTableInsertRequests,
  buildWriteControl,
  clearInheritedFormattingRequest,
  validateInsertIndex,
} from "../lib/docs-structure.js";
import {
  borderSchema,
  buildTableCellStyle,
  findTableAt,
  findTableStartingAt,
  hasMergedCells,
  nestingRefusal,
  normalizeTableData,
  styleFailureWarning,
  tableStartFor,
  MAX_CELL_CHARS,
  MAX_TABLE_COLUMNS,
  MAX_TABLE_ROWS,
} from "../lib/tableModel.js";
import { docLocation, endOfBodyIndex } from "../lib/tabs.js";
import { log, errorFields } from "../lib/log.js";
import {
  DOCUMENTS_SCOPE,
  REQUIRED_REVISION_DESCRIPTION,
  TAB_ID_DESCRIPTION,
  docsBatchUpdate,
  docsOnly,
  fetchResolvedTab,
  toolErrorResponse,
  toolResult,
} from './shared.js';

interface InsertedTable {
  rows: number;
  columns: number;
  tableStartIndex: number;
  startIndexVerified: boolean;
  warning?: string;
}

/**
 * Insert a table and fill its cells, then optionally clear the formatting it
 * inherited from the insertion point.
 *
 * Once the insert batch returns the table exists, so nothing past that point
 * may throw: a retry would create a second table. Failures after it come back
 * as a warning on a successful result instead.
 */
async function insertTableAt(
  documentId: string,
  accessToken: string,
  tableData: string[][],
  index: number,
  clearInheritedFormatting: boolean,
  tabId?: string,
): Promise<InsertedTable> {
  const columns = tableData[0].length;

  await docsBatchUpdate(documentId, accessToken, [
    { insertTable: { location: docLocation(index, tabId), rows: tableData.length, columns } },
    ...buildTableInsertRequests(tableData, index, tabId),
  ]);

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
    const { tabBody } = await fetchResolvedTab(documentId, accessToken, tabId);
    table = findTableAt(tabBody, index);

    if (table) {
      await docsBatchUpdate(documentId, accessToken, [
        clearInheritedFormattingRequest(table.startIndex + 1, table.endIndex, tabId),
      ]);
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

const tableOutput = (id: string, table: InsertedTable, message: string) => ({
  id,
  rows: table.rows,
  columns: table.columns,
  table_start_index: table.tableStartIndex,
  table_start_index_verified: table.startIndexVerified,
  message,
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
  append_table: {
    description: 'Insert a table with data at the end of a Google Doc. Supports ragged rows (will be padded with empty cells). Optionally pass `required_revision_id` (from get_document with include_structure=true) to fail the initial insert if the document was edited concurrently; the follow-up formatting clear is not revision-gated since it operates on the just-inserted table. If this call fails, check the document with get_document before retrying — a table that was already created would otherwise be inserted twice. On success the result reports table_start_index plus table_start_index_verified; if a warning comes back, the table already exists — never insert it again. Native Google Docs only: Word (.doc/.docx) uploads are read-only and this tool refuses them.',
    outputSchema: {
      id: z.string(),
      rows: z.number(),
      columns: z.number(),
      table_start_index: z.number(),
      table_start_index_verified: z.boolean().describe('True when the table was found at this index after insertion. False means the index is inferred from the API contract and was not confirmed — re-read the document before using it.'),
      message: z.string(),
      warning: z.string().optional(),
    },
    schema: {
      document_id: z.string().describe('Google Doc ID'),
      rows: z.array(z.array(z.string().max(MAX_CELL_CHARS)).max(MAX_TABLE_COLUMNS)).max(MAX_TABLE_ROWS).describe('Table data as array of rows, each row is array of cell strings'),
      tab_id: z.string().optional().describe(TAB_ID_DESCRIPTION),
      clear_inherited_formatting: z.boolean().optional().describe('Clear inherited formatting on the newly appended table content. Defaults to true.'),
      required_revision_id: z.string().optional().describe('If provided, the initial table insert fails with a 400 if the document revision has changed since this revisionId. Use the `revisionId` returned by get_document(include_structure=true) for optimistic concurrency control.'),
    },
    handler: requirePermissionSecure(DOCUMENTS_SCOPE, docsOnly(async ({ document_id, rows, tab_id, clear_inherited_formatting, required_revision_id }: any, context: any) => {
      try {
        const { accessToken } = context;

        const tableData = normalizeTableData(rows);
        const maxCols = tableData[0].length;

        const { resolved } = await fetchResolvedTab(document_id, accessToken, tab_id);
        const insertIndex = endOfBodyIndex(resolved.body);

        // Build requests: first insert empty table, then populate cells in reverse order
        const requests: any[] = [
          {
            insertTable: {
              location: docLocation(insertIndex, resolved.tabId),
              rows: tableData.length,
              columns: maxCols,
            },
          },
          ...buildTableInsertRequests(tableData, insertIndex, resolved.tabId),
        ];

        await docsBatchUpdate(document_id, accessToken, requests, buildWriteControl(required_revision_id));

        // Clear inherited formatting on the newly inserted table — best-effort.
        // The table is already inserted at this point; if this second
        // batchUpdate fails, we still want to surface success and warn
        // about the formatting step rather than fail the whole operation.
        //
        // Locate OUR table by matching the known insertion point AND shape.
        // A naive `lastTable` lookup would strip formatting from a
        // different table if a concurrent edit appended one after our
        // insert; this also accounts for the rare case where another
        // client inserted at the same stale end position by additionally
        // matching rows/columns. If more than one table still satisfies
        // both filters, we refuse to guess and surface a warning instead
        // of silently formatting the wrong table.
        let formattingWarning: string | undefined;
        // The API contract puts the table one past the insertion index;
        // locating it below upgrades this from inferred to confirmed.
        let tableStartIndex = tableStartFor(insertIndex);
        let tableStartIndexVerified = false;
        if (clear_inherited_formatting !== false) {
          try {
            const { resolved: updated } = await fetchResolvedTab(document_id, accessToken, resolved.tabId);
            const allTables = updated.body.filter((el: any) => el.table);
            const candidates = allTables.filter((t: any) =>
              typeof t.startIndex === 'number' &&
              typeof t.endIndex === 'number' &&
              t.startIndex >= insertIndex - 1 &&
              t.startIndex <= insertIndex + 1 &&
              t.table?.rows === tableData.length &&
              t.table?.columns === maxCols
            );
            if (candidates.length === 1) {
              const ourTable = candidates[0];
              tableStartIndex = ourTable.startIndex;
              tableStartIndexVerified = true;
              await docsBatchUpdate(document_id, accessToken, [
                clearInheritedFormattingRequest(ourTable.startIndex + 1, ourTable.endIndex, resolved.tabId),
              ]);
            } else if (candidates.length === 0) {
              formattingWarning = 'Could not locate the inserted table to clear formatting (document may have been edited concurrently).';
            } else {
              formattingWarning = `Skipped clearing formatting: ${candidates.length} tables match the insertion point and shape; refusing to guess which is ours.`;
            }
          } catch (formattingErr) {
            formattingWarning = formattingErr instanceof Error ? formattingErr.message : String(formattingErr);
          }
        }

        const output: {
          id: string;
          rows: number;
          columns: number;
          table_start_index: number;
          table_start_index_verified: boolean;
          message: string;
          warning?: string;
        } = {
          id: document_id,
          rows: tableData.length,
          columns: maxCols,
          table_start_index: tableStartIndex,
          table_start_index_verified: tableStartIndexVerified,
          message: `Table inserted with ${tableData.length} rows and ${maxCols} columns`,
        };
        if (formattingWarning) {
          output.warning = `Table inserted, but clearing inherited formatting failed: ${formattingWarning}. Do not insert the table again.`;
        }
        return toolResult(output);
      } catch (err) {
        return toolErrorResponse(err);
      }
    })),
  },

  insert_table: {
    description: 'Insert a table with data at a specific index in a Google Doc — use this for any position other than the end of the document; index 1 is the top of the document body. To place a table between existing content, call get_document with include_structure=true and pass the startIndex of the paragraph the table should appear above. Google accepts most indices, so a wrong one usually misplaces the table instead of failing: an index inside a paragraph splits that paragraph, and an index inside a table cell nests the new table there (nested tables cannot be found afterwards, so their inherited formatting is left alone and table_start_index comes back unverified). If this call fails, check the document with get_document before retrying — a table that was already created would otherwise be inserted twice. On success the result reports table_start_index plus table_start_index_verified; if a warning comes back, the table already exists — never insert it again. Supports ragged rows (will be padded with empty cells). Native Google Docs only: Word (.doc/.docx) uploads are read-only and this tool refuses them.',
    outputSchema: tableToolOutputSchema,
    schema: {
      document_id: z.string().describe('Google Doc ID'),
      rows: z.array(z.array(z.string().max(MAX_CELL_CHARS)).max(MAX_TABLE_COLUMNS)).max(MAX_TABLE_ROWS).describe('Table data as array of rows, each row is array of cell strings. Pass empty rows (e.g. [[], [], []]) with columns set to create an empty table.'),
      index: z.number().int().min(1).describe('Document index to insert the table at. 1 is the top of the document body. Use get_document with include_structure=true to find indices.'),
      tab_id: z.string().optional().describe(TAB_ID_DESCRIPTION),
      columns: z.number().int().min(1).max(MAX_TABLE_COLUMNS).optional().describe('Force the table width; short rows are padded with empty cells. Defaults to the widest row. Must not be smaller than the widest row.'),
      clear_inherited_formatting: z.boolean().optional().describe('Clear formatting inherited from the insertion point on the new table content. Defaults to true.'),
      allow_nested: z.boolean().optional().describe('Allow an index that falls inside an existing table, which nests the new table in that cell. Defaults to false, which refuses such an index and tells you which ones would work.'),
    },
    handler: requirePermissionSecure(DOCUMENTS_SCOPE, docsOnly(async ({ document_id, rows, index, tab_id, columns, clear_inherited_formatting, allow_nested }: any, context: any) => {
      try {
        const { accessToken } = context;

        if (!Number.isInteger(index)) {
          throw new Error('index must be a whole number. Use get_document with include_structure=true to find valid indices.');
        }
        validateInsertIndex(index);

        const tableData = normalizeTableData(rows, columns);

        if (allow_nested !== true) {
          const { tabBody } = await fetchResolvedTab(document_id, accessToken, tab_id);
          const refusal = nestingRefusal(tabBody, index);
          if (refusal) {
            throw new Error(refusal);
          }
        }

        const table = await insertTableAt(
          document_id,
          accessToken,
          tableData,
          index,
          clear_inherited_formatting !== false,
          tab_id,
        );

        return toolResult(tableOutput(
          document_id,
          table,
          `Table inserted with ${table.rows} rows and ${table.columns} columns` +
            (table.startIndexVerified ? `, starting at index ${table.tableStartIndex}` : ''),
        ));
      } catch (err) {
        return toolErrorResponse(err);
      }
    })),
  },

  update_table_style: {
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
      document_id: z.string().describe('Google Doc ID'),
      table_start_index: z.number().int().min(1).describe('Start index of the table itself, from insert_table/append_table or a table element in get_document structure. Not the index passed to insert_table, which is one lower.'),
      tab_id: z.string().optional().describe(TAB_ID_DESCRIPTION),
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
    },
    handler: requirePermissionSecure(DOCUMENTS_SCOPE, docsOnly(async (args: any, context: any) => {
      try {
        const { document_id, table_start_index, tab_id, row_index, column_index, row_span, column_span } = args;
        const { accessToken } = context;

        if (!Number.isInteger(table_start_index)) {
          throw new Error('table_start_index must be a whole number. Use get_document with include_structure=true and take the startIndex of a table element.');
        }

        const { tableCellStyle, fields } = buildTableCellStyle(args);
        if (fields.length === 0) {
          throw new Error('At least one style property must be provided (background_color, border, border_top, border_right, border_bottom, border_left, content_alignment).');
        }

        const { resolved, tabBody } = await fetchResolvedTab(document_id, accessToken, tab_id);
        const element = findTableStartingAt(tabBody, table_start_index);
        if (!element) {
          const tables = (tabBody.body?.content ?? []).filter((c) => c.table);
          // An index inside a table is a nested table, which the Docs API does not
          // expose as its own element, so no lookup can ever reach it.
          const enclosing = tables.find(
            (c) => (c.startIndex ?? 0) < table_start_index && table_start_index < c.endIndex
          );
          if (enclosing) {
            throw new Error(
              `Index ${table_start_index} is inside the table at ${enclosing.startIndex}-${enclosing.endIndex}, ` +
              `so it refers to a table nested in one of its cells. Nested tables cannot be styled. ` +
              `Style the outer table at ${enclosing.startIndex} instead.`
            );
          }
          const starts = tables.map((c) => c.startIndex);
          const listed = starts.slice(0, 20).join(', ');
          const rest = starts.length > 20 ? ` and ${starts.length - 20} more` : '';
          throw new Error(
            `No table starts at index ${table_start_index}. ` +
            (starts.length
              ? `Tables here start at ${listed}${rest}.`
              : 'No tables were found.') +
            (resolved.tabId
              ? ` Only the tab '${resolved.tabId}' was read; pass a different tab_id for tables in other tabs.`
              : '')
          );
        }

        const rows = element.table?.rows ?? 0;
        const columns = element.table?.columns ?? 0;
        const startRow = row_index ?? 0;
        const startColumn = column_index ?? 0;
        const size = `The table at index ${table_start_index} has ${rows} rows and ${columns} columns`;

        // Check the start before deriving a span from it, or the default span
        // goes negative and the range reads backwards.
        if (startRow >= rows || startColumn >= columns) {
          throw new Error(
            `${size}, so row_index ${startRow} and column_index ${startColumn} are out of range. ` +
            `Both are zero-based, so the last cell is row ${rows - 1}, column ${columns - 1}.`
          );
        }

        const rowSpan = row_span ?? rows - startRow;
        const columnSpan = column_span ?? columns - startColumn;

        if (startRow + rowSpan > rows || startColumn + columnSpan > columns) {
          throw new Error(
            `The requested range (rows ${startRow}-${startRow + rowSpan - 1}, columns ${startColumn}-${startColumn + columnSpan - 1}) ` +
            `does not fit. ${size}.`
          );
        }

        // The bounds above were checked against this revision, so pin the write to it
        // rather than let Google apply the range to a table a collaborator has resized.
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
        }], buildWriteControl(tabBody.revisionId));

        const requested = rowSpan * columnSpan;
        const output: any = {
          id: document_id,
          table_start_index,
          rows,
          columns,
          cells_styled: requested,
          applied: fields,
          message: `Styled ${requested} cell(s) (${fields.join(', ')}) in the table at index ${table_start_index}`,
        };
        const warnings: string[] = [];
        if (!tabBody.revisionId) {
          warnings.push(
            'The document reported no revision, so this write could not be pinned to the revision it was validated against. ' +
            'A concurrent edit could have changed the table between the check and the write.'
          );
        }
        if (hasMergedCells(element)) {
          warnings.push(
            'This table contains merged cells, so Google may have styled cells outside the requested range. ' +
            'cells_styled counts the range that was asked for, not necessarily the cells affected. ' +
            'Read the document back if the exact extent matters.'
          );
        }
        if (warnings.length) {
          output.warning = warnings.join(' ');
        }
        return toolResult(output);
      } catch (err) {
        return toolErrorResponse(err);
      }
    })),
  },
};
