import { z } from 'zod';
import { buildDocumentStyle, PAGE_SIZES } from "../lib/documentStyle.js";
import {
  docsBatchUpdate,
  documentIdField,
  editTool,
  fetchResolvedTab,
  keysEnum,
  tabIdField,
  toolResult,
} from './shared.js';

export const documentPropertyTools = {
  update_document_style: editTool({
    description: `Change page setup of a Google Doc (or of one tab with tab_id): margins in points (72 per inch; Google's default is 72 on every side), page size (${Object.keys(PAGE_SIZES).join(', ')}), orientation, whether the first page has its own header and footer, and the number the page count starts at. Only the properties passed change. Native Google Docs only.`,
    outputSchema: {
      id: z.string(),
      applied: z.array(z.string()).describe('The document style fields set'),
      message: z.string(),
    },
    schema: {
      document_id: documentIdField,
      tab_id: tabIdField,
      margin_top_pt: z.number().min(0).max(1000).optional().describe('Top margin in points'),
      margin_bottom_pt: z.number().min(0).max(1000).optional().describe('Bottom margin in points'),
      margin_left_pt: z.number().min(0).max(1000).optional().describe('Left margin in points'),
      margin_right_pt: z.number().min(0).max(1000).optional().describe('Right margin in points'),
      page_size: keysEnum(PAGE_SIZES).optional().describe('Paper size; portrait unless orientation says otherwise'),
      orientation: z.enum(['portrait', 'landscape']).optional().describe('Page orientation; alone, it rotates the current page size'),
      first_page_header_footer: z.boolean().optional().describe('Give the first page its own header and footer. The first page then shows none until one is added in the Docs editor, since Google creates no FIRST_PAGE segment here and set_header_footer manages only the default ones'),
      page_number_start: z.number().int().min(0).max(100_000).optional().describe('Number of the first page'),
    },
    run: async (args, { accessToken }) => {
      const { document_id, tab_id } = args;

      // Orientation alone rotates the size just read, so pin to that revision
      const fromReadSize = args.orientation !== undefined && args.page_size === undefined;
      const resolved = fromReadSize ? (await fetchResolvedTab(document_id, accessToken, tab_id)).resolved : undefined;
      const { documentStyle, fields } = buildDocumentStyle(args, resolved?.documentStyle?.pageSize);
      const tabId = resolved?.tabId ?? tab_id;

      await docsBatchUpdate(document_id, accessToken, [{
        updateDocumentStyle: {
          documentStyle,
          fields: fields.join(','),
          ...(tabId ? { tabId } : {}),
        },
      }], { required: resolved?.revisionId });

      return toolResult({
        id: document_id,
        applied: fields,
        message: `Updated document style (${fields.join(', ')})${tabId ? ` in tab '${tabId}'` : ''}`,
      });
    },
  }),
};
