/** Document property tools: rename_document and update_document_style */

import { z } from 'zod';
import { withGoogleAuth as requirePermissionSecure } from "../auth.js";
import { PAGE_SIZES, buildDocumentStyle } from "../lib/documentStyle.js";
import { wordWriteRefusal } from "../lib/word.js";
import { isWordMime } from "../lib/wordText.js";
import { makeDriveRequest } from '../lib/google.js';
import {
  DOCUMENTS_SCOPE,
  DRIVE_FILE_SCOPE,
  TAB_ID_DESCRIPTION,
  docsBatchUpdate,
  docsOnly,
  fetchResolvedTab,
  toolErrorResponse,
  toolResult,
} from './shared.js';

const MAX_NAME_LENGTH = 500;

export const documentPropertyTools = {
  rename_document: {
    description: 'Rename a Google Doc (its Drive file name, which is the title shown in Docs). The content and the document id do not change. Under this connector\'s drive.file scope the rename works on documents it created or that were opened with it; another app\'s document answers 403, and the Drive connector renames those. Word uploads are refused like every other write.',
    outputSchema: {
      id: z.string(),
      name: z.string(),
      previous_name: z.string().optional(),
      webViewLink: z.string().optional(),
      message: z.string(),
    },
    schema: {
      document_id: z.string().describe('Google Doc ID'),
      name: z.string().max(MAX_NAME_LENGTH).describe('The new name'),
    },
    handler: requirePermissionSecure(DRIVE_FILE_SCOPE, async ({ document_id, name }: any, context: any) => {
      try {
        const { accessToken } = context;
        const trimmed = String(name ?? '').trim();
        if (!trimmed) throw new Error('name must not be empty.');

        const meta = await makeDriveRequest(
          `/files/${encodeURIComponent(document_id)}?fields=name,mimeType,webViewLink&supportsAllDrives=true`,
          accessToken
        );
        if (isWordMime(meta.mimeType || '')) {
          throw new Error(wordWriteRefusal(meta.name || document_id, meta.mimeType, meta.webViewLink || `https://drive.google.com/file/d/${document_id}/view`));
        }

        const updated = await makeDriveRequest(
          `/files/${encodeURIComponent(document_id)}?supportsAllDrives=true&fields=id,name,webViewLink`,
          accessToken,
          { method: 'PATCH', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ name: trimmed }) }
        );

        return toolResult({
          id: document_id,
          name: updated.name ?? trimmed,
          previous_name: meta.name,
          webViewLink: updated.webViewLink ?? meta.webViewLink,
          message: `Renamed '${meta.name}' to '${updated.name ?? trimmed}'`,
        });
      } catch (err) {
        return toolErrorResponse(err);
      }
    }),
  },

  update_document_style: {
    description: `Change page setup of a Google Doc (or of one tab with tab_id): margins in points (72 per inch; Google's default is 72 on every side), page size (${Object.keys(PAGE_SIZES).join(', ')}), orientation, whether the first page has its own header and footer, and the number the page count starts at. Only the properties passed change. Native Google Docs only.`,
    outputSchema: {
      id: z.string(),
      applied: z.array(z.string()).describe('The document style fields set'),
      message: z.string(),
    },
    schema: {
      document_id: z.string().describe('Google Doc ID'),
      tab_id: z.string().min(1).optional().describe(TAB_ID_DESCRIPTION),
      margin_top_pt: z.number().min(0).max(1000).optional().describe('Top margin in points'),
      margin_bottom_pt: z.number().min(0).max(1000).optional().describe('Bottom margin in points'),
      margin_left_pt: z.number().min(0).max(1000).optional().describe('Left margin in points'),
      margin_right_pt: z.number().min(0).max(1000).optional().describe('Right margin in points'),
      page_size: z.enum(Object.keys(PAGE_SIZES) as [string, ...string[]]).optional().describe('Paper size; portrait unless orientation says otherwise'),
      orientation: z.enum(['portrait', 'landscape']).optional().describe('Page orientation; alone, it rotates the current page size'),
      first_page_header_footer: z.boolean().optional().describe('Give the first page its own header and footer (Google then creates FIRST_PAGE segments, listed by get_document)'),
      page_number_start: z.number().int().min(0).max(100_000).optional().describe('Number of the first page'),
    },
    handler: requirePermissionSecure(DOCUMENTS_SCOPE, docsOnly(async (args: any, context: any) => {
      try {
        const { document_id, tab_id } = args;
        const { accessToken } = context;

        const { resolved } = await fetchResolvedTab(document_id, accessToken, tab_id);
        const { documentStyle, fields } = buildDocumentStyle(args, resolved.documentStyle?.pageSize);
        if (fields.length === 0) {
          throw new Error('At least one property must be provided (margin_top_pt, margin_bottom_pt, margin_left_pt, margin_right_pt, page_size, orientation, first_page_header_footer, page_number_start).');
        }

        await docsBatchUpdate(document_id, accessToken, [{
          updateDocumentStyle: {
            documentStyle,
            fields: fields.join(','),
            ...(resolved.tabId ? { tabId: resolved.tabId } : {}),
          },
        }]);

        return toolResult({
          id: document_id,
          applied: fields,
          message: `Updated document style (${fields.join(', ')})${resolved.tabId ? ` in tab '${resolved.tabId}'` : ''}`,
        });
      } catch (err) {
        return toolErrorResponse(err);
      }
    })),
  },
};
