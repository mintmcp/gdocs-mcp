/** Style tools: update_text_style and update_paragraph_style */

import { z } from 'zod';
import { withGoogleAuth as requirePermissionSecure } from "../auth.js";
import {
  buildWriteControl,
  listStyleRequest,
  buildTextStyle,
  validateIndexRange,
} from "../lib/docs-structure.js";
import { docRange } from "../lib/tabs.js";
import {
  DOCUMENTS_SCOPE,
  REQUIRED_REVISION_DESCRIPTION,
  TAB_ID_DESCRIPTION,
  docsBatchUpdate,
  docsOnly,
  toolErrorResponse,
  toolResult,
} from './shared.js';

export const styleTools = {
  update_text_style: {
    description: 'Apply formatting (bold, italic, underline, strikethrough, font size, colour, link) to a text range in a Google Doc. Use get_document with include_structure=true to find the correct indices. NOTE: If you also need to set a heading level via update_paragraph_style, do that first — heading changes reset character formatting. Optionally pass `required_revision_id` (from get_document with include_structure=true) to fail the write if the document was edited concurrently.',
    outputSchema: {
      id: z.string(),
      message: z.string(),
    },
    schema: {
      document_id: z.string().describe('Google Doc ID'),
      startIndex: z.coerce.number().int().describe('Start index of text range'),
      endIndex: z.coerce.number().int().describe('End index of text range (exclusive)'),
      tab_id: z.string().optional().describe(TAB_ID_DESCRIPTION),
      bold: z.boolean().optional().describe('Set bold'),
      italic: z.boolean().optional().describe('Set italic'),
      underline: z.boolean().optional().describe('Set underline'),
      strikethrough: z.boolean().optional().describe('Set strikethrough'),
      font_size: z.coerce.number().positive().max(400).optional().describe('Font size in points, e.g. 14'),
      color: z.string().optional().describe('Text colour as 6-digit hex, e.g. #1A73E8'),
      link_url: z.string().optional().describe('Set hyperlink URL (empty string removes the link)'),
      required_revision_id: z.string().optional().describe(REQUIRED_REVISION_DESCRIPTION),
    },
    handler: requirePermissionSecure(DOCUMENTS_SCOPE, docsOnly(async ({ document_id, startIndex, endIndex, tab_id, bold, italic, underline, strikethrough, font_size, color, link_url, required_revision_id }: any, context: any) => {
      try {
        const { accessToken } = context;

        validateIndexRange(startIndex, endIndex);
        const { textStyle, fields } = buildTextStyle({ bold, italic, underline, strikethrough, font_size, color, link_url });

        await docsBatchUpdate(document_id, accessToken, [{
          updateTextStyle: {
            range: docRange(startIndex, endIndex, tab_id),
            textStyle,
            fields: fields.join(','),
          },
        }], buildWriteControl(required_revision_id));

        return toolResult({
          id: document_id,
          message: `Applied text style (${fields.join(', ')}) to range ${startIndex}-${endIndex}`,
        });
      } catch (err) {
        return toolErrorResponse(err);
      }
    })),
  },

  update_paragraph_style: {
    description: 'Apply paragraph formatting (heading level, alignment, list style) to a range in a Google Doc. list_style turns every paragraph in the range into bullet, numbered or checkbox list items, or removes bullets with "none"; nesting comes from leading tab characters, which Google removes, so indices after the range can shift; re-read the structure with get_document before further index-based edits. Use get_document with include_structure=true to find the correct indices. WARNING: Setting heading_level applies a named style that resets character-level formatting (bold, italic, etc.). If you also need to apply text styles, call update_paragraph_style first, then apply text styles after. Optionally pass `required_revision_id` (from get_document with include_structure=true) to fail the write if the document was edited concurrently.',
    outputSchema: {
      id: z.string(),
      message: z.string(),
    },
    schema: {
      document_id: z.string().describe('Google Doc ID'),
      startIndex: z.coerce.number().int().describe('Start index of paragraph range'),
      endIndex: z.coerce.number().int().describe('End index of paragraph range (exclusive)'),
      tab_id: z.string().optional().describe(TAB_ID_DESCRIPTION),
      heading_level: z.coerce.number().int().min(0).max(6).optional().describe('Heading level: 0=normal text, 1-6=heading levels'),
      alignment: z.enum(['START', 'CENTER', 'END', 'JUSTIFIED']).optional().describe('Paragraph alignment'),
      list_style: z.enum(['bullet', 'numbered', 'checkbox', 'none']).optional().describe('Make the paragraphs in the range list items: "bullet" (disc), "numbered" (1. 2. 3.), "checkbox", or "none" to remove existing bullets'),
      required_revision_id: z.string().optional().describe(REQUIRED_REVISION_DESCRIPTION),
    },
    handler: requirePermissionSecure(DOCUMENTS_SCOPE, docsOnly(async ({ document_id, startIndex, endIndex, tab_id, heading_level, alignment, list_style, required_revision_id }: any, context: any) => {
      try {
        const { accessToken } = context;

        validateIndexRange(startIndex, endIndex);

        const paragraphStyle: any = {};
        const fields: string[] = [];

        if (heading_level !== undefined) {
          const headingMap: Record<number, string> = {
            0: 'NORMAL_TEXT',
            1: 'HEADING_1',
            2: 'HEADING_2',
            3: 'HEADING_3',
            4: 'HEADING_4',
            5: 'HEADING_5',
            6: 'HEADING_6',
          };
          const namedStyle = headingMap[heading_level];
          if (!namedStyle) {
            throw new Error('heading_level must be 0 (normal) or 1-6');
          }
          paragraphStyle.namedStyleType = namedStyle;
          fields.push('namedStyleType');
        }

        if (alignment !== undefined) {
          paragraphStyle.alignment = alignment;
          fields.push('alignment');
        }

        if (fields.length === 0 && !list_style) {
          throw new Error('At least one style property must be provided (heading_level, alignment, list_style)');
        }

        const requests: any[] = [];
        if (fields.length > 0) {
          requests.push({
            updateParagraphStyle: {
              range: docRange(startIndex, endIndex, tab_id),
              paragraphStyle,
              fields: fields.join(','),
            },
          });
        }
        if (list_style) {
          requests.push(listStyleRequest(startIndex, endIndex, list_style, tab_id));
        }
        const applied = [...fields, ...(list_style ? [`list_style=${list_style}`] : [])];

        await docsBatchUpdate(document_id, accessToken, requests, buildWriteControl(required_revision_id));

        return toolResult({
          id: document_id,
          message: `Applied paragraph style (${applied.join(', ')}) to range ${startIndex}-${endIndex}`,
        });
      } catch (err) {
        return toolErrorResponse(err);
      }
    })),
  },
};
