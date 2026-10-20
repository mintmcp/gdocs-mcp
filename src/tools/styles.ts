import { z } from 'zod';
import {
  buildParagraphRequests,
  buildTextStyle,
  BULLET_PRESETS,
  docRange,
  LIST_STYLES,
  NAMED_STYLES,
  validateIndexRange,
} from "../lib/docRequests.js";
import {
  docsBatchUpdate,
  documentIdField,
  doneMessage,
  editResultShape,
  editTool,
  requiredRevisionField,
  segmentIdField,
  suggestField,
  tabIdField,
  toolResult,
} from './shared.js';

export const styleTools = {
  update_text_style: editTool({
    description: 'Apply formatting (bold, italic, underline, strikethrough, small caps, font size and family, colour, highlight, superscript/subscript, link) to a text range in a Google Doc. Use get_document with include_structure=true to find the correct indices. NOTE: If you also need to set a heading level via update_paragraph_style, do that first — heading changes reset character formatting. Optionally pass `required_revision_id` (from get_document with include_structure=true) to fail the write if the document was edited concurrently.',
    outputSchema: editResultShape,
    schema: {
      document_id: documentIdField,
      startIndex: z.coerce.number().int().describe('Start index of text range'),
      endIndex: z.coerce.number().int().describe('End index of text range (exclusive)'),
      tab_id: tabIdField,
      bold: z.boolean().optional().describe('Set bold'),
      italic: z.boolean().optional().describe('Set italic'),
      underline: z.boolean().optional().describe('Set underline'),
      strikethrough: z.boolean().optional().describe('Set strikethrough'),
      small_caps: z.boolean().optional().describe('Set small capitals'),
      font_size: z.coerce.number().positive().max(400).optional().describe('Font size in points, e.g. 14'),
      font_family: z.string().min(1).max(100).optional().describe('Font family as Google Docs names it, e.g. "Arial", "Roboto", "Times New Roman"; an unknown name is accepted and falls back when rendered'),
      color: z.string().optional().describe('Text colour as 6-digit hex, e.g. #1A73E8'),
      background_color: z.string().optional().describe('Highlight colour as 6-digit hex, e.g. #FFFF00; empty string removes the highlight'),
      baseline_offset: z.enum(['NONE', 'SUPERSCRIPT', 'SUBSCRIPT']).optional().describe('Superscript, subscript, or NONE to reset'),
      link_url: z.string().optional().describe('Set hyperlink URL (empty string removes the link)'),
      segment_id: segmentIdField,
      required_revision_id: requiredRevisionField,
      suggest: suggestField,
    },
    run: async (args, { accessToken }) => {
      const { document_id, startIndex, endIndex, tab_id, segment_id, required_revision_id, suggest } = args;

      validateIndexRange(startIndex, endIndex, Boolean(segment_id));
      const { textStyle, fields } = buildTextStyle(args);

      await docsBatchUpdate(document_id, accessToken, [{
        updateTextStyle: {
          range: docRange(startIndex, endIndex, tab_id, segment_id),
          textStyle,
          fields: fields.join(','),
        },
      }], { required: required_revision_id, suggest });

      return toolResult({
        id: document_id,
        message: doneMessage(`Applied text style (${fields.join(', ')}) to range ${startIndex}-${endIndex}`, suggest),
      });
    },
  }),

  update_paragraph_style: editTool({
    description: 'Apply paragraph formatting (heading level or named style such as TITLE, alignment, list style, line spacing, space above and below, indents, keep-with-next, page break before) to a range in a Google Doc. list_style turns every paragraph in the range into bullet, numbered or checkbox list items, or removes bullets with "none"; bullet_preset picks a specific Google glyph scheme instead of the default for that list_style; nesting comes from leading tab characters, which Google removes, so indices after the range can shift; re-read the structure with get_document before further index-based edits. Use get_document with include_structure=true to find the correct indices. WARNING: Setting heading_level or named_style applies a named style that resets character-level formatting (bold, italic, etc.). If you also need to apply text styles, call update_paragraph_style first, then apply text styles after. Optionally pass `required_revision_id` (from get_document with include_structure=true) to fail the write if the document was edited concurrently.',
    outputSchema: editResultShape,
    schema: {
      document_id: documentIdField,
      startIndex: z.coerce.number().int().describe('Start index of paragraph range'),
      endIndex: z.coerce.number().int().describe('End index of paragraph range (exclusive)'),
      tab_id: tabIdField,
      heading_level: z.coerce.number().int().min(0).max(6).optional().describe('Heading level: 0=normal text, 1-6=heading levels'),
      named_style: z.enum(NAMED_STYLES).optional().describe('Named paragraph style, for TITLE and SUBTITLE (headings can use heading_level). Not together with heading_level.'),
      alignment: z.enum(['START', 'CENTER', 'END', 'JUSTIFIED']).optional().describe('Paragraph alignment'),
      list_style: z.enum(LIST_STYLES).optional().describe('Make the paragraphs in the range list items: "bullet" (disc), "numbered" (1. 2. 3.), "checkbox", or "none" to remove existing bullets, which also resets the indent the list gave them unless indent_start_pt or indent_first_line_pt is set'),
      bullet_preset: z.enum(BULLET_PRESETS).optional().describe('The exact glyph scheme for list_style, which must match it: NUMBERED_* presets go with "numbered", BULLET_CHECKBOX with "checkbox" and the other BULLET_* with "bullet". E.g. NUMBERED_DECIMAL_NESTED for 1.1.1 numbering or BULLET_ARROW_DIAMOND_DISC'),
      line_spacing: z.number().min(6).max(1000).optional().describe('Line spacing as a percentage: 100 single, 115 Google\'s default, 150, 200 double'),
      space_above_pt: z.number().min(0).max(1000).optional().describe('Space before the paragraph, in points'),
      space_below_pt: z.number().min(0).max(1000).optional().describe('Space after the paragraph, in points'),
      indent_start_pt: z.number().min(0).max(1000).optional().describe('Left indent (start side) in points; 36 is half an inch'),
      indent_end_pt: z.number().min(0).max(1000).optional().describe('Right indent (end side) in points'),
      indent_first_line_pt: z.number().min(0).max(1000).optional().describe('Where the first line starts, in points from the margin like indent_start_pt: 54 with indent_start_pt 36 starts it 18pt further in, and less than indent_start_pt gives a hanging indent'),
      keep_with_next: z.boolean().optional().describe('Keep the paragraph on the same page as the next one'),
      page_break_before: z.boolean().optional().describe('Always start the paragraph on a new page. Body text only: not with segment_id, and not inside a table'),
      segment_id: segmentIdField,
      required_revision_id: requiredRevisionField,
      suggest: suggestField,
    },
    run: async (args, { accessToken }) => {
      const { document_id, startIndex, endIndex, tab_id, segment_id, required_revision_id, suggest } = args;

      validateIndexRange(startIndex, endIndex, Boolean(segment_id));
      const { requests, applied } = buildParagraphRequests(args, docRange(startIndex, endIndex, tab_id, segment_id));

      await docsBatchUpdate(document_id, accessToken, requests, { required: required_revision_id, suggest });

      return toolResult({
        id: document_id,
        message: doneMessage(`Applied paragraph style (${applied.join(', ')}) to range ${startIndex}-${endIndex}`, suggest),
      });
    },
  }),
};
