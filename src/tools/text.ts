import { z } from 'zod';
import {
  appendTextRequests,
  clearInheritedFormattingRequest,
  docLocation,
  docRange,
  validateIndexRange,
  validateInsertIndex,
} from "../lib/docRequests.js";
import { tabsCriteriaFor } from "../lib/tabs.js";
import {
  docsBatchUpdate,
  documentIdField,
  doneMessage,
  editResultShape,
  editTool,
  fetchResolvedTab,
  fetchTabList,
  requiredRevisionField,
  segmentIdField,
  suggestField,
  tabIdField,
  toolResult,
} from './shared.js';

export const textTools = {
  append_text: editTool({
    description: 'Append plain text to the end of a Google Doc (or, with tab_id, to the end of that tab) as a new paragraph: a newline is added before the text unless the document already ends with an empty paragraph or the text starts with one. Use get_document first to verify the document exists and see current content. Optionally pass `required_revision_id` (from get_document with include_structure=true) to fail the write if the document was edited concurrently; without it the text still goes at the end of the revision just read, even if someone edits in between.',
    outputSchema: editResultShape,
    schema: {
      document_id: documentIdField,
      text: z.string().min(1).describe('Plain text to append to the end of the document'),
      tab_id: tabIdField,
      clear_inherited_formatting: z.boolean().optional().describe('Clear formatting inherited from the preceding text (character styles, and heading or list paragraph styles) on the newly appended content. Defaults to true.'),
      required_revision_id: requiredRevisionField,
      suggest: suggestField,
    },
    run: async ({ document_id, text, tab_id, clear_inherited_formatting, required_revision_id, suggest }, { accessToken }) => {
      const { resolved } = await fetchResolvedTab(document_id, accessToken, tab_id);
      const requests = appendTextRequests(resolved.body, text, resolved.tabId, clear_inherited_formatting !== false);

      await docsBatchUpdate(document_id, accessToken, requests, { required: required_revision_id, read: resolved.revisionId, suggest });

      return toolResult({
        id: document_id,
        message: doneMessage('Text appended successfully', suggest),
      });
    },
  }),

  replace_text: editTool({
    description: 'Replace all occurrences of a text string in one tab of a Google Doc: the tab named by tab_id, or the first tab when omitted, matching where the other editing tools write. Use get_document first to see current content and verify the text to replace exists. Use empty new_text to delete occurrences. With use_regex=true, old_text is a regular expression (RE2 syntax, as Google Docs find-and-replace uses) and new_text is literal: capture groups cannot be referenced in the replacement. To replace across every tab of a multi-tab document, call once per tab. Optionally pass `required_revision_id` (from get_document with include_structure=true) to fail the write if the document was edited concurrently.',
    outputSchema: {
      id: z.string(),
      occurrencesChanged: z.number(),
      message: z.string(),
    },
    schema: {
      document_id: documentIdField,
      old_text: z.string().min(1).describe('Text to find (all occurrences will be replaced)'),
      new_text: z.string().describe('Replacement text (empty string to delete)'),
      tab_id: tabIdField,
      match_case: z.boolean().optional().describe('Whether to match case (default true)'),
      use_regex: z.boolean().optional().describe('Treat old_text as a regular expression (RE2 syntax). Defaults to false.'),
      required_revision_id: requiredRevisionField,
      suggest: suggestField,
    },
    run: async ({ document_id, old_text, new_text, tab_id, match_case, use_regex, required_revision_id, suggest }, { accessToken }) => {
      const { tabs } = await fetchTabList(document_id, accessToken);
      const tabsCriteria = tabsCriteriaFor(tabs, tab_id);

      const result = await docsBatchUpdate(document_id, accessToken, [{
        replaceAllText: {
          replaceText: new_text,
          containsText: {
            text: old_text,
            matchCase: match_case !== false,
            ...(use_regex ? { searchByRegex: true } : {}),
          },
          ...(tabsCriteria ? { tabsCriteria } : {}),
        },
      }], { required: required_revision_id, suggest }) as { replies: Array<{ replaceAllText?: { occurrencesChanged: number } }> };

      const occurrencesChanged = result.replies?.[0]?.replaceAllText?.occurrencesChanged || 0;

      return toolResult({
        id: document_id,
        occurrencesChanged,
        message: occurrencesChanged > 0
          ? doneMessage(`Replaced ${occurrencesChanged} occurrence(s)`, suggest)
          : 'No occurrences found',
      });
    },
  }),

  delete_content: editTool({
    description: 'Delete content in a Google Doc by index range. Use get_document with include_structure=true first to find the correct startIndex and endIndex. Optionally pass `required_revision_id` (from get_document with include_structure=true) to fail the write if the document was edited concurrently.',
    destructiveHint: true,
    outputSchema: editResultShape,
    schema: {
      document_id: documentIdField,
      startIndex: z.coerce.number().int().describe('Start index of content to delete (use get_document with include_structure to find indices)'),
      endIndex: z.coerce.number().int().describe('End index of content to delete (exclusive)'),
      tab_id: tabIdField,
      segment_id: segmentIdField,
      required_revision_id: requiredRevisionField,
      suggest: suggestField,
    },
    run: async ({ document_id, startIndex, endIndex, tab_id, segment_id, required_revision_id, suggest }, { accessToken }) => {
      validateIndexRange(startIndex, endIndex, Boolean(segment_id));

      await docsBatchUpdate(document_id, accessToken, [{
        deleteContentRange: {
          range: docRange(startIndex, endIndex, tab_id, segment_id),
        },
      }], { required: required_revision_id, suggest });

      return toolResult({
        id: document_id,
        message: doneMessage(`Deleted content from index ${startIndex} to ${endIndex}`, suggest),
      });
    },
  }),

  insert_text: editTool({
    description: 'Insert text at a specific position in a Google Doc. Use get_document with include_structure=true to find the correct index. NOTE: Inserted text inherits formatting from surrounding text at the insertion point. Use clear_inherited_formatting=true to reset every character style update_text_style can set (bold, italic, font, colour, highlight, link and the rest); paragraph styles such as a heading or list are kept. Optionally pass `required_revision_id` (from get_document with include_structure=true) to fail the write if the document was edited concurrently.',
    outputSchema: editResultShape,
    schema: {
      document_id: documentIdField,
      text: z.string().min(1).describe('Text to insert'),
      index: z.coerce.number().int().describe('Position to insert at (use get_document with include_structure to find indices)'),
      tab_id: tabIdField,
      clear_inherited_formatting: z.boolean().optional().describe('Clear character formatting inherited from surrounding text on the newly inserted content; heading and list paragraph styles are kept. Defaults to false.'),
      segment_id: segmentIdField,
      required_revision_id: requiredRevisionField,
      suggest: suggestField,
    },
    run: async ({ document_id, text, index, tab_id, clear_inherited_formatting, segment_id, required_revision_id, suggest }, { accessToken }) => {
      validateInsertIndex(index, Boolean(segment_id));

      const requests: any[] = [{
        insertText: {
          location: docLocation(index, tab_id, segment_id),
          text,
        },
      }];

      if (clear_inherited_formatting === true) {
        requests.push(clearInheritedFormattingRequest(index, index + text.length, tab_id, segment_id));
      }

      await docsBatchUpdate(document_id, accessToken, requests, { required: required_revision_id, suggest });

      return toolResult({
        id: document_id,
        message: doneMessage(`Text inserted at index ${index}`, suggest),
      });
    },
  }),
};
