/** Text editing tools: append_text, replace_text, delete_content and insert_text */

import { z } from 'zod';
import { withGoogleAuth as requirePermissionSecure } from "../auth.js";
import {
  buildWriteControl,
  clearInheritedFormattingRequest,
  clearInheritedParagraphStyleRequests,
  summarizeTabs,
  validateIndexRange,
  validateInsertIndex,
} from "../lib/docs-structure.js";
import { appendInsertion, docLocation, docRange, tabsCriteriaFor } from "../lib/tabs.js";
import { makeDocsRequest } from '../lib/google.js';
import {
  DOCUMENTS_SCOPE,
  REQUIRED_REVISION_DESCRIPTION,
  TAB_ID_DESCRIPTION,
  TAB_LIST_FIELDS,
  docsBatchUpdate,
  docsOnly,
  fetchResolvedTab,
  toolErrorResponse,
  toolResult,
} from './shared.js';

export const textTools = {
  append_text: {
    description: 'Append plain text to the end of a Google Doc (or, with tab_id, to the end of that tab) as a new paragraph: a newline is added before the text unless the document already ends with an empty paragraph or the text starts with one. Use get_document first to verify the document exists and see current content. Optionally pass `required_revision_id` (from get_document with include_structure=true) to fail the write if the document was edited concurrently.',
    outputSchema: {
      id: z.string(),
      message: z.string(),
    },
    schema: {
      document_id: z.string().describe('Google Doc ID'),
      text: z.string().min(1).describe('Plain text to append to the end of the document'),
      tab_id: z.string().optional().describe(TAB_ID_DESCRIPTION),
      clear_inherited_formatting: z.boolean().optional().describe('Clear formatting inherited from the preceding text (character styles, and heading or list paragraph styles) on the newly appended content. Defaults to true.'),
      required_revision_id: z.string().optional().describe(REQUIRED_REVISION_DESCRIPTION),
    },
    handler: requirePermissionSecure(DOCUMENTS_SCOPE, docsOnly(async ({ document_id, text, tab_id, clear_inherited_formatting, required_revision_id }: any, context: any) => {
      try {
        const { accessToken } = context;

        const { resolved } = await fetchResolvedTab(document_id, accessToken, tab_id);
        const insertion = appendInsertion(resolved.body, text);

        const requests: any[] = [{
          insertText: {
            location: docLocation(insertion.index, resolved.tabId),
            text: insertion.text,
          },
        }];

        if (clear_inherited_formatting !== false) {
          const end = insertion.index + insertion.text.length;
          // The separator newline still belongs to the previous paragraph, whose style must stay
          requests.push(
            clearInheritedFormattingRequest(insertion.index, end, resolved.tabId),
            ...clearInheritedParagraphStyleRequests(insertion.paragraphStart, end, resolved.tabId),
          );
        }

        await docsBatchUpdate(document_id, accessToken, requests, buildWriteControl(required_revision_id));

        return toolResult({
          id: document_id,
          message: 'Text appended successfully',
        });
      } catch (err) {
        return toolErrorResponse(err);
      }
    })),
  },

  replace_text: {
    description: 'Replace all occurrences of a text string in one tab of a Google Doc: the tab named by tab_id, or the first tab when omitted, matching where the other editing tools write. Use get_document first to see current content and verify the text to replace exists. Use empty new_text to delete occurrences. To replace across every tab of a multi-tab document, call once per tab. Optionally pass `required_revision_id` (from get_document with include_structure=true) to fail the write if the document was edited concurrently.',
    outputSchema: {
      id: z.string(),
      occurrencesChanged: z.number(),
      message: z.string(),
    },
    schema: {
      document_id: z.string().describe('Google Doc ID'),
      old_text: z.string().describe('Text to find (all occurrences will be replaced)'),
      new_text: z.string().describe('Replacement text (empty string to delete)'),
      tab_id: z.string().optional().describe(TAB_ID_DESCRIPTION),
      match_case: z.boolean().optional().describe('Whether to match case (default true)'),
      required_revision_id: z.string().optional().describe(REQUIRED_REVISION_DESCRIPTION),
    },
    handler: requirePermissionSecure(DOCUMENTS_SCOPE, docsOnly(async ({ document_id, old_text, new_text, tab_id, match_case, required_revision_id }: any, context: any) => {
      try {
        const { accessToken } = context;

        // Docs API returns 400 on empty `old_text`; reject up-front with a
        // clear structured error.
        if (typeof old_text !== 'string' || old_text.length === 0) {
          throw new Error('old_text must be a non-empty string');
        }

        const listing = await makeDocsRequest(`/${encodeURIComponent(document_id)}?fields=${TAB_LIST_FIELDS}`, accessToken, { method: 'GET' });
        const tabsCriteria = tabsCriteriaFor(summarizeTabs(listing.tabs), tab_id);

        const result = await docsBatchUpdate(document_id, accessToken, [{
          replaceAllText: {
            replaceText: new_text,
            containsText: {
              text: old_text,
              matchCase: match_case !== false,
            },
            ...(tabsCriteria ? { tabsCriteria } : {}),
          },
        }], buildWriteControl(required_revision_id)) as { replies: Array<{ replaceAllText?: { occurrencesChanged: number } }> };

        const occurrencesChanged = result.replies?.[0]?.replaceAllText?.occurrencesChanged || 0;

        return toolResult({
          id: document_id,
          occurrencesChanged,
          message: occurrencesChanged > 0
            ? `Replaced ${occurrencesChanged} occurrence(s)`
            : 'No occurrences found',
        });
      } catch (err) {
        return toolErrorResponse(err);
      }
    })),
  },

  delete_content: {
    description: 'Delete content in a Google Doc by index range. Use get_document with include_structure=true first to find the correct startIndex and endIndex. Optionally pass `required_revision_id` (from get_document with include_structure=true) to fail the write if the document was edited concurrently.',
    destructiveHint: true,
    outputSchema: {
      id: z.string(),
      message: z.string(),
    },
    schema: {
      document_id: z.string().describe('Google Doc ID'),
      startIndex: z.coerce.number().int().describe('Start index of content to delete (use get_document with include_structure to find indices)'),
      endIndex: z.coerce.number().int().describe('End index of content to delete (exclusive)'),
      tab_id: z.string().optional().describe(TAB_ID_DESCRIPTION),
      required_revision_id: z.string().optional().describe(REQUIRED_REVISION_DESCRIPTION),
    },
    handler: requirePermissionSecure(DOCUMENTS_SCOPE, docsOnly(async ({ document_id, startIndex, endIndex, tab_id, required_revision_id }: any, context: any) => {
      try {
        const { accessToken } = context;

        validateIndexRange(startIndex, endIndex);

        await docsBatchUpdate(document_id, accessToken, [{
          deleteContentRange: {
            range: {
              ...docRange(startIndex, endIndex, tab_id),
              segmentId: '',
            },
          },
        }], buildWriteControl(required_revision_id));

        return toolResult({
          id: document_id,
          message: `Deleted content from index ${startIndex} to ${endIndex}`,
        });
      } catch (err) {
        return toolErrorResponse(err);
      }
    })),
  },

  insert_text: {
    description: 'Insert text at a specific position in a Google Doc. Use get_document with include_structure=true to find the correct index. NOTE: Inserted text inherits formatting from surrounding text at the insertion point. Use clear_inherited_formatting=true to reset character formatting (bold, italic, underline, strikethrough, font size, colour, link); paragraph styles such as a heading or list are kept. Optionally pass `required_revision_id` (from get_document with include_structure=true) to fail the write if the document was edited concurrently.',
    outputSchema: {
      id: z.string(),
      message: z.string(),
    },
    schema: {
      document_id: z.string().describe('Google Doc ID'),
      text: z.string().describe('Text to insert'),
      index: z.coerce.number().int().describe('Position to insert at (use get_document with include_structure to find indices)'),
      tab_id: z.string().optional().describe(TAB_ID_DESCRIPTION),
      clear_inherited_formatting: z.boolean().optional().describe('Clear character formatting inherited from surrounding text on the newly inserted content; heading and list paragraph styles are kept. Defaults to false.'),
      required_revision_id: z.string().optional().describe(REQUIRED_REVISION_DESCRIPTION),
    },
    handler: requirePermissionSecure(DOCUMENTS_SCOPE, docsOnly(async ({ document_id, text, index, tab_id, clear_inherited_formatting, required_revision_id }: any, context: any) => {
      try {
        const { accessToken } = context;

        // Docs body content starts at index 1 (index 0 is the document
        // start sentinel); inserting at 0 always 400s.
        validateInsertIndex(index);

        const requests: any[] = [{
          insertText: {
            location: docLocation(index, tab_id),
            text,
          },
        }];

        if (clear_inherited_formatting === true) {
          requests.push(clearInheritedFormattingRequest(index, index + text.length, tab_id));
        }

        await docsBatchUpdate(document_id, accessToken, requests, buildWriteControl(required_revision_id));

        return toolResult({
          id: document_id,
          message: `Text inserted at index ${index}`,
        });
      } catch (err) {
        return toolErrorResponse(err);
      }
    })),
  },
};
