import { z } from 'zod';
import {
  docsBatchUpdate,
  documentIdField,
  editTool,
  keysEnum,
  toolResult,
  withRefusalHint,
} from './shared.js';

const ACTIONS = {
  accept: { request: 'acceptSuggestion', done: 'Accepted' },
  reject: { request: 'rejectSuggestion', done: 'Rejected' },
  delete: { request: 'deleteSuggestion', done: 'Deleted' },
} as const;

const actionField = keysEnum(ACTIONS);

// Google answers an unknown or already resolved suggestion id with the same 404 as a missing document
const explainMissingSuggestion = withRefusalHint(
  () => true,
  'If the document opens, a suggestion id is unknown or already resolved: re-read the ids with get_document include_structure=true.',
  404,
);

export const suggestionTools = {
  review_suggestion: editTool({
    description: 'Accept, reject or delete pending suggestions (Google Docs "Suggesting" mode changes) in a Google Doc. Find suggestion ids with get_document include_structure=true: each element of `structure` touched by a suggestion lists them in `suggestions`, as do the headers, footers and footnotes, and the write tools create them when called with suggest=true. Accepting and rejecting need edit access to the document; rejecting also works for the suggestion\'s own author, and deleting only for the author. All ids in one call are applied together: if any fails, none is applied. Native Google Docs only.',
    outputSchema: {
      id: z.string(),
      action: actionField,
      count: z.number(),
      message: z.string(),
    },
    schema: {
      document_id: documentIdField,
      suggestion_ids: z.array(z.string().min(1)).min(1).max(100).describe('Suggestion ids from get_document structure, e.g. "suggest.abc123"'),
      action: actionField.describe('accept applies the suggestion, reject discards it, delete removes it (author only)'),
    },
    run: async ({ document_id, suggestion_ids, action }, { accessToken }) => {
      const { request, done } = ACTIONS[action];
      const ids = [...new Set(suggestion_ids)];

      await docsBatchUpdate(document_id, accessToken, ids.map((suggestionId) => ({ [request]: { suggestionId } })))
        .catch(explainMissingSuggestion);

      return toolResult({
        id: document_id,
        action,
        count: ids.length,
        message: `${done} ${ids.length} suggestion(s)`,
      });
    },
  }),
};
