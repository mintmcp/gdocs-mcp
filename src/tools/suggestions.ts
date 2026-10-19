/** review_suggestion: accept, reject or delete pending suggestions by id */

import { z } from 'zod';
import { withGoogleAuth as requirePermissionSecure } from "../auth.js";
import { DOCUMENTS_SCOPE, docsBatchUpdate, docsOnly, toolErrorResponse, toolResult } from './shared.js';

const ACTIONS = {
  accept: 'acceptSuggestion',
  reject: 'rejectSuggestion',
  delete: 'deleteSuggestion',
} as const;

export const suggestionTools = {
  review_suggestion: {
    description: 'Accept, reject or delete pending suggestions (Google Docs "Suggesting" mode changes) in a Google Doc. Find suggestion ids with get_document include_structure=true: each paragraph or table touched by a suggestion lists them in `suggestions`, and the write tools create them when called with suggest=true. Accepting and rejecting need edit access to the document; rejecting also works for the suggestion\'s own author, and deleting only for the author. All ids in one call are applied together: if any fails, none is applied. Native Google Docs only.',
    outputSchema: {
      id: z.string(),
      action: z.enum(['accept', 'reject', 'delete']),
      count: z.number(),
      message: z.string(),
    },
    schema: {
      document_id: z.string().describe('Google Doc ID'),
      suggestion_ids: z.array(z.string().min(1)).min(1).max(100).describe('Suggestion ids from get_document structure, e.g. "suggest.abc123"'),
      action: z.enum(['accept', 'reject', 'delete']).describe('accept applies the suggestion, reject discards it, delete removes it (author only)'),
    },
    handler: requirePermissionSecure(DOCUMENTS_SCOPE, docsOnly(async ({ document_id, suggestion_ids, action }: any, context: any) => {
      try {
        const { accessToken } = context;
        const kind = ACTIONS[action as keyof typeof ACTIONS];
        const ids: string[] = [...new Set<string>(suggestion_ids)];

        await docsBatchUpdate(document_id, accessToken, ids.map((suggestionId) => ({ [kind]: { suggestionId } })));

        return toolResult({
          id: document_id,
          action,
          count: ids.length,
          message: `${action === 'accept' ? 'Accepted' : action === 'reject' ? 'Rejected' : 'Deleted'} ${ids.length} suggestion(s)`,
        });
      } catch (err) {
        return toolErrorResponse(err);
      }
    })),
  },
};
