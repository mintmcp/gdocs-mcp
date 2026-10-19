/** manage_tab: add, rename or delete a tab of a Google Doc */

import { z } from 'zod';
import { withGoogleAuth as requirePermissionSecure } from "../auth.js";
import { summarizeTabs } from "../lib/docs-structure.js";
import { unknownTabError } from "../lib/tabs.js";
import { makeDocsRequest } from '../lib/google.js';
import { DOCUMENTS_SCOPE, TAB_LIST_FIELDS, docsBatchUpdate, docsOnly, toolErrorResponse, toolResult } from './shared.js';

const tabSchema = z.object({
  tabId: z.string(),
  title: z.string().optional(),
  index: z.number().int().optional(),
});

export const tabTools = {
  manage_tab: {
    description: 'Add, rename or delete a tab of a Google Doc. "add" creates a top-level tab with `title` and returns its id, ready for the editing tools\' tab_id; "rename" sets the title of tab_id; "delete" removes tab_id and every tab nested under it, with all their content. A document always keeps at least one tab. get_document lists the tabs. Native Google Docs only.',
    destructiveHint: true,
    outputSchema: {
      id: z.string(),
      action: z.enum(['add', 'rename', 'delete']),
      tab: tabSchema.optional().describe('The tab added or renamed'),
      message: z.string(),
    },
    schema: {
      document_id: z.string().describe('Google Doc ID'),
      action: z.enum(['add', 'rename', 'delete']),
      tab_id: z.string().min(1).optional().describe('rename / delete: the tab, as tabProperties.tabId or the value after ?tab= in its URL'),
      title: z.string().min(1).max(200).optional().describe('add / rename: the tab title'),
    },
    handler: requirePermissionSecure(DOCUMENTS_SCOPE, docsOnly(async ({ document_id, action, tab_id, title }: any, context: any) => {
      try {
        const { accessToken } = context;

        if (action === 'add') {
          if (!title) throw new Error('add needs a title for the new tab.');
          const reply = await docsBatchUpdate(document_id, accessToken, [{ addDocumentTab: { tabProperties: { title } } }]);
          const props = reply?.replies?.[0]?.addDocumentTab?.tabProperties;
          const tab = props?.tabId ? { tabId: props.tabId, ...(props.title !== undefined ? { title: props.title } : {}), ...(props.index !== undefined ? { index: props.index } : {}) } : undefined;
          return toolResult({
            id: document_id,
            action,
            ...(tab ? { tab } : {}),
            message: tab ? `Added tab '${title}' with id ${tab.tabId}` : `Added tab '${title}'; read get_document for its id`,
          });
        }

        if (!tab_id) throw new Error(`${action} needs tab_id.`);
        const listing = await makeDocsRequest(`/${encodeURIComponent(document_id)}?fields=${TAB_LIST_FIELDS}`, accessToken, { method: 'GET' });
        const tabs = summarizeTabs(listing.tabs);
        const target = tabs.find((t) => t.tabId === tab_id);
        if (!target) throw unknownTabError(tab_id, tabs);

        if (action === 'rename') {
          if (!title) throw new Error('rename needs the new title.');
          await docsBatchUpdate(document_id, accessToken, [
            { updateDocumentTabProperties: { tabProperties: { tabId: tab_id, title }, fields: 'title' } },
          ]);
          return toolResult({
            id: document_id,
            action,
            tab: { tabId: tab_id, title, ...(target.index !== undefined ? { index: target.index } : {}) },
            message: `Renamed tab ${tab_id} from '${target.title ?? ''}' to '${title}'`,
          });
        }

        // Deleting a tab takes its nested tabs with it, so only top-level tabs count as survivors
        const rootTabs = tabs.filter((t) => (t.nestingLevel ?? 0) === 0);
        if ((target.nestingLevel ?? 0) === 0 && rootTabs.length <= 1) {
          throw new Error(`'${tab_id}' is the only top-level tab, and a document cannot be left without one. Delete its content with delete_content instead.`);
        }
        await docsBatchUpdate(document_id, accessToken, [{ deleteTab: { tabId: tab_id } }]);
        return toolResult({ id: document_id, action, message: `Deleted tab ${tab_id} ('${target.title ?? ''}') and its nested tabs` });
      } catch (err) {
        return toolErrorResponse(err);
      }
    })),
  },
};
