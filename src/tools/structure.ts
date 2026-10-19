/** Document structure tools: insert_break, set_header_footer and insert_footnote */

import { z } from 'zod';
import { withGoogleAuth as requirePermissionSecure } from "../auth.js";
import { buildWriteControl, validateInsertIndex } from "../lib/docs-structure.js";
import { docLocation, insertionPoint } from "../lib/tabs.js";
import { HEADER_FOOTER, replaceSegmentTextRequests, type HeaderFooterKind } from "../lib/segments.js";
import {
  DOCUMENTS_SCOPE,
  SUGGEST_DESCRIPTION,
  TAB_ID_DESCRIPTION,
  docsBatchUpdate,
  docsOnly,
  doneMessage,
  fetchResolvedTab,
  partialFailure,
  toolErrorResponse,
  toolResult,
} from './shared.js';

type Placement = ReturnType<typeof insertionPoint>;

const BREAK_REQUESTS: Record<string, (at: Placement) => any> = {
  page: (at) => ({ insertPageBreak: at }),
  section_continuous: (at) => ({ insertSectionBreak: { sectionType: 'CONTINUOUS', ...at } }),
  section_next_page: (at) => ({ insertSectionBreak: { sectionType: 'NEXT_PAGE', ...at } }),
};

export const structureTools = {
  insert_break: {
    description: 'Insert a page break or a section break into a Google Doc, at `index` or at the end of the document (or tab) when index is omitted. Use get_document with include_structure=true to find the index; the break goes before the content at that index, and inside a paragraph it splits the paragraph. "page" starts a new page. "section_next_page" starts a new section on a new page and "section_continuous" starts one on the same page; sections can carry their own headers and footers. Google refuses breaks inside a table, header, footer or footnote. Native Google Docs only.',
    outputSchema: {
      id: z.string(),
      message: z.string(),
    },
    schema: {
      document_id: z.string().describe('Google Doc ID'),
      type: z.enum(['page', 'section_continuous', 'section_next_page']).describe('The kind of break'),
      index: z.number().int().optional().describe('Index to insert the break at; 1 is the top of the body. Omit to append it at the end.'),
      tab_id: z.string().min(1).optional().describe(TAB_ID_DESCRIPTION),
      suggest: z.boolean().optional().describe(SUGGEST_DESCRIPTION),
    },
    handler: requirePermissionSecure(DOCUMENTS_SCOPE, docsOnly(async ({ document_id, type, index, tab_id, suggest }: any, context: any) => {
      try {
        const { accessToken } = context;
        if (index !== undefined) validateInsertIndex(index);

        await docsBatchUpdate(document_id, accessToken, [
          BREAK_REQUESTS[type](insertionPoint(index, tab_id)),
        ], buildWriteControl(undefined, suggest));

        const what = type === 'page' ? 'Page break' : 'Section break';
        return toolResult({
          id: document_id,
          message: doneMessage(index === undefined ? `${what} appended` : `${what} inserted at index ${index}`, suggest),
        });
      } catch (err) {
        return toolErrorResponse(err);
      }
    })),
  },

  set_header_footer: {
    description: 'Set, replace or remove the default header or footer of a Google Doc (or of one tab with tab_id). Pass `text` to make the header or footer read exactly that, creating it first when the document has none; pass an empty string to clear it; pass remove=true to delete it entirely. The text is plain: style it afterwards with update_text_style or update_paragraph_style and the `segment_id` this returns (indices inside a header start at 0). get_document with include_structure=true lists existing headers and footers with their ids and text. Only the DEFAULT header and footer are managed here; first-page, even-page and per-section ones are listed by get_document and can be edited with segment_id, but not created. Native Google Docs only.',
    outputSchema: {
      id: z.string(),
      kind: z.enum(['header', 'footer']),
      segment_id: z.string().optional().describe('Id of the header or footer, for segment_id on the editing tools. Absent after a removal.'),
      created: z.boolean().optional().describe('True when the header or footer did not exist and was created'),
      removed: z.boolean().optional(),
      message: z.string(),
    },
    schema: {
      document_id: z.string().describe('Google Doc ID'),
      kind: z.enum(['header', 'footer']).describe('Which one to change'),
      text: z.string().max(10_000).optional().describe('The new text; replaces whatever is there. Empty string clears it.'),
      remove: z.boolean().optional().describe('Delete the header or footer instead of writing text'),
      tab_id: z.string().min(1).optional().describe(TAB_ID_DESCRIPTION),
    },
    handler: requirePermissionSecure(DOCUMENTS_SCOPE, docsOnly(async ({ document_id, kind, text, remove, tab_id }: any, context: any) => {
      try {
        const { accessToken } = context;
        const which = kind as HeaderFooterKind;
        if (text === undefined && !remove) {
          throw new Error(`Pass text or remove: text writes the ${which} (empty text clears it) and remove=true deletes it.`);
        }

        const spec = HEADER_FOOTER[which];
        const { resolved } = await fetchResolvedTab(document_id, accessToken, tab_id);
        const existingId: string | undefined = resolved.documentStyle?.[spec.defaultIdField];
        const tabId = resolved.tabId;
        const location = tabId ? ` in tab '${tabId}'` : '';

        if (remove) {
          if (!existingId) throw new Error(`There is no ${which} to remove${location}.`);
          await docsBatchUpdate(document_id, accessToken, [
            { [spec.deleteRequest]: { [spec.idKey]: existingId, ...(tabId ? { tabId } : {}) } },
          ]);
          return toolResult({ id: document_id, kind: which, removed: true, message: `Removed the ${which}${location}` });
        }

        if (existingId) {
          const requests = replaceSegmentTextRequests(spec.segments(resolved)[existingId], existingId, text, tabId);
          if (requests.length) await docsBatchUpdate(document_id, accessToken, requests);
          return toolResult({
            id: document_id,
            kind: which,
            segment_id: existingId,
            created: false,
            message: text ? `Replaced the ${which} text${location}` : `Cleared the ${which}${location}`,
          });
        }

        // Google 400s on a second DEFAULT header, so the id must be read from
        // the create reply rather than guessed.
        const reply = await docsBatchUpdate(document_id, accessToken, [{
          [spec.createRequest]: { type: 'DEFAULT', ...(tabId ? { sectionBreakLocation: docLocation(0, tabId) } : {}) },
        }]);
        const createdId: string | undefined = reply?.replies?.[0]?.[spec.createRequest]?.[spec.idKey];
        if (!createdId) {
          throw new Error(`Google created the ${which} but did not return its id. Read the document with get_document include_structure=true to find it, then write with insert_text and segment_id.`);
        }
        if (text) {
          try {
            await docsBatchUpdate(document_id, accessToken, [
              { insertText: { location: docLocation(0, tabId, createdId), text } },
            ]);
          } catch (err) {
            return partialFailure(err, `The ${which} was created but writing its text failed`, { kind: which, segment_id: createdId, created: true });
          }
        }
        return toolResult({
          id: document_id,
          kind: which,
          segment_id: createdId,
          created: true,
          message: `Created the ${which}${location}${text ? ' and wrote its text' : ''}`,
        });
      } catch (err) {
        return toolErrorResponse(err);
      }
    })),
  },

  insert_footnote: {
    description: 'Add a footnote to a Google Doc: a footnote reference goes in at `index` and `text` becomes the note at the bottom of the page. Use get_document with include_structure=true to find the index, typically the endIndex of the sentence minus one so the mark sits before the paragraph\'s newline. Google refuses footnotes inside a table, header, footer or another footnote. The reply carries the footnote id, which is a segment_id for the editing tools (indices inside it start at 0). If the result reports partialSuccess, the footnote exists but is empty: write its text with insert_text and segment_id instead of calling this again. If the call fails in any other way, read the document with get_document include_structure=true before retrying: a footnoteReference at the index means the footnote was created and only its text is missing. Native Google Docs only.',
    outputSchema: {
      id: z.string(),
      footnote_id: z.string(),
      message: z.string(),
    },
    schema: {
      document_id: z.string().describe('Google Doc ID'),
      index: z.number().int().min(1).describe('Body index for the footnote reference mark'),
      text: z.string().min(1).max(10_000).describe('The footnote text'),
      tab_id: z.string().min(1).optional().describe(TAB_ID_DESCRIPTION),
    },
    handler: requirePermissionSecure(DOCUMENTS_SCOPE, docsOnly(async ({ document_id, index, text, tab_id }: any, context: any) => {
      try {
        const { accessToken } = context;
        validateInsertIndex(index);

        const reply = await docsBatchUpdate(document_id, accessToken, [
          { createFootnote: { location: docLocation(index, tab_id) } },
        ]);
        const footnoteId: string | undefined = reply?.replies?.[0]?.createFootnote?.footnoteId;
        if (!footnoteId) {
          throw new Error('Google created the footnote but did not return its id. Read the document with get_document include_structure=true before retrying, or the footnote will be duplicated.');
        }

        try {
          await docsBatchUpdate(document_id, accessToken, [
            { insertText: { location: docLocation(0, tab_id, footnoteId), text } },
          ]);
        } catch (err) {
          return partialFailure(err, 'The footnote was created but writing its text failed', { footnote_id: footnoteId });
        }

        return toolResult({
          id: document_id,
          footnote_id: footnoteId,
          message: `Footnote inserted at index ${index}`,
        });
      } catch (err) {
        return toolErrorResponse(err);
      }
    })),
  },
};
