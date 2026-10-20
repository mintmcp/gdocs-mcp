import { z } from 'zod';
import { insertionPoint, validateInsertIndex } from "../lib/docRequests.js";
import { IMAGE_FETCH_REFUSAL, IMAGE_LIMITS_NOTE, MAX_IMAGE_URL_LENGTH, imageObjectSize, imageSourceUrl } from "../lib/images.js";
import { collectInlineObjects, findInlineObjectTab } from "../lib/tabs.js";
import {
  docsBatchUpdate,
  documentIdField,
  doneMessage,
  editTool,
  fetchResolvedTab,
  suggestField,
  tabIdField,
  toolResult,
  withRefusalHint,
} from './shared.js';

const MAX_IMAGE_DIMENSION_PT = 2000;

const explainImageRefusal = withRefusalHint(
  (message) => IMAGE_FETCH_REFUSAL.test(message.replace(/insertInlineImage|replaceImage/g, '')),
  IMAGE_LIMITS_NOTE,
);

export const imageTools = {
  insert_image: editTool({
    description: `Insert an image from a URL into a Google Doc as an inline image, at \`index\` or at the end of the document (or tab) when index is omitted. ${IMAGE_LIMITS_NOTE} A Google Drive share link works only when the file is shared with "anyone with the link"; it is turned into a download URL here. width_pt and height_pt size the image in points (72 per inch; a letter page body is about 468pt wide); set one to keep the aspect ratio, neither for the image's natural size. Google refuses images inside footnotes. The reply carries the object id, which replace_image and get_document structure (inlineImage elements) use. Native Google Docs only.`,
    outputSchema: {
      id: z.string(),
      object_id: z.string().optional().describe('Id of the new inline image, for replace_image'),
      message: z.string(),
    },
    schema: {
      document_id: documentIdField,
      image_url: z.string().max(MAX_IMAGE_URL_LENGTH * 2).describe('https URL of a PNG, JPEG or GIF that Google can fetch without logging in'),
      index: z.number().int().optional().describe('Index to insert the image at; 1 is the top of the body. Omit to append it at the end.'),
      tab_id: tabIdField,
      width_pt: z.number().positive().max(MAX_IMAGE_DIMENSION_PT).optional().describe('Width in points'),
      height_pt: z.number().positive().max(MAX_IMAGE_DIMENSION_PT).optional().describe('Height in points'),
      suggest: suggestField,
    },
    run: async ({ document_id, image_url, index, tab_id, width_pt, height_pt, suggest }, { accessToken }) => {
      if (index !== undefined) validateInsertIndex(index);
      const uri = imageSourceUrl(image_url);

      const reply = await docsBatchUpdate(document_id, accessToken, [
        { insertInlineImage: { uri, ...insertionPoint(index, tab_id), ...imageObjectSize(width_pt, height_pt) } },
      ], { suggest }).catch(explainImageRefusal);

      const objectId: string | undefined = reply?.replies?.[0]?.insertInlineImage?.objectId;
      return toolResult({
        id: document_id,
        ...(objectId ? { object_id: objectId } : {}),
        message: doneMessage(index === undefined ? 'Image appended' : `Image inserted at index ${index}`, suggest),
      });
    },
  }),

  replace_image: editTool({
    description: `Replace an existing inline image in a Google Doc with one fetched from a URL, keeping the image's position and size (the new image is center-cropped to fit). Find the object id in get_document with include_structure=true (inlineImage elements carry inlineObjectId). ${IMAGE_LIMITS_NOTE} Native Google Docs only.`,
    outputSchema: {
      id: z.string(),
      object_id: z.string(),
      message: z.string(),
    },
    schema: {
      document_id: documentIdField,
      object_id: z.string().min(1).describe('inlineObjectId of the image to replace, from get_document structure'),
      image_url: z.string().max(MAX_IMAGE_URL_LENGTH * 2).describe('https URL of the replacement PNG, JPEG or GIF'),
      tab_id: z.string().min(1).optional().describe('Tab holding the image, when the document has tabs. Omit to search every tab.'),
    },
    run: async ({ document_id, object_id, image_url, tab_id }, { accessToken }) => {
      const uri = imageSourceUrl(image_url);

      const { doc } = await fetchResolvedTab(document_id, accessToken, tab_id);
      const inlineObjects = collectInlineObjects(doc, tab_id);
      if (!inlineObjects[object_id]) {
        const known = Object.keys(inlineObjects);
        throw new Error(
          `No inline image with id '${object_id}'${tab_id ? ` in tab '${tab_id}'` : ''}. ` +
          (known.length ? `This document's images are: ${known.slice(0, 20).join(', ')}${known.length > 20 ? ` and ${known.length - 20} more` : ''}.` : 'This document has no inline images; use insert_image.')
        );
      }

      // Without tab_id the image may sit in any tab, and Google resolves the id within one tab
      const tabId = tab_id ?? findInlineObjectTab(doc, object_id);
      await docsBatchUpdate(document_id, accessToken, [
        { replaceImage: { imageObjectId: object_id, uri, imageReplaceMethod: 'CENTER_CROP', ...(tabId ? { tabId } : {}) } },
      ]).catch(explainImageRefusal);

      return toolResult({
        id: document_id,
        object_id,
        message: `Replaced image ${object_id}`,
      });
    },
  }),
};
