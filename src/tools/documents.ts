import { z } from 'zod';
import {
  MARKDOWN_UPLOAD_MIME,
  MAX_MARKDOWN_BYTES,
  MULTIPART_UPLOAD_URL,
  buildMultipartUpload,
  isMarkdownSource,
  markdownByteLength,
} from "../lib/markdownImport.js";
import { fetchWordMagic, NATIVE_DOC_MIME } from "../lib/word.js";
import { isWordMime, sniffWordFormat } from "../lib/wordText.js";
import { GOOGLE_DRIVE_API, makeDriveRequest } from '../lib/google.js';
import { log, errorFields } from "../lib/log.js";
import { SCOPES } from "../scopes.js";
import {
  defineTool,
  type DriveFile,
  isConversionRefusal,
  toolResult,
} from './shared.js';

const docUrl = (id: string) => `https://docs.google.com/document/d/${id}/edit`;

function assertNativeDoc(file: DriveFile, describeSource: string, strayUrl: string): void {
  if (file.mimeType === NATIVE_DOC_MIME) return;
  throw new Error(
    `Drive stored ${describeSource} as ${file.mimeType || 'an unreported type'} instead of ` +
    `converting it to a Google Doc. The file is at ${strayUrl} — delete it or convert it ` +
    `manually rather than creating another.`
  );
}

// Drive converts in the same call that creates the file, so a body is never half-written
async function importDoc(
  metadata: object,
  body: string,
  format: 'markdown' | 'plain',
  accessToken: string
): Promise<DriveFile> {
  const upload = buildMultipartUpload(metadata, body, format === 'markdown' ? MARKDOWN_UPLOAD_MIME : 'text/plain; charset=UTF-8');
  let file: DriveFile;
  try {
    file = await makeDriveRequest(MULTIPART_UPLOAD_URL, accessToken, {
      method: 'POST',
      headers: { 'Content-Type': upload.contentType },
      body: upload.body,
    });
  } catch (err) {
    log("error", "document_import_failed", { format, ...errorFields(err) });
    if (format !== 'markdown' || !isConversionRefusal(err)) throw err;
    throw new Error(
      `Drive could not convert this markdown into a Google Doc. Retry with body_format "plain" ` +
      `to store the text as written, or simplify the markdown.`
    );
  }

  assertNativeDoc(file, 'the upload', `https://drive.google.com/file/d/${file.id}/view`);
  return file;
}

export const documentTools = {
  create_document: defineTool(SCOPES.DOCUMENTS, {
    description:
      'Create a new Google Doc with optional initial content. The body is treated as markdown ' +
      'by default and imported as real Google Docs formatting — headings, lists, links, bold, ' +
      'italic, inline code, horizontal rules and tables all become native styles, so there is ' +
      'no need to follow up with styling tools. Pass body_format "plain" to keep the text ' +
      'verbatim instead. Optionally place it in a specific folder (including shared drive folders).',
    outputSchema: {
      id: z.string(),
      title: z.string(),
      webViewLink: z.string(),
      message: z.string(),
    },
    schema: {
      title: z.string().describe('Title for the new document'),
      body: z.string().optional().describe(
        'Optional initial content. Markdown unless body_format is "plain". As markdown, ' +
        'separate paragraphs with a blank line — single newlines are joined into one paragraph.'
      ),
      body_format: z.enum(['markdown', 'plain']).optional().describe(
        'How to interpret body. Defaults to "markdown", which reflows the text and consumes ' +
        'markdown punctuation such as *, _ and #. Use "plain" whenever the literal characters ' +
        'and line breaks must be preserved exactly, such as pasted code or a document about ' +
        'markdown syntax.'
      ),
      parent_folder_id: z.string().optional().describe('ID of the folder to create the document in (supports shared drive folders)'),
    },
    run: async ({ title, body, body_format, parent_folder_id }, { accessToken }) => {
      const fileMetadata: any = {
        name: title,
        mimeType: NATIVE_DOC_MIME,
      };
      if (parent_folder_id) {
        fileMetadata.parents = [parent_folder_id];
      }

      const format = body_format ?? 'markdown';
      const bytes = body ? markdownByteLength(body) : 0;

      if (bytes > MAX_MARKDOWN_BYTES) {
        throw new Error(
          `Body is ${bytes} bytes, above the ${MAX_MARKDOWN_BYTES} limit for create_document, ` +
          `which applies to both body formats. Split it across documents, or upload the file to ` +
          `Drive and call convert_to_google_doc.`
        );
      }

      const file: DriveFile = body
        ? await importDoc(fileMetadata, body, format, accessToken)
        : await makeDriveRequest(`/files?supportsAllDrives=true`, accessToken, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify(fileMetadata),
        });

      return toolResult({
        id: file.id,
        title: file.name ?? title,
        webViewLink: docUrl(file.id),
        message: body && format === 'markdown'
          ? 'Document created with markdown converted to native Google Docs formatting.'
          : 'Document created successfully',
      });
    },
  }),

  convert_to_google_doc: defineTool(SCOPES.DRIVE_FILE, {
    description:
      'Convert an uploaded Word (.doc or .docx) or Markdown (.md) file already in Drive into a ' +
      'NEW, editable native Google Doc. The original file is left untouched. Drive performs the ' +
      'conversion, so formatting, tables and images are preserved. ' +
      'For Word, use this when the user wants to edit a file that search_documents or ' +
      'get_document reported with a Word mimeType. ' +
      'For Markdown, note that search_documents does not list .md files, so the user must supply ' +
      'the Drive file id (from a Drive link or the Drive connector); prefer this over re-typing a ' +
      'large markdown file through create_document, since Drive reads the bytes directly and ' +
      'nothing is paraphrased or truncated.',
    outputSchema: {
      id: z.string().describe('ID of the new native Google Doc'),
      name: z.string(),
      webViewLink: z.string(),
      sourceId: z.string().describe('The source file this was converted from, unchanged'),
      message: z.string(),
    },
    schema: {
      file_id: z.string().describe('Drive file ID of the .doc, .docx or .md file to convert'),
      name: z.string().optional().describe('Name for the new Doc (defaults to the original name)'),
    },
    run: async ({ file_id, name }, { accessToken }) => {
      const meta = await makeDriveRequest(
        `/files/${encodeURIComponent(file_id)}?fields=name,mimeType,size,webViewLink&supportsAllDrives=true`,
        accessToken
      );
      const sourceMime: string = meta.mimeType || '';

      if (sourceMime === NATIVE_DOC_MIME) {
        throw new Error(
          `'${meta.name}' is already a native Google Doc and is editable as-is. Nothing to convert.`
        );
      }
      const fromMarkdown = isMarkdownSource(sourceMime);
      if (!fromMarkdown && !isWordMime(sourceMime)) {
        const namedMarkdown = /\.(md|markdown|mdown|mkd)$/i.test(meta.name || '');
        throw new Error(
          namedMarkdown
            ? `'${meta.name}' is stored as ${sourceMime}, not text/markdown, so Drive would copy it ` +
              `literally instead of applying markdown formatting. Re-upload it as markdown, or pass ` +
              `its text to create_document instead.`
            : `'${meta.name}' is not a Word or Markdown file (${sourceMime}), so it cannot be converted to a Google Doc.`
        );
      }
      if (fromMarkdown) {
        const sourceBytes = Number(meta.size);
        if (meta.size == null || meta.size === '' || !Number.isFinite(sourceBytes)) {
          throw new Error(
            `Drive did not report a size for '${meta.name}', so it cannot be checked against the ` +
            `${MAX_MARKDOWN_BYTES} byte limit for markdown conversion. Open it directly: ${meta.webViewLink}`
          );
        }
        if (sourceBytes > MAX_MARKDOWN_BYTES) {
          throw new Error(
            `'${meta.name}' is ${sourceBytes} bytes, above the ${MAX_MARKDOWN_BYTES} limit for markdown conversion.`
          );
        }
      }
      if (!fromMarkdown && !sniffWordFormat(await fetchWordMagic(file_id, accessToken, GOOGLE_DRIVE_API))) {
        throw new Error(
          `'${meta.name}' is named like a Word file but its contents are not a Word ` +
          `document at all, so there is nothing to convert. Open it directly: ${meta.webViewLink}`
        );
      }

      let result: DriveFile;
      try {
        result = await makeDriveRequest(
          `/files/${encodeURIComponent(file_id)}/copy?supportsAllDrives=true&fields=id,name,mimeType,webViewLink`,
          accessToken,
          {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
              mimeType: NATIVE_DOC_MIME,
              ...(name ? { name } : {}),
            }),
          }
        );
      } catch (err) {
        log("error", "markdown_convert_failed", { ...errorFields(err), mime: sourceMime.split(";")[0].trim() });
        // Only a refusal is about the file; anything else must not send the user off re-uploading it
        if (!isConversionRefusal(err)) throw err;
        throw new Error(
          `Drive could not convert '${meta.name}' into a Google Doc. ` +
          (fromMarkdown
            ? `The file may not be text Drive can read as markdown. `
            : `It may be a different Office format stored under a Word name. `) +
          `Open it directly: ${meta.webViewLink}`
        );
      }

      assertNativeDoc(result, `'${meta.name}'`, `https://drive.google.com/file/d/${result.id}/view`);

      return toolResult({
        id: result.id,
        name: result.name,
        webViewLink: result.webViewLink || docUrl(result.id),
        sourceId: file_id,
        message:
          `Converted to a new native Google Doc, which is fully editable. ` +
          `The original '${meta.name}' is unchanged.`,
      });
    },
  }),
};
