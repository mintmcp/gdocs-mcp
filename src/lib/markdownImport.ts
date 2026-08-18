/**
 * Markdown -> native Google Doc, by way of Drive's own importer.
 *
 * Drive converts an uploaded source file into a Doc when the metadata names the
 * target type and the media part names the source type, so nothing here parses
 * markdown itself.
 */

export const MARKDOWN_MIME = 'text/markdown';

/** Drive rejects oversized imports; measured in bytes because that is what goes on the wire. */
export const MAX_MARKDOWN_BYTES = 400_000;

/** text/* with no charset is ambiguous, and fetch serializes the body as UTF-8. */
export const MARKDOWN_UPLOAD_MIME = `${MARKDOWN_MIME}; charset=UTF-8`;

export const MULTIPART_UPLOAD_URL =
  'https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&supportsAllDrives=true' +
  '&fields=id,name,mimeType,webViewLink';

const MARKDOWN_MIMES = new Set([MARKDOWN_MIME, 'text/x-markdown']);

export function markdownByteLength(content: string): number {
  return new TextEncoder().encode(content).byteLength;
}

/**
 * Mime only: a .md name over text/plain is still stored as plain text, and Drive
 * converts it literally, so accepting it would promise formatting we cannot deliver.
 */
export function isMarkdownSource(mimeType: string): boolean {
  return MARKDOWN_MIMES.has(mimeType.split(';')[0].trim().toLowerCase());
}

/**
 * A random boundary keeps document text from being read as framing: the caller
 * controls the content but cannot predict the delimiter.
 */
export function buildMultipartUpload(
  metadata: object,
  content: string,
  contentMime: string
): { body: string; contentType: string } {
  const boundary = `mdimport-${crypto.randomUUID()}`;
  const body =
    `--${boundary}\r\n` +
    `Content-Type: application/json; charset=UTF-8\r\n\r\n` +
    `${JSON.stringify(metadata)}\r\n` +
    `--${boundary}\r\n` +
    `Content-Type: ${contentMime}\r\n\r\n` +
    `${content}\r\n` +
    `--${boundary}--`;

  return { body, contentType: `multipart/related; boundary=${boundary}` };
}
