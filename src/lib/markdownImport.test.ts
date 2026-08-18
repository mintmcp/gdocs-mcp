/**
 * Markdown source detection and the multipart envelope Drive needs to import
 * one. The conversion itself is Drive's, so what matters here is that we hand
 * it a well-formed request and never let document text escape its part.
 */

import { describe, it, expect } from 'vitest';

const DOCX_MIME = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
import {
  MARKDOWN_MIME,
  MARKDOWN_UPLOAD_MIME,
  isMarkdownSource,
  markdownByteLength,
  buildMultipartUpload,
} from './markdownImport.js';

describe('markdown source detection', () => {
  it('accepts the markdown mime types Drive stores uploads under', () => {
    expect(isMarkdownSource('text/markdown')).toBe(true);
    expect(isMarkdownSource('text/x-markdown')).toBe(true);
  });

  it('ignores charset parameters and casing on the mime type', () => {
    expect(isMarkdownSource('text/markdown; charset=utf-8')).toBe(true);
    expect(isMarkdownSource('TEXT/MARKDOWN')).toBe(true);
  });

  it('refuses text/plain, which Drive would copy literally however it is named', () => {
    expect(isMarkdownSource('text/plain')).toBe(false);
  });

  it('refuses Word and native Docs', () => {
    expect(isMarkdownSource(DOCX_MIME)).toBe(false);
    expect(isMarkdownSource('application/vnd.google-apps.document')).toBe(false);
  });
});

describe('byte measurement', () => {
  it('counts UTF-8 bytes, not UTF-16 code units, so multibyte text is not undercounted', () => {
    expect(markdownByteLength('abc')).toBe(3);
    expect(markdownByteLength('東京')).toBe(6);
    expect(markdownByteLength('🎉')).toBe(4);
  });
});

describe('multipart upload envelope', () => {
  const metadata = { name: 'Brief', mimeType: 'application/vnd.google-apps.document' };

  it('carries the metadata and the content in separate parts', () => {
    const { body } = buildMultipartUpload(metadata, '# Title', MARKDOWN_MIME);

    expect(body).toContain('Content-Type: application/json; charset=UTF-8');
    expect(body).toContain(JSON.stringify(metadata));
    expect(body).toContain(`Content-Type: ${MARKDOWN_MIME}`);
    expect(body).toContain('# Title');
  });

  it('reports a content type whose boundary matches the body', () => {
    const { body, contentType } = buildMultipartUpload(metadata, 'x', MARKDOWN_MIME);
    const boundary = contentType.match(/boundary=(.+)$/)?.[1];

    expect(boundary).toBeTruthy();
    expect(body.startsWith(`--${boundary}\r\n`)).toBe(true);
    expect(body.endsWith(`\r\n--${boundary}--`)).toBe(true);
  });

  it('uses a fresh boundary per call so one document cannot fix another', () => {
    const a = buildMultipartUpload(metadata, 'x', MARKDOWN_MIME);
    const b = buildMultipartUpload(metadata, 'x', MARKDOWN_MIME);
    expect(a.contentType).not.toBe(b.contentType);
  });

  it('keeps content that mimics a boundary inside its own part', () => {
    const { body, contentType } = buildMultipartUpload(metadata, '--not-a-boundary--', MARKDOWN_MIME);
    const boundary = contentType.match(/boundary=(.+)$/)?.[1] as string;

    expect(body.split(`--${boundary}`).length - 1).toBe(3);
  });

  it('declares the charset so accented and CJK markdown is not mis-decoded on import', () => {
    const { body } = buildMultipartUpload(metadata, 'café — 東京 🎉', MARKDOWN_UPLOAD_MIME);

    expect(body).toContain('Content-Type: text/markdown; charset=UTF-8');
    expect(body).toContain('café — 東京 🎉');
  });
});
