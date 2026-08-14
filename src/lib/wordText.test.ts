/**
 * Functional coverage against the checked-in corpus, parameterized over both
 * containers since they hold the same content. Hostile inputs: wordText.security.test.ts.
 */

import { describe, it, expect } from 'vitest';
import {
  extractWordText,
  sanitizeWordText,
  sniffWordFormat,
  isWordMime,
  WordInvalidError,
  DOCX_MIME,
  LEGACY_DOC_MIME,
  MAX_TEXT_CHARS,
} from './wordText.js';
import { fixture, buildDocx } from './__fixtures__/index.js';

describe.each([
  ['docx', 'sample.docx', 'large.docx'],
  ['legacy doc', 'sample.doc', 'large.doc'],
] as const)('extractWordText (%s)', (label, small, large) => {
  const format = label === 'docx' ? 'docx' : 'doc';

  it('extracts text in document order', async () => {
    const r = await extractWordText(fixture(small));
    expect(r.truncated).toBe(false);
    expect(r.text).toContain('CANARY-ALPHA-7731');
    expect(r.text).toContain('CANARY-OMEGA-9914');
    expect(r.text.indexOf('Section One')).toBeLessThan(r.text.indexOf('Section Three'));
  });

  it('reports the format the bytes say, and the char count it returned', async () => {
    const r = await extractWordText(fixture(small));
    expect(r.format).toBe(format);
    expect(r.chars).toBe(r.text.length);
  });

  it('keeps table cell values', async () => {
    const r = await extractWordText(fixture(small));
    for (const cell of ['Invoice number', 'INV-2026-0042', 'Amount', '1250.00 EUR']) {
      expect(r.text).toContain(cell);
    }
  });

  it('reads a whole 50-page document without dropping content', async () => {
    const r = await extractWordText(fixture(large));
    // Every page is stamped, so a gap or an early stop shows up as a miscount.
    expect(r.text.match(/PAGE-\d{3}-START-CANARY/g)).toHaveLength(50);
    expect(r.text).toContain('DOC-FINAL-CANARY-9999');
    expect(r.truncated).toBe(false);
  });

  it('caps output and flags it rather than returning unbounded text', async () => {
    const r = await extractWordText(fixture(large), { maxChars: 5_000 });
    expect(r.truncated).toBe(true);
    expect(r.text.length).toBeLessThanOrEqual(5_000);
    expect(r.chars).toBe(r.text.length);
    expect(r.text).toContain('DOC-START-CANARY-0000');
    expect(r.text).not.toContain('DOC-FINAL-CANARY-9999');
  });
});

describe('extractWordText (edge cases)', () => {
  it('keeps a header and footer that carry content, and labels them', async () => {
    const r = await extractWordText(fixture('header-content.docx'));
    expect(r.text).toMatch(/\[Header\]\nCONFIDENTIAL - Acme Corp/);
    expect(r.text).toMatch(/\[Footer\]\nContract ref ACME-2026-77/);
    // The header leads so a classification marking survives a truncated read.
    expect(r.text.indexOf('CONFIDENTIAL')).toBeLessThan(r.text.indexOf('BODY-CANARY'));
  });

  it('drops a header or footer that is only a page number', async () => {
    const r = await extractWordText(fixture('header-furniture.docx'));
    expect(r.text).toContain('BODY-CANARY');
    expect(r.text).not.toContain('[Header]');
    expect(r.text).not.toContain('[Footer]');
  });

  it('keeps a classification marking even when the body is truncated away', async () => {
    const r = await extractWordText(fixture('header-content.docx'), { maxChars: 30 });
    expect(r.truncated).toBe(true);
    expect(r.text).toContain('CONFIDENTIAL');
  });

  it('labels the streams it appends after the body', async () => {
    // They land after the body's closing line, so unlabelled they read as body prose.
    const r = await extractWordText(fixture('textbox.docx'));
    expect(r.text).toMatch(/\[Text boxes\]\nINSIDE-TEXTBOX/);
  });

  it('includes text box content, which lives outside the body stream', async () => {
    const r = await extractWordText(fixture('textbox.docx'));
    for (const marker of [
      'BEFORE-TEXTBOX',
      'AFTER-TEXTBOX-SAME-PARAGRAPH',
      'INSIDE-TEXTBOX',
      'NEXT-PARAGRAPH',
    ]) {
      expect(r.text).toContain(marker);
    }
  });

  it('decodes XML entities instead of leaking markup into the text', async () => {
    const r = await extractWordText(buildDocx(['Smith &amp; Co &lt;tag&gt; &#65;&#x42;']));
    expect(r.text).toBe('Smith & Co <tag> AB');
  });

  it('throws WordInvalidError on bytes that are neither container', async () => {
    await expect(extractWordText(new Uint8Array([1, 2, 3, 4, 5, 6, 7, 8])))
      .rejects.toThrow(WordInvalidError);
  });
});

describe('sniffWordFormat', () => {
  it('routes on the container bytes, which is what picks the parser', () => {
    expect(sniffWordFormat(fixture('sample.docx'))).toBe('docx');
    expect(sniffWordFormat(fixture('sample.doc'))).toBe('doc');
    expect(sniffWordFormat(new Uint8Array([0x25, 0x50, 0x44, 0x46]))).toBeNull();
    expect(sniffWordFormat(new Uint8Array([0x50, 0x4b]))).toBeNull();
    expect(sniffWordFormat(new Uint8Array())).toBeNull();
  });
});

describe('sanitizeWordText', () => {
  it('normalizes Word paragraph marks to newlines', () => {
    expect(sanitizeWordText('a\r\nb\rc')).toBe('a\nb\nc');
  });

  it('drops the NUL word-extractor emits for deleted ranges, and other C0 controls', () => {
    expect(sanitizeWordText('a\u0000b\u0007c\u001fd\u007fe')).toBe('abcde');
  });

  it('keeps tab and newline, which carry real layout', () => {
    expect(sanitizeWordText('a\tb\nc')).toBe('a\tb\nc');
  });

  it('drops bidi overrides, which can render text as something it is not', () => {
    expect(sanitizeWordText('a\u202ab\u202ec\u2066d\u2069e')).toBe('abcde');
  });

  it('drops lone surrogates while keeping whole pairs', () => {
    expect(sanitizeWordText('a\ud800b\udfffc')).toBe('abc');
    expect(sanitizeWordText('ok 😀')).toBe('ok 😀');
  });
});

describe('isWordMime', () => {
  it('accepts both Word containers and nothing else', () => {
    expect(isWordMime(DOCX_MIME)).toBe(true);
    expect(isWordMime(LEGACY_DOC_MIME)).toBe(true);
    expect(isWordMime('application/vnd.google-apps.document')).toBe(false);
    expect(isWordMime('application/pdf')).toBe(false);
    expect(isWordMime('')).toBe(false);
  });

  it('exposes a default cap', () => {
    expect(MAX_TEXT_CHARS).toBeGreaterThan(0);
  });
});
