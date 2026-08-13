import { describe, it, expect } from 'vitest';
import {
  docsOnly,
  wordErrorMessage,
  wordWriteRefusal,
  NATIVE_DOC_MIME,
} from './word.js';
import {
  WordEncryptedError,
  WordInvalidError,
  WordTimeoutError,
  WordTooLargeError,
  DOCX_MIME,
  LEGACY_DOC_MIME,
} from './wordText.js';

describe('wordErrorMessage', () => {
  it('explains each classified failure and always offers the link', () => {
    const cases: Array<[Error, RegExp]> = [
      [new WordEncryptedError('x'), /password-protected or encrypted/],
      [new WordTooLargeError('too big'), /too large to read safely/],
      [new WordTimeoutError('x'), /took too long/],
      [new WordInvalidError('bad zip'), /not a readable Word file/],
    ];
    for (const [err, pattern] of cases) {
      const msg = wordErrorMessage(err, 'Report.docx', 'https://x/1');
      expect(msg).toMatch(pattern);
      expect(msg).toContain('https://x/1');
    }
  });

  it('returns null for anything unclassified, so the caller rethrows unchanged', () => {
    expect(wordErrorMessage(new Error('socket hang up'), 'a', 'b')).toBeNull();
    expect(wordErrorMessage(undefined, 'a', 'b')).toBeNull();
  });
});

describe('wordWriteRefusal', () => {
  it('names the right extension and points at convert_to_google_doc', () => {
    expect(wordWriteRefusal('A.docx', DOCX_MIME, 'https://x/1')).toContain('(.docx)');
    expect(wordWriteRefusal('A.doc', LEGACY_DOC_MIME, 'https://x/1')).toContain('(.doc)');
    expect(wordWriteRefusal('A.docx', DOCX_MIME, 'https://x/1')).toContain('convert_to_google_doc');
  });
});

describe('docsOnly', () => {
  const meta = (mimeType: string) => async () => ({
    name: 'File', mimeType, webViewLink: 'https://x/1',
  });
  const ok = async () => 'handler ran';
  const ctx = { accessToken: 'tok' };

  it('refuses a .docx before the write reaches the Docs API', async () => {
    const guarded = docsOnly(ok, meta(DOCX_MIME) as any);
    await expect(guarded({ document_id: 'd1' }, ctx)).rejects.toThrow(/convert_to_google_doc/);
  });

  it('refuses a legacy .doc too', async () => {
    const guarded = docsOnly(ok, meta(LEGACY_DOC_MIME) as any);
    await expect(guarded({ document_id: 'd1' }, ctx)).rejects.toThrow(/\(\.doc\)/);
  });

  it('lets a native Doc through', async () => {
    const guarded = docsOnly(ok, meta(NATIVE_DOC_MIME) as any);
    await expect(guarded({ document_id: 'd1' }, ctx)).resolves.toBe('handler ran');
  });

  it('skips the lookup entirely when there is no document_id to check', async () => {
    let called = false;
    const guarded = docsOnly(ok, (async () => { called = true; return {}; }) as any);
    await expect(guarded({}, ctx)).resolves.toBe('handler ran');
    expect(called).toBe(false);
  });

  it('does not swallow a metadata lookup failure', async () => {
    const boom = async () => { throw new Error('drive 500'); };
    const guarded = docsOnly(ok, boom as any);
    await expect(guarded({ document_id: 'd1' }, ctx)).rejects.toThrow('drive 500');
  });
});
