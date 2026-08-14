/**
 * Adversarial tests for the container preflight: any Drive collaborator controls
 * these bytes, so every bound here is a live DoS boundary. The bombs are
 * hand-built rather than mocked — a size check that only ever sees honest
 * headers proves nothing.
 */

import { describe, it, expect } from 'vitest';
import { zipSync, strToU8 } from 'fflate';
import {
  extractWordText,
  sniffWordFormat,
  WordInvalidError,
  WordEncryptedError,
  WordTooLargeError,
  MAX_ZIP_ENTRIES,
  MAX_ENTRY_BYTES,
  MAX_TOTAL_UNPACKED_BYTES,
  MAX_WORD_BYTES,
} from './wordText.js';
import { buildDocx, buildOle, documentXml, fixture } from './__fixtures__/index.js';

const compressible = (bytes: number) => new Uint8Array(bytes);

const CD_SIG = 0x02014b50;

/** Rewrites the uncompressed size a single-entry zip declares, in both headers. */
function lieAboutSize(zip: Uint8Array, declared: number): Uint8Array {
  const out = zip.slice();
  const view = new DataView(out.buffer);
  view.setUint32(22, declared, true);
  for (let at = 0; at + 4 <= out.byteLength; at++) {
    if (view.getUint32(at, true) === CD_SIG) {
      view.setUint32(at + 24, declared, true);
      break;
    }
  }
  return out;
}

describe('container routing: bytes decide, not the mimeType', () => {
  it('picks the parser from the bytes, not from the mimeType', () => {
    expect(sniffWordFormat(buildOle(['WordDocument']))).toBe('doc');
    expect(sniffWordFormat(fixture('sample.docx'))).toBe('docx');
  });

  it('rejects bytes that are neither container', async () => {
    for (const bytes of [
      new Uint8Array([0x25, 0x50, 0x44, 0x46, 0x2d]),
      strToU8('<html><body>hi</body></html>'),
      new Uint8Array(0),
    ]) {
      await expect(extractWordText(bytes)).rejects.toThrow(WordInvalidError);
    }
  });

  it('fails an OLE file as a .doc rather than as a broken zip', async () => {
    await expect(extractWordText(buildOle(['NotWordAtAll'])))
      .rejects.toThrow(/damaged, or it is not really a Word document/);
  });
});

describe('container mismatch', () => {
  it('names the disagreement instead of letting the wrong parser fail obscurely', async () => {
    await expect(extractWordText(buildOle(['WordDocument']), { expectedFormat: 'docx' }))
      .rejects.toThrow(/contents are a legacy \.doc .* stored as a \.docx/);
  });

  it('says nothing when the container matches the mimeType', async () => {
    const r = await extractWordText(fixture('sample.docx'), { expectedFormat: 'docx' });
    expect(r.format).toBe('docx');
  });

  // Office wraps an encrypted .docx in OLE, so the mismatch must not claim the
  // user renamed a .doc when the file is simply password-protected.
  it('still reports an encrypted .docx as encrypted, not as a renamed .doc', async () => {
    await expect(
      extractWordText(buildOle(['EncryptedPackage', 'EncryptionInfo']), { expectedFormat: 'docx' })
    ).rejects.toThrow(WordEncryptedError);
  });
});

describe('encrypted containers', () => {
  it('classifies an OLE holding EncryptedPackage instead of failing obscurely', async () => {
    await expect(extractWordText(buildOle(['EncryptedPackage', 'EncryptionInfo'])))
      .rejects.toThrow(WordEncryptedError);
  });

  it('does not call a document encrypted just because its text says so', async () => {
    // .doc body text is UTF-16LE, so a raw byte scan would false-positive here.
    const named = buildOle(['WordDocument', 'EncryptedPackageNot']);
    await expect(extractWordText(named)).rejects.toThrow(WordInvalidError);
  });

  it('rejects a zip entry flagged as encrypted', async () => {
    const zip = buildDocx(['hi']);
    const view = new DataView(zip.buffer, zip.byteOffset, zip.byteLength);
    for (let at = 0; at + 4 <= zip.byteLength; at++) {
      if (view.getUint32(at, true) === CD_SIG) {
        view.setUint16(at + 8, view.getUint16(at + 8, true) | 0x1, true);
        break;
      }
    }
    await expect(extractWordText(zip)).rejects.toThrow(WordEncryptedError);
  });
});

describe('zip preflight: measured bytes, not declared', () => {
  it('rejects an archive with far too many entries', async () => {
    const parts: Record<string, Uint8Array> = {};
    for (let i = 0; i < 10_000; i++) parts[`e${i}.bin`] = strToU8('x');
    await expect(extractWordText(zipSync(parts))).rejects.toThrow(
      new RegExp(`above the ${MAX_ZIP_ENTRIES}-entry limit`)
    );
  });

  it('rejects a real bomb on MEASURED bytes, not on what it declares', async () => {
    // Honest headers, ~9MB of zeros in a few KB: only counting output catches it.
    const bomb = buildDocx(['hi'], { 'word/bomb.bin': compressible(MAX_ENTRY_BYTES + 1_000_000) });
    expect(bomb.byteLength).toBeLessThan(64 * 1024);
    await expect(extractWordText(bomb)).rejects.toThrow(WordTooLargeError);
  });

  it('still rejects a bomb whose declared size lies about a small payload', async () => {
    const zip = zipSync({ 'word/document.xml': compressible(MAX_ENTRY_BYTES + 1_000_000) });
    await expect(extractWordText(lieAboutSize(zip, 128)))
      .rejects.toThrow(/unpacks past the 128 bytes it declares/);
  });

  it('rejects a header that under-declares what unpacked', async () => {
    const zip = zipSync({ 'word/document.xml': strToU8('x'.repeat(1_000)) });
    await expect(extractWordText(lieAboutSize(zip, 500)))
      .rejects.toThrow(/unpacks past the 500 bytes it declares/);
  });

  // Over-declaring never trips the in-loop abort, so only the post-loop equality
  // check catches it — the one branch the other size tests cannot reach.
  it('rejects a header that over-declares what unpacked', async () => {
    const zip = zipSync({ 'word/document.xml': strToU8('x'.repeat(1_000)) });
    await expect(extractWordText(lieAboutSize(zip, 2_000)))
      .rejects.toThrow(/unpacks to 1000 bytes but declares 2000/);
  });

  it('rejects an honest oversize header before inflating anything', async () => {
    const zip = zipSync({ 'word/document.xml': strToU8('x'.repeat(1_000)) });
    await expect(extractWordText(lieAboutSize(zip, MAX_ENTRY_BYTES + 1)))
      .rejects.toThrow(/declares \d+ bytes, past the \d+ byte per-entry limit/);
  });

  // Whichever record is chosen, every entry the parser reads must be one preflight
  // measured. Verified yauzl resolves this file the same way, so it is refused.
  it('never parses entries the preflight did not measure', async () => {
    const zip = buildDocx(['hi']);
    const fake = new Uint8Array(22);
    const fakeView = new DataView(fake.buffer);
    fakeView.setUint32(0, 0x06054b50, true);
    fakeView.setUint16(10, MAX_ZIP_ENTRIES + 1_000, true);

    const crafted = new Uint8Array(zip.byteLength + fake.byteLength);
    crafted.set(zip);
    crafted.set(fake, zip.byteLength);

    const view = new DataView(crafted.buffer);
    for (let at = zip.byteLength - 22; at >= 0; at--) {
      if (view.getUint32(at, true) === 0x06054b50) {
        view.setUint16(at + 20, fake.byteLength, true);
        break;
      }
    }

    const error = await extractWordText(crafted).then(() => null, (e) => e as Error);
    expect(error).toBeInstanceOf(WordInvalidError);
  });

  it('rejects a total that only breaches the cap across entries', async () => {
    const per = MAX_ENTRY_BYTES - 1;
    const extra: Record<string, Uint8Array> = {};
    for (let i = 0; i < Math.ceil(MAX_TOTAL_UNPACKED_BYTES / per) + 1; i++) {
      extra[`word/pad${i}.bin`] = compressible(per);
    }
    await expect(extractWordText(buildDocx(['hi'], extra)))
      .rejects.toThrow(new RegExp(`${MAX_TOTAL_UNPACKED_BYTES} byte total limit`));
  });

  it('rejects a compression method Word never emits', async () => {
    const zip = buildDocx(['hi']);
    const view = new DataView(zip.buffer, zip.byteOffset, zip.byteLength);
    for (let at = 0; at + 4 <= zip.byteLength; at++) {
      if (view.getUint32(at, true) === CD_SIG) {
        view.setUint16(at + 10, 12, true);
        break;
      }
    }
    await expect(extractWordText(zip)).rejects.toThrow(/compression method 12/);
  });

  it('rejects a directory that points outside the file', async () => {
    const zip = buildDocx(['hi']);
    const view = new DataView(zip.buffer, zip.byteOffset, zip.byteLength);
    view.setUint32(zip.byteLength - 22 + 16, 0xfffffff0, true);
    await expect(extractWordText(zip)).rejects.toThrow(WordInvalidError);
  });

  it('lets an ordinary Word document through untouched', async () => {
    const r = await extractWordText(fixture('sample.docx'));
    expect(r.text).toContain('CANARY-ALPHA-7731');
  });
});

describe('DTD rejection', () => {
  const bomb =
    `<?xml version="1.0"?><!DOCTYPE lolz [<!ENTITY lol "lol">]>` +
    `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body/></w:document>`;

  it('rejects a DTD in the main document part', async () => {
    const zip = buildDocx([], { 'word/document.xml': strToU8(bomb) });
    await expect(extractWordText(zip)).rejects.toThrow(/declares an XML DTD/);
  });

  it('rejects an external entity hidden in a side part', async () => {
    const xxe = `<?xml version="1.0"?><!DOCTYPE r [<!ENTITY x SYSTEM "file:///etc/passwd">]><r>&x;</r>`;
    const zip = buildDocx(['hi'], { 'word/footnotes.xml': strToU8(xxe) });
    await expect(extractWordText(zip)).rejects.toThrow(/declares an XML DTD/);
  });

  it('rejects it in a .rels part too', async () => {
    const zip = buildDocx(['hi'], { 'word/_rels/document.xml.rels': strToU8(`<!DOCTYPE x><r/>`) });
    await expect(extractWordText(zip)).rejects.toThrow(/declares an XML DTD/);
  });

  it('sees a declaration split across two decompressed chunks', async () => {
    // The scanner carries a tail; without it a chunk boundary is a bypass.
    const padded = `<?xml version="1.0"?><!--${'p'.repeat(200_000)}--><!DOCTYPE x><r/>`;
    const zip = buildDocx(['hi'], { 'word/notes.xml': strToU8(padded) });
    await expect(extractWordText(zip)).rejects.toThrow(/declares an XML DTD/);
  });

  it('leaves a document that merely mentions the word alone', async () => {
    const r = await extractWordText(buildDocx(['We reject any &lt;!DOCTYPE in uploads.']));
    expect(r.text).toContain('<!DOCTYPE');
  });

  // word-extractor picks parts to XML-parse from [Content_Types].xml, not from the
  // entry name, so a name-gated scan is bypassable both of these ways.
  it('rejects a DTD in a part whose name does not look like XML', async () => {
    const remapped =
      `<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
      `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
      `<Default Extension="xml" ContentType="application/xml"/>` +
      `<Override PartName="/word/evil.bin" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>` +
      `</Types>`;
    const zip = buildDocx(['hi'], {
      '[Content_Types].xml': strToU8(remapped),
      'word/evil.bin': strToU8(bomb),
    });
    await expect(extractWordText(zip)).rejects.toThrow(/declares an XML DTD/);
  });

  it('rejects a DTD in a part whose name is longer than the name buffer', async () => {
    // decodeName caps at 200 chars, which used to drop the .xml suffix the gate read.
    const longName = `word/${'n'.repeat(220)}.xml`;
    const zip = buildDocx(['hi'], { [longName]: strToU8(bomb) });
    await expect(extractWordText(zip)).rejects.toThrow(/declares an XML DTD/);
  });
});

describe('output sanitizing and char budget', () => {
  it('strips bidi overrides from extracted text', async () => {
    const r = await extractWordText(buildDocx(['pay\u202egro.exe\u202c now']));
    expect(r.text).toBe('paygro.exe now');
  });

  it('strips C1 controls and the invisible directional marks too', async () => {
    // C1 sits above the C0 range, and LRM/RLM/ALM reorder rendering while invisible.
    const r = await extractWordText(buildDocx(['a\u0085b\u009fc\u200ed\u200fe\u061cf']));
    expect(r.text).toBe('abcdef');
  });

  it('keeps joiners that carry meaning in real scripts', async () => {
    const r = await extractWordText(buildDocx(['\u0915\u094d\u200d\u0937']));
    expect(r.text).toContain('\u200d');
  });

  it('enforces the char budget and says it truncated', async () => {
    const r = await extractWordText(buildDocx(['y'.repeat(50_000)]), { maxChars: 1_000 });
    expect(r.truncated).toBe(true);
    expect(r.text.length).toBe(1_000);
    expect(r.chars).toBe(1_000);
  });

  it('does not leave half a surrogate pair at the cut', async () => {
    const r = await extractWordText(buildDocx(['ab😀😀😀']), { maxChars: 3 });
    expect(r.truncated).toBe(true);
    expect(r.text).toBe('ab');
    expect(/[\ud800-\udfff]/.test(r.text)).toBe(false);
  });
});

describe('byte cap', () => {
  it('refuses an oversize array before any parsing happens', async () => {
    const huge = new Uint8Array(MAX_WORD_BYTES + 1);
    huge.set([0x50, 0x4b, 0x03, 0x04]);
    await expect(extractWordText(huge)).rejects.toThrow(WordTooLargeError);
  });
});

describe('deadline', () => {
  it('stops the preflight when the clock has already run out', async () => {
    const zip = buildDocx(['hi'], { 'word/pad.bin': compressible(1_000) });
    await expect(extractWordText(zip, { timeoutMs: -1 })).rejects.toThrow(/ran out of time/);
  });
});

describe('a docx with no readable text', () => {
  it('reports empty content rather than pretending it parsed something', async () => {
    const empty = buildDocx([], { 'word/document.xml': strToU8(documentXml()) });
    const r = await extractWordText(empty);
    expect(r.text).toBe('');
    expect(r.chars).toBe(0);
    expect(r.truncated).toBe(false);
  });
});
