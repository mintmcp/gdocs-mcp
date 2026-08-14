/**
 * Reads .docx and legacy .doc uploads. word-extractor parses; the container
 * bounds are ours because its zip reader does not limit decompression. Text only,
 * so Word's hyperlink fields never surface the javascript:/data: vector.
 */

import { Buffer } from 'node:buffer';
import WordExtractor, { type Document } from 'word-extractor';

export const MAX_WORD_BYTES = 20 * 1024 * 1024;
export const MAX_TEXT_CHARS = 200_000;
export const EXTRACT_TIMEOUT_MS = 10_000;

export const MAX_ZIP_ENTRIES = 512;
export const MAX_ENTRY_BYTES = 8 * 1024 * 1024;
export const MAX_TOTAL_UNPACKED_BYTES = 50 * 1024 * 1024;

const MAX_OLE_DIR_ENTRIES = 4_096;
const MAX_OLE_DIR_SECTORS = 4_096;

// Compound-file sector values at or above this are reserved markers, not sectors.
const OLE_FIRST_RESERVED_SECTOR = 0xfffffffa;
const OLE_END_OF_CHAIN = 0xfffffffe;
const OLE_HEADER_SIZE = 512;
const OLE_DIR_ENTRY_SIZE = 128;
const OLE_HEADER_DIFAT_SLOTS = 109;
// Sector shift 7..12 is 128B..4KB, the only sizes the format allows.
const OLE_MIN_SECTOR_SHIFT = 7;
const OLE_MAX_SECTOR_SHIFT = 12;
// Deflate peaks near 1032:1, so an 8KB slice cannot expand past ~8.3MB even if
// the runtime emits it as one chunk: the worst-case transient is arithmetic.
const INFLATE_CHUNK_BYTES = 8 * 1024;

export const DOCX_MIME =
  'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
export const LEGACY_DOC_MIME = 'application/msword';

export type WordFormat = 'docx' | 'doc';

export function isWordMime(mimeType: string): boolean {
  return mimeType === DOCX_MIME || mimeType === LEGACY_DOC_MIME;
}

export class WordInvalidError extends Error {
  /** Parser internals for the log; never for the user-facing message. */
  readonly detail?: string;
  constructor(message: string, detail?: string) {
    super(message);
    this.name = 'WordInvalidError';
    this.detail = detail;
  }
}
export class WordEncryptedError extends Error {
  constructor(message: string) { super(message); this.name = 'WordEncryptedError'; }
}
export class WordTooLargeError extends Error {
  constructor(message: string) { super(message); this.name = 'WordTooLargeError'; }
}
export class WordTimeoutError extends Error {
  constructor(message: string) { super(message); this.name = 'WordTimeoutError'; }
}

export interface WordText {
  text: string;
  format: WordFormat;
  truncated: boolean;
  chars: number;
}

export interface WordOptions {
  maxChars?: number;
  timeoutMs?: number;
  /** The container the caller's mimeType implies, so a mismatch can be named. */
  expectedFormat?: WordFormat;
}

const FORMAT_LABEL: Record<WordFormat, string> = {
  docx: 'a .docx (Word 2007 or later)',
  doc: 'a legacy .doc (Word 97-2003)',
};

const ZIP_MAGIC = [0x50, 0x4b, 0x03, 0x04];
const OLE_MAGIC = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];

const startsWith = (bytes: Uint8Array, sig: number[]) =>
  bytes.byteLength >= sig.length && sig.every((b, i) => bytes[i] === b);

export function sniffWordFormat(bytes: Uint8Array): WordFormat | null {
  if (startsWith(bytes, ZIP_MAGIC)) return 'docx';
  if (startsWith(bytes, OLE_MAGIC)) return 'doc';
  return null;
}

const EOCD_SIG = 0x06054b50;
const CD_SIG = 0x02014b50;
const LFH_SIG = 0x04034b50;
const ZIP64_MARKER = 0xffff;
const ZIP64_SIZE_MARKER = 0xffffffff;

// Fixed record sizes, not offsets: each is the length of a zip structure.
const EOCD_SIZE = 22;
const CD_HEADER_SIZE = 46;
const LFH_SIZE = 30;
const MAX_ZIP_COMMENT = 65_535;
const EOCD_SEARCH_SPAN = EOCD_SIZE + MAX_ZIP_COMMENT;

const OPEN_ANGLE = 0x3c;
const MAX_ENTRY_NAME_CHARS = 200;
const MAX_ERROR_DETAIL_CHARS = 200;

const DOCTYPE_PATTERN = asciiLower('<!doctype');
const ENTITY_PATTERN = asciiLower('<!entity');
const MAX_PATTERN_LEN = Math.max(DOCTYPE_PATTERN.length, ENTITY_PATTERN.length);

function asciiLower(text: string): Uint8Array {
  const out = new Uint8Array(text.length);
  for (let i = 0; i < text.length; i++) out[i] = text.charCodeAt(i) | 0x20;
  return out;
}

function matchAt(buf: Uint8Array, at: number, pattern: Uint8Array): boolean {
  if (at + pattern.length > buf.length) return false;
  for (let i = 0; i < pattern.length; i++) {
    if ((buf[at + i] | 0x20) !== pattern[i]) return false;
  }
  return true;
}

/**
 * Rejects a DTD, where billion-laughs and XXE start. Every entry, not just
 * XML-looking names: [Content_Types].xml can aim 'word/x.bin' at the body.
 */
class DtdScanner {
  private tail = new Uint8Array(0);

  push(chunk: Uint8Array, entryName: string): void {
    const buf = this.tail.length === 0 ? chunk : concat(this.tail, chunk);
    for (let i = 0; i < buf.length; i++) {
      if (buf[i] !== OPEN_ANGLE) continue;
      if (matchAt(buf, i, DOCTYPE_PATTERN) || matchAt(buf, i, ENTITY_PATTERN)) {
        throw new WordInvalidError(
          `the part '${entryName}' declares an XML DTD, which Word never writes and which is ` +
          `used to attack XML parsers. If the file is genuine, open it and re-save it from Word`
        );
      }
    }
    this.tail = buf.slice(Math.max(0, buf.length - (MAX_PATTERN_LEN - 1)));
  }
}

function concat(a: Uint8Array, b: Uint8Array): Uint8Array {
  const out = new Uint8Array(a.length + b.length);
  out.set(a);
  out.set(b, a.length);
  return out;
}

// Requiring the comment to reach EOF stops a record planted inside the real
// one's comment from winning the backward scan.
function findEocd(bytes: Uint8Array, view: DataView): number {
  const floor = Math.max(0, bytes.byteLength - EOCD_SEARCH_SPAN);
  for (let at = bytes.byteLength - EOCD_SIZE; at >= floor; at--) {
    if (view.getUint32(at, true) !== EOCD_SIG) continue;
    if (at + EOCD_SIZE + view.getUint16(at + 20, true) === bytes.byteLength) return at;
  }
  throw new WordInvalidError('it is not a readable .docx (no zip end-of-directory record)');
}

interface ZipEntry {
  name: string;
  method: number;
  compressedSize: number;
  declaredSize: number;
  localOffset: number;
  encrypted: boolean;
}

function readCentralDirectory(bytes: Uint8Array, view: DataView): ZipEntry[] {
  const eocd = findEocd(bytes, view);
  const declaredEntries = view.getUint16(eocd + 10, true);
  const cdOffset = view.getUint32(eocd + 16, true);

  if (declaredEntries === ZIP64_MARKER || cdOffset === ZIP64_SIZE_MARKER) {
    throw new WordInvalidError('it uses zip64, which no Word document needs');
  }
  if (declaredEntries > MAX_ZIP_ENTRIES) {
    throw new WordInvalidError(
      `it holds ${declaredEntries} zip entries, above the ${MAX_ZIP_ENTRIES}-entry limit`
    );
  }
  if (cdOffset + CD_HEADER_SIZE > bytes.byteLength) {
    throw new WordInvalidError('its zip directory points outside the file');
  }

  const entries: ZipEntry[] = [];
  let at = cdOffset;
  while (entries.length < declaredEntries) {
    if (at + CD_HEADER_SIZE > bytes.byteLength || view.getUint32(at, true) !== CD_SIG) {
      throw new WordInvalidError('its zip directory is truncated or corrupt');
    }
    const nameLen = view.getUint16(at + 28, true);
    const extraLen = view.getUint16(at + 30, true);
    const commentLen = view.getUint16(at + 32, true);
    entries.push({
      name: decodeName(bytes, at + CD_HEADER_SIZE, nameLen),
      method: view.getUint16(at + 10, true),
      compressedSize: view.getUint32(at + 20, true),
      declaredSize: view.getUint32(at + 24, true),
      localOffset: view.getUint32(at + 42, true),
      encrypted: (view.getUint16(at + 8, true) & 0x1) !== 0,
    });
    at += CD_HEADER_SIZE + nameLen + extraLen + commentLen;
  }
  return entries;
}

function decodeName(bytes: Uint8Array, at: number, len: number): string {
  let name = '';
  for (let i = 0; i < len && at + i < bytes.byteLength; i++) {
    const b = bytes[at + i];
    name += b >= 0x20 && b < 0x7f ? String.fromCharCode(b) : '?';
  }
  return name.slice(0, MAX_ENTRY_NAME_CHARS);
}

function entryData(bytes: Uint8Array, view: DataView, entry: ZipEntry): Uint8Array {
  const lfh = entry.localOffset;
  if (lfh + LFH_SIZE > bytes.byteLength || view.getUint32(lfh, true) !== LFH_SIG) {
    throw new WordInvalidError(`its entry '${entry.name}' has no local header`);
  }
  const start = lfh + LFH_SIZE + view.getUint16(lfh + 26, true) + view.getUint16(lfh + 28, true);
  const end = start + entry.compressedSize;
  if (end > bytes.byteLength) {
    throw new WordInvalidError(`its entry '${entry.name}' runs past the end of the file`);
  }
  return bytes.subarray(start, end);
}

/**
 * Counts what each entry ACTUALLY produces: the declared size is attacker-controlled
 * and deflate reaches ~1032:1, so a 20MB upload can expand past 20GB.
 */
async function preflightZip(bytes: Uint8Array, deadline: number): Promise<void> {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const entries = readCentralDirectory(bytes, view);
  let total = 0;

  for (const entry of entries) {
    if (Date.now() > deadline) {
      throw new WordTimeoutError('the .docx preflight ran out of time');
    }
    if (entry.encrypted) {
      throw new WordEncryptedError('the .docx is password-protected');
    }
    if (entry.method !== 0 && entry.method !== 8) {
      throw new WordInvalidError(
        `its entry '${entry.name}' uses compression method ${entry.method}, which Word never emits`
      );
    }

    if (entry.declaredSize > MAX_ENTRY_BYTES) {
      throw new WordTooLargeError(
        `its entry '${entry.name}' declares ${entry.declaredSize} bytes, past the ` +
        `${MAX_ENTRY_BYTES} byte per-entry limit`
      );
    }
    if (total + entry.declaredSize > MAX_TOTAL_UNPACKED_BYTES) {
      throw new WordTooLargeError(
        `it unpacks past the ${MAX_TOTAL_UNPACKED_BYTES} byte total limit`
      );
    }

    const data = entryData(bytes, view, entry);
    const scanner = new DtdScanner();
    let actual = 0;

    for await (const chunk of inflate(data, entry.method)) {
      actual += chunk.byteLength;
      if (actual > entry.declaredSize) {
        throw new WordInvalidError(
          `its entry '${entry.name}' unpacks past the ${entry.declaredSize} bytes it declares`
        );
      }
      scanner.push(chunk, entry.name);
    }

    if (actual !== entry.declaredSize) {
      throw new WordInvalidError(
        `its entry '${entry.name}' unpacks to ${actual} bytes but declares ${entry.declaredSize}`
      );
    }
    total += actual;
  }
}

async function* inflate(data: Uint8Array, method: number): AsyncGenerator<Uint8Array> {
  if (method === 0) {
    if (data.byteLength) yield data;
    return;
  }
  const source = new ReadableStream<Uint8Array>({
    start(controller) {
      for (let at = 0; at < data.byteLength; at += INFLATE_CHUNK_BYTES) {
        controller.enqueue(data.subarray(at, Math.min(at + INFLATE_CHUNK_BYTES, data.byteLength)));
      }
      controller.close();
    },
  });
  const reader = source.pipeThrough(new DecompressionStream('deflate-raw')).getReader();
  try {
    for (;;) {
      const { done, value } = await reader.read();
      if (done) return;
      if (value) yield value;
    }
  } catch (err) {
    throw new WordInvalidError(
      `one of its zip entries could not be decompressed (${(err as Error)?.message ?? err})`
    );
  } finally {
    await reader.cancel().catch(() => {});
  }
}

// Matches directory names, not raw bytes: .doc body text is UTF-16LE, so a
// document merely mentioning EncryptedPackage would otherwise look encrypted.
// Returns [] on anything it cannot walk, which reads as "not encrypted" and
// defers to the parser. The cost is a worse message, never a weaker check.
function oleDirectoryNames(bytes: Uint8Array): string[] {
  if (bytes.byteLength < OLE_HEADER_SIZE) return [];
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const sectorShift = view.getUint16(30, true);
  if (sectorShift < OLE_MIN_SECTOR_SHIFT || sectorShift > OLE_MAX_SECTOR_SHIFT) return [];
  const sectorSize = 1 << sectorShift;
  const offsetOf = (sector: number) => (sector + 1) * sectorSize;

  // Only the header's DIFAT slots; parsing classifies anything deeper.
  const fatSectors: number[] = [];
  for (let i = 0; i < OLE_HEADER_DIFAT_SLOTS; i++) {
    const loc = view.getUint32(76 + i * 4, true);
    if (loc >= OLE_FIRST_RESERVED_SECTOR) break;
    fatSectors.push(loc);
  }
  const perFatSector = sectorSize / 4;
  const nextSector = (sector: number): number => {
    const fatIndex = Math.floor(sector / perFatSector);
    if (fatIndex >= fatSectors.length) return OLE_END_OF_CHAIN;
    const at = offsetOf(fatSectors[fatIndex]) + (sector % perFatSector) * 4;
    return at + 4 <= bytes.byteLength ? view.getUint32(at, true) : OLE_END_OF_CHAIN;
  };

  const names: string[] = [];
  const seen = new Set<number>();
  let sector = view.getUint32(48, true);
  while (
    sector < OLE_FIRST_RESERVED_SECTOR &&
    !seen.has(sector) &&
    names.length < MAX_OLE_DIR_ENTRIES &&
    seen.size < MAX_OLE_DIR_SECTORS
  ) {
    seen.add(sector);
    const start = offsetOf(sector);
    if (start + sectorSize > bytes.byteLength) break;
    for (let at = start; at + OLE_DIR_ENTRY_SIZE <= start + sectorSize; at += OLE_DIR_ENTRY_SIZE) {
      const nameLen = view.getUint16(at + 64, true);
      if (nameLen < 4 || nameLen > 64) continue;
      let name = '';
      for (let i = 0; i < nameLen - 2; i += 2) name += String.fromCharCode(view.getUint16(at + i, true));
      if (name) names.push(name);
    }
    sector = nextSector(sector);
  }
  return names;
}

const ENCRYPTED_OLE_STREAMS = new Set(['EncryptedPackage', 'EncryptionInfo']);

function preflightOle(bytes: Uint8Array): void {
  if (oleDirectoryNames(bytes).some((name) => ENCRYPTED_OLE_STREAMS.has(name))) {
    throw new WordEncryptedError('the file is an encrypted Office container');
  }
}

// word-extractor emits \x00 for deleted ranges; bidi controls and the invisible
// LRM/RLM/ALM can make text render as something it is not. ZWJ/ZWNJ stay, they
// carry meaning in several scripts.
const CONTROLS = /[\u0000-\u0008\u000b\u000c\u000e-\u001f\u007f-\u009f]/g;
const BIDI_OVERRIDES = /[\u061c\u200e\u200f\u202a-\u202e\u2066-\u2069]/g;
const LONE_SURROGATES = /[\ud800-\udbff](?![\udc00-\udfff])|(?<![\ud800-\udbff])[\udc00-\udfff]/g;

export function sanitizeWordText(text: string): string {
  return text
    .replace(/\r\n?/g, '\n')
    .replace(CONTROLS, '')
    .replace(BIDI_OVERRIDES, '')
    .replace(LONE_SURROGATES, '');
}

function sliceChars(text: string, room: number): string {
  const cut = text.slice(0, room);
  const last = cut.charCodeAt(cut.length - 1);
  return last >= 0xd800 && last <= 0xdbff ? cut.slice(0, -1) : cut;
}

const headerOrFooter = (text: string) => (isFurniture(text || '') ? '' : text);

// Letters mean content the body never repeats — letterhead, a CONFIDENTIAL mark.
const isFurniture = (text: string) => !/\p{L}/u.test(text);

/**
 * Separate parser streams, appended rather than dropped and labelled so they do
 * not read as body prose. Header leads and footer trails to mirror the page,
 * which also keeps a classification marking on the surviving side of truncation.
 */
function assemble(doc: Document, maxChars: number, deadline: number): { text: string; truncated: boolean } {
  const streams: Array<[string, string]> = [
    ['[Header]\n', headerOrFooter(doc.getHeaders({ includeFooters: false }))],
    ['', doc.getBody()],
    ['[Footer]\n', headerOrFooter(doc.getFooters())],
    ['[Text boxes]\n', doc.getTextboxes()],
    ['[Footnotes]\n', doc.getFootnotes()],
    ['[Endnotes]\n', doc.getEndnotes()],
  ];
  let text = '';
  let truncated = false;

  for (const [label, stream] of streams) {
    if (Date.now() > deadline) {
      throw new WordTimeoutError('assembling the extracted text ran out of time');
    }
    const body = sanitizeWordText(stream || '').trim();
    if (!body) continue;
    const part = label + body;

    const separator = text ? '\n' : '';
    const room = maxChars - text.length - separator.length;
    if (room <= 0) {
      truncated = true;
      break;
    }
    if (part.length > room) {
      text += separator + sliceChars(part, room);
      truncated = true;
      break;
    }
    text += separator + part;
  }

  return { text: text.trimEnd(), truncated };
}

const extractor = new WordExtractor();

/**
 * The deadline bounds our own loops only. word-extractor's parse is synchronous
 * and cannot be preempted from JS; the byte and unpacked-size caps bound that.
 */
export async function extractWordText(
  bytes: Uint8Array,
  options: WordOptions = {}
): Promise<WordText> {
  const maxChars = options.maxChars ?? MAX_TEXT_CHARS;
  const deadline = Date.now() + (options.timeoutMs ?? EXTRACT_TIMEOUT_MS);

  if (bytes.byteLength > MAX_WORD_BYTES) {
    throw new WordTooLargeError(
      `the file is ${bytes.byteLength} bytes, above the ${MAX_WORD_BYTES} byte limit`
    );
  }

  const format = sniffWordFormat(bytes);
  if (!format) {
    throw new WordInvalidError(
      'its contents are not a Word document at all, whatever its name says'
    );
  }
  // Preflight first: Office wraps an encrypted .docx in an OLE container, so the
  // mismatch below would otherwise blame the name for a password-protected file.
  if (format === 'docx') {
    await preflightZip(bytes, deadline);
  } else {
    preflightOle(bytes);
  }
  if (options.expectedFormat && format !== options.expectedFormat) {
    throw new WordInvalidError(
      `its contents are ${FORMAT_LABEL[format]} even though it is stored as ` +
      `${FORMAT_LABEL[options.expectedFormat]}. Renaming a Word file does not convert it: ` +
      `open it and re-save it in the format its name claims`
    );
  }

  let doc: Document;
  try {
    doc = await extractor.extract(Buffer.from(bytes));
  } catch (err: any) {
    const message = String(err?.message ?? err);
    if (/password|encrypt/i.test(message)) {
      throw new WordEncryptedError('the file is password-protected');
    }
    throw new WordInvalidError(
      'it is damaged, or it is not really a Word document despite its name',
      message.slice(0, MAX_ERROR_DETAIL_CHARS)
    );
  }
  if (Date.now() > deadline) {
    throw new WordTimeoutError('reading the Word file ran out of time');
  }

  const { text, truncated } = assemble(doc, maxChars, deadline);
  return { text, format, truncated, chars: text.length };
}
