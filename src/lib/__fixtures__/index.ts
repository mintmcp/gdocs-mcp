import { readFileSync } from 'node:fs';
import { zipSync, strToU8 } from 'fflate';

/** Loads a checked-in Word fixture as the bytes extractWordText receives. */
export const fixture = (name: string) =>
  new Uint8Array(readFileSync(new URL(`./${name}`, import.meta.url)));

const CONTENT_TYPES =
  `<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
  `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
  `<Default Extension="xml" ContentType="application/xml"/>` +
  `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>` +
  `</Types>`;

const PACKAGE_RELS =
  `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
  `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>` +
  `</Relationships>`;

export function documentXml(...paragraphs: string[]): string {
  return `<?xml version="1.0" encoding="UTF-8"?>` +
    `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>` +
    paragraphs.map((t) => `<w:p><w:r><w:t xml:space="preserve">${t}</w:t></w:r></w:p>`).join('') +
    `</w:body></w:document>`;
}

/** A minimal .docx word-extractor accepts, so hostile parts can be added around it. */
export function buildDocx(
  paragraphs: string[],
  extra: Record<string, Uint8Array> = {}
): Uint8Array {
  return zipSync({
    '[Content_Types].xml': strToU8(CONTENT_TYPES),
    '_rels/.rels': strToU8(PACKAGE_RELS),
    'word/document.xml': strToU8(documentXml(...paragraphs)),
    ...extra,
  }, { level: 9 });
}

const END_OF_CHAIN = 0xfffffffe;
const FAT_SECTOR = 0xfffffffd;
const FREE_SECTOR = 0xffffffff;
const SECTOR = 512;

/**
 * A compound file whose directory holds exactly the given stream names, for
 * driving the OLE classification without shipping a real encrypted document.
 *
 * CFB header offsets written below: 24/26 version, 28 byte order, 30/32 sector
 * shifts, 44 FAT count, 48 first directory sector, 56 mini cutoff, 60 mini FAT,
 * 68 DIFAT extension, 76 DIFAT[0].
 */
export function buildOle(streamNames: string[]): Uint8Array {
  const bytes = new Uint8Array(SECTOR * 3);
  const view = new DataView(bytes.buffer);

  bytes.set([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1], 0);
  view.setUint16(24, 0x3e, true);
  view.setUint16(26, 3, true);
  view.setUint16(28, 0xfffe, true);
  view.setUint16(30, 9, true);
  view.setUint16(32, 6, true);
  view.setUint32(44, 1, true);
  view.setUint32(48, 1, true);
  view.setUint32(56, 4096, true);
  view.setUint32(60, END_OF_CHAIN, true);
  view.setUint32(68, END_OF_CHAIN, true);
  view.setUint32(76, 0, true);
  for (let i = 1; i < 109; i++) view.setUint32(76 + i * 4, FREE_SECTOR, true);

  const fat = SECTOR;
  for (let i = 0; i < SECTOR / 4; i++) view.setUint32(fat + i * 4, FREE_SECTOR, true);
  view.setUint32(fat, FAT_SECTOR, true);
  view.setUint32(fat + 4, END_OF_CHAIN, true);

  const dir = SECTOR * 2;
  ['Root Entry', ...streamNames].forEach((name, index) => {
    const at = dir + index * 128;
    for (let i = 0; i < name.length; i++) view.setUint16(at + i * 2, name.charCodeAt(i), true);
    view.setUint16(at + 64, (name.length + 1) * 2, true);
    bytes[at + 66] = index === 0 ? 5 : 2;
    view.setUint32(at + 68, FREE_SECTOR, true);
    view.setUint32(at + 72, FREE_SECTOR, true);
    view.setUint32(at + 76, FREE_SECTOR, true);
  });

  return bytes;
}
