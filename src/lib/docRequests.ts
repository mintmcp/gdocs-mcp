export interface WriteControl {
  requiredRevisionId?: string;
  targetRevisionId?: string;
  writeMode?: 'SUGGEST';
}

export interface WriteOptions {
  required?: string;
  read?: string;
  suggest?: boolean;
}

export function buildWriteControl({ required, read, suggest = false }: WriteOptions): WriteControl | undefined {
  const revision = required
    ? { requiredRevisionId: required }
    : read ? { targetRevisionId: read } : {};
  const control = { ...revision, ...(suggest ? { writeMode: 'SUGGEST' as const } : {}) };
  return Object.keys(control).length ? control : undefined;
}

export interface DocRange {
  startIndex: number;
  endIndex: number;
  tabId?: string;
  segmentId?: string;
}

export function docLocation(index: number, tabId?: string, segmentId?: string): { index: number; tabId?: string; segmentId?: string } {
  return { index, ...(tabId ? { tabId } : {}), ...(segmentId ? { segmentId } : {}) };
}

export function docRange(startIndex: number, endIndex: number, tabId?: string, segmentId?: string): DocRange {
  return { startIndex, endIndex, ...(tabId ? { tabId } : {}), ...(segmentId ? { segmentId } : {}) };
}

export function insertionPoint(index: number | undefined, tabId?: string) {
  return index === undefined
    ? { endOfSegmentLocation: { ...(tabId ? { tabId } : {}) } }
    : { location: docLocation(index, tabId) };
}

export function endOfBodyIndex(body: any[]): number {
  const last = body[body.length - 1];
  if (!last || typeof last.endIndex !== 'number') {
    throw new Error('Document body is empty or unreadable; cannot find the end index.');
  }
  return Math.max(1, last.endIndex - 1);
}

export function validateIndexRange(startIndex: number, endIndex: number, inSegment = false): void {
  const min = inSegment ? 0 : 1;
  if (!Number.isFinite(startIndex) || startIndex < min) {
    throw new Error(`startIndex must be >= ${min}${inSegment ? '' : ' (Docs body starts at index 1)'}`);
  }
  if (!Number.isFinite(endIndex) || endIndex <= startIndex) {
    throw new Error('endIndex must be greater than startIndex');
  }
}

export function validateInsertIndex(index: number, inSegment = false): void {
  const min = inSegment ? 0 : 1;
  if (!Number.isFinite(index) || index < min) {
    throw new Error(inSegment
      ? 'index must be >= 0 (header, footer and footnote text starts at index 0)'
      : 'index must be >= 1 (Docs body starts at index 1)');
  }
}

export interface StyleField<A> {
  arg: keyof A & string;
  field: string;
  encode?: (value: any) => unknown;
}

export const pt = (magnitude: number) => ({ magnitude, unit: 'PT' });

export function hexToRgbColor(hex: string, label: string) {
  const match = /^#?([0-9a-fA-F]{6})$/.exec(String(hex));
  if (!match) {
    throw new Error(
      `${label} must be a 6-digit hex colour such as #1A73E8, received '${String(hex).slice(0, 20)}'.`
    );
  }
  const value = parseInt(match[1], 16);
  return {
    rgbColor: {
      red: ((value >> 16) & 255) / 255,
      green: ((value >> 8) & 255) / 255,
      blue: (value & 255) / 255,
    },
  };
}

export function buildStyle<A>(table: Array<StyleField<A>>, args: A): { style: Record<string, unknown>; fields: string[] } {
  const style: Record<string, unknown> = {};
  const fields: string[] = [];
  for (const { arg, field, encode } of table) {
    const value = args[arg];
    if (value === undefined) continue;
    style[field] = encode ? encode(value) : value;
    fields.push(field);
  }
  return { style, fields };
}

export interface TextStyleArgs {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strikethrough?: boolean;
  small_caps?: boolean;
  font_size?: number;
  font_family?: string;
  color?: string;
  background_color?: string;
  baseline_offset?: 'NONE' | 'SUPERSCRIPT' | 'SUBSCRIPT';
  link_url?: string;
}

const TEXT_STYLE_FIELDS: Array<StyleField<TextStyleArgs>> = [
  { arg: 'bold', field: 'bold' },
  { arg: 'italic', field: 'italic' },
  { arg: 'underline', field: 'underline' },
  { arg: 'strikethrough', field: 'strikethrough' },
  { arg: 'font_size', field: 'fontSize', encode: pt },
  { arg: 'font_family', field: 'weightedFontFamily', encode: (fontFamily) => ({ fontFamily }) },
  { arg: 'color', field: 'foregroundColor', encode: (hex) => ({ color: hexToRgbColor(hex, 'color') }) },
  // An empty colour object is how Google clears a highlight
  { arg: 'background_color', field: 'backgroundColor', encode: (hex) => (hex === '' ? {} : { color: hexToRgbColor(hex, 'background_color') }) },
  { arg: 'baseline_offset', field: 'baselineOffset' },
  { arg: 'small_caps', field: 'smallCaps' },
  { arg: 'link_url', field: 'link', encode: (url) => (url === '' ? null : { url }) },
];

export function buildTextStyle(args: TextStyleArgs): { textStyle: Record<string, unknown>; fields: string[] } {
  const { style, fields } = buildStyle(TEXT_STYLE_FIELDS, args);
  if (fields.length === 0) {
    throw new Error(`At least one style property must be provided (${TEXT_STYLE_FIELDS.map((f) => f.arg).join(', ')})`);
  }
  return { textStyle: style, fields };
}

export function clearInheritedFormattingRequest(startIndex: number, endIndex: number, tabId?: string, segmentId?: string): any {
  return {
    updateTextStyle: {
      range: docRange(startIndex, endIndex, tabId, segmentId),
      textStyle: {},
      fields: TEXT_STYLE_FIELDS.map((f) => f.field).join(','),
    },
  };
}

export const NAMED_STYLES = ['NORMAL_TEXT', 'TITLE', 'SUBTITLE', 'HEADING_1', 'HEADING_2', 'HEADING_3', 'HEADING_4', 'HEADING_5', 'HEADING_6'] as const;

export const LIST_STYLES = ['bullet', 'numbered', 'checkbox', 'none'] as const;
export type ListStyle = typeof LIST_STYLES[number];

const LIST_PRESETS: Record<Exclude<ListStyle, 'none'>, string> = {
  bullet: 'BULLET_DISC_CIRCLE_SQUARE',
  numbered: 'NUMBERED_DECIMAL_ALPHA_ROMAN',
  checkbox: 'BULLET_CHECKBOX',
};

const presetListStyle = (preset: string): Exclude<ListStyle, 'none'> =>
  preset === 'BULLET_CHECKBOX' ? 'checkbox' : preset.startsWith('NUMBERED_') ? 'numbered' : 'bullet';

export const BULLET_PRESETS = [
  'BULLET_DISC_CIRCLE_SQUARE', 'BULLET_DIAMONDX_ARROW3D_SQUARE', 'BULLET_CHECKBOX', 'BULLET_ARROW_DIAMOND_DISC',
  'BULLET_STAR_CIRCLE_SQUARE', 'BULLET_ARROW3D_CIRCLE_SQUARE', 'BULLET_LEFTTRIANGLE_DIAMOND_DISC',
  'BULLET_DIAMONDX_HOLLOWDIAMOND_SQUARE', 'BULLET_DIAMOND_CIRCLE_SQUARE', 'NUMBERED_DECIMAL_ALPHA_ROMAN',
  'NUMBERED_DECIMAL_ALPHA_ROMAN_PARENS', 'NUMBERED_DECIMAL_NESTED', 'NUMBERED_UPPERALPHA_ALPHA_ROMAN',
  'NUMBERED_UPPERROMAN_UPPERALPHA_DECIMAL', 'NUMBERED_ZERODECIMAL_ALPHA_ROMAN',
] as const;

export interface ParagraphStyleArgs {
  heading_level?: number;
  named_style?: string;
  alignment?: string;
  line_spacing?: number;
  space_above_pt?: number;
  space_below_pt?: number;
  indent_start_pt?: number;
  indent_end_pt?: number;
  indent_first_line_pt?: number;
  keep_with_next?: boolean;
  page_break_before?: boolean;
  list_style?: ListStyle;
  bullet_preset?: string;
}

function headingStyle(level: number): string {
  if (!Number.isInteger(level) || level < 0 || level > 6) {
    throw new Error('heading_level must be 0 (normal) or 1-6');
  }
  return level === 0 ? 'NORMAL_TEXT' : `HEADING_${level}`;
}

const PARAGRAPH_STYLE_FIELDS: Array<StyleField<ParagraphStyleArgs>> = [
  { arg: 'heading_level', field: 'namedStyleType', encode: headingStyle },
  { arg: 'named_style', field: 'namedStyleType' },
  { arg: 'alignment', field: 'alignment' },
  { arg: 'line_spacing', field: 'lineSpacing' },
  { arg: 'space_above_pt', field: 'spaceAbove', encode: pt },
  { arg: 'space_below_pt', field: 'spaceBelow', encode: pt },
  { arg: 'indent_start_pt', field: 'indentStart', encode: pt },
  { arg: 'indent_end_pt', field: 'indentEnd', encode: pt },
  { arg: 'indent_first_line_pt', field: 'indentFirstLine', encode: pt },
  { arg: 'keep_with_next', field: 'keepWithNext' },
  { arg: 'page_break_before', field: 'pageBreakBefore' },
];

export function buildParagraphRequests(args: ParagraphStyleArgs, range: DocRange): { requests: any[]; applied: string[] } {
  if (args.heading_level !== undefined && args.named_style !== undefined) {
    throw new Error('Pass one of heading_level or named_style, not both: both set the paragraph\'s named style.');
  }
  if (range.segmentId && args.page_break_before !== undefined) {
    throw new Error('page_break_before applies to body paragraphs only; Google does not allow it in a header, footer or footnote.');
  }
  const { list_style: listStyle, bullet_preset: preset } = args;
  if (preset && presetListStyle(preset) !== listStyle) {
    throw new Error(`bullet_preset ${preset} draws a ${presetListStyle(preset)} list, so it goes with list_style "${presetListStyle(preset)}"${listStyle ? `, not "${listStyle}"` : ''}.`);
  }
  const { style, fields } = buildStyle(PARAGRAPH_STYLE_FIELDS, args);
  if (fields.length === 0 && !listStyle) {
    throw new Error(`At least one style property must be provided (${[...PARAGRAPH_STYLE_FIELDS.map((f) => f.arg), 'list_style'].join(', ')})`);
  }
  const applied = [...fields, ...(listStyle ? [`list_style=${listStyle}`] : [])];

  const mask = listStyle === 'none'
    ? [...new Set([...fields, 'indentStart', 'indentFirstLine'])]
    : fields;
  const styleRequests = mask.length > 0
    ? [{ updateParagraphStyle: { range, paragraphStyle: style, fields: mask.join(',') } }]
    : [];

  if (listStyle === 'none') return { requests: [{ deleteParagraphBullets: { range } }, ...styleRequests], applied };
  if (listStyle) {
    const bullets = { createParagraphBullets: { range, bulletPreset: preset ?? LIST_PRESETS[listStyle] } };
    return { requests: [...styleRequests, bullets], applied };
  }
  return { requests: styleRequests, applied };
}

export function clearInheritedParagraphStyleRequests(range: DocRange): any[] {
  const fields = [...new Set(PARAGRAPH_STYLE_FIELDS.map((f) => f.field))];
  return [
    { deleteParagraphBullets: { range } },
    { updateParagraphStyle: { range, paragraphStyle: { namedStyleType: 'NORMAL_TEXT' }, fields: fields.join(',') } },
  ];
}

export function appendTextRequests(body: any[], text: string, tabId: string | undefined, clearInherited: boolean): any[] {
  const index = endOfBodyIndex(body);
  const elements: any[] = body[body.length - 1]?.paragraph?.elements ?? [];
  const lastHasContent = elements.some((el) => (el.textRun ? el.textRun.content.replace(/\n$/, '') !== '' : true));
  const inserted = lastHasContent && !text.startsWith('\n') ? `\n${text}` : text;

  const requests: any[] = [{ insertText: { location: docLocation(index, tabId), text: inserted } }];
  if (!clearInherited) return requests;

  const newText = docRange(inserted.startsWith('\n') ? index + 1 : index, index + inserted.length, tabId);
  if (newText.endIndex <= newText.startIndex) return requests;
  return [
    ...requests,
    clearInheritedFormattingRequest(newText.startIndex, newText.endIndex, tabId),
    ...clearInheritedParagraphStyleRequests(newText),
  ];
}
