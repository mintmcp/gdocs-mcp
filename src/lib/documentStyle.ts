import { buildStyle, pt, type StyleField } from './docRequests.js';

export const PAGE_SIZES = {
  LETTER: { width: 612, height: 792 },
  LEGAL: { width: 612, height: 1008 },
  TABLOID: { width: 792, height: 1224 },
  A3: { width: 841.9, height: 1190.6 },
  A4: { width: 595.3, height: 841.9 },
  A5: { width: 419.5, height: 595.3 },
} as const;

export type PageSizeName = keyof typeof PAGE_SIZES;

export interface DocumentStyleArgs {
  margin_top_pt?: number;
  margin_bottom_pt?: number;
  margin_left_pt?: number;
  margin_right_pt?: number;
  page_size?: PageSizeName;
  orientation?: 'portrait' | 'landscape';
  first_page_header_footer?: boolean;
  page_number_start?: number;
}

const DOCUMENT_STYLE_FIELDS: Array<StyleField<DocumentStyleArgs>> = [
  { arg: 'margin_top_pt', field: 'marginTop', encode: pt },
  { arg: 'margin_bottom_pt', field: 'marginBottom', encode: pt },
  { arg: 'margin_left_pt', field: 'marginLeft', encode: pt },
  { arg: 'margin_right_pt', field: 'marginRight', encode: pt },
  { arg: 'first_page_header_footer', field: 'useFirstPageHeaderFooter' },
  { arg: 'page_number_start', field: 'pageNumberStart' },
];

export function buildDocumentStyle(
  args: DocumentStyleArgs,
  current: { width?: { magnitude?: number }; height?: { magnitude?: number } } | undefined,
): { documentStyle: Record<string, unknown>; fields: string[] } {
  const { style: documentStyle, fields } = buildStyle(DOCUMENT_STYLE_FIELDS, args);

  if (args.page_size !== undefined || args.orientation !== undefined) {
    let { width, height } = args.page_size
      ? PAGE_SIZES[args.page_size]
      : { width: current?.width?.magnitude, height: current?.height?.magnitude };
    if (width === undefined || height === undefined) {
      throw new Error('The document reported no page size, so orientation alone cannot be set; pass page_size as well.');
    }
    if (args.orientation) {
      const landscape = args.orientation === 'landscape';
      if (landscape !== width > height) [width, height] = [height, width];
    }
    documentStyle.pageSize = { width: pt(width), height: pt(height) };
    fields.push('pageSize');
  }
  if (fields.length === 0) {
    throw new Error(`At least one property must be provided (${[...DOCUMENT_STYLE_FIELDS.map((f) => f.arg), 'page_size', 'orientation'].join(', ')}).`);
  }

  return { documentStyle, fields };
}
