/** The updateDocumentStyle payload: page size, margins and header/footer options. */

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

const MARGINS: Array<[keyof DocumentStyleArgs, string]> = [
  ['margin_top_pt', 'marginTop'],
  ['margin_bottom_pt', 'marginBottom'],
  ['margin_left_pt', 'marginLeft'],
  ['margin_right_pt', 'marginRight'],
];

const pt = (magnitude: number) => ({ magnitude, unit: 'PT' });

/**
 * `current` is the document's pageSize, needed when only the orientation
 * changes: the page keeps its dimensions and swaps them.
 */
export function buildDocumentStyle(
  args: DocumentStyleArgs,
  current: { width?: { magnitude?: number }; height?: { magnitude?: number } } | undefined,
): { documentStyle: Record<string, unknown>; fields: string[] } {
  const documentStyle: Record<string, unknown> = {};
  const fields: string[] = [];

  for (const [arg, field] of MARGINS) {
    const value = args[arg];
    if (value === undefined) continue;
    documentStyle[field] = pt(value as number);
    fields.push(field);
  }

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

  if (args.first_page_header_footer !== undefined) {
    documentStyle.useFirstPageHeaderFooter = args.first_page_header_footer;
    fields.push('useFirstPageHeaderFooter');
  }
  if (args.page_number_start !== undefined) {
    documentStyle.pageNumberStart = args.page_number_start;
    fields.push('pageNumberStart');
  }

  return { documentStyle, fields };
}
