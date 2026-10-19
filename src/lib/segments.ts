/**
 * Headers, footers and footnotes are "segments": bodies of their own with
 * indices that start at 0 and end on a newline Google will not let anyone
 * delete. These helpers read them and build the requests that rewrite them.
 */

import { docLocation, docRange } from './tabs.js';

export interface Segment {
  content?: any[];
}

export type HeaderFooterKind = 'header' | 'footer';

/** Everything that differs between a header and a footer in Google's requests and replies. */
export const HEADER_FOOTER = {
  header: {
    defaultIdField: 'defaultHeaderId',
    idKey: 'headerId',
    createRequest: 'createHeader',
    deleteRequest: 'deleteHeader',
    segments: (tab: { headers: Record<string, Segment> }) => tab.headers,
  },
  footer: {
    defaultIdField: 'defaultFooterId',
    idKey: 'footerId',
    createRequest: 'createFooter',
    deleteRequest: 'deleteFooter',
    segments: (tab: { footers: Record<string, Segment> }) => tab.footers,
  },
} as const;

export interface HeaderFooterSummary {
  id: string;
  type: 'DEFAULT' | 'FIRST_PAGE' | 'EVEN_PAGE' | 'SECTION';
  text: string;
}

/** The ids documentStyle points at; any other header or footer belongs to a section. */
const ROLE_FIELDS: Record<HeaderFooterKind, Array<[string, HeaderFooterSummary['type']]>> = {
  header: [['defaultHeaderId', 'DEFAULT'], ['firstPageHeaderId', 'FIRST_PAGE'], ['evenPageHeaderId', 'EVEN_PAGE']],
  footer: [['defaultFooterId', 'DEFAULT'], ['firstPageFooterId', 'FIRST_PAGE'], ['evenPageFooterId', 'EVEN_PAGE']],
};

export function segmentText(segment: Segment | undefined): string {
  let out = '';
  for (const el of segment?.content ?? []) {
    for (const run of el.paragraph?.elements ?? []) {
      out += run.textRun?.content ?? '';
    }
  }
  return out;
}

/** Index of the segment's final newline, which is also the end of what may be deleted. */
export function segmentEndIndex(segment: Segment | undefined): number {
  const content = segment?.content ?? [];
  const last = content[content.length - 1];
  return typeof last?.endIndex === 'number' ? Math.max(0, last.endIndex - 1) : 0;
}

export function listHeaderFooters(
  kind: HeaderFooterKind,
  documentStyle: any,
  segments: Record<string, Segment> | undefined,
): HeaderFooterSummary[] {
  const roles = new Map<string, HeaderFooterSummary['type']>();
  for (const [field, type] of ROLE_FIELDS[kind]) {
    const id = documentStyle?.[field];
    if (typeof id === 'string') roles.set(id, type);
  }
  const out: HeaderFooterSummary[] = [];
  for (const [id, segment] of Object.entries(segments ?? {})) {
    out.push({ id, type: roles.get(id) ?? 'SECTION', text: segmentText(segment) });
  }
  // DEFAULT first, then the rest in Google's order
  return out.sort((a, b) => Number(b.type === 'DEFAULT') - Number(a.type === 'DEFAULT'));
}

/** Replace everything in a segment with `text`; empty text clears it. */
export function replaceSegmentTextRequests(segment: Segment | undefined, segmentId: string, text: string, tabId?: string): any[] {
  const requests: any[] = [];
  const end = segmentEndIndex(segment);
  if (end > 0) {
    requests.push({ deleteContentRange: { range: docRange(0, end, tabId, segmentId) } });
  }
  if (text) {
    requests.push({ insertText: { location: docLocation(0, tabId, segmentId), text } });
  }
  return requests;
}
