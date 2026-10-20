import { docLocation, docRange } from './docRequests.js';
import { contentSuggestions, renderBodyText, type Suggestion } from './docs-structure.js';

export interface Segment {
  content?: any[];
}

export type HeaderFooterKind = 'header' | 'footer';

export const HEADER_FOOTER = {
  header: {
    roleIdFields: { DEFAULT: 'defaultHeaderId', FIRST_PAGE: 'firstPageHeaderId', EVEN_PAGE: 'evenPageHeaderId' },
    idKey: 'headerId',
    createRequest: 'createHeader',
    deleteRequest: 'deleteHeader',
    segments: (tab: { headers: Record<string, Segment> }) => tab.headers,
  },
  footer: {
    roleIdFields: { DEFAULT: 'defaultFooterId', FIRST_PAGE: 'firstPageFooterId', EVEN_PAGE: 'evenPageFooterId' },
    idKey: 'footerId',
    createRequest: 'createFooter',
    deleteRequest: 'deleteFooter',
    segments: (tab: { footers: Record<string, Segment> }) => tab.footers,
  },
} as const;

export interface SegmentSummary {
  id: string;
  text: string;
  suggestions?: Suggestion[];
}

export interface HeaderFooterSummary extends SegmentSummary {
  type: 'DEFAULT' | 'FIRST_PAGE' | 'EVEN_PAGE' | 'SECTION';
}

export function summarizeSegment(id: string, segment: Segment | undefined): SegmentSummary {
  const suggestions = contentSuggestions(segment?.content ?? []);
  return { id, text: renderBodyText(segment?.content ?? []), ...(suggestions.length ? { suggestions } : {}) };
}

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
  for (const [type, field] of Object.entries(HEADER_FOOTER[kind].roleIdFields)) {
    const id = documentStyle?.[field];
    if (typeof id === 'string') roles.set(id, type as HeaderFooterSummary['type']);
  }
  const out: HeaderFooterSummary[] = [];
  for (const [id, segment] of Object.entries(segments ?? {})) {
    out.push({ ...summarizeSegment(id, segment), type: roles.get(id) ?? 'SECTION' });
  }
  return out.sort((a, b) => Number(b.type === 'DEFAULT') - Number(a.type === 'DEFAULT'));
}

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
