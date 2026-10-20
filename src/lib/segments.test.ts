import { describe, expect, it } from 'vitest';
import { listHeaderFooters, replaceSegmentTextRequests, segmentEndIndex, summarizeSegment } from './segments.js';

const segment = (text: string) => ({
  content: [{ startIndex: 0, endIndex: text.length, paragraph: { elements: [{ textRun: { content: text } }] } }],
});

describe('segment text and end index', () => {
  it('joins the runs and ends before the final newline', () => {
    expect(summarizeSegment('h', segment('Hi\n')).text).toBe('Hi\n');
    expect(segmentEndIndex(segment('Hi\n'))).toBe(2);
    expect(segmentEndIndex(segment('\n'))).toBe(0);
    expect(segmentEndIndex(undefined)).toBe(0);
  });

  it('includes the cells of a table in the segment', () => {
    const cell = (text: string, startIndex: number) => ({ content: [{ startIndex, endIndex: startIndex + text.length, paragraph: { elements: [{ textRun: { content: text } }] } }] });
    const header = {
      content: [
        { startIndex: 0, endIndex: 9, table: { rows: 1, columns: 2, tableRows: [{ tableCells: [cell('Logo\n', 2), cell('Q3\n', 8)] }] } },
        { startIndex: 9, endIndex: 15, paragraph: { elements: [{ textRun: { content: 'Draft\n' } }] } },
      ],
    };
    expect(summarizeSegment('h', header).text).toBe('Logo\tQ3\nDraft\n');
  });
});

describe('listHeaderFooters', () => {
  it('names the role documentStyle gives each id and marks the rest as section headers', () => {
    const out = listHeaderFooters('header', { defaultHeaderId: 'a', firstPageHeaderId: 'b' }, {
      c: segment('Section\n'),
      b: segment('First\n'),
      a: segment('Default\n'),
    });
    expect(out).toEqual([
      { id: 'a', type: 'DEFAULT', text: 'Default\n' },
      { id: 'c', type: 'SECTION', text: 'Section\n' },
      { id: 'b', type: 'FIRST_PAGE', text: 'First\n' },
    ]);
    expect(listHeaderFooters('footer', {}, undefined)).toEqual([]);
  });
});

describe('replaceSegmentTextRequests', () => {
  it('deletes the old text up to the final newline, then inserts at 0', () => {
    expect(replaceSegmentTextRequests(segment('Old\n'), 'kix.h', 'New', 't.0')).toEqual([
      { deleteContentRange: { range: { startIndex: 0, endIndex: 3, tabId: 't.0', segmentId: 'kix.h' } } },
      { insertText: { location: { index: 0, tabId: 't.0', segmentId: 'kix.h' }, text: 'New' } },
    ]);
  });

  it('skips the delete on an empty segment and the insert on empty text', () => {
    expect(replaceSegmentTextRequests(segment('\n'), 'kix.h', 'New')).toEqual([
      { insertText: { location: { index: 0, segmentId: 'kix.h' }, text: 'New' } },
    ]);
    expect(replaceSegmentTextRequests(segment('Old\n'), 'kix.h', '')).toEqual([
      { deleteContentRange: { range: { startIndex: 0, endIndex: 3, segmentId: 'kix.h' } } },
    ]);
  });
});

describe('summarizeSegment', () => {
  it('lists the suggestions inside a header or footnote, and none when there are none', () => {
    const suggested = { content: [{ startIndex: 0, endIndex: 4, paragraph: { elements: [{ textRun: { content: 'New\n', suggestedInsertionIds: ['suggest.h'] } }] } }] };
    expect(summarizeSegment('kix.h', suggested)).toEqual({ id: 'kix.h', text: 'New\n', suggestions: [{ id: 'suggest.h', kind: 'insertion' }] });
    expect(summarizeSegment('kix.f', segment('Plain\n'))).toEqual({ id: 'kix.f', text: 'Plain\n' });
  });
});
