import { describe, expect, it } from 'vitest';
import { listHeaderFooters, replaceSegmentTextRequests, segmentEndIndex, segmentText } from './segments.js';

const segment = (text: string) => ({
  content: [{ startIndex: 0, endIndex: text.length, paragraph: { elements: [{ textRun: { content: text } }] } }],
});

describe('segmentText / segmentEndIndex', () => {
  it('joins the runs and ends before the final newline', () => {
    expect(segmentText(segment('Hi\n'))).toBe('Hi\n');
    expect(segmentEndIndex(segment('Hi\n'))).toBe(2);
    expect(segmentEndIndex(segment('\n'))).toBe(0);
    expect(segmentEndIndex(undefined)).toBe(0);
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
