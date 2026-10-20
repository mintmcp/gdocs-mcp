import { describe, expect, it } from 'vitest';
import {
  escapeDriveQueryName,
  headingsOf,
  isGoogleServedUrl,
  namedStyleToHeadingLevel,
  parseDocumentStructure,
} from './docs-structure.js';

/**
 * Helper: build a Docs paragraph element with a single textRun.
 */
function paragraph(
  text: string,
  opts: { startIndex?: number; endIndex?: number; namedStyleType?: string; elements?: any[] } = {},
) {
  const start = opts.startIndex ?? 1;
  return {
    startIndex: start,
    endIndex: opts.endIndex ?? start + text.length,
    paragraph: {
      elements: opts.elements ?? [{ textRun: { content: text } }],
      paragraphStyle: opts.namedStyleType ? { namedStyleType: opts.namedStyleType } : undefined,
    },
  };
}

describe('namedStyleToHeadingLevel', () => {
  it('maps HEADING_1..HEADING_6 to numeric levels', () => {
    expect(namedStyleToHeadingLevel('HEADING_1')).toBe(1);
    expect(namedStyleToHeadingLevel('HEADING_6')).toBe(6);
  });

  it('returns undefined for non-heading named styles', () => {
    expect(namedStyleToHeadingLevel('NORMAL_TEXT')).toBeUndefined();
    expect(namedStyleToHeadingLevel('TITLE')).toBeUndefined();
    expect(namedStyleToHeadingLevel('SUBTITLE')).toBeUndefined();
  });

  it('returns undefined for undefined / unknown / out-of-range styles', () => {
    expect(namedStyleToHeadingLevel(undefined)).toBeUndefined();
    expect(namedStyleToHeadingLevel('HEADING_7')).toBeUndefined();
    expect(namedStyleToHeadingLevel('HEADING_0')).toBeUndefined();
    expect(namedStyleToHeadingLevel('GARBAGE')).toBeUndefined();
  });
});

describe('headingsOf', () => {
  const extractHeadings = (content: any[]) => headingsOf(parseDocumentStructure(content).elements);

  it('returns headings in document order with level/text/indices', () => {
    const content = [
      paragraph('Title\n', { startIndex: 1, endIndex: 7, namedStyleType: 'HEADING_1' }),
      paragraph('Body paragraph\n', { startIndex: 7, endIndex: 22, namedStyleType: 'NORMAL_TEXT' }),
      paragraph('Subsection\n', { startIndex: 22, endIndex: 33, namedStyleType: 'HEADING_2' }),
    ];
    const headings = extractHeadings(content);
    expect(headings).toEqual([
      { level: 1, text: 'Title', startIndex: 1, endIndex: 7 },
      { level: 2, text: 'Subsection', startIndex: 22, endIndex: 33 },
    ]);
  });

  it('skips non-paragraph elements and non-heading paragraphs', () => {
    const content = [
      paragraph('Plain', { namedStyleType: 'NORMAL_TEXT' }),
      { startIndex: 10, endIndex: 20, table: { rows: 1, columns: 1 } },
      paragraph('Heading', { startIndex: 20, endIndex: 27, namedStyleType: 'HEADING_3' }),
    ];
    const headings = extractHeadings(content);
    expect(headings).toHaveLength(1);
    expect(headings[0].level).toBe(3);
  });

  it('trims only trailing newlines from heading text', () => {
    const content = [
      paragraph('Heading with   spaces\n\n', { namedStyleType: 'HEADING_1' }),
    ];
    expect(extractHeadings(content)[0].text).toBe('Heading with   spaces');
  });

  it('handles missing endIndex / empty elements gracefully', () => {
    const content = [
      { startIndex: 1, endIndex: 2, paragraph: { paragraphStyle: { namedStyleType: 'HEADING_1' } } },
    ];
    const headings = extractHeadings(content);
    expect(headings).toEqual([{ level: 1, text: '', startIndex: 1, endIndex: 2 }]);
  });
});

describe('parseDocumentStructure', () => {
  it('flattens paragraphs, tables, section breaks, and TOCs', () => {
    const content = [
      paragraph('Hello\n', { startIndex: 1, endIndex: 7 }),
      { startIndex: 7, endIndex: 8, sectionBreak: {} },
      { startIndex: 8, endIndex: 50, table: { rows: 2, columns: 2 } },
      { startIndex: 50, endIndex: 60, tableOfContents: {} },
    ];
    const { elements: structure } = parseDocumentStructure(content);
    expect(structure.map((e) => e.type)).toEqual([
      'paragraph',
      'sectionBreak',
      'table',
      'tableOfContents',
    ]);
    expect(structure[0].text).toBe('Hello\n');
  });

  it('emits a separate inlineImage element for each inlineObjectElement in a paragraph', () => {
    const content = [
      {
        startIndex: 1,
        endIndex: 10,
        paragraph: {
          elements: [
            { textRun: { content: 'Pic ' } },
            { startIndex: 5, endIndex: 6, inlineObjectElement: { inlineObjectId: 'img-1' } },
            { textRun: { content: ' end' } },
          ],
        },
      },
    ];
    const { elements: structure } = parseDocumentStructure(content);
    expect(structure).toHaveLength(2);
    expect(structure[0].type).toBe('paragraph');
    expect(structure[0].text).toBe('Pic  end');
    expect(structure[1]).toEqual({
      type: 'inlineImage',
      startIndex: 5,
      endIndex: 6,
      inlineObjectId: 'img-1',
    });
  });

  it('records headingLevel only when the paragraph has a heading namedStyleType', () => {
    const content = [
      paragraph('H', { namedStyleType: 'HEADING_2' }),
      paragraph('P', { namedStyleType: 'NORMAL_TEXT' }),
    ];
    const { elements: structure } = parseDocumentStructure(content);
    expect(structure[0].headingLevel).toBe(2);
    expect(structure[1].headingLevel).toBeUndefined();
  });

  it('skips unknown element types silently', () => {
    const content = [
      paragraph('keep'),
      { startIndex: 5, endIndex: 6, mysteryFutureElement: {} },
    ];
    expect(parseDocumentStructure(content).elements).toHaveLength(1);
  });

  it('returns an empty array for empty content', () => {
    expect(parseDocumentStructure([]).elements).toEqual([]);
  });
});

describe('escapeDriveQueryName', () => {
  it('returns undefined for empty / undefined input', () => {
    expect(escapeDriveQueryName(undefined)).toBeUndefined();
    expect(escapeDriveQueryName('')).toBeUndefined();
  });

  it('escapes backslash BEFORE single-quote so the apostrophe escape survives', () => {
    // Input:  a\b'c
    // Expected: a\\b\'c
    expect(escapeDriveQueryName("a\\b'c")).toBe("a\\\\b\\'c");
  });

  it('escapes a lone backslash', () => {
    expect(escapeDriveQueryName('back\\slash')).toBe('back\\\\slash');
  });

  it('escapes a lone apostrophe', () => {
    expect(escapeDriveQueryName("O'Brien")).toBe("O\\'Brien");
  });

  it('leaves unrelated characters untouched', () => {
    expect(escapeDriveQueryName('plain name')).toBe('plain name');
  });
});

describe('isGoogleServedUrl', () => {
  it('allows Google-served image and API hosts', () => {
    expect(isGoogleServedUrl('https://lh7-us.googleusercontent.com/docsz/abc')).toBe(true);
    expect(isGoogleServedUrl('https://googleusercontent.com/x')).toBe(true);
    expect(isGoogleServedUrl('https://docs.googleapis.com/v1/x')).toBe(true);
  });

  it('refuses third-party hosts, lookalikes and non-https schemes', () => {
    expect(isGoogleServedUrl('https://evil.example/x.png')).toBe(false);
    expect(isGoogleServedUrl('https://evilgoogleusercontent.com/x')).toBe(false);
    expect(isGoogleServedUrl('https://googleusercontent.com.evil.example/x')).toBe(false);
    expect(isGoogleServedUrl('http://lh7-us.googleusercontent.com/x')).toBe(false);
    expect(isGoogleServedUrl('not a url')).toBe(false);
  });
});

describe('parseDocumentStructure list items', () => {
  it('reports the bullet nesting level on list paragraphs only', () => {
    const content = [
      paragraph('plain\n'),
      { ...paragraph('item\n', { startIndex: 7 }), paragraph: { ...paragraph('item\n').paragraph, bullet: { listId: 'kix.1' } } },
      { ...paragraph('nested\n', { startIndex: 12 }), paragraph: { ...paragraph('nested\n').paragraph, bullet: { listId: 'kix.1', nestingLevel: 1 } } },
    ];
    const { elements } = parseDocumentStructure(content);
    expect(elements[0].bullet).toBeUndefined();
    expect(elements[1].bullet).toEqual({ nestingLevel: 0 });
    expect(elements[2].bullet).toEqual({ nestingLevel: 1 });
  });
});

describe('parseDocumentStructure suggestions', () => {
  it('reports a suggested edit to the text inside a table cell on the table', () => {
    const cellText = { startIndex: 4, endIndex: 7, paragraph: { elements: [{ startIndex: 4, endIndex: 7, textRun: { content: 'ab\n', suggestedInsertionIds: ['suggest.cell'] } }] } };
    const table = { startIndex: 2, endIndex: 9, table: { rows: 1, columns: 1, tableRows: [{ tableCells: [{ content: [cellText] }] }] } };
    expect(parseDocumentStructure([table]).elements[0].suggestions).toEqual([{ id: 'suggest.cell', kind: 'insertion' }]);
  });

  it('reports suggestions on every kind of paragraph element, not only text runs', () => {
    const p = { startIndex: 1, endIndex: 4, paragraph: { elements: [
      { startIndex: 1, endIndex: 2, horizontalRule: { suggestedInsertionIds: ['suggest.rule'] } },
      { startIndex: 2, endIndex: 3, richLink: { suggestedDeletionIds: ['suggest.chip'] } },
      { startIndex: 3, endIndex: 4, textRun: { content: '\n' } },
    ] } };
    expect(parseDocumentStructure([p]).elements[0].suggestions).toEqual([
      { id: 'suggest.rule', kind: 'insertion' },
      { id: 'suggest.chip', kind: 'deletion' },
    ]);
  });
});

describe('parseDocumentStructure suggestions on breaks', () => {
  it('reports a suggested section break and a suggestion inside a table of contents', () => {
    const { elements } = parseDocumentStructure([
      { startIndex: 0, endIndex: 1, sectionBreak: { suggestedInsertionIds: ['suggest.sec'] } },
      { startIndex: 1, endIndex: 9, tableOfContents: { content: [paragraph('Intro\n', { startIndex: 2, endIndex: 8 })], suggestedDeletionIds: ['suggest.toc'] } },
    ]);
    expect(elements[0].suggestions).toEqual([{ id: 'suggest.sec', kind: 'insertion' }]);
    expect(elements[1].suggestions).toEqual([{ id: 'suggest.toc', kind: 'deletion' }]);
  });
});
