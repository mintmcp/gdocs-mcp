import { describe, expect, it } from 'vitest';
import {
  buildTableInsertRequests,
  buildWriteControl,
  clearInheritedFormattingRequest,
  escapeDriveQueryName,
  extractHeadings,
  findLetterByContent,
  namedStyleToHeadingLevel,
  normalizeForMatch,
  parseDocumentStructure,
  parseExportArtifacts,
  summarizeTabs,
  validateIndexRange,
  validateInsertIndex,
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

describe('extractHeadings', () => {
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

describe('clearInheritedFormattingRequest', () => {
  it('builds an updateTextStyle that clears the exact set of inheritable fields', () => {
    const req = clearInheritedFormattingRequest(5, 10);
    expect(req).toEqual({
      updateTextStyle: {
        range: { startIndex: 5, endIndex: 10 },
        textStyle: {},
        fields: 'bold,italic,underline,strikethrough,link',
      },
    });
  });

  it('passes through indices as given (no clamping)', () => {
    const req = clearInheritedFormattingRequest(0, 1);
    expect(req.updateTextStyle.range).toEqual({ startIndex: 0, endIndex: 1 });
  });
});

describe('buildTableInsertRequests', () => {
  it('returns an empty array for empty input', () => {
    expect(buildTableInsertRequests([], 10)).toEqual([]);
    expect(buildTableInsertRequests([[]], 10)).toEqual([]);
  });

  it('emits requests in REVERSE document order so earlier inserts do not shift later indices', () => {
    // 2 rows x 2 cols, all cells filled.
    const requests = buildTableInsertRequests([['A', 'B'], ['C', 'D']], 1);
    // Strictly descending insertText.location.index.
    const indices = requests.map((r) => r.insertText.location.index);
    const sortedDesc = [...indices].sort((a, b) => b - a);
    expect(indices).toEqual(sortedDesc);
    expect(requests).toHaveLength(4);
  });

  it('computes per-cell indices with the documented row+col formula', () => {
    // tableInsertIndex=10, 1 row x 2 cols.
    // rowIndex = 3 + 0 * 2 * 2 + 0 = 3
    // col 0 -> 3 + 0*2 + 1 + 10 = 14
    // col 1 -> 3 + 1*2 + 1 + 10 = 16
    const requests = buildTableInsertRequests([['X', 'Y']], 10);
    // Reverse order: Y first, then X.
    expect(requests[0]).toEqual({ insertText: { text: 'Y', location: { index: 16 } } });
    expect(requests[1]).toEqual({ insertText: { text: 'X', location: { index: 14 } } });
  });

  it('skips empty cells so no insertText is emitted for them', () => {
    const requests = buildTableInsertRequests([['A', '', 'B']], 1);
    const texts = requests.map((r) => r.insertText.text);
    expect(texts).toContain('A');
    expect(texts).toContain('B');
    expect(texts).not.toContain('');
    expect(requests).toHaveLength(2);
  });

  it('coerces non-string cell values via String(...)', () => {
    // Cells *should* be strings, but defend in depth.
    const requests = buildTableInsertRequests([[42 as any, true as any]], 1);
    const texts = requests.map((r) => r.insertText.text);
    expect(texts).toContain('42');
    expect(texts).toContain('true');
  });

  it('handles a single-cell table', () => {
    const requests = buildTableInsertRequests([['only']], 5);
    expect(requests).toHaveLength(1);
    expect(requests[0].insertText.text).toBe('only');
    // rowIndex=3, col=0 -> 3 + 0 + 1 + 5 = 9
    expect(requests[0].insertText.location.index).toBe(9);
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

describe('buildWriteControl', () => {
  it('returns undefined for missing / non-string revisionId', () => {
    expect(buildWriteControl(undefined)).toBeUndefined();
    expect(buildWriteControl('')).toBeUndefined();
    expect(buildWriteControl(123 as any)).toBeUndefined();
  });

  it('wraps a string revisionId into the Docs writeControl shape', () => {
    expect(buildWriteControl('rev-abc')).toEqual({ requiredRevisionId: 'rev-abc' });
  });
});

describe('validateIndexRange', () => {
  it('accepts a valid range', () => {
    expect(() => validateIndexRange(0, 1)).not.toThrow();
    expect(() => validateIndexRange(5, 10)).not.toThrow();
  });

  it('rejects negative startIndex', () => {
    expect(() => validateIndexRange(-1, 5)).toThrow(/startIndex/);
  });

  it('rejects endIndex <= startIndex', () => {
    expect(() => validateIndexRange(5, 5)).toThrow(/endIndex/);
    expect(() => validateIndexRange(5, 4)).toThrow(/endIndex/);
  });

  it('rejects non-finite values', () => {
    expect(() => validateIndexRange(NaN, 5)).toThrow();
    expect(() => validateIndexRange(0, Infinity)).toThrow();
  });
});

describe('validateInsertIndex', () => {
  it('accepts an index >= 1', () => {
    expect(() => validateInsertIndex(1)).not.toThrow();
    expect(() => validateInsertIndex(42)).not.toThrow();
  });

  it('rejects an index < 1 (the document-start sentinel)', () => {
    expect(() => validateInsertIndex(0)).toThrow(/Docs body starts at index 1/);
    expect(() => validateInsertIndex(-5)).toThrow();
  });

  it('rejects non-finite values', () => {
    expect(() => validateInsertIndex(NaN)).toThrow();
  });
});

describe('summarizeTabs', () => {
  it('returns an empty array for missing / empty tabs', () => {
    expect(summarizeTabs(undefined)).toEqual([]);
    expect(summarizeTabs([])).toEqual([]);
  });

  it('flattens nested childTabs in pre-order with depth-derived nestingLevel fallback', () => {
    const tabs = [
      {
        tabProperties: { tabId: 'root', title: 'Root', index: 0 },
        childTabs: [
          { tabProperties: { tabId: 'child', title: 'Child', index: 0 } },
        ],
      },
      { tabProperties: { tabId: 'sibling', title: 'Sibling', index: 1, nestingLevel: 0 } },
    ];
    const out = summarizeTabs(tabs);
    expect(out.map((t) => t.tabId)).toEqual(['root', 'child', 'sibling']);
    // root: no provided nestingLevel -> falls back to depth (0).
    expect(out[0].nestingLevel).toBe(0);
    // child: no provided nestingLevel -> falls back to depth (1).
    expect(out[1].nestingLevel).toBe(1);
    // sibling: explicit nestingLevel preserved.
    expect(out[2].nestingLevel).toBe(0);
  });

  it('skips tabs without a tabId', () => {
    const tabs = [
      { tabProperties: { title: 'No id' } },
      { tabProperties: { tabId: 'has-id' } },
    ];
    expect(summarizeTabs(tabs).map((t) => t.tabId)).toEqual(['has-id']);
  });
});

describe('parseExportArtifacts', () => {
  it('returns the input unchanged when there is no trailing letter block', () => {
    const parsed = parseExportArtifacts('Just a plain doc.\nNo markers.');
    expect(parsed.body).toBe('Just a plain doc.\nNo markers.');
    expect(parsed.letterContent.size).toBe(0);
    expect(parsed.letterReactions.size).toBe(0);
    expect(parsed.commentLetters.size).toBe(0);
  });

  it('strips the trailing letter block and captures comment text + reactions', () => {
    const text = [
      'Body paragraph one.',
      'Body paragraph two with [a] marker.',
      '',
      '[a]First comment',
      'Alice reacted with 👍 at 2024-01-02 14:30',
      'Bob reacted with 🎉 at 2024-01-02 15:00 PM',
      '[b]Second comment text',
    ].join('\n');
    const parsed = parseExportArtifacts(text);
    expect(parsed.body).toBe('Body paragraph one.\nBody paragraph two with [a] marker.');
    expect(parsed.letterContent.get('a')).toBe('First comment');
    expect(parsed.letterContent.get('b')).toBe('Second comment text');
    expect(parsed.commentLetters).toEqual(new Set(['a', 'b']));
    const rx = parsed.letterReactions.get('a')!;
    expect(rx).toHaveLength(2);
    expect(rx[0]).toEqual({ author_name: 'Alice', emoji: '👍', timestamp: '2024-01-02T14:30' });
    // AM/PM suffix dropped; HH is already 24h.
    expect(rx[1]).toEqual({ author_name: 'Bob', emoji: '🎉', timestamp: '2024-01-02T15:00' });
    expect(parsed.letterReactions.get('b')).toEqual([]);
  });
});

describe('normalizeForMatch', () => {
  it('collapses runs of whitespace and trims', () => {
    expect(normalizeForMatch('  a   b\n\tc  ')).toBe('a b c');
  });
});

describe('findLetterByContent', () => {
  it('finds the letter whose normalised content matches', () => {
    const map = new Map([
      ['a', 'Hello world'],
      ['b', '  Goodbye  world  '],
    ]);
    expect(findLetterByContent(map, 'Hello\tworld')).toBe('a');
    expect(findLetterByContent(map, 'Goodbye world')).toBe('b');
  });

  it('returns undefined for empty / no-match input', () => {
    const map = new Map([['a', 'Hello']]);
    expect(findLetterByContent(map, '')).toBeUndefined();
    expect(findLetterByContent(map, 'nope')).toBeUndefined();
  });
});
