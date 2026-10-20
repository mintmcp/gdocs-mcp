import { describe, expect, it } from 'vitest';
import {
  appendTextRequests,
  buildParagraphRequests,
  buildTextStyle,
  buildWriteControl,
  clearInheritedFormattingRequest,
  clearInheritedParagraphStyleRequests,
  docLocation,
  docRange,
  endOfBodyIndex,
  validateIndexRange,
  validateInsertIndex,
} from './docRequests.js';
import { paragraph } from "./__fixtures__/index.js";

describe('docLocation / docRange', () => {
  it('omits tabId entirely when not targeting a tab', () => {
    expect(docLocation(5)).toEqual({ index: 5 });
    expect(docRange(1, 9)).toEqual({ startIndex: 1, endIndex: 9 });
  });

  it('carries tabId when targeting a tab', () => {
    expect(docLocation(5, 't.1')).toEqual({ index: 5, tabId: 't.1' });
    expect(docRange(1, 9, 't.1')).toEqual({ startIndex: 1, endIndex: 9, tabId: 't.1' });
  });
});

describe('segment-aware locations', () => {
  it('adds segmentId only when given, after the tab', () => {
    expect(docLocation(3, 't.0', 'kix.h')).toEqual({ index: 3, tabId: 't.0', segmentId: 'kix.h' });
    expect(docLocation(3, undefined, 'kix.h')).toEqual({ index: 3, segmentId: 'kix.h' });
    expect(docRange(1, 4, undefined, 'kix.h')).toEqual({ startIndex: 1, endIndex: 4, segmentId: 'kix.h' });
    expect(docRange(1, 4, 't.0')).toEqual({ startIndex: 1, endIndex: 4, tabId: 't.0' });
  });
});

describe('endOfBodyIndex', () => {
  it('returns one before the final endIndex', () => {
    expect(endOfBodyIndex([paragraph('a\n', 1), paragraph('b\n', 3)])).toBe(4);
  });

  it('never goes below the body start sentinel', () => {
    expect(endOfBodyIndex([{ endIndex: 1 }])).toBe(1);
  });

  it('throws on an empty body instead of crashing', () => {
    expect(() => endOfBodyIndex([])).toThrow(/empty or unreadable/);
  });
});

describe('buildWriteControl', () => {
  it('returns undefined when there is nothing to control', () => {
    expect(buildWriteControl({})).toBeUndefined();
    expect(buildWriteControl({ required: '' })).toBeUndefined();
    expect(buildWriteControl({ read: '' })).toBeUndefined();
  });

  it('targets the revision the indices were read from unless one is required', () => {
    expect(buildWriteControl({ read: 'read-1' })).toEqual({ targetRevisionId: 'read-1' });
    expect(buildWriteControl({ read: 'read-1', suggest: true })).toEqual({ targetRevisionId: 'read-1', writeMode: 'SUGGEST' });
    expect(buildWriteControl({ required: 'rev-1', read: 'read-1' })).toEqual({ requiredRevisionId: 'rev-1' });
  });

  it('wraps a string revisionId into the Docs writeControl shape', () => {
    expect(buildWriteControl({ required: 'rev-abc' })).toEqual({ requiredRevisionId: 'rev-abc' });
  });
});

describe('buildWriteControl suggest mode', () => {
  it('adds writeMode SUGGEST, with or without a revision pin', () => {
    expect(buildWriteControl({ suggest: true })).toEqual({ writeMode: 'SUGGEST' });
    expect(buildWriteControl({ required: 'rev-1', suggest: true })).toEqual({ requiredRevisionId: 'rev-1', writeMode: 'SUGGEST' });
    expect(buildWriteControl({ required: 'rev-1' })).toEqual({ requiredRevisionId: 'rev-1' });
  });
});

describe('validateIndexRange', () => {
  it('accepts a valid range', () => {
    expect(() => validateIndexRange(1, 2)).not.toThrow();
    expect(() => validateIndexRange(5, 10)).not.toThrow();
  });

  it('starts a body range at 1 and a segment range at 0', () => {
    expect(() => validateIndexRange(0, 2)).toThrow(/>= 1 \(Docs body starts at index 1\)/);
    expect(() => validateIndexRange(0, 2, true)).not.toThrow();
    expect(() => validateIndexRange(-1, 2, true)).toThrow(/>= 0/);
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
    expect(() => validateIndexRange(1, Infinity)).toThrow();
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

describe('validateInsertIndex in a segment', () => {
  it('allows index 0, where header and footer text starts', () => {
    expect(() => validateInsertIndex(0, true)).not.toThrow();
    expect(() => validateInsertIndex(-1, true)).toThrow(/index must be >= 0/);
  });
});

describe('buildTextStyle', () => {
  it('maps every supported property to a field', () => {
    expect(buildTextStyle({ bold: true, italic: false, font_size: 14, color: '#FF0000', link_url: 'https://x.y' })).toEqual({
      textStyle: {
        bold: true,
        italic: false,
        fontSize: { magnitude: 14, unit: 'PT' },
        foregroundColor: { color: { rgbColor: { red: 1, green: 0, blue: 0 } } },
        link: { url: 'https://x.y' },
      },
      fields: ['bold', 'italic', 'fontSize', 'foregroundColor', 'link'],
    });
  });

  it('clears a link with an empty url and rejects a bad colour', () => {
    expect(buildTextStyle({ link_url: '' })).toEqual({ textStyle: { link: null }, fields: ['link'] });
    expect(() => buildTextStyle({ color: 'red' })).toThrow(/6-digit hex/);
  });

  it('rejects an empty request', () => {
    expect(() => buildTextStyle({})).toThrow(/At least one/);
  });
});

describe('clearInheritedFormattingRequest', () => {
  it('builds an updateTextStyle that clears the exact set of inheritable fields', () => {
    const req = clearInheritedFormattingRequest(5, 10);
    expect(req).toEqual({
      updateTextStyle: {
        range: { startIndex: 5, endIndex: 10 },
        textStyle: {},
        fields: 'bold,italic,underline,strikethrough,fontSize,weightedFontFamily,foregroundColor,backgroundColor,baselineOffset,smallCaps,link',
      },
    });
  });

  it('clears every field update_text_style can set', () => {
    const { fields } = buildTextStyle({
      bold: true, italic: true, underline: true, strikethrough: true, small_caps: true, font_size: 12,
      font_family: 'Arial', color: '#000000', background_color: '#FFFF00', baseline_offset: 'SUPERSCRIPT', link_url: 'https://x.y',
    });
    expect(clearInheritedFormattingRequest(1, 2).updateTextStyle.fields.split(',')).toEqual(fields);
  });

  it('carries a segment', () => {
    expect(clearInheritedFormattingRequest(0, 3, 't.1', 'kix.h').updateTextStyle.range).toEqual({ startIndex: 0, endIndex: 3, tabId: 't.1', segmentId: 'kix.h' });
  });

  it('passes through indices as given (no clamping)', () => {
    const req = clearInheritedFormattingRequest(0, 1);
    expect(req.updateTextStyle.range).toEqual({ startIndex: 0, endIndex: 1 });
  });
});


describe('buildParagraphRequests', () => {
  const range = docRange(5, 20, 't.1');

  it('maps heading levels and dimensions in field-mask order', () => {
    expect(buildParagraphRequests({ heading_level: 2, space_above_pt: 6, keep_with_next: true }, range)).toEqual({
      requests: [{
        updateParagraphStyle: {
          range,
          paragraphStyle: { namedStyleType: 'HEADING_2', spaceAbove: { magnitude: 6, unit: 'PT' }, keepWithNext: true },
          fields: 'namedStyleType,spaceAbove,keepWithNext',
        },
      }],
      applied: ['namedStyleType', 'spaceAbove', 'keepWithNext'],
    });
    expect(buildParagraphRequests({ heading_level: 0 }, range).requests[0].updateParagraphStyle.paragraphStyle).toEqual({ namedStyleType: 'NORMAL_TEXT' });
  });

  it('creates bullets after the paragraph style, with the preset for each list style', () => {
    const { requests, applied } = buildParagraphRequests({ heading_level: 1, list_style: 'numbered' }, range);
    expect(requests.map((r) => Object.keys(r)[0])).toEqual(['updateParagraphStyle', 'createParagraphBullets']);
    expect(requests[1].createParagraphBullets).toEqual({ range, bulletPreset: 'NUMBERED_DECIMAL_ALPHA_ROMAN' });
    expect(applied).toEqual(['namedStyleType', 'list_style=numbered']);
    expect(buildParagraphRequests({ list_style: 'bullet' }, range).requests).toEqual([
      { createParagraphBullets: { range, bulletPreset: 'BULLET_DISC_CIRCLE_SQUARE' } },
    ]);
    expect(buildParagraphRequests({ list_style: 'bullet', bullet_preset: 'BULLET_ARROW_DIAMOND_DISC' }, range).requests[0].createParagraphBullets.bulletPreset).toBe('BULLET_ARROW_DIAMOND_DISC');
    expect(buildParagraphRequests({ list_style: 'numbered', bullet_preset: 'NUMBERED_DECIMAL_NESTED' }, range).requests[0].createParagraphBullets.bulletPreset).toBe('NUMBERED_DECIMAL_NESTED');
  });

  it('removes bullets first, then resets the indent Google adds to keep the old nesting', () => {
    expect(buildParagraphRequests({ heading_level: 1, list_style: 'none' }, range).requests).toEqual([
      { deleteParagraphBullets: { range } },
      { updateParagraphStyle: { range, paragraphStyle: { namedStyleType: 'HEADING_1' }, fields: 'namedStyleType,indentStart,indentFirstLine' } },
    ]);
  });

  it('keeps an indent the caller set while removing bullets', () => {
    const [, style] = buildParagraphRequests({ list_style: 'none', indent_start_pt: 18 }, range).requests;
    expect(style.updateParagraphStyle).toEqual({
      range,
      paragraphStyle: { indentStart: { magnitude: 18, unit: 'PT' } },
      fields: 'indentStart,indentFirstLine',
    });
  });

  it('rejects empty, conflicting and preset-only requests', () => {
    expect(() => buildParagraphRequests({}, range)).toThrow(/At least one style property.*list_style/);
    expect(() => buildParagraphRequests({ heading_level: 1, named_style: 'TITLE' }, range)).toThrow(/not both/);
    expect(() => buildParagraphRequests({ bullet_preset: 'BULLET_CHECKBOX' }, range)).toThrow(/goes with list_style "checkbox"\.$/);
    expect(() => buildParagraphRequests({ list_style: 'numbered', bullet_preset: 'BULLET_CHECKBOX' }, range)).toThrow(/checkbox list.*not "numbered"/);
    expect(() => buildParagraphRequests({ list_style: 'none', bullet_preset: 'NUMBERED_DECIMAL_NESTED' }, range)).toThrow(/not "none"/);
    expect(() => buildParagraphRequests({ heading_level: 7 }, range)).toThrow(/heading_level must be/);
  });
});

describe('clearInheritedParagraphStyleRequests', () => {
  it('removes bullets, then resets every paragraph style update_paragraph_style can set', () => {
    const range = docRange(8, 19, 't.1');
    expect(clearInheritedParagraphStyleRequests(range)).toEqual([
      { deleteParagraphBullets: { range } },
      { updateParagraphStyle: { range, paragraphStyle: { namedStyleType: 'NORMAL_TEXT' }, fields: 'namedStyleType,alignment,lineSpacing,spaceAbove,spaceBelow,indentStart,indentEnd,indentFirstLine,keepWithNext,pageBreakBefore' } },
    ]);
  });
});

describe('appendTextRequests', () => {
  const ranges = (requests: any[]) => requests.slice(1).map((r) => (Object.values(r)[0] as any).range);

  it('starts a new paragraph when the last one has text, and resets only the new one', () => {
    const requests = appendTextRequests([paragraph('alpha\n', 1)], 'beta', 't.1', true);
    expect(requests[0]).toEqual({ insertText: { location: { index: 6, tabId: 't.1' }, text: '\nbeta' } });
    expect(ranges(requests)).toEqual(Array(3).fill({ startIndex: 7, endIndex: 11, tabId: 't.1' }));
  });

  it('keeps the previous paragraph out of the resets when the text brings its own newline', () => {
    const requests = appendTextRequests([paragraph('alpha\n', 1)], '\nbeta', undefined, true);
    expect(requests[0].insertText.text).toBe('\nbeta');
    expect(ranges(requests)).toEqual(Array(3).fill({ startIndex: 7, endIndex: 11 }));
  });

  it('writes into an empty last paragraph without a separator', () => {
    const requests = appendTextRequests([paragraph('alpha\n', 1), paragraph('\n', 7)], 'beta', undefined, true);
    expect(requests[0].insertText).toEqual({ location: { index: 7 }, text: 'beta' });
    expect(ranges(requests)[0]).toEqual({ startIndex: 7, endIndex: 11 });
  });

  it('writes into a table-terminated body without a separator', () => {
    const body = [{ startIndex: 1, endIndex: 10, table: {} }, paragraph('\n', 10)];
    expect(appendTextRequests(body, 'beta', undefined, false)).toEqual([{ insertText: { location: { index: 10 }, text: 'beta' } }]);
  });

  it('treats an image-only paragraph as content', () => {
    const body = [{ startIndex: 1, endIndex: 3, paragraph: { elements: [{ inlineObjectElement: { inlineObjectId: 'kix.img' } }, { textRun: { content: '\n' } }] } }];
    expect(appendTextRequests(body, 'caption', undefined, false)[0].insertText).toEqual({ location: { index: 2 }, text: '\ncaption' });
  });

  it('sends no resets when nothing but a separator was added', () => {
    expect(appendTextRequests([paragraph('alpha\n', 1)], '\n', undefined, true)).toHaveLength(1);
  });
});

describe('appendTextRequests next to an empty last paragraph', () => {
  it('leaves the empty paragraph alone when the text brings its own newline', () => {
    const requests = appendTextRequests([paragraph('alpha\n', 1), paragraph('\n', 7)], '\nbeta', undefined, true);
    expect(requests[0]).toEqual({ insertText: { location: { index: 7 }, text: '\nbeta' } });
    expect(requests.slice(1).map((r) => (Object.values(r)[0] as any).range)).toEqual(Array(3).fill({ startIndex: 8, endIndex: 12 }));
  });

  it('resets the empty paragraph the text joins', () => {
    const requests = appendTextRequests([paragraph('alpha\n', 1), paragraph('\n', 7)], 'beta', undefined, true);
    expect(requests.slice(1).map((r) => (Object.values(r)[0] as any).range)).toEqual(Array(3).fill({ startIndex: 7, endIndex: 11 }));
  });
});

describe('buildParagraphRequests in a segment', () => {
  it('refuses page_break_before, which Google does not allow outside the body', () => {
    expect(() => buildParagraphRequests({ page_break_before: true }, docRange(0, 3, undefined, 'kix.h'))).toThrow(/body paragraphs only/);
    expect(() => buildParagraphRequests({ page_break_before: true }, docRange(1, 3))).not.toThrow();
  });
});
