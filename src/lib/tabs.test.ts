import { describe, expect, it } from 'vitest';
import { renderBodyText } from './docs-structure.js';
import {
  collectInlineObjects,
  resolveTab,
  summarizeTabs,
  tabsCriteriaFor,
  unknownTabError,
} from './tabs.js';
import { paragraph } from "./__fixtures__/index.js";

const tabbedDoc = {
  tabs: [
    {
      tabProperties: { tabId: 't.0', title: 'First', index: 0 },
      documentTab: {
        body: { content: [paragraph('alpha\n', 1)] },
        inlineObjects: { 'kix.a': { marker: 1 } },
      },
      childTabs: [
        {
          tabProperties: { tabId: 't.child', title: 'Nested', index: 0, nestingLevel: 1 },
          documentTab: { body: { content: [paragraph('nested\n', 1)] } },
        },
      ],
    },
    {
      tabProperties: { tabId: 't.1', title: 'Second', index: 1 },
      documentTab: { body: { content: [paragraph('beta\n', 1)] } },
    },
  ],
};

describe('resolveTab', () => {
  it('returns the root body with no tabId for untabbed documents', () => {
    const doc = { body: { content: [paragraph('hi\n', 1)] }, inlineObjects: { 'kix.x': {} } };
    const resolved = resolveTab(doc);
    expect(resolved.tabId).toBeUndefined();
    expect(resolved.body).toHaveLength(1);
    expect(resolved.inlineObjects).toHaveProperty('kix.x');
    expect(resolved.tabs).toEqual([]);
  });

  it('rejects a tab_id on an untabbed document', () => {
    expect(() => resolveTab({ body: { content: [] } }, 't.0'))
      .toThrow(/has no tabs/);
  });

  it('defaults to the first tab with an explicit tabId on tabbed documents', () => {
    const resolved = resolveTab(tabbedDoc);
    expect(resolved.tabId).toBe('t.0');
    expect(resolved.title).toBe('First');
    expect(resolved.body[0].paragraph.elements[0].textRun.content).toBe('alpha\n');
    expect(resolved.inlineObjects).toHaveProperty('kix.a');
  });

  it('resolves a top-level tab by id', () => {
    const resolved = resolveTab(tabbedDoc, 't.1');
    expect(resolved.tabId).toBe('t.1');
    expect(resolved.body[0].paragraph.elements[0].textRun.content).toBe('beta\n');
  });

  it('resolves a nested child tab by id', () => {
    const resolved = resolveTab(tabbedDoc, 't.child');
    expect(resolved.tabId).toBe('t.child');
    expect(resolved.title).toBe('Nested');
    expect(resolved.body[0].paragraph.elements[0].textRun.content).toBe('nested\n');
  });

  it('lists the available tabs for an unknown tab_id', () => {
    expect(() => resolveTab(tabbedDoc, 't.nope'))
      .toThrow(/t\.0 \(First\).*t\.child \(Nested\).*t\.1 \(Second\)/);
  });

  it('summarizes all tabs including nested ones', () => {
    const resolved = resolveTab(tabbedDoc);
    expect(resolved.tabs.map((t) => t.tabId)).toEqual(['t.0', 't.child', 't.1']);
  });

  it('tolerates a tab without content', () => {
    const doc = { tabs: [{ tabProperties: { tabId: 't.0' } }] };
    const resolved = resolveTab(doc, 't.0');
    expect(resolved.body).toEqual([]);
    expect(resolved.inlineObjects).toEqual({});
  });
});

describe('renderBodyText', () => {
  const table = (rows: string[][]) => ({
    startIndex: 5, endIndex: 20,
    table: { tableRows: rows.map((row) => ({ tableCells: row.map((text) => ({ content: [paragraph(`${text}\n`, 0)] })) })) },
  });

  it('concatenates paragraph text', () => {
    expect(renderBodyText([paragraph('one\n', 1), paragraph('two\n', 5)])).toBe('one\ntwo\n');
  });

  it('renders tables as tab-separated rows', () => {
    expect(renderBodyText([paragraph('hi\n', 1), table([['a', 'b'], ['c', 'd']])])).toBe('hi\na\tb\nc\td\n');
  });

  it('keeps every character of a long cell, which the structure caps', () => {
    const long = 'x'.repeat(5000);
    expect(renderBodyText([table([[long]])])).toBe(`${long}\n`);
  });

  it('skips a section break', () => {
    expect(renderBodyText([{ startIndex: 0, endIndex: 1, sectionBreak: {} }, paragraph('x\n', 1)])).toBe('x\n');
  });
});

describe('collectInlineObjects', () => {
  it('merges the root field with every tab, including nested ones', () => {
    const merged = collectInlineObjects({
      inlineObjects: { 'kix.root': {} },
      tabs: tabbedDoc.tabs,
    });
    expect(Object.keys(merged).sort()).toEqual(['kix.a', 'kix.root']);
  });

  it('returns an empty object for a doc without images', () => {
    expect(collectInlineObjects({})).toEqual({});
  });

  it('keeps to one tab when given its id, and refuses an unknown one', () => {
    expect(Object.keys(collectInlineObjects(tabbedDoc, 't.0'))).toEqual(['kix.a']);
    expect(collectInlineObjects(tabbedDoc, 't.1')).toEqual({});
    expect(() => collectInlineObjects(tabbedDoc, 't.nope')).toThrow(/No tab with id/);
  });
});

describe('unknownTabError', () => {
  it('mentions the URL param so agents can self-correct', () => {
    const err = unknownTabError('t.x', [{ tabId: 't.0', title: 'Main' }]);
    expect(err.message).toMatch(/\?tab=/);
  });
});

describe('tabsCriteriaFor', () => {
  const tabs = [{ tabId: 't.0', title: 'One' }, { tabId: 't.1', title: 'Two' }];

  it('returns no criteria for an untabbed document', () => {
    expect(tabsCriteriaFor([])).toBeUndefined();
  });

  it('rejects a tab_id on an untabbed document', () => {
    expect(() => tabsCriteriaFor([], 't.0')).toThrow(/has no tabs/);
  });

  it('pins to the first tab when tab_id is omitted', () => {
    expect(tabsCriteriaFor(tabs)).toEqual({ tabIds: ['t.0'] });
  });

  it('pins to the requested tab', () => {
    expect(tabsCriteriaFor(tabs, 't.1')).toEqual({ tabIds: ['t.1'] });
  });

  it('lists the available tabs for an unknown tab_id', () => {
    expect(() => tabsCriteriaFor(tabs, 't.nope')).toThrow(/t\.0 \(One\).*t\.1 \(Two\)/);
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
