import { describe, expect, it } from 'vitest';
import {
  collectInlineObjects,
  docLocation,
  docRange,
  endOfBodyIndex,
  renderStructureText,
  renderTabText,
  resolveTab,
  unknownTabError,
} from './tabs.js';

const para = (text: string, startIndex: number, endIndex: number) => ({
  startIndex,
  endIndex,
  paragraph: { elements: [{ startIndex, endIndex, textRun: { content: text } }] },
});

const tabbedDoc = {
  tabs: [
    {
      tabProperties: { tabId: 't.0', title: 'First', index: 0 },
      documentTab: {
        body: { content: [para('alpha\n', 1, 7)] },
        inlineObjects: { 'kix.a': { marker: 1 } },
      },
      childTabs: [
        {
          tabProperties: { tabId: 't.child', title: 'Nested', index: 0, nestingLevel: 1 },
          documentTab: { body: { content: [para('nested\n', 1, 8)] } },
        },
      ],
    },
    {
      tabProperties: { tabId: 't.1', title: 'Second', index: 1 },
      documentTab: { body: { content: [para('beta\n', 1, 6)] } },
    },
  ],
};

describe('resolveTab', () => {
  it('returns the root body with no tabId for untabbed documents', () => {
    const doc = { body: { content: [para('hi\n', 1, 4)] }, inlineObjects: { 'kix.x': {} } };
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

describe('endOfBodyIndex', () => {
  it('returns one before the final endIndex', () => {
    expect(endOfBodyIndex([para('a\n', 1, 3), para('b\n', 3, 5)])).toBe(4);
  });

  it('never goes below the body start sentinel', () => {
    expect(endOfBodyIndex([{ endIndex: 1 }])).toBe(1);
  });

  it('throws on an empty body instead of crashing', () => {
    expect(() => endOfBodyIndex([])).toThrow(/empty or unreadable/);
  });
});

describe('renderStructureText / renderTabText', () => {
  it('concatenates paragraph text', () => {
    expect(renderTabText([para('one\n', 1, 5), para('two\n', 5, 9)])).toBe('one\ntwo\n');
  });

  it('renders tables as tab-separated rows', () => {
    const text = renderStructureText([
      { type: 'paragraph', startIndex: 1, endIndex: 5, text: 'hi\n' },
      { type: 'table', startIndex: 5, endIndex: 20, cells: [['a', 'b'], ['c', 'd']] },
    ]);
    expect(text).toBe('hi\na\tb\nc\td\n');
  });

  it('skips non-text elements', () => {
    const text = renderStructureText([
      { type: 'sectionBreak', startIndex: 0, endIndex: 1 },
      { type: 'inlineImage', startIndex: 3, endIndex: 4 },
      { type: 'paragraph', startIndex: 1, endIndex: 5, text: 'x\n' },
    ]);
    expect(text).toBe('x\n');
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
});

describe('unknownTabError', () => {
  it('mentions the URL param so agents can self-correct', () => {
    const err = unknownTabError('t.x', [{ tabId: 't.0', title: 'Main' }]);
    expect(err.message).toMatch(/\?tab=/);
  });
});
