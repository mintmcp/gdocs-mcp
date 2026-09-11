/**
 * Pure helpers for Google Docs Tabs support: resolving which tab a read or
 * write targets, building tab-aware Location/Range objects, and rendering a
 * tab's plain text from its parsed structure.
 *
 * Like docs-structure.ts, everything here is referentially transparent so it
 * can be unit-tested without mocking the network.
 */

import {
  parseDocumentStructure,
  summarizeTabs,
  type StructureElement,
  type TabSummary,
} from './docs-structure.js';

/**
 * The body a read/write targets after tab resolution. `tabId` is undefined
 * only for documents without tabs; on tabbed documents it is always explicit
 * (defaulting to the first tab) so outbound requests never rely on Google's
 * inconsistent implicit defaults (location omits-tabId → first tab, but
 * replaceAllText omits-tabsCriteria → ALL tabs).
 */
export interface ResolvedTab {
  tabId?: string;
  title?: string;
  body: any[];
  inlineObjects: Record<string, any>;
  tabs: TabSummary[];
}

/** Depth-first search of the nested tab tree for a tabId. */
function findTabNode(tabs: any[] | undefined, tabId: string): any | undefined {
  for (const tab of tabs || []) {
    if (tab?.tabProperties?.tabId === tabId) return tab;
    const child = findTabNode(tab?.childTabs, tabId);
    if (child) return child;
  }
  return undefined;
}

/**
 * Error for an unknown tab_id, listing the tabs that do exist so the caller
 * can correct itself without another read.
 */
export function unknownTabError(tabId: string, tabs: TabSummary[]): Error {
  const listed = tabs.map((t) => t.title ? `${t.tabId} (${t.title})` : t.tabId).join(', ');
  return new Error(
    tabs.length === 0
      ? `tab_id '${tabId}' was passed but this document has no tabs. Omit tab_id.`
      : `No tab with id '${tabId}'. This document's tabs are: ${listed}. ` +
        `The tab id is the value after ?tab= in the document URL.`
  );
}

/**
 * Resolve which body a read/write targets. The document must have been
 * fetched with `includeTabsContent=true` for tabbed documents to carry
 * their content.
 *
 * - untabbed doc, no tabId  → root body, tabId stays undefined
 * - untabbed doc, tabId     → throws
 * - tabbed doc, no tabId    → first tab, with its tabId made explicit
 * - tabbed doc, tabId       → that tab, or throws listing the tabs
 */
export function resolveTab(doc: any, tabId?: string): ResolvedTab {
  const tabs = summarizeTabs(doc?.tabs);

  if (tabs.length === 0) {
    if (tabId) throw unknownTabError(tabId, tabs);
    return { body: doc?.body?.content || [], inlineObjects: doc?.inlineObjects || {}, tabs };
  }

  const target = tabId ?? tabs[0].tabId;
  const node = findTabNode(doc.tabs, target);
  if (!node) throw unknownTabError(target, tabs);

  return {
    tabId: target,
    title: node.tabProperties?.title,
    body: node.documentTab?.body?.content || [],
    inlineObjects: node.documentTab?.inlineObjects || {},
    tabs,
  };
}

/** Location for insertText/insertTable; tabId only when targeting a tab. */
export function docLocation(index: number, tabId?: string): { index: number; tabId?: string } {
  return { index, ...(tabId ? { tabId } : {}) };
}

/** Range for delete/style requests; tabId only when targeting a tab. */
export function docRange(
  startIndex: number,
  endIndex: number,
  tabId?: string,
): { startIndex: number; endIndex: number; tabId?: string } {
  return { startIndex, endIndex, ...(tabId ? { tabId } : {}) };
}

/**
 * End-of-body insertion index (one before the final newline sentinel).
 * Throws when the body is empty rather than letting callers index into
 * `content[length - 1]` and crash.
 */
export function endOfBodyIndex(body: any[]): number {
  const last = body[body.length - 1];
  if (!last || typeof last.endIndex !== 'number') {
    throw new Error('Document body is empty or unreadable; cannot find the end index.');
  }
  return Math.max(1, last.endIndex - 1);
}

/**
 * Render a tab's plain text from its parsed structure, so per-tab `content`
 * comes from the same source as the indices the mutating tools take. Tables
 * render as tab-separated rows; images, section breaks and tables of contents
 * contribute no text.
 */
export function renderStructureText(elements: StructureElement[]): string {
  let out = '';
  for (const el of elements) {
    if (el.type === 'paragraph') {
      out += el.text ?? '';
    } else if (el.type === 'table' && el.cells) {
      out += el.cells.map((row) => row.join('\t')).join('\n') + '\n';
    }
  }
  return out;
}

/** Convenience: parse + render a resolved tab's text in one step. */
export function renderTabText(body: any[]): string {
  return renderStructureText(parseDocumentStructure(body).elements);
}

/**
 * Gather inline objects from every tab (plus the legacy root field), for
 * whole-document image extraction. Tabbed documents keep per-tab
 * `documentTab.inlineObjects`; the root field only exists on responses
 * fetched without `includeTabsContent`.
 */
export function collectInlineObjects(doc: any): Record<string, any> {
  const out: Record<string, any> = { ...(doc?.inlineObjects || {}) };
  const walk = (tabs: any[] | undefined) => {
    for (const tab of tabs || []) {
      Object.assign(out, tab?.documentTab?.inlineObjects || {});
      walk(tab?.childTabs);
    }
  };
  walk(doc?.tabs);
  return out;
}
