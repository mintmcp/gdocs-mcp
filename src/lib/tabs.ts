/** Pure helpers for Google Docs Tabs: tab resolution, tab-aware Location/Range objects, per-tab text */

import {
  summarizeTabs,
  type StructureElement,
  type TabSummary,
} from './docs-structure.js';

/**
 * On tabbed documents `tabId` is always explicit (defaulting to the first tab)
 * so requests never rely on Google's inconsistent implicit defaults: a location
 * without tabId targets the first tab, but replaceAllText without tabsCriteria
 * targets ALL tabs
 */
export interface ResolvedTab {
  tabId?: string;
  title?: string;
  body: any[];
  inlineObjects: Record<string, any>;
  tabs: TabSummary[];
}

function findTabNode(tabs: any[] | undefined, tabId: string): any | undefined {
  for (const tab of tabs || []) {
    if (tab?.tabProperties?.tabId === tabId) return tab;
    const child = findTabNode(tab?.childTabs, tabId);
    if (child) return child;
  }
  return undefined;
}

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
 * Resolve which body a read/write targets. The doc must have been fetched with
 * includeTabsContent=true, or tabbed documents arrive without their content
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

export function docLocation(index: number, tabId?: string): { index: number; tabId?: string } {
  return { index, ...(tabId ? { tabId } : {}) };
}

export function docRange(
  startIndex: number,
  endIndex: number,
  tabId?: string,
): { startIndex: number; endIndex: number; tabId?: string } {
  return { startIndex, endIndex, ...(tabId ? { tabId } : {}) };
}

/** End-of-body insertion index, one before the final newline sentinel */
export function endOfBodyIndex(body: any[]): number {
  const last = body[body.length - 1];
  if (!last || typeof last.endIndex !== 'number') {
    throw new Error('Document body is empty or unreadable; cannot find the end index.');
  }
  return Math.max(1, last.endIndex - 1);
}

/**
 * Per-tab text is rendered from the same parsed structure the editing indices
 * come from, because the Drive plain-text export cannot be scoped to a tab
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

/**
 * Tabbed documents keep inline objects per tab under documentTab.inlineObjects;
 * the root field only appears on fetches without includeTabsContent
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
