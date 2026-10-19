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
  documentStyle: any;
  headers: Record<string, any>;
  footers: Record<string, any>;
  footnotes: Record<string, any>;
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

/** Inline objects, headers and footers live beside the body, on the tab or the root. */
function tabSegments(container: any): Pick<ResolvedTab, 'inlineObjects' | 'documentStyle' | 'headers' | 'footers' | 'footnotes'> {
  return {
    inlineObjects: container?.inlineObjects || {},
    documentStyle: container?.documentStyle || {},
    headers: container?.headers || {},
    footers: container?.footers || {},
    footnotes: container?.footnotes || {},
  };
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
    return { body: doc?.body?.content || [], ...tabSegments(doc), tabs };
  }

  const target = tabId ?? tabs[0].tabId;
  const node = findTabNode(doc.tabs, target);
  if (!node) throw unknownTabError(target, tabs);

  return {
    tabId: target,
    title: node.tabProperties?.title,
    body: node.documentTab?.body?.content || [],
    ...tabSegments(node.documentTab),
    tabs,
  };
}

/**
 * Unscoped replaceAllText hits ALL tabs, unlike every other request type;
 * pin it to one explicit tab so replace writes where the other tools write
 */
export function tabsCriteriaFor(tabs: TabSummary[], tabId?: string): { tabIds: string[] } | undefined {
  if (tabs.length === 0) {
    if (tabId) throw unknownTabError(tabId, tabs);
    return undefined;
  }
  const target = tabId ?? tabs[0].tabId;
  if (!tabs.some((t) => t.tabId === target)) throw unknownTabError(target, tabs);
  return { tabIds: [target] };
}

/** Where an insert goes: a location when the index is given, else the end of the body. */
export function insertionPoint(index: number | undefined, tabId?: string) {
  return index === undefined
    ? { endOfSegmentLocation: { ...(tabId ? { tabId } : {}) } }
    : { location: docLocation(index, tabId) };
}

/** The tab whose inlineObjects hold `objectId`; undefined on untabbed documents or when absent. */
export function findInlineObjectTab(doc: any, objectId: string): string | undefined {
  const walk = (tabs: any[] | undefined): string | undefined => {
    for (const tab of tabs || []) {
      if (tab?.documentTab?.inlineObjects?.[objectId]) return tab.tabProperties?.tabId;
      const nested = walk(tab?.childTabs);
      if (nested) return nested;
    }
    return undefined;
  };
  return walk(doc?.tabs);
}

/** segmentId names a header, footer or footnote; omitted it means the body. */
export function docLocation(index: number, tabId?: string, segmentId?: string): { index: number; tabId?: string; segmentId?: string } {
  return { index, ...(tabId ? { tabId } : {}), ...(segmentId ? { segmentId } : {}) };
}

export function docRange(
  startIndex: number,
  endIndex: number,
  tabId?: string,
  segmentId?: string,
): { startIndex: number; endIndex: number; tabId?: string; segmentId?: string } {
  return { startIndex, endIndex, ...(tabId ? { tabId } : {}), ...(segmentId ? { segmentId } : {}) };
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
 * Google inserts end-of-body text before the final newline, so without a
 * separator it joins the last paragraph. Callers who pass their own leading
 * newline keep control of the spacing
 */
export function appendInsertion(body: any[], text: string): { index: number; text: string; paragraphStart: number } {
  const index = endOfBodyIndex(body);
  const elements: any[] = body[body.length - 1]?.paragraph?.elements ?? [];
  const lastHasContent = elements.some((el) => el.textRun ? el.textRun.content.replace(/\n$/, '') !== '' : true);
  const separate = lastHasContent && !text.startsWith('\n');
  // A caller's own leading newline also ends the existing paragraph, so the new one starts after it
  return { index, text: separate ? `\n${text}` : text, paragraphStart: lastHasContent ? index + 1 : index };
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
