export interface TabSummary {
  tabId: string;
  title?: string;
  index?: number;
  nestingLevel?: number;
}

function flattenTabs(tabs: any[] | undefined, depth = 0): Array<{ tab: any; depth: number }> {
  return (tabs ?? []).flatMap((tab) => [{ tab, depth }, ...flattenTabs(tab?.childTabs, depth + 1)]);
}

export function summarizeTabs(tabs: any[] | undefined): TabSummary[] {
  return flattenTabs(tabs).flatMap(({ tab, depth }) => {
    const props = tab?.tabProperties || {};
    if (!props.tabId) return [];
    return [{
      tabId: props.tabId,
      ...(props.title !== undefined ? { title: props.title } : {}),
      ...(props.index !== undefined ? { index: props.index } : {}),
      nestingLevel: props.nestingLevel ?? depth,
    }];
  });
}

/**
 * On tabbed documents `tabId` is always explicit (defaulting to the first tab)
 * so requests never rely on Google's inconsistent implicit defaults: a location
 * without tabId targets the first tab, but replaceAllText without tabsCriteria
 * targets ALL tabs
 */
export interface ResolvedTab {
  tabId?: string;
  revisionId?: string;
  title?: string;
  body: any[];
  inlineObjects: Record<string, any>;
  documentStyle: any;
  headers: Record<string, any>;
  footers: Record<string, any>;
  footnotes: Record<string, any>;
  tabs: TabSummary[];
}

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

function targetTabId(tabs: TabSummary[], tabId?: string): string | undefined {
  if (tabs.length === 0) {
    if (tabId) throw unknownTabError(tabId, tabs);
    return undefined;
  }
  const target = tabId ?? tabs[0].tabId;
  if (!tabs.some((t) => t.tabId === target)) throw unknownTabError(target, tabs);
  return target;
}

/**
 * Resolve which body a read/write targets. The doc must have been fetched with
 * includeTabsContent=true, or tabbed documents arrive without their content
 */
export function resolveTab(doc: any, tabId?: string): ResolvedTab {
  const tabs = summarizeTabs(doc?.tabs);
  const target = targetTabId(tabs, tabId);
  if (target === undefined) {
    return { revisionId: doc?.revisionId, body: doc?.body?.content || [], ...tabSegments(doc), tabs };
  }

  const node = flattenTabs(doc.tabs).find(({ tab }) => tab?.tabProperties?.tabId === target)?.tab;
  if (!node) throw unknownTabError(target, tabs);

  return {
    tabId: target,
    revisionId: doc.revisionId,
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
  const target = targetTabId(tabs, tabId);
  return target === undefined ? undefined : { tabIds: [target] };
}

export function findInlineObjectTab(doc: any, objectId: string): string | undefined {
  return flattenTabs(doc?.tabs).find(({ tab }) => tab?.documentTab?.inlineObjects?.[objectId])?.tab.tabProperties?.tabId;
}

/**
 * Tabbed documents keep inline objects per tab under documentTab.inlineObjects;
 * the root field only appears on fetches without includeTabsContent
 */
export function collectInlineObjects(doc: any, tabId?: string): Record<string, any> {
  if (tabId) return resolveTab(doc, tabId).inlineObjects;
  return Object.assign({}, doc?.inlineObjects, ...flattenTabs(doc?.tabs).map(({ tab }) => tab?.documentTab?.inlineObjects));
}
