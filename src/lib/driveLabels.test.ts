import { describe, it, expect, beforeEach, afterEach, vi } from 'vitest';
import { getLabelSchema, getFileLabels, clearLabelSchemaCache } from './driveLabels.js';

/**
 * Minimal fetch stub: routes by substring match on the URL, in order.
 * Each route's handler returns a Response; unmatched URLs fail the test.
 */
function stubFetch(routes: Array<[string, (url: string) => Response]>) {
  const calls: string[] = [];
  vi.stubGlobal('fetch', vi.fn(async (input: any) => {
    const url = String(input);
    calls.push(url);
    for (const [substr, handler] of routes) {
      if (url.includes(substr)) return handler(url);
    }
    throw new Error(`unexpected fetch: ${url}`);
  }));
  return calls;
}

function jsonResponse(body: unknown, status = 200): Response {
  return new Response(JSON.stringify(body), {
    status,
    headers: { 'Content-Type': 'application/json' },
  });
}

const LABEL_SCHEMA_BODY = {
  fields: [
    {
      id: 'field1',
      selectionOptions: {
        choices: [
          { id: 'choiceA', properties: { displayName: 'Confidential' } },
          { id: 'choiceB', properties: { displayName: 'Internal' } },
          { id: 'choiceNoName', properties: {} },
        ],
      },
    },
    { id: 'textField' }, // non-selection field — no choices to map
  ],
};

describe('getLabelSchema', () => {
  beforeEach(() => clearLabelSchemaCache());
  afterEach(() => vi.unstubAllGlobals());

  it('maps fieldId -> choiceId -> displayName and skips nameless choices', async () => {
    stubFetch([['drivelabels.googleapis.com', () => jsonResponse(LABEL_SCHEMA_BODY)]]);
    const schema = await getLabelSchema('lbl1', 'rev7', 'tok');
    expect(schema).toEqual({ field1: { choiceA: 'Confidential', choiceB: 'Internal' } });
  });

  it('requests the exact applied revision with LABEL_VIEW_FULL', async () => {
    const calls = stubFetch([['drivelabels.googleapis.com', () => jsonResponse(LABEL_SCHEMA_BODY)]]);
    await getLabelSchema('lbl1', 'rev7', 'tok');
    expect(calls[0]).toContain('/labels/lbl1@rev7?view=LABEL_VIEW_FULL');
  });

  it('caches by label@revision — second call does not refetch', async () => {
    const calls = stubFetch([['drivelabels.googleapis.com', () => jsonResponse(LABEL_SCHEMA_BODY)]]);
    await getLabelSchema('lbl1', 'rev7', 'tok');
    await getLabelSchema('lbl1', 'rev7', 'tok');
    expect(calls.length).toBe(1);
    await getLabelSchema('lbl1', 'rev8', 'tok');
    expect(calls.length).toBe(2);
  });
});

describe('getFileLabels', () => {
  beforeEach(() => clearLabelSchemaCache());
  afterEach(() => vi.unstubAllGlobals());

  const appliedSelection = {
    id: 'lbl1',
    revisionId: 'rev7',
    fields: {
      field1: { valueType: 'selection', selection: ['choiceA'] },
    },
  };

  it('returns [] and no error for an unlabeled file', async () => {
    stubFetch([['listLabels', () => jsonResponse({ labels: [] })]]);
    const res = await getFileLabels('f1', 'tok');
    expect(res).toEqual({ labels: [] });
  });

  it('resolves selection choice ids to display names', async () => {
    stubFetch([
      ['listLabels', () => jsonResponse({ labels: [appliedSelection] })],
      ['drivelabels.googleapis.com', () => jsonResponse(LABEL_SCHEMA_BODY)],
    ]);
    const res = await getFileLabels('f1', 'tok');
    expect(res).toEqual({ labels: ['Confidential'] });
  });

  it('surfaces text-field values directly and dedupes across labels', async () => {
    const applied = [
      appliedSelection,
      {
        id: 'lbl2',
        fields: { t: { valueType: 'text', text: ['Confidential', 'Legal Hold'] } },
      },
    ];
    stubFetch([
      ['listLabels', () => jsonResponse({ labels: applied })],
      ['drivelabels.googleapis.com', () => jsonResponse(LABEL_SCHEMA_BODY)],
    ]);
    const res = await getFileLabels('f1', 'tok');
    expect(res.labels.sort()).toEqual(['Confidential', 'Legal Hold']);
    expect(res.error).toBeUndefined();
  });

  it('skips date/integer/user fields as non-classificatory', async () => {
    const applied = [{
      id: 'lbl3',
      fields: {
        d: { valueType: 'dateString', dateString: ['2026-01-01'] },
        i: { valueType: 'integer', integer: ['5'] },
      },
    }];
    stubFetch([['listLabels', () => jsonResponse({ labels: applied })]]);
    const res = await getFileLabels('f1', 'tok');
    expect(res).toEqual({ labels: [] });
  });

  it('follows listLabels pagination', async () => {
    let page = 0;
    stubFetch([
      ['listLabels', () => {
        page += 1;
        return page === 1
          ? jsonResponse({ labels: [], nextPageToken: 'p2' })
          : jsonResponse({ labels: [{ id: 'lbl2', fields: { t: { valueType: 'text', text: ['Internal'] } } }] });
      }],
    ]);
    const res = await getFileLabels('f1', 'tok');
    expect(page).toBe(2);
    expect(res.labels).toEqual(['Internal']);
  });

  it('flags "label read failed" when listLabels errors (e.g. missing scope 403)', async () => {
    stubFetch([
      ['listLabels', () => jsonResponse({ error: { message: 'insufficient scope' } }, 403)],
    ]);
    const res = await getFileLabels('f1', 'tok');
    expect(res.labels).toEqual([]);
    expect(res.error).toBe('label read failed');
  });

  it('surfaces the raw choice id and flags "incomplete label resolution" when the schema lacks the choice', async () => {
    const applied = [{
      id: 'lbl1',
      revisionId: 'rev7',
      fields: { field1: { valueType: 'selection', selection: ['choiceUnknown'] } },
    }];
    stubFetch([
      ['listLabels', () => jsonResponse({ labels: applied })],
      ['drivelabels.googleapis.com', () => jsonResponse(LABEL_SCHEMA_BODY)],
    ]);
    const res = await getFileLabels('f1', 'tok');
    expect(res.labels).toEqual(['choiceUnknown']);
    expect(res.error).toBe('incomplete label resolution');
  });

  it('keeps the stronger "label read failed" signal when a later choice is also unresolved', async () => {
    let page = 0;
    stubFetch([
      ['listLabels', () => {
        page += 1;
        return page === 1
          ? jsonResponse({
              labels: [{ id: 'lbl1', revisionId: 'rev7', fields: { field1: { valueType: 'selection', selection: ['choiceUnknown'] } } }],
              nextPageToken: 'p2',
            })
          : jsonResponse({ error: { message: 'insufficient scope' } }, 403);
      }],
      ['drivelabels.googleapis.com', () => jsonResponse(LABEL_SCHEMA_BODY)],
    ]);
    const res = await getFileLabels('f1', 'tok');
    expect(res.labels).toEqual(['choiceUnknown']);
    expect(res.error).toBe('label read failed');
  });

  it('keeps values from other labels when one schema fetch fails', async () => {
    const applied = [
      { id: 'lblGood', fields: { t: { valueType: 'text', text: ['Internal'] } } },
      appliedSelection,
    ];
    stubFetch([
      ['listLabels', () => jsonResponse({ labels: applied })],
      ['drivelabels.googleapis.com', () => jsonResponse({ error: { message: 'nope' } }, 403)],
    ]);
    const res = await getFileLabels('f1', 'tok');
    expect(res.labels).toEqual(['Internal']);
    expect(res.error).toBe('incomplete label resolution');
  });
});
