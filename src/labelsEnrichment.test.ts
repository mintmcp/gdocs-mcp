import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import { GoogleDocsTools } from './tools.js';
import { requestContext } from './auth.js';
import { stubFetch, jsonResponse, LABEL_SCHEMA_BODY } from './lib/labelStubs.js';

const LISTED_LABEL = {
  labels: [{
    id: 'lbl1', revisionId: 'rev7',
    fields: { field1: { valueType: 'selection', selection: ['choiceA'] } },
  }],
};

const tools = GoogleDocsTools.getTools() as any;
const call = (tool: string, args: any) =>
  requestContext.run({ accessToken: 'tok' }, () => tools[tool].handler(args));

describe('label enrichment on document reads', () => {
  const savedProfile = process.env.PROFILE;
  beforeEach(() => {
    delete process.env.PROFILE;
  });
  afterEach(() => {
    vi.unstubAllGlobals();
    vi.unstubAllEnvs();
    if (savedProfile === undefined) delete process.env.PROFILE;
    else process.env.PROFILE = savedProfile;
  });

  it('get_document_images carries _meta.applied when enrichment is on', async () => {
    stubFetch([
      ['listLabels', () => jsonResponse(LISTED_LABEL)],
      ['drivelabels.googleapis.com', () => jsonResponse({ properties: { title: 'Classification' }, ...LABEL_SCHEMA_BODY })],
      ['docs.googleapis.com', () => jsonResponse({})],
    ]);
    const res: any = await call('get_document_images', { document_id: 'd1' });
    expect(res.isError).toBeUndefined();
    expect(res.content[0].text).toMatch(/No images found/);
    expect(res._meta.applied).toEqual([{
      labelId: 'lbl1', revisionId: 'rev7', title: 'Classification', resolved: true,
      values: [{ fieldId: 'field1', valueType: 'selection', choiceId: 'choiceA', displayName: 'Confidential', resolved: true }],
    }]);
  });

  it('a standard deployment makes no label calls and returns no _meta', async () => {
    vi.stubEnv('PROFILE', 'standard');
    const calls = stubFetch([
      ['docs.googleapis.com', () => jsonResponse({})],
    ]);
    const res: any = await call('get_document_images', { document_id: 'd1' });
    expect(res.isError).toBeUndefined();
    expect(res._meta).toBeUndefined();
    expect(calls.some((c) => c.url.includes('listLabels'))).toBe(false);
  });

  it('get_document error envelopes still carry _meta', async () => {
    stubFetch([
      ['listLabels', () => jsonResponse(LISTED_LABEL)],
      ['drivelabels.googleapis.com', () => jsonResponse({ properties: { title: 'Classification' }, ...LABEL_SCHEMA_BODY })],
      ['www.googleapis.com/drive', () => jsonResponse({ error: { message: 'boom' } }, 500)],
    ]);
    const res: any = await call('get_document', { document_id: 'd1' });
    expect(res.isError).toBe(true);
    expect(res.structuredContent).toBeUndefined();
    expect(res._meta.applied).toHaveLength(1);
  });

  it('a malformed label wire degrades instead of failing the read', async () => {
    stubFetch([
      ['listLabels', () => jsonResponse({ labels: [null] })],
      ['docs.googleapis.com', () => jsonResponse({})],
    ]);
    const res: any = await call('get_document_images', { document_id: 'd1' });
    expect(res.isError).toBeUndefined();
    expect(res._meta.labelsError).toBe('label read failed');
    expect(res._meta.applied).toEqual([]);
  });
});
