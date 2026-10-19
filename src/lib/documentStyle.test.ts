import { describe, expect, it } from 'vitest';
import { buildDocumentStyle } from './documentStyle.js';

const LETTER = { width: { magnitude: 612 }, height: { magnitude: 792 } };

describe('buildDocumentStyle', () => {
  it('builds nothing from no arguments', () => {
    expect(buildDocumentStyle({}, LETTER)).toEqual({ documentStyle: {}, fields: [] });
  });

  it('rotates a named size for landscape and leaves portrait alone', () => {
    expect(buildDocumentStyle({ page_size: 'A4', orientation: 'landscape' }, undefined).documentStyle.pageSize)
      .toEqual({ width: { magnitude: 841.9, unit: 'PT' }, height: { magnitude: 595.3, unit: 'PT' } });
    expect(buildDocumentStyle({ page_size: 'A4', orientation: 'portrait' }, undefined).documentStyle.pageSize)
      .toEqual({ width: { magnitude: 595.3, unit: 'PT' }, height: { magnitude: 841.9, unit: 'PT' } });
  });

  it('keeps an already-landscape page when asked for landscape again', () => {
    const landscape = { width: { magnitude: 792 }, height: { magnitude: 612 } };
    expect(buildDocumentStyle({ orientation: 'landscape' }, landscape).documentStyle.pageSize)
      .toEqual({ width: { magnitude: 792, unit: 'PT' }, height: { magnitude: 612, unit: 'PT' } });
  });

  it('needs a page size to rotate', () => {
    expect(() => buildDocumentStyle({ orientation: 'landscape' }, undefined)).toThrow(/pass page_size/);
  });
});
