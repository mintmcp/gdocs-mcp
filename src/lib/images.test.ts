import { describe, expect, it } from 'vitest';
import { imageObjectSize, imageSourceUrl } from './images.js';

describe('imageSourceUrl', () => {
  it('passes a plain https URL through', () => {
    expect(imageSourceUrl('https://example.com/a.png')).toBe('https://example.com/a.png');
  });

  it('rewrites every Drive link shape to the download URL', () => {
    const want = 'https://drive.google.com/uc?export=download&id=1AbC_dEf';
    expect(imageSourceUrl('https://drive.google.com/file/d/1AbC_dEf/view?usp=sharing')).toBe(want);
    expect(imageSourceUrl('https://drive.google.com/open?id=1AbC_dEf')).toBe(want);
    expect(imageSourceUrl('https://drive.google.com/uc?id=1AbC_dEf&export=view')).toBe(want);
  });

  it('refuses http, data: and malformed URLs', () => {
    expect(() => imageSourceUrl('http://example.com/a.png')).toThrow(/https/);
    expect(() => imageSourceUrl('data:image/png;base64,AAAA')).toThrow(/https/);
    expect(() => imageSourceUrl('not a url')).toThrow(/absolute https URL/);
  });

  it('refuses URLs over 2 kB', () => {
    expect(() => imageSourceUrl(`https://example.com/${'a'.repeat(2100)}`)).toThrow(/2048/);
  });
});

describe('imageObjectSize', () => {
  it('builds only the dimensions given', () => {
    expect(imageObjectSize(100, undefined)).toEqual({ objectSize: { width: { magnitude: 100, unit: 'PT' } } });
    expect(imageObjectSize(undefined, 50)).toEqual({ objectSize: { height: { magnitude: 50, unit: 'PT' } } });
    expect(imageObjectSize()).toEqual({});
  });
});
