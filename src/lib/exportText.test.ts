import { describe, expect, it } from 'vitest';
import {
  anchorThreads,
  compactMarkdownExport,
  findLetterByContent,
  normalizeForMatch,
  parseExportArtifacts,
} from './exportText.js';

describe('parseExportArtifacts', () => {
  it('returns the input unchanged when there is no trailing letter block', () => {
    const parsed = parseExportArtifacts('Just a plain doc.\nNo markers.');
    expect(parsed.body).toBe('Just a plain doc.\nNo markers.');
    expect(parsed.letterContent.size).toBe(0);
    expect(parsed.letterReactions.size).toBe(0);
    expect(parsed.commentLetters.size).toBe(0);
  });

  it('strips the trailing letter block and captures comment text + reactions', () => {
    const text = [
      'Body paragraph one.',
      'Body paragraph two with [a] marker.',
      '',
      '[a]First comment',
      'Alice reacted with 👍 at 2024-01-02 14:30',
      'Bob reacted with 🎉 at 2024-01-02 15:00 PM',
      '[b]Second comment text',
    ].join('\n');
    const parsed = parseExportArtifacts(text);
    expect(parsed.body).toBe('Body paragraph one.\nBody paragraph two with [a] marker.');
    expect(parsed.letterContent.get('a')).toBe('First comment');
    expect(parsed.letterContent.get('b')).toBe('Second comment text');
    expect(parsed.commentLetters).toEqual(new Set(['a', 'b']));
    const rx = parsed.letterReactions.get('a')!;
    expect(rx).toHaveLength(2);
    expect(rx[0]).toEqual({ author_name: 'Alice', emoji: '👍', timestamp: '2024-01-02T14:30' });
    // AM/PM suffix dropped; HH is already 24h.
    expect(rx[1]).toEqual({ author_name: 'Bob', emoji: '🎉', timestamp: '2024-01-02T15:00' });
    expect(parsed.letterReactions.get('b')).toEqual([]);
  });
});

describe('normalizeForMatch', () => {
  it('collapses runs of whitespace and trims', () => {
    expect(normalizeForMatch('  a   b\n\tc  ')).toBe('a b c');
  });
});

describe('findLetterByContent', () => {
  it('finds the letter whose normalised content matches', () => {
    const map = new Map([
      ['a', 'Hello world'],
      ['b', '  Goodbye  world  '],
    ]);
    expect(findLetterByContent(map, 'Hello\tworld')).toBe('a');
    expect(findLetterByContent(map, 'Goodbye world')).toBe('b');
  });

  it('returns undefined for empty / no-match input', () => {
    const map = new Map([['a', 'Hello']]);
    expect(findLetterByContent(map, '')).toBeUndefined();
    expect(findLetterByContent(map, 'nope')).toBeUndefined();
  });
});

describe('compactMarkdownExport', () => {
  it('replaces embedded base64 images with a placeholder and counts them', () => {
    const md = '![][image1]\n\n[image1]: <data:image/png;base64,iVBORw0KGgo=>\n![x](data:image/jpeg;base64,/9j/4AAQ+/=)';
    expect(compactMarkdownExport(md)).toEqual({
      content: '![][image1]\n\n[image1]: <embedded-image-omitted>\n![x](embedded-image-omitted)',
      imagesOmitted: 2,
      truncated: false,
    });
  });

  it('caps the text after stripping images', () => {
    expect(compactMarkdownExport('abcdef', 4)).toEqual({ content: 'abcd', imagesOmitted: 0, truncated: true });
    expect(compactMarkdownExport('abcd', 4)).toEqual({ content: 'abcd', imagesOmitted: 0, truncated: false });
  });
});

describe('anchorThreads', () => {
  const comment = (id: string, content: string, quote?: string) => ({
    id, content, createdTime: '2024-01-01T00:00:00Z', ...(quote ? { quotedFileContent: { value: quote } } : {}),
  });

  it('anchors each thread to its quote, strips the markers after it and orders threads by position', () => {
    const text = 'Alpha[a] beta gamma[b]\n\n[a]On alpha\n[b]On gamma';
    const { cleanedContent, threads } = anchorThreads(text, [comment('c2', 'On gamma', 'gamma'), comment('c1', 'On alpha', 'Alpha')]);
    expect(cleanedContent).toBe('Alpha beta gamma');
    expect(threads.map((t) => [t.thread_id, t.anchor_offset])).toEqual([
      ['c1', { start: 0, end: 5 }],
      ['c2', { start: 11, end: 16 }],
    ]);
  });

  it('keeps text rendered from the structure whole, since it has no export footer', () => {
    const text = 'Alpha\n[a]Looks like a footer line\n';
    const { cleanedContent, threads } = anchorThreads(text, [comment('c1', 'On alpha', 'Alpha')], false);
    expect(cleanedContent).toBe(text);
    expect(threads[0].anchor_offset).toEqual({ start: 0, end: 5 });
  });

  it('leaves a thread without a findable quote unanchored and last', () => {
    const { cleanedContent, threads } = anchorThreads('Plain text', [comment('c1', 'General'), comment('c2', 'Gone', 'deleted words')]);
    expect(cleanedContent).toBe('Plain text');
    expect(threads.every((t) => t.anchor_offset === undefined)).toBe(true);
  });

  it('returns the body alone when there are no comments', () => {
    expect(anchorThreads('Body\n\n[a]x', [])).toEqual({ cleanedContent: 'Body', threads: [] });
  });
});
