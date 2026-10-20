export interface Reaction {
  author_name: string;
  emoji: string;
  timestamp: string;
}

export interface DriveCommentReply {
  id: string;
  content: string;
  author?: { displayName?: string; emailAddress?: string };
  createdTime: string;
}

export interface DriveComment extends DriveCommentReply {
  quotedFileContent?: { value: string };
  resolved?: boolean;
  replies?: DriveCommentReply[];
}

export interface ThreadOutput {
  // Half-open span into `content`; absent for document-level comments and quotes no longer in the body
  anchor_offset?: { start: number; end: number };
  /** Drive's comment id, which is also the `disco=` deep-link parameter. */
  thread_id: string;
  resolved: boolean;
  comments: Array<{
    author_email?: string;
    author_name?: string;
    timestamp: string;
    content: string;
    reactions?: Reaction[];
  }>;
}

export interface ParsedExport {
  body: string;
  letterContent: Map<string, string>;
  letterReactions: Map<string, Reaction[]>;
  commentLetters: Set<string>;
}

export function normalizeForMatch(text: string): string {
  return text.replace(/\s+/g, ' ').trim();
}

// Reactions exist only in the export footer, as `<name> reacted with <emoji> at YYYY-MM-DD HH:MM (AM|PM)`, HH already 24h
const REACTION_RE = /^(.+?) reacted with (\S+) at (\d{4}-\d{2}-\d{2} \d{2}:\d{2})(?:\s*[AP]M)?\s*$/;

// A `[a]` typed in the body looks like a comment marker, so only the footer is stripped here
export function parseExportArtifacts(plainText: string): ParsedExport {
  const lines = plainText.split('\n');

  let blockStart = -1;
  for (let candidate = 0; candidate < lines.length; candidate++) {
    if (!/^\[[a-z]+\]/.test(lines[candidate])) continue;
    let inMarker = false;
    let valid = true;
    for (let i = candidate; i < lines.length; i++) {
      const line = lines[i];
      if (line.trim() === '') continue;
      if (/^\[[a-z]+\]/.test(line)) { inMarker = true; continue; }
      if (!inMarker) { valid = false; break; }
    }
    if (valid) { blockStart = candidate; break; }
  }

  const letterContent = new Map<string, string>();
  const letterReactions = new Map<string, Reaction[]>();
  const commentLetters = new Set<string>();

  if (blockStart >= 0) {
    let currentLetter: string | null = null;
    for (let i = blockStart; i < lines.length; i++) {
      const line = lines[i].replace(/\r$/, '');
      const m = line.match(/^\[([a-z]+)\](.*)$/);
      if (m) {
        currentLetter = m[1];
        commentLetters.add(currentLetter);
        letterContent.set(currentLetter, m[2].trim());
        letterReactions.set(currentLetter, []);
      } else if (currentLetter !== null && line.trim() !== '') {
        const rx = line.match(REACTION_RE);
        if (rx) {
          letterReactions.get(currentLetter)!.push({
            author_name: rx[1].trim(),
            emoji: rx[2],
            timestamp: rx[3].replace(' ', 'T'),
          });
        }
      }
    }
  }

  const body = blockStart >= 0
    ? lines.slice(0, blockStart).join('\n')
    : plainText;

  return {
    body: body.replace(/[\s\r\n]+$/, ''),
    letterContent,
    letterReactions,
    commentLetters,
  };
}

export function findLetterByContent(
  letterContent: Map<string, string>,
  apiContent: string,
): string | undefined {
  const target = normalizeForMatch(apiContent);
  if (!target) return undefined;
  for (const [letter, content] of letterContent) {
    if (normalizeForMatch(content) === target) return letter;
  }
  return undefined;
}

export function anchorThreads(
  plainText: string,
  allComments: DriveComment[],
  fromExport = true,
): { cleanedContent: string; threads: ThreadOutput[] } {
  const { body: bodyOnly, letterContent, letterReactions, commentLetters }: ParsedExport = fromExport
    ? parseExportArtifacts(plainText)
    : { body: plainText, letterContent: new Map(), letterReactions: new Map(), commentLetters: new Set() };

  const reactionsFor = (apiContent: string): Reaction[] | undefined => {
    const letter = findLetterByContent(letterContent, apiContent);
    const list = letter ? letterReactions.get(letter) : undefined;
    return list && list.length > 0 ? list : undefined;
  };

  type Positioned = {
    cmt: DriveComment;
    quoteStart: number | null;
    quoteEnd: number | null;
    commentCount: number;
  };
  const positioned: Positioned[] = allComments.map((cmt) => {
    const quote = cmt.quotedFileContent?.value;
    const idx = quote ? bodyOnly.indexOf(quote) : -1;
    return {
      cmt,
      quoteStart: (quote && idx >= 0) ? idx : null,
      quoteEnd: (quote && idx >= 0) ? idx + quote.length : null,
      commentCount: 1 + (cmt.replies?.length || 0),
    };
  });

  type Group = { pos: number; expectedMarkers: number; entries: Positioned[] };
  const groups = new Map<number, Group>();
  for (const p of positioned) {
    if (p.quoteEnd === null) continue;
    let g = groups.get(p.quoteEnd);
    if (!g) {
      g = { pos: p.quoteEnd, expectedMarkers: 0, entries: [] };
      groups.set(p.quoteEnd, g);
    }
    g.expectedMarkers += p.commentCount;
    g.entries.push(p);
  }

  const ascendingPositions = [...groups.keys()].sort((a, b) => a - b);
  let shift = 0;
  let body = bodyOnly;
  const finalOffsets = new Map<Positioned, { start: number; end: number }>();

  for (const pos of ascendingPositions) {
    const group = groups.get(pos)!;
    const adjustedEnd = pos + shift;
    let end = adjustedEnd;
    let consumed = 0;
    while (consumed < group.expectedMarkers && end < body.length) {
      const m = body.slice(end).match(/^\[([a-z]+)\]/);
      if (!m || !commentLetters.has(m[1])) break;
      end += m[0].length;
      consumed++;
    }
    const removedLen = end - adjustedEnd;
    body = body.slice(0, adjustedEnd) + body.slice(end);

    for (const entry of group.entries) {
      finalOffsets.set(entry, {
        start: entry.quoteStart! + shift,
        end: adjustedEnd,
      });
    }
    shift -= removedLen;
  }

  const ordered = [...positioned].sort((a, b) => {
    if (a.quoteEnd === null && b.quoteEnd === null) return 0;
    if (a.quoteEnd === null) return 1;
    if (b.quoteEnd === null) return -1;
    return a.quoteEnd! - b.quoteEnd!;
  });

  const threads: ThreadOutput[] = ordered.map((p) => ({
    thread_id: p.cmt.id,
    anchor_offset: finalOffsets.get(p),
    resolved: !!p.cmt.resolved,
    comments: [
      {
        author_email: p.cmt.author?.emailAddress,
        author_name: p.cmt.author?.displayName,
        timestamp: p.cmt.createdTime,
        content: p.cmt.content,
        reactions: reactionsFor(p.cmt.content),
      },
      ...(p.cmt.replies || []).map((r) => ({
        author_email: r.author?.emailAddress,
        author_name: r.author?.displayName,
        timestamp: r.createdTime,
        content: r.content,
        reactions: reactionsFor(r.content),
      })),
    ],
  }));

  return { cleanedContent: body, threads };
}

export const MAX_MARKDOWN_CHARS = 200_000;

const EMBEDDED_IMAGE_RE = /data:image\/[\w.+-]+;base64,[A-Za-z0-9+/=]+/g;

export function compactMarkdownExport(markdown: string, maxChars = MAX_MARKDOWN_CHARS): { content: string; imagesOmitted: number; truncated: boolean } {
  let imagesOmitted = 0;
  const stripped = markdown.replace(EMBEDDED_IMAGE_RE, () => {
    imagesOmitted++;
    return 'embedded-image-omitted';
  });
  const truncated = stripped.length > maxChars;
  return { content: truncated ? stripped.slice(0, maxChars) : stripped, imagesOmitted, truncated };
}

