/**
 * word-extractor ships no type declarations. Only the accessors this codebase
 * uses are declared; see https://github.com/morungos/node-word-extractor.
 */
declare module 'word-extractor' {
  import type { Buffer } from 'node:buffer';

  /** Each accessor returns the text of one stream, or '' when absent. */
  class Document {
    getBody(): string;
    getFootnotes(): string;
    getEndnotes(): string;
    getHeaders(options?: { includeFooters?: boolean }): string;
    getFooters(): string;
    getAnnotations(): string;
    getTextboxes(options?: { includeHeadersAndFooters?: boolean; includeBody?: boolean }): string;
  }

  class WordExtractor {
    extract(source: string | Buffer): Promise<Document>;
  }

  export default WordExtractor;
  export { Document };
}
