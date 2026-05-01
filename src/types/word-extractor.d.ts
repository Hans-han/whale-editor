declare module 'word-extractor' {
  export interface ExtractedDocument {
    getBody(): string;
    getHeaders(options?: { includeFooters?: boolean }): string;
    getFooters(): string;
    getAnnotations(): string;
    getEndnotes(): string;
    getFootnotes(): string;
    getTextboxes(options?: { includeHeadersAndFooters?: boolean }): string;
  }

  export default class WordExtractor {
    extract(input: string | Buffer): Promise<ExtractedDocument>;
  }
}
