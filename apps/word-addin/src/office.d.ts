/**
 * Type declarations for Office.js / Word API globals.
 * These are loaded from CDN in index.html, so they are available at runtime
 * but not discoverable by TypeScript without explicit declarations.
 */

declare namespace Word {
  function run<T>(callback: (context: Word.RequestContext) => Promise<T>): Promise<T>;

  interface RequestContext {
    document: Word.Document;
    sync(): Promise<void>;
  }

  interface Document {
    body: Word.Body;
    getSelection(): Word.Range;
  }

  interface Body extends Word.Range {
    insertText(text: string, location: Word.InsertLocation): Word.Range;
    insertParagraph(paragraphText: string, insertLocation: Word.InsertLocation): Word.Paragraph;
    paragraphs: Word.ParagraphCollection;
  }

  interface Range {
    text: string;
    font: Word.Font;
    insertText(text: string, location: Word.InsertLocation): Word.Range;
    insertParagraph(paragraphText: string, insertLocation: Word.InsertLocation): Word.Paragraph;
    load(option: string | string[]): void;
    search(searchText: string, options?: Word.SearchOptions): Word.SearchRangeCollection;
    paragraphs: Word.ParagraphCollection;
    getFirst(): Word.Range;
  }

  interface Font {
    name: string;
    size: number;
    color: string;
    bold: boolean | string;
    italic: boolean | string;
    underline: string;
    load(option: string | string[]): void;
  }

  interface Paragraph {
    text: string;
    style: string;
    isListItem: boolean;
    insertParagraph(paragraphText: string, insertLocation: Word.InsertLocation): Word.Paragraph;
    getRange(rangeLocation: string): Word.Range;
    delete(): void;
    load(option: string | string[]): void;
  }

  interface ParagraphCollection {
    items: Word.Paragraph[];
    load(option: string | string[]): void;
    getFirst(): Word.Paragraph;
    getLast(): Word.Paragraph;
  }

  interface SearchOptions {
    matchCase?: boolean;
    matchWholeWord?: boolean;
    matchPrefix?: boolean;
    matchSuffix?: boolean;
  }

  interface SearchRangeCollection {
    items: Word.Range[];
    load(option: string | string[]): void;
  }

  type InsertLocation =
    | "Start"
    | "End"
    | "Before"
    | "After"
    | "Replace";
}

declare namespace Office {
  function onReady(callback?: (info: { host: string; platform: string }) => void): Promise<{ host: string; platform: string }>;
}

declare const Office: {
  onReady(callback?: (info?: { host: string; platform: string }) => void): Promise<{ host: string; platform: string }>;
};