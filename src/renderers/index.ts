import type { DocumentRenderer } from '../types';
import { CsvRenderer } from './CsvRenderer';
import { DocxRenderer } from './DocxRenderer';
import { ImageRenderer } from './ImageRenderer';
import { LegacyPptRenderer } from './LegacyPptRenderer';
import { MarkdownRenderer } from './MarkdownRenderer';
import { MediaRenderer } from './MediaRenderer';
import { PdfRenderer } from './PdfRenderer';
import { PptxRenderer } from './PptxRenderer';
import { SpreadsheetRenderer } from './SpreadsheetRenderer';
import { TextRenderer } from './TextRenderer';
import { UnsupportedRenderer } from './UnsupportedRenderer';

export const builtInRenderers: DocumentRenderer[] = [
  PdfRenderer,
  ImageRenderer,
  MarkdownRenderer,
  CsvRenderer,
  DocxRenderer,
  LegacyPptRenderer,
  PptxRenderer,
  SpreadsheetRenderer,
  TextRenderer,
  MediaRenderer,
  UnsupportedRenderer,
];

export {
  CsvRenderer,
  DocxRenderer,
  ImageRenderer,
  LegacyPptRenderer,
  MarkdownRenderer,
  MediaRenderer,
  PdfRenderer,
  PptxRenderer,
  SpreadsheetRenderer,
  TextRenderer,
  UnsupportedRenderer,
};
