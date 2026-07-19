import JSZip from 'jszip';

export type SpreadsheetCellValue = string | number | boolean | Date | null;

export interface SpreadsheetCellStyle {
  backgroundColor?: string;
  borderBottom?: string;
  borderLeft?: string;
  borderRight?: string;
  borderTop?: string;
  color?: string;
  direction?: 'ltr' | 'rtl';
  fontFamily?: string;
  fontSize?: string;
  fontStyle?: string;
  fontWeight?: string;
  paddingInlineStart?: string;
  textAlign?: 'left' | 'center' | 'right' | 'justify';
  textDecoration?: string;
  verticalAlign?: 'top' | 'middle' | 'bottom';
  whiteSpace?: 'normal' | 'nowrap' | 'pre-wrap';
  writingMode?: 'horizontal-tb' | 'vertical-rl';
}

export interface SpreadsheetComment {
  author?: string;
  text: string;
  date?: string;
  kind: 'comment' | 'note';
}

export interface SpreadsheetDrawingAnchor {
  fromColumn: number;
  fromRow: number;
  columnOffsetPx?: number;
  rowOffsetPx?: number;
  toColumn?: number;
  toRow?: number;
  toColumnOffsetPx?: number;
  toRowOffsetPx?: number;
  widthPx?: number;
  heightPx?: number;
}

export interface SpreadsheetImageDrawing {
  anchor: SpreadsheetDrawingAnchor;
  dataUrl: string;
  kind: 'image';
  name?: string;
}

export interface SpreadsheetChartSeries {
  name?: string;
  values: number[];
}

export interface SpreadsheetChartDrawing {
  anchor: SpreadsheetDrawingAnchor;
  categories: string[];
  chartType: 'area' | 'bar' | 'column' | 'doughnut' | 'line' | 'pie' | 'unknown';
  kind: 'chart';
  series: SpreadsheetChartSeries[];
  title?: string;
  valueFormatCode?: string;
}

export type SpreadsheetDrawing = SpreadsheetChartDrawing | SpreadsheetImageDrawing;

export interface ParsedSpreadsheetCell {
  value: SpreadsheetCellValue;
  formattedValue?: string;
  style?: SpreadsheetCellStyle;
  comments?: SpreadsheetComment[];
  colSpan?: number;
  hiddenByMerge?: boolean;
  rowSpan?: number;
}

export interface ParsedXlsxSheet {
  columnWidths?: number[];
  drawings?: SpreadsheetDrawing[];
  name: string;
  rowHeights?: number[];
  rows: ParsedSpreadsheetCell[][];
}

export interface ParsedXlsxWorkbook {
  sheets: ParsedXlsxSheet[];
}

interface ParsedStyle {
  style?: SpreadsheetCellStyle;
  numFmtId?: number;
}

interface ParsedWorksheetData {
  columnWidths?: number[];
  rowHeights?: number[];
  rows: ParsedSpreadsheetCell[][];
}

interface MergeRange {
  startRow: number;
  startColumn: number;
  endRow: number;
  endColumn: number;
}

interface ParsedSheetWithChartDescriptors extends ParsedXlsxSheet {
  chartDescriptors: SpreadsheetChartDescriptor[];
}

interface SpreadsheetChartDescriptor {
  anchor: SpreadsheetDrawingAnchor;
  chartXml: string;
}

type SpreadsheetBorderProperty = 'borderBottom' | 'borderLeft' | 'borderRight' | 'borderTop';

const builtinDateFormats = new Set([14, 15, 16, 17, 22, 27, 30, 36, 45, 46, 47, 50, 57]);
const builtinNumberFormats = new Map<number, string>([
  [1, '0'],
  [2, '0.00'],
  [3, '#,##0'],
  [4, '#,##0.00'],
  [9, '0%'],
  [10, '0.00%'],
  [11, '0.00E+00'],
  [14, 'm/d/yy'],
  [15, 'd-mmm-yy'],
  [16, 'd-mmm'],
  [17, 'mmm-yy'],
  [18, 'h:mm AM/PM'],
  [19, 'h:mm:ss AM/PM'],
  [20, 'h:mm'],
  [21, 'h:mm:ss'],
  [22, 'm/d/yy h:mm'],
  [37, '#,##0;(#,##0)'],
  [38, '#,##0;[Red](#,##0)'],
  [39, '#,##0.00;(#,##0.00)'],
  [40, '#,##0.00;[Red](#,##0.00)'],
  [45, 'mm:ss'],
  [46, '[h]:mm:ss'],
  [47, 'mmss.0'],
  [48, '##0.0E+0'],
  [49, '@'],
]);
const emuPerPixel = 9525;

const indexedColors = [
  '000000',
  'FFFFFF',
  'FF0000',
  '00FF00',
  '0000FF',
  'FFFF00',
  'FF00FF',
  '00FFFF',
  '000000',
  'FFFFFF',
  'FF0000',
  '00FF00',
  '0000FF',
  'FFFF00',
  'FF00FF',
  '00FFFF',
  '800000',
  '008000',
  '000080',
  '808000',
  '800080',
  '008080',
  'C0C0C0',
  '808080',
  '9999FF',
  '993366',
  'FFFFCC',
  'CCFFFF',
  '660066',
  'FF8080',
  '0066CC',
  'CCCCFF',
  '000080',
  'FF00FF',
  'FFFF00',
  '00FFFF',
  '800080',
  '800000',
  '008080',
  '0000FF',
  '00CCFF',
  'CCFFFF',
  'CCFFCC',
  'FFFF99',
  '99CCFF',
  'FF99CC',
  'CC99FF',
  'FFCC99',
  '3366FF',
  '33CCCC',
  '99CC00',
  'FFCC00',
  'FF9900',
  'FF6600',
  '666699',
  '969696',
  '003366',
  '339966',
  '003300',
  '333300',
  '993300',
  '993366',
  '333399',
  '333333',
];

function parseXml(xml: string): Document {
  return new DOMParser().parseFromString(xml, 'application/xml');
}

function elementsByLocalName(root: ParentNode, localName: string): Element[] {
  return Array.from(root.querySelectorAll('*')).filter((element) => element.localName === localName);
}

function childByLocalName(root: ParentNode, localName: string): Element | undefined {
  return Array.from(root.children).find((element) => element.localName === localName);
}

function normalizeWorkbookTarget(target: string): string {
  const cleanTarget = target.replace(/^\/+/, '');
  return cleanTarget.startsWith('xl/') ? cleanTarget : `xl/${cleanTarget}`;
}

function normalizePath(path: string): string {
  const parts: string[] = [];

  path.split('/').forEach((part) => {
    if (!part || part === '.') {
      return;
    }

    if (part === '..') {
      parts.pop();
      return;
    }

    parts.push(part);
  });

  return parts.join('/');
}

function resolveRelationshipTarget(sourcePath: string, target: string): string {
  if (target.startsWith('/')) {
    return target.replace(/^\/+/, '');
  }

  const basePath = sourcePath.split('/').slice(0, -1).join('/');
  return normalizePath(`${basePath}/${target}`);
}

function relationshipPathForPart(path: string): string {
  const parts = path.split('/');
  const fileName = parts.pop();
  return `${parts.join('/')}/_rels/${fileName}.rels`;
}

function emuToPixels(value: string | null): number {
  const emu = Number(value ?? 0);
  return Number.isFinite(emu) ? emu / emuPerPixel : 0;
}

function relationshipAttribute(element: Element | undefined, localName: string): string | undefined {
  if (!element) {
    return undefined;
  }

  return (
    element.getAttribute(`r:${localName}`) ??
    Array.from(element.attributes).find((attribute) => attribute.localName === localName)?.value
  );
}

function cellRefToIndexes(cellRef: string): { rowIndex: number; columnIndex: number } | null {
  const match = /^\$?([A-Z]+)\$?(\d+)$/i.exec(cellRef);
  if (!match) {
    return null;
  }

  const letters = match[1].toUpperCase();
  let columnIndex = 0;

  for (const letter of letters) {
    columnIndex = columnIndex * 26 + (letter.charCodeAt(0) - 64);
  }

  return {
    rowIndex: Number(match[2]) - 1,
    columnIndex: columnIndex - 1,
  };
}

function parseSimpleRangeFormula(
  formula: string | undefined,
  fallbackSheetName: string,
): { sheetName: string; start: { rowIndex: number; columnIndex: number }; end: { rowIndex: number; columnIndex: number } } | undefined {
  const cleanFormula = formula?.trim().replace(/^=/, '');

  if (!cleanFormula) {
    return undefined;
  }

  const match = /^(?:'((?:[^']|'')+)'|([^!]+))!(\$?[A-Z]+\$?\d+)(?::(\$?[A-Z]+\$?\d+))?$/i.exec(cleanFormula);
  const localMatch = /^(\$?[A-Z]+\$?\d+)(?::(\$?[A-Z]+\$?\d+))?$/i.exec(cleanFormula);
  const sheetName = match ? (match[1] ? match[1].replace(/''/g, "'") : match[2]) : fallbackSheetName;
  const startRef = match?.[3] ?? localMatch?.[1] ?? '';
  const endRef = match?.[4] ?? localMatch?.[2] ?? startRef;
  const start = cellRefToIndexes(startRef);
  const end = cellRefToIndexes(endRef);

  if (!start || !end || !sheetName) {
    return undefined;
  }

  return {
    sheetName,
    start: {
      rowIndex: Math.min(start.rowIndex, end.rowIndex),
      columnIndex: Math.min(start.columnIndex, end.columnIndex),
    },
    end: {
      rowIndex: Math.max(start.rowIndex, end.rowIndex),
      columnIndex: Math.max(start.columnIndex, end.columnIndex),
    },
  };
}

function imageMimeTypeFromPath(path: string): string {
  const extension = path.split('.').pop()?.toLowerCase();

  switch (extension) {
    case 'gif':
      return 'image/gif';
    case 'jpg':
    case 'jpeg':
      return 'image/jpeg';
    case 'svg':
      return 'image/svg+xml';
    case 'webp':
      return 'image/webp';
    case 'tif':
    case 'tiff':
      return 'image/tiff';
    case 'png':
    default:
      return 'image/png';
  }
}

function excelColumnWidthToPixels(width: number): number {
  return Math.max(24, Math.min(640, Math.round(width * 7 + 5)));
}

function pointsToPixels(points: number): number {
  return Math.max(1, Math.min(640, Math.round((points * 96) / 72)));
}

function parseColumnWidths(document: Document): number[] {
  const widths: number[] = [];

  elementsByLocalName(document, 'col').forEach((columnNode) => {
    const min = Number(columnNode.getAttribute('min'));
    const max = Number(columnNode.getAttribute('max'));
    const width = Number(columnNode.getAttribute('width'));

    if (!Number.isFinite(min) || !Number.isFinite(max) || !Number.isFinite(width)) {
      return;
    }

    for (let columnIndex = Math.max(0, min - 1); columnIndex <= max - 1; columnIndex += 1) {
      widths[columnIndex] = excelColumnWidthToPixels(width);
    }
  });

  return widths;
}

function parseRowHeights(document: Document): number[] {
  const heights: number[] = [];

  elementsByLocalName(document, 'row').forEach((rowNode, index) => {
    const rowNumber = Number(rowNode.getAttribute('r') ?? index + 1);
    const height = Number(rowNode.getAttribute('ht'));

    if (!Number.isFinite(rowNumber) || !Number.isFinite(height)) {
      return;
    }

    heights[Math.max(0, rowNumber - 1)] = pointsToPixels(height);
  });

  return heights;
}

function parseMergeRange(ref: string): MergeRange | null {
  const [startRef, endRef] = ref.split(':');
  const start = cellRefToIndexes(startRef ?? '');
  const end = cellRefToIndexes(endRef ?? startRef ?? '');

  if (!start || !end) {
    return null;
  }

  return {
    startRow: Math.min(start.rowIndex, end.rowIndex),
    startColumn: Math.min(start.columnIndex, end.columnIndex),
    endRow: Math.max(start.rowIndex, end.rowIndex),
    endColumn: Math.max(start.columnIndex, end.columnIndex),
  };
}

function ensureCell(rows: ParsedSpreadsheetCell[][], rowIndex: number, columnIndex: number): ParsedSpreadsheetCell {
  rows[rowIndex] ??= [];
  rows[rowIndex][columnIndex] ??= { value: null };
  return rows[rowIndex][columnIndex];
}

function firstBorderInRange(
  rows: ParsedSpreadsheetCell[][],
  range: MergeRange,
  borderProperty: SpreadsheetBorderProperty,
  rowStart: number,
  rowEnd: number,
  columnStart: number,
  columnEnd: number,
): string | undefined {
  for (let rowIndex = rowStart; rowIndex <= rowEnd; rowIndex += 1) {
    for (let columnIndex = columnStart; columnIndex <= columnEnd; columnIndex += 1) {
      const border = rows[rowIndex]?.[columnIndex]?.style?.[borderProperty];

      if (border) {
        return border;
      }
    }
  }

  return rows[range.startRow]?.[range.startColumn]?.style?.[borderProperty];
}

function mergeRangeBorderStyle(rows: ParsedSpreadsheetCell[][], range: MergeRange): SpreadsheetCellStyle | undefined {
  const borderStyle: SpreadsheetCellStyle = {};
  const top = firstBorderInRange(rows, range, 'borderTop', range.startRow, range.startRow, range.startColumn, range.endColumn);
  const right = firstBorderInRange(rows, range, 'borderRight', range.startRow, range.endRow, range.endColumn, range.endColumn);
  const bottom = firstBorderInRange(rows, range, 'borderBottom', range.endRow, range.endRow, range.startColumn, range.endColumn);
  const left = firstBorderInRange(rows, range, 'borderLeft', range.startRow, range.endRow, range.startColumn, range.startColumn);

  if (top) {
    borderStyle.borderTop = top;
  }

  if (right) {
    borderStyle.borderRight = right;
  }

  if (bottom) {
    borderStyle.borderBottom = bottom;
  }

  if (left) {
    borderStyle.borderLeft = left;
  }

  return Object.keys(borderStyle).length > 0 ? borderStyle : undefined;
}

function applyMergeRanges(rows: ParsedSpreadsheetCell[][], ranges: MergeRange[]) {
  ranges.forEach((range) => {
    const rowSpan = range.endRow - range.startRow + 1;
    const colSpan = range.endColumn - range.startColumn + 1;

    if (rowSpan <= 1 && colSpan <= 1) {
      return;
    }

    const originCell = ensureCell(rows, range.startRow, range.startColumn);
    originCell.rowSpan = rowSpan;
    originCell.colSpan = colSpan;
    originCell.style = {
      ...(originCell.style ?? {}),
      ...(mergeRangeBorderStyle(rows, range) ?? {}),
    };

    for (let rowIndex = range.startRow; rowIndex <= range.endRow; rowIndex += 1) {
      for (let columnIndex = range.startColumn; columnIndex <= range.endColumn; columnIndex += 1) {
        if (rowIndex === range.startRow && columnIndex === range.startColumn) {
          continue;
        }

        const mergedCell = ensureCell(rows, rowIndex, columnIndex);
        mergedCell.hiddenByMerge = true;
      }
    }
  });
}

function normalizeHexColor(value: string): string | undefined {
  const clean = value.replace(/[^a-f0-9]/gi, '').toUpperCase();

  if (clean.length === 8) {
    return clean.slice(0, 2) === '00' ? undefined : `#${clean.slice(2)}`;
  }

  if (clean.length === 6) {
    return `#${clean}`;
  }

  return undefined;
}

function applyTint(hex: string, tintValue: string | null): string {
  const tint = tintValue ? Number(tintValue) : 0;
  if (!Number.isFinite(tint) || tint === 0) {
    return hex;
  }

  const clean = hex.replace('#', '');
  const channels = [clean.slice(0, 2), clean.slice(2, 4), clean.slice(4, 6)].map((channel) => Number.parseInt(channel, 16));
  const tinted = channels.map((channel) => {
    const nextValue = tint < 0 ? channel * (1 + tint) : channel + (255 - channel) * tint;
    return Math.max(0, Math.min(255, Math.round(nextValue))).toString(16).padStart(2, '0');
  });

  return `#${tinted.join('').toUpperCase()}`;
}

function parseThemeColors(themeXml?: string): string[] {
  if (!themeXml) {
    return [];
  }

  const document = parseXml(themeXml);
  const colorScheme = elementsByLocalName(document, 'clrScheme')[0];
  if (!colorScheme) {
    return [];
  }

  const colorsBySchemeName = new Map<string, string>();
  Array.from(colorScheme.children).forEach((colorNode) => {
    const srgbColor = elementsByLocalName(colorNode, 'srgbClr')[0]?.getAttribute('val');
    const systemColor = elementsByLocalName(colorNode, 'sysClr')[0]?.getAttribute('lastClr');
    colorsBySchemeName.set(colorNode.localName, normalizeHexColor(srgbColor ?? systemColor ?? '') ?? '');
  });
  const spreadsheetThemeOrder = [
    'lt1',
    'dk1',
    'lt2',
    'dk2',
    'accent1',
    'accent2',
    'accent3',
    'accent4',
    'accent5',
    'accent6',
    'hlink',
    'folHlink',
  ];

  return spreadsheetThemeOrder.map((schemeName) => colorsBySchemeName.get(schemeName) ?? '');
}

function colorFromElement(element: Element | undefined, themeColors: string[]): string | undefined {
  if (!element) {
    return undefined;
  }

  const rgb = element.getAttribute('rgb');
  if (rgb) {
    return normalizeHexColor(rgb);
  }

  const indexed = element.getAttribute('indexed');
  if (indexed) {
    const color = indexedColors[Number(indexed)];
    return color ? `#${color}` : undefined;
  }

  const theme = element.getAttribute('theme');
  if (theme) {
    const color = themeColors[Number(theme)];
    return color ? applyTint(color, element.getAttribute('tint')) : undefined;
  }

  return undefined;
}

function parseSharedStrings(xml?: string): string[] {
  if (!xml) {
    return [];
  }

  const document = parseXml(xml);
  return elementsByLocalName(document, 'si').map((item) => {
    return elementsByLocalName(item, 't')
      .map((textNode) => textNode.textContent ?? '')
      .join('');
  });
}

function parseNumberFormats(stylesDocument: Document): Map<number, string> {
  const numberFormats = new Map<number, string>();

  elementsByLocalName(stylesDocument, 'numFmt').forEach((formatNode) => {
    const id = Number(formatNode.getAttribute('numFmtId'));
    const code = formatNode.getAttribute('formatCode');

    if (Number.isFinite(id) && code) {
      numberFormats.set(id, code);
    }
  });

  return numberFormats;
}

function isDateFormat(numFmtId: number | undefined, numberFormats: Map<number, string>): boolean {
  if (typeof numFmtId !== 'number') {
    return false;
  }

  if (builtinDateFormats.has(numFmtId)) {
    return true;
  }

  const format = numberFormats.get(numFmtId)?.toLowerCase() ?? '';
  return /(^|[^\\])[dmyhsa]/.test(format) && !/[0#?]\/[0#?]/.test(format);
}

function formatCodeForId(numFmtId: number | undefined, numberFormats: Map<number, string>): string | undefined {
  if (typeof numFmtId !== 'number') {
    return undefined;
  }

  return numberFormats.get(numFmtId) ?? builtinNumberFormats.get(numFmtId);
}

function excelSerialDate(value: number, date1904: boolean): Date {
  const epoch = Date.UTC(date1904 ? 1904 : 1899, date1904 ? 0 : 11, date1904 ? 1 : 30);
  return new Date(epoch + value * 86_400_000);
}

function splitFormatSections(formatCode: string): string[] {
  const sections: string[] = [];
  let current = '';
  let inQuote = false;

  for (let index = 0; index < formatCode.length; index += 1) {
    const character = formatCode[index];

    if (character === '"') {
      inQuote = !inQuote;
      current += character;
      continue;
    }

    if (character === ';' && !inQuote) {
      sections.push(current);
      current = '';
      continue;
    }

    current += character;
  }

  sections.push(current);
  return sections;
}

function selectFormatSection(formatCode: string, value: number): string {
  const sections = splitFormatSections(formatCode);

  if (value < 0 && sections[1]) {
    return sections[1];
  }

  if (value === 0 && sections[2]) {
    return sections[2];
  }

  return sections[0] ?? formatCode;
}

function stripFormatMetadata(section: string): string {
  return section
    .replace(/\[[^\]]+\]/g, '')
    .replace(/"([^"]*)"/g, '$1')
    .replace(/\\(.)/g, '$1')
    .replace(/_./g, '')
    .replace(/\*./g, '');
}

function decimalPlacesFromFormat(section: string): { maximum: number; minimum: number } {
  const numericSection = stripFormatMetadata(section);
  const decimalPart = numericSection.split('.')[1]?.match(/[0#?]+/)?.[0] ?? '';

  return {
    maximum: decimalPart.length,
    minimum: Array.from(decimalPart).filter((character) => character === '0').length,
  };
}

function currencySymbolFromFormat(section: string): string {
  const bracketCurrency = /\[\$([^\]-]+)(?:-[^\]]+)?\]/.exec(section)?.[1];

  if (bracketCurrency) {
    return bracketCurrency;
  }

  return /[₹$€£¥]/.exec(stripFormatMetadata(section))?.[0] ?? '';
}

function formatNumberByCode(value: number, formatCode: string | undefined): string | undefined {
  if (!formatCode || /^general$/i.test(formatCode.trim())) {
    return undefined;
  }

  const section = selectFormatSection(formatCode, value);
  const cleanSection = stripFormatMetadata(section);
  const percent = cleanSection.includes('%');
  const currency = currencySymbolFromFormat(section);
  const placeholderIndex = cleanSection.search(/[0#?]/);
  const currencyIndex = currency ? cleanSection.indexOf(currency) : -1;
  const prefix = currency && (currencyIndex < placeholderIndex || placeholderIndex === -1) ? currency : '';
  const suffix = `${percent ? '%' : ''}${currency && currencyIndex > placeholderIndex ? currency : ''}`;
  const decimalPlaces = decimalPlacesFromFormat(section);
  const useGrouping = /[0#?],[0#?]/.test(cleanSection);
  const hasParentheses = value < 0 && cleanSection.includes('(') && cleanSection.includes(')');
  const adjustedValue = Math.abs(percent ? value * 100 : value);
  const formattedNumber = new Intl.NumberFormat('en-US', {
    maximumFractionDigits: decimalPlaces.maximum,
    minimumFractionDigits: decimalPlaces.minimum,
    useGrouping,
  }).format(adjustedValue);
  const sign = value < 0 && !hasParentheses ? '-' : '';
  const formatted = `${sign}${prefix}${formattedNumber}${suffix}`;

  return hasParentheses ? `(${formatted})` : formatted;
}

function pad2(value: number): string {
  return String(value).padStart(2, '0');
}

function formatDateByCode(value: Date, formatCode: string | undefined): string | undefined {
  if (!formatCode) {
    return undefined;
  }

  const cleanCode = stripFormatMetadata(selectFormatSection(formatCode, 1)).toLowerCase();
  const year = value.getUTCFullYear();
  const month = value.getUTCMonth() + 1;
  const day = value.getUTCDate();
  const shortYear = String(year).slice(-2);
  const monthNames = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

  if (cleanCode.includes('yyyy-mm-dd')) {
    return `${year}-${pad2(month)}-${pad2(day)}`;
  }

  if (cleanCode.includes('d-mmm-yy')) {
    return `${day}-${monthNames[month - 1]}-${shortYear}`;
  }

  if (cleanCode.includes('mmm-yy')) {
    return `${monthNames[month - 1]}-${shortYear}`;
  }

  if (cleanCode.includes('m/d/yy')) {
    return `${month}/${day}/${shortYear}`;
  }

  return undefined;
}

function formatCellDisplayValue(value: SpreadsheetCellValue, numFmtId: number | undefined, numberFormats: Map<number, string>): string | undefined {
  const formatCode = formatCodeForId(numFmtId, numberFormats);

  if (typeof value === 'number') {
    return formatNumberByCode(value, formatCode);
  }

  if (value instanceof Date) {
    return formatDateByCode(value, formatCode);
  }

  return undefined;
}

function booleanAttribute(value: string | null): boolean {
  return value === '1' || value === 'true';
}

function parseAlignmentStyle(alignmentNode: Element | undefined): SpreadsheetCellStyle {
  if (!alignmentNode) {
    return {};
  }

  const style: SpreadsheetCellStyle = {};
  const horizontal = alignmentNode.getAttribute('horizontal');
  const vertical = alignmentNode.getAttribute('vertical');
  const indent = Number(alignmentNode.getAttribute('indent') ?? 0);
  const textRotation = alignmentNode.getAttribute('textRotation');
  const readingOrder = alignmentNode.getAttribute('readingOrder');

  if (horizontal === 'center' || horizontal === 'centerContinuous') {
    style.textAlign = 'center';
  } else if (horizontal === 'right') {
    style.textAlign = 'right';
  } else if (horizontal === 'justify' || horizontal === 'distributed') {
    style.textAlign = 'justify';
  } else if (horizontal === 'left') {
    style.textAlign = 'left';
  }

  if (vertical === 'center') {
    style.verticalAlign = 'middle';
  } else if (vertical === 'bottom') {
    style.verticalAlign = 'bottom';
  } else if (vertical === 'top') {
    style.verticalAlign = 'top';
  }

  if (alignmentNode.hasAttribute('wrapText')) {
    style.whiteSpace = booleanAttribute(alignmentNode.getAttribute('wrapText')) ? 'pre-wrap' : 'nowrap';
  }

  if (Number.isFinite(indent) && indent > 0) {
    style.paddingInlineStart = `${Math.min(indent * 12, 120)}px`;
  }

  if (textRotation === '255') {
    style.writingMode = 'vertical-rl';
  }

  if (readingOrder === '1') {
    style.direction = 'ltr';
  } else if (readingOrder === '2') {
    style.direction = 'rtl';
  }

  return style;
}

function excelBorderLineStyle(value: string | null): { lineStyle: string; width: string } | undefined {
  switch (value) {
    case 'dashDot':
    case 'dashDotDot':
    case 'dashed':
    case 'mediumDashDot':
    case 'mediumDashDotDot':
    case 'mediumDashed':
    case 'slantDashDot':
      return { lineStyle: 'dashed', width: value.startsWith('medium') ? '2px' : '1px' };
    case 'dotted':
    case 'hair':
      return { lineStyle: 'dotted', width: '1px' };
    case 'double':
      return { lineStyle: 'double', width: '3px' };
    case 'medium':
      return { lineStyle: 'solid', width: '2px' };
    case 'thick':
      return { lineStyle: 'solid', width: '3px' };
    case 'thin':
      return { lineStyle: 'solid', width: '1px' };
    default:
      return undefined;
  }
}

function parseBorderSideStyle(sideNode: Element | undefined, themeColors: string[]): string | undefined {
  if (!sideNode) {
    return undefined;
  }

  const borderLine = excelBorderLineStyle(sideNode.getAttribute('style'));
  if (!borderLine) {
    return undefined;
  }

  const color = colorFromElement(childByLocalName(sideNode, 'color'), themeColors) ?? '#000000';
  return `${borderLine.width} ${borderLine.lineStyle} ${color}`;
}

function parseBorderStyle(borderNode: Element, themeColors: string[]): SpreadsheetCellStyle {
  const style: SpreadsheetCellStyle = {};
  const left = parseBorderSideStyle(childByLocalName(borderNode, 'left'), themeColors);
  const right = parseBorderSideStyle(childByLocalName(borderNode, 'right'), themeColors);
  const top = parseBorderSideStyle(childByLocalName(borderNode, 'top'), themeColors);
  const bottom = parseBorderSideStyle(childByLocalName(borderNode, 'bottom'), themeColors);

  if (left) {
    style.borderLeft = left;
  }

  if (right) {
    style.borderRight = right;
  }

  if (top) {
    style.borderTop = top;
  }

  if (bottom) {
    style.borderBottom = bottom;
  }

  return style;
}

function parseStyles(stylesXml: string | undefined, themeColors: string[]): ParsedStyle[] {
  if (!stylesXml) {
    return [];
  }

  const document = parseXml(stylesXml);
  const numberFormats = parseNumberFormats(document);
  const fontStyles = elementsByLocalName(document, 'fonts')[0]
    ? Array.from(elementsByLocalName(document, 'fonts')[0].children)
        .filter((element) => element.localName === 'font')
        .map((fontNode) => {
          const style: SpreadsheetCellStyle = {};
          const color = colorFromElement(childByLocalName(fontNode, 'color'), themeColors);
          const fontSize = childByLocalName(fontNode, 'sz')?.getAttribute('val');
          const fontFamily = childByLocalName(fontNode, 'name')?.getAttribute('val');

          if (color) {
            style.color = color;
          }

          if (fontSize) {
            style.fontSize = `${fontSize}pt`;
          }

          if (fontFamily) {
            const safeFontFamily = fontFamily.replace(/["\\]/g, '').trim();
            style.fontFamily = safeFontFamily.includes(' ')
              ? `"${safeFontFamily}", Arial, Helvetica, sans-serif`
              : `${safeFontFamily}, Arial, Helvetica, sans-serif`;
          }

          if (childByLocalName(fontNode, 'b')) {
            style.fontWeight = '700';
          }

          if (childByLocalName(fontNode, 'i')) {
            style.fontStyle = 'italic';
          }

          if (childByLocalName(fontNode, 'u')) {
            style.textDecoration = 'underline';
          }

          return style;
        })
    : [];

  const fillStyles = elementsByLocalName(document, 'fills')[0]
    ? Array.from(elementsByLocalName(document, 'fills')[0].children)
        .filter((element) => element.localName === 'fill')
        .map((fillNode) => {
          const patternFill = childByLocalName(fillNode, 'patternFill');
          const patternType = patternFill?.getAttribute('patternType');

          if (!patternFill || patternType === 'none' || patternType === 'gray125') {
            return {};
          }

          const backgroundColor =
            colorFromElement(childByLocalName(patternFill, 'fgColor'), themeColors) ??
            colorFromElement(childByLocalName(patternFill, 'bgColor'), themeColors);

          return backgroundColor ? { backgroundColor } : {};
        })
    : [];

  const borderStyles = elementsByLocalName(document, 'borders')[0]
    ? Array.from(elementsByLocalName(document, 'borders')[0].children)
        .filter((element) => element.localName === 'border')
        .map((borderNode) => parseBorderStyle(borderNode, themeColors))
    : [];

  const cellFormats = elementsByLocalName(document, 'cellXfs')[0]
    ? Array.from(elementsByLocalName(document, 'cellXfs')[0].children).filter((element) => element.localName === 'xf')
    : [];

  return cellFormats.map((formatNode) => {
    const fontId = Number(formatNode.getAttribute('fontId') ?? 0);
    const fillId = Number(formatNode.getAttribute('fillId') ?? 0);
    const borderId = Number(formatNode.getAttribute('borderId') ?? 0);
    const numFmtId = Number(formatNode.getAttribute('numFmtId') ?? 0);
    const style: SpreadsheetCellStyle = {
      ...(fillStyles[fillId] ?? {}),
      ...(fontStyles[fontId] ?? {}),
      ...(borderStyles[borderId] ?? {}),
      ...parseAlignmentStyle(childByLocalName(formatNode, 'alignment')),
    };

    return {
      style: Object.keys(style).length > 0 ? style : undefined,
      numFmtId: Number.isFinite(numFmtId) ? numFmtId : undefined,
    };
  });
}

function parseWorkbookSheets(workbookXml: string, relationshipsXml?: string): Array<{ name: string; path: string }> {
  const workbook = parseXml(workbookXml);
  const relationshipMap = new Map<string, string>();

  if (relationshipsXml) {
    const relationships = parseXml(relationshipsXml);
    elementsByLocalName(relationships, 'Relationship').forEach((relationship) => {
      const id = relationship.getAttribute('Id');
      const target = relationship.getAttribute('Target');

      if (id && target) {
        relationshipMap.set(id, normalizeWorkbookTarget(target));
      }
    });
  }

  return elementsByLocalName(workbook, 'sheet')
    .map((sheet, index) => {
      const relationId = sheet.getAttribute('r:id') ?? sheet.getAttribute('id') ?? '';
      return {
        name: sheet.getAttribute('name') ?? `Sheet${index + 1}`,
        path: relationshipMap.get(relationId) ?? `xl/worksheets/sheet${index + 1}.xml`,
      };
    })
    .filter((sheet) => sheet.path);
}

function parseRelationships(relationshipsXml?: string, sourcePath = ''): Array<{ id: string; type: string; target: string }> {
  if (!relationshipsXml) {
    return [];
  }

  const relationships = parseXml(relationshipsXml);
  return elementsByLocalName(relationships, 'Relationship')
    .map((relationship) => {
      const id = relationship.getAttribute('Id') ?? '';
      const type = relationship.getAttribute('Type') ?? '';
      const target = relationship.getAttribute('Target') ?? '';

      return {
        id,
        type,
        target: sourcePath ? resolveRelationshipTarget(sourcePath, target) : target,
      };
    })
    .filter((relationship) => relationship.id && relationship.target);
}

function parseCommentText(element: Element): string {
  return elementsByLocalName(element, 't')
    .map((textNode) => textNode.textContent ?? '')
    .join('');
}

function parseClassicComments(xml?: string): Map<string, SpreadsheetComment[]> {
  const commentsByRef = new Map<string, SpreadsheetComment[]>();

  if (!xml) {
    return commentsByRef;
  }

  const document = parseXml(xml);
  const authors = elementsByLocalName(document, 'author').map((author) => author.textContent ?? '');

  elementsByLocalName(document, 'comment').forEach((comment) => {
    const ref = comment.getAttribute('ref');
    const text = parseCommentText(comment).trim();

    if (!ref || !text) {
      return;
    }

    const authorId = Number(comment.getAttribute('authorId') ?? 0);
    const items = commentsByRef.get(ref) ?? [];
    items.push({
      author: authors[authorId],
      kind: 'note',
      text,
    });
    commentsByRef.set(ref, items);
  });

  return commentsByRef;
}

function parsePersons(xml?: string): Map<string, string> {
  const persons = new Map<string, string>();

  if (!xml) {
    return persons;
  }

  const document = parseXml(xml);
  elementsByLocalName(document, 'person').forEach((person) => {
    const id = person.getAttribute('id');
    const displayName = person.getAttribute('displayName') ?? person.getAttribute('userId') ?? '';

    if (id && displayName) {
      persons.set(id, displayName);
    }
  });

  return persons;
}

function parseThreadedComments(xml: string | undefined, persons: Map<string, string>): Map<string, SpreadsheetComment[]> {
  const commentsByRef = new Map<string, SpreadsheetComment[]>();

  if (!xml) {
    return commentsByRef;
  }

  const document = parseXml(xml);
  elementsByLocalName(document, 'threadedComment').forEach((comment) => {
    const ref = comment.getAttribute('ref');
    const text = parseCommentText(comment).trim();

    if (!ref || !text) {
      return;
    }

    const personId = comment.getAttribute('personId') ?? '';
    const items = commentsByRef.get(ref) ?? [];
    items.push({
      author: persons.get(personId) ?? personId,
      date: comment.getAttribute('dT') ?? undefined,
      kind: 'comment',
      text,
    });
    commentsByRef.set(ref, items);
  });

  return commentsByRef;
}

function mergeCommentMaps(...maps: Array<Map<string, SpreadsheetComment[]>>): Map<string, SpreadsheetComment[]> {
  const merged = new Map<string, SpreadsheetComment[]>();

  maps.forEach((map) => {
    map.forEach((comments, ref) => {
      merged.set(ref, [...(merged.get(ref) ?? []), ...comments]);
    });
  });

  return merged;
}

function parseDrawingPoint(element: Element | undefined): {
  column: number;
  columnOffsetPx: number;
  row: number;
  rowOffsetPx: number;
} {
  if (!element) {
    return {
      column: 0,
      columnOffsetPx: 0,
      row: 0,
      rowOffsetPx: 0,
    };
  }

  return {
    column: Number(childByLocalName(element, 'col')?.textContent ?? 0),
    columnOffsetPx: emuToPixels(childByLocalName(element, 'colOff')?.textContent ?? '0'),
    row: Number(childByLocalName(element, 'row')?.textContent ?? 0),
    rowOffsetPx: emuToPixels(childByLocalName(element, 'rowOff')?.textContent ?? '0'),
  };
}

function parseDrawingAnchor(anchorElement: Element): SpreadsheetDrawingAnchor | undefined {
  const from = parseDrawingPoint(childByLocalName(anchorElement, 'from'));

  if (!Number.isFinite(from.column) || !Number.isFinite(from.row)) {
    return undefined;
  }

  const anchor: SpreadsheetDrawingAnchor = {
    columnOffsetPx: from.columnOffsetPx,
    fromColumn: Math.max(0, from.column),
    fromRow: Math.max(0, from.row),
    rowOffsetPx: from.rowOffsetPx,
  };
  const toElement = childByLocalName(anchorElement, 'to');
  const extentElement = childByLocalName(anchorElement, 'ext');

  if (toElement) {
    const to = parseDrawingPoint(toElement);
    anchor.toColumn = Math.max(anchor.fromColumn, to.column);
    anchor.toColumnOffsetPx = to.columnOffsetPx;
    anchor.toRow = Math.max(anchor.fromRow, to.row);
    anchor.toRowOffsetPx = to.rowOffsetPx;
  }

  if (extentElement) {
    anchor.widthPx = emuToPixels(extentElement.getAttribute('cx'));
    anchor.heightPx = emuToPixels(extentElement.getAttribute('cy'));
  }

  return anchor;
}

function cachedTextValues(element: Element | undefined): string[] {
  if (!element) {
    return [];
  }

  const cache = childByLocalName(element, 'strCache') ?? childByLocalName(element, 'multiLvlStrCache');
  return elementsByLocalName(cache ?? element, 'pt').map((point) => childByLocalName(point, 'v')?.textContent ?? '');
}

function cachedNumberValues(element: Element | undefined): number[] {
  if (!element) {
    return [];
  }

  return elementsByLocalName(element, 'pt')
    .map((point) => Number(childByLocalName(point, 'v')?.textContent ?? ''))
    .filter((value) => Number.isFinite(value));
}

function cellsFromRangeFormula(
  formula: string | undefined,
  fallbackSheetName: string,
  sheetLookup: Map<string, ParsedSpreadsheetCell[][]>,
): ParsedSpreadsheetCell[] {
  const range = parseSimpleRangeFormula(formula, fallbackSheetName);

  if (!range) {
    return [];
  }

  const rows = sheetLookup.get(range.sheetName);
  if (!rows) {
    return [];
  }

  const cells: ParsedSpreadsheetCell[] = [];

  for (let rowIndex = range.start.rowIndex; rowIndex <= range.end.rowIndex; rowIndex += 1) {
    for (let columnIndex = range.start.columnIndex; columnIndex <= range.end.columnIndex; columnIndex += 1) {
      cells.push(rows[rowIndex]?.[columnIndex] ?? { value: null });
    }
  }

  return cells;
}

function valuesFromRangeFormula(
  formula: string | undefined,
  fallbackSheetName: string,
  sheetLookup: Map<string, ParsedSpreadsheetCell[][]>,
): SpreadsheetCellValue[] {
  return cellsFromRangeFormula(formula, fallbackSheetName, sheetLookup).map((cell) => cell.value);
}

function textValuesFromChartReference(
  element: Element | undefined,
  fallbackSheetName: string,
  sheetLookup: Map<string, ParsedSpreadsheetCell[][]>,
): string[] {
  if (!element) {
    return [];
  }

  const cached = cachedTextValues(element);

  if (cached.length > 0) {
    return cached;
  }

  const formula = elementsByLocalName(element, 'f')[0]?.textContent ?? undefined;
  return cellsFromRangeFormula(formula, fallbackSheetName, sheetLookup).map((cell) => {
    const value = cell.formattedValue ?? cell.value;

    if (value instanceof Date) {
      return value.toLocaleDateString();
    }

    return value === null || typeof value === 'undefined' ? '' : String(value);
  });
}

function numberValuesFromChartReference(
  element: Element | undefined,
  fallbackSheetName: string,
  sheetLookup: Map<string, ParsedSpreadsheetCell[][]>,
): number[] {
  if (!element) {
    return [];
  }

  const cached = cachedNumberValues(element);

  if (cached.length > 0) {
    return cached;
  }

  const formula = elementsByLocalName(element, 'f')[0]?.textContent ?? undefined;
  return valuesFromRangeFormula(formula, fallbackSheetName, sheetLookup)
    .map((value) => (typeof value === 'number' ? value : Number(value)))
    .filter((value) => Number.isFinite(value));
}

function parseChartType(document: Document): SpreadsheetChartDrawing['chartType'] {
  const chartTypeElement = elementsByLocalName(document, 'barChart')[0];

  if (chartTypeElement) {
    return childByLocalName(chartTypeElement, 'barDir')?.getAttribute('val') === 'bar' ? 'bar' : 'column';
  }

  if (elementsByLocalName(document, 'lineChart')[0]) {
    return 'line';
  }

  if (elementsByLocalName(document, 'areaChart')[0]) {
    return 'area';
  }

  if (elementsByLocalName(document, 'pieChart')[0]) {
    return 'pie';
  }

  if (elementsByLocalName(document, 'doughnutChart')[0]) {
    return 'doughnut';
  }

  return 'unknown';
}

function parseChartTitle(document: Document): string | undefined {
  const titleElement = elementsByLocalName(document, 'title')[0];
  const text = titleElement
    ? elementsByLocalName(titleElement, 't')
        .map((node) => node.textContent ?? '')
        .join(' ')
        .trim()
    : '';

  return text || undefined;
}

function parseChartValueFormat(document: Document, seriesElements: Element[]): string | undefined {
  const axisFormat = elementsByLocalName(document, 'valAx')
    .map((axisElement) => childByLocalName(axisElement, 'numFmt')?.getAttribute('formatCode') ?? undefined)
    .find((formatCode) => formatCode && !/^general$/i.test(formatCode));

  if (axisFormat) {
    return axisFormat;
  }

  return seriesElements
    .map((seriesElement) => {
      const valueElement = childByLocalName(seriesElement, 'val');
      return valueElement ? elementsByLocalName(valueElement, 'formatCode')[0]?.textContent?.trim() : undefined;
    })
    .find((formatCode) => formatCode && !/^general$/i.test(formatCode));
}

function parseChartDrawing(
  descriptor: SpreadsheetChartDescriptor,
  fallbackSheetName: string,
  sheetLookup: Map<string, ParsedSpreadsheetCell[][]>,
): SpreadsheetChartDrawing {
  const chartDocument = parseXml(descriptor.chartXml);
  const seriesElements = elementsByLocalName(chartDocument, 'ser');
  const series = seriesElements.map((seriesElement, index) => {
    const textElement = childByLocalName(seriesElement, 'tx');
    const name =
      (textElement ? childByLocalName(textElement, 'v')?.textContent : undefined) ??
      textValuesFromChartReference(textElement, fallbackSheetName, sheetLookup)[0] ??
      `Series ${index + 1}`;
    const values = numberValuesFromChartReference(childByLocalName(seriesElement, 'val'), fallbackSheetName, sheetLookup);

    return {
      name,
      values,
    };
  });
  const categories = textValuesFromChartReference(childByLocalName(seriesElements[0], 'cat'), fallbackSheetName, sheetLookup);

  return {
    anchor: descriptor.anchor,
    categories,
    chartType: parseChartType(chartDocument),
    kind: 'chart',
    series,
    title: parseChartTitle(chartDocument),
    valueFormatCode: parseChartValueFormat(chartDocument, seriesElements),
  };
}

async function parseWorksheetDrawings(
  drawingPath: string,
  drawingXml: string,
  relationshipsXml: string | undefined,
  readText: (path: string) => Promise<string | undefined>,
  readBase64: (path: string) => Promise<string | undefined>,
): Promise<{ chartDescriptors: SpreadsheetChartDescriptor[]; images: SpreadsheetImageDrawing[] }> {
  const document = parseXml(drawingXml);
  const relationships = parseRelationships(relationshipsXml, drawingPath);
  const relationshipsById = new Map(relationships.map((relationship) => [relationship.id, relationship]));
  const chartDescriptors: SpreadsheetChartDescriptor[] = [];
  const images: SpreadsheetImageDrawing[] = [];
  const drawingRoot = document.documentElement.localName === 'wsDr' ? document.documentElement : elementsByLocalName(document, 'wsDr')[0];

  await Promise.all(
    drawingRoot
      ? Array.from(drawingRoot.children)
          .filter((element) => element.localName === 'oneCellAnchor' || element.localName === 'twoCellAnchor')
          .map(async (anchorElement) => {
            const anchor = parseDrawingAnchor(anchorElement);

            if (!anchor) {
              return;
            }

            const chartRelationId = relationshipAttribute(elementsByLocalName(anchorElement, 'chart')[0], 'id');
            const imageRelationId = relationshipAttribute(elementsByLocalName(anchorElement, 'blip')[0], 'embed');

            if (chartRelationId) {
              const chartRelationship = relationshipsById.get(chartRelationId);
              const chartXml = chartRelationship ? await readText(chartRelationship.target) : undefined;

              if (chartXml) {
                chartDescriptors.push({ anchor, chartXml });
              }
            }

            if (imageRelationId) {
              const imageRelationship = relationshipsById.get(imageRelationId);
              const base64 = imageRelationship ? await readBase64(imageRelationship.target) : undefined;

              if (imageRelationship && base64) {
                images.push({
                  anchor,
                  dataUrl: `data:${imageMimeTypeFromPath(imageRelationship.target)};base64,${base64}`,
                  kind: 'image',
                  name: elementsByLocalName(anchorElement, 'cNvPr')[0]?.getAttribute('name') ?? undefined,
                });
              }
            }
          })
      : [],
  );

  return { chartDescriptors, images };
}

function parseCellValue(
  cell: Element,
  sharedStrings: string[],
  style: ParsedStyle | undefined,
  numberFormats: Map<number, string>,
  date1904: boolean,
): SpreadsheetCellValue {
  const type = cell.getAttribute('t');
  const rawValue = childByLocalName(cell, 'v')?.textContent ?? '';

  if (type === 's') {
    return sharedStrings[Number(rawValue)] ?? '';
  }

  if (type === 'inlineStr') {
    const inlineString = childByLocalName(cell, 'is');
    return inlineString
      ? elementsByLocalName(inlineString, 't')
          .map((textNode) => textNode.textContent ?? '')
          .join('')
      : '';
  }

  if (type === 'b') {
    return rawValue === '1';
  }

  if (type === 'str') {
    return rawValue;
  }

  if (type === 'e') {
    return rawValue ? `#${rawValue}` : null;
  }

  if (!rawValue) {
    return null;
  }

  const numericValue = Number(rawValue);
  if (!Number.isFinite(numericValue)) {
    return rawValue;
  }

  return isDateFormat(style?.numFmtId, numberFormats) ? excelSerialDate(numericValue, date1904) : numericValue;
}

function setCell(rows: ParsedSpreadsheetCell[][], rowIndex: number, columnIndex: number, cell: ParsedSpreadsheetCell) {
  if (!rows[rowIndex]) {
    rows[rowIndex] = [];
  }

  rows[rowIndex][columnIndex] = cell;
}

function mergeCell(rows: ParsedSpreadsheetCell[][], rowIndex: number, columnIndex: number, cell: Partial<ParsedSpreadsheetCell>) {
  if (!rows[rowIndex]) {
    rows[rowIndex] = [];
  }

  const current = rows[rowIndex][columnIndex] ?? { value: null };
  rows[rowIndex][columnIndex] = {
    ...current,
    ...cell,
    comments: cell.comments ? [...(current.comments ?? []), ...cell.comments] : current.comments,
    style: {
      ...(current.style ?? {}),
      ...(cell.style ?? {}),
    },
  };
}

function parseWorksheet(
  worksheetXml: string,
  sharedStrings: string[],
  styles: ParsedStyle[],
  numberFormats: Map<number, string>,
  date1904: boolean,
  commentsByRef: Map<string, SpreadsheetComment[]>,
): ParsedWorksheetData {
  const document = parseXml(worksheetXml);
  const rows: ParsedSpreadsheetCell[][] = [];
  const mergeRanges = elementsByLocalName(document, 'mergeCell')
    .map((mergeNode) => parseMergeRange(mergeNode.getAttribute('ref') ?? ''))
    .filter((range): range is MergeRange => Boolean(range));

  elementsByLocalName(document, 'c').forEach((cell) => {
    const ref = cell.getAttribute('r');
    if (!ref) {
      return;
    }

    const position = cellRefToIndexes(ref);
    if (!position) {
      return;
    }

    const styleIndex = Number(cell.getAttribute('s') ?? 0);
    const parsedStyle = styles[styleIndex];
    const value = parseCellValue(cell, sharedStrings, parsedStyle, numberFormats, date1904);
    const formattedValue = formatCellDisplayValue(value, parsedStyle?.numFmtId, numberFormats);

    setCell(rows, position.rowIndex, position.columnIndex, {
      formattedValue,
      value,
      style: parsedStyle?.style,
    });
  });

  commentsByRef.forEach((comments, ref) => {
    const position = cellRefToIndexes(ref);

    if (position) {
      mergeCell(rows, position.rowIndex, position.columnIndex, { comments });
    }
  });

  applyMergeRanges(rows, mergeRanges);

  return {
    columnWidths: parseColumnWidths(document),
    rowHeights: parseRowHeights(document),
    rows,
  };
}

export async function parseXlsxWorkbook(arrayBuffer: ArrayBuffer): Promise<ParsedXlsxWorkbook> {
  const zip = await JSZip.loadAsync(arrayBuffer);
  const readText = async (path: string): Promise<string | undefined> => {
    const file = zip.file(path);
    return file ? file.async('text') : undefined;
  };
  const readBase64 = async (path: string): Promise<string | undefined> => {
    const file = zip.file(path);
    return file ? file.async('base64') : undefined;
  };
  const workbookXml = await readText('xl/workbook.xml');

  if (!workbookXml) {
    throw new Error('Unable to find workbook metadata in this XLSX file.');
  }

  const workbookDocument = parseXml(workbookXml);
  const date1904 = elementsByLocalName(workbookDocument, 'workbookPr')[0]?.getAttribute('date1904') === '1';
  const stylesXml = await readText('xl/styles.xml');
  const stylesDocument = stylesXml ? parseXml(stylesXml) : undefined;
  const numberFormats = stylesDocument ? parseNumberFormats(stylesDocument) : new Map<number, string>();
  const themeColors = parseThemeColors(await readText('xl/theme/theme1.xml'));
  const styles = parseStyles(stylesXml, themeColors);
  const sharedStrings = parseSharedStrings(await readText('xl/sharedStrings.xml'));
  const sheets = parseWorkbookSheets(workbookXml, await readText('xl/_rels/workbook.xml.rels'));
  const persons = new Map([
    ...parsePersons(await readText('xl/persons/person.xml')),
    ...parsePersons(await readText('xl/persons/persons.xml')),
  ]);

  const parsedSheets: ParsedSheetWithChartDescriptors[] = await Promise.all(
    sheets.map(async (sheet) => {
      const sheetRelationships = parseRelationships(await readText(relationshipPathForPart(sheet.path)), sheet.path);
      const noteRelationships = sheetRelationships.filter((relationship) => relationship.type.toLowerCase().endsWith('/comments'));
      const threadedRelationships = sheetRelationships.filter((relationship) =>
        relationship.type.toLowerCase().includes('threadedcomment'),
      );
      const drawingRelationships = sheetRelationships.filter((relationship) => relationship.type.toLowerCase().endsWith('/drawing'));
      const noteComments = await Promise.all(noteRelationships.map(async (relationship) => parseClassicComments(await readText(relationship.target))));
      const threadedComments = await Promise.all(
        threadedRelationships.map(async (relationship) => parseThreadedComments(await readText(relationship.target), persons)),
      );
      const worksheetDrawings = await Promise.all(
        drawingRelationships.map(async (relationship) => {
          const drawingXml = await readText(relationship.target);

          return drawingXml
            ? parseWorksheetDrawings(
                relationship.target,
                drawingXml,
                await readText(relationshipPathForPart(relationship.target)),
                readText,
                readBase64,
              )
            : { chartDescriptors: [], images: [] };
        }),
      );
      const worksheet = parseWorksheet(
        (await readText(sheet.path)) ?? '',
        sharedStrings,
        styles,
        numberFormats,
        date1904,
        mergeCommentMaps(...noteComments, ...threadedComments),
      );

      return {
        chartDescriptors: worksheetDrawings.flatMap((drawing) => drawing.chartDescriptors),
        columnWidths: worksheet.columnWidths,
        drawings: worksheetDrawings.flatMap((drawing) => drawing.images),
        name: sheet.name,
        rowHeights: worksheet.rowHeights,
        rows: worksheet.rows,
      };
    }),
  );
  const sheetLookup = new Map(parsedSheets.map((sheet) => [sheet.name, sheet.rows]));

  return {
    sheets: parsedSheets.map((sheet) => {
      const chartDrawings = sheet.chartDescriptors.map((descriptor) => parseChartDrawing(descriptor, sheet.name, sheetLookup));
      const { chartDescriptors, ...parsedSheet } = sheet;

      return {
        ...parsedSheet,
        drawings: [...(parsedSheet.drawings ?? []), ...chartDrawings],
      };
    }),
  };
}
