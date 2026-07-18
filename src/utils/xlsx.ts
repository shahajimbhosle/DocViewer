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

export interface ParsedSpreadsheetCell {
  value: SpreadsheetCellValue;
  style?: SpreadsheetCellStyle;
  comments?: SpreadsheetComment[];
  colSpan?: number;
  hiddenByMerge?: boolean;
  rowSpan?: number;
}

export interface ParsedXlsxSheet {
  columnWidths?: number[];
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

type SpreadsheetBorderProperty = 'borderBottom' | 'borderLeft' | 'borderRight' | 'borderTop';

const builtinDateFormats = new Set([14, 15, 16, 17, 22, 27, 30, 36, 45, 46, 47, 50, 57]);

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

function cellRefToIndexes(cellRef: string): { rowIndex: number; columnIndex: number } | null {
  const match = /^([A-Z]+)(\d+)$/i.exec(cellRef);
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

function excelSerialDate(value: number, date1904: boolean): Date {
  const epoch = Date.UTC(date1904 ? 1904 : 1899, date1904 ? 0 : 11, date1904 ? 1 : 30);
  return new Date(epoch + value * 86_400_000);
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

          if (color) {
            style.color = color;
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

    setCell(rows, position.rowIndex, position.columnIndex, {
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

  return {
    sheets: await Promise.all(
      sheets.map(async (sheet) => {
        const sheetRelationships = parseRelationships(await readText(relationshipPathForPart(sheet.path)), sheet.path);
        const noteRelationships = sheetRelationships.filter((relationship) => relationship.type.toLowerCase().endsWith('/comments'));
        const threadedRelationships = sheetRelationships.filter((relationship) =>
          relationship.type.toLowerCase().includes('threadedcomment'),
        );
        const noteComments = await Promise.all(noteRelationships.map(async (relationship) => parseClassicComments(await readText(relationship.target))));
        const threadedComments = await Promise.all(
          threadedRelationships.map(async (relationship) => parseThreadedComments(await readText(relationship.target), persons)),
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
          columnWidths: worksheet.columnWidths,
          name: sheet.name,
          rowHeights: worksheet.rowHeights,
          rows: worksheet.rows,
        };
      }),
    ),
  };
}
