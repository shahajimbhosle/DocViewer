import JSZip from 'jszip';
import type {
  ParsedSpreadsheetCell,
  ParsedXlsxWorkbook,
  SpreadsheetCellStyle,
  SpreadsheetCellValue,
  SpreadsheetComment,
} from './xlsx';

interface OdsStyle {
  cellStyle?: SpreadsheetCellStyle;
  columnWidth?: number;
  rowHeight?: number;
}

interface ParsedOdsColumns {
  defaultCellStyles: Array<SpreadsheetCellStyle | undefined>;
  widths: number[];
}

const maxParsedRows = 5000;
const maxParsedColumns = 512;
const maxRepeatedBlankRows = 50;
const maxRepeatedDataRows = 250;
const maxRepeatedStyledRows = 100;
const maxRepeatedCells = 256;
const maxRepeatedStyledCells = 64;
const maxRowSpan = 200;

function parseXml(xml: string): Document {
  return new DOMParser().parseFromString(xml, 'application/xml');
}

function localAttribute(element: Element | undefined, localName: string): string | null {
  if (!element) {
    return null;
  }

  const directValue = element.getAttribute(localName);
  if (directValue !== null) {
    return directValue;
  }

  return Array.from(element.attributes).find((attribute) => attribute.localName === localName)?.value ?? null;
}

function childElementsByLocalName(element: Element, localName: string): Element[] {
  return Array.from(element.children).filter((child) => child.localName === localName);
}

function childElementByLocalName(element: Element, localName: string): Element | undefined {
  return childElementsByLocalName(element, localName)[0];
}

function elementsByLocalName(root: ParentNode, localName: string): Element[] {
  const searchableRoot = root as ParentNode & {
    getElementsByTagNameNS?: (namespace: string, localName: string) => HTMLCollectionOf<Element>;
  };

  if (typeof searchableRoot.getElementsByTagNameNS === 'function') {
    return Array.from(searchableRoot.getElementsByTagNameNS('*', localName));
  }

  return Array.from(root.querySelectorAll('*')).filter((element) => element.localName === localName);
}

function numberAttribute(element: Element, localName: string, fallback = 1): number {
  const value = Number(localAttribute(element, localName) ?? fallback);
  return Number.isFinite(value) && value > 0 ? Math.floor(value) : fallback;
}

function measurementToPixels(value: string | null): number | undefined {
  if (!value) {
    return undefined;
  }

  const match = /^(-?\d+(?:\.\d+)?)(cm|mm|in|pt|pc|px)$/i.exec(value.trim());
  if (!match) {
    return undefined;
  }

  const amount = Number(match[1]);
  const unit = match[2].toLowerCase();

  if (!Number.isFinite(amount) || amount <= 0) {
    return undefined;
  }

  const pixels =
    unit === 'cm'
      ? (amount * 96) / 2.54
      : unit === 'mm'
        ? (amount * 96) / 25.4
        : unit === 'in'
          ? amount * 96
          : unit === 'pt'
            ? (amount * 96) / 72
            : unit === 'pc'
              ? amount * 16
              : amount;

  return Math.max(1, Math.min(640, Math.round(pixels)));
}

function mapTextAlign(value: string | null): SpreadsheetCellStyle['textAlign'] | undefined {
  if (value === 'center') {
    return 'center';
  }

  if (value === 'end' || value === 'right') {
    return 'right';
  }

  if (value === 'justify') {
    return 'justify';
  }

  if (value === 'start' || value === 'left') {
    return 'left';
  }

  return undefined;
}

function mapVerticalAlign(value: string | null): SpreadsheetCellStyle['verticalAlign'] | undefined {
  if (value === 'middle' || value === 'center') {
    return 'middle';
  }

  if (value === 'bottom') {
    return 'bottom';
  }

  if (value === 'top') {
    return 'top';
  }

  return undefined;
}

function normalizeOdsBorder(value: string | null): string | undefined {
  const border = value?.trim();

  if (!border || border === 'none' || border === 'hidden') {
    return undefined;
  }

  if (/^0(?:\.0+)?(?:cm|mm|in|pt|pc|px)?\s/i.test(border)) {
    return undefined;
  }

  return border;
}

function applyOdsBorderStyles(cellStyle: SpreadsheetCellStyle, tableCellProperties: Element | undefined) {
  const border = normalizeOdsBorder(localAttribute(tableCellProperties, 'border'));
  const borderTop = normalizeOdsBorder(localAttribute(tableCellProperties, 'border-top')) ?? border;
  const borderRight = normalizeOdsBorder(localAttribute(tableCellProperties, 'border-right')) ?? border;
  const borderBottom = normalizeOdsBorder(localAttribute(tableCellProperties, 'border-bottom')) ?? border;
  const borderLeft = normalizeOdsBorder(localAttribute(tableCellProperties, 'border-left')) ?? border;

  if (borderTop) {
    cellStyle.borderTop = borderTop;
  }

  if (borderRight) {
    cellStyle.borderRight = borderRight;
  }

  if (borderBottom) {
    cellStyle.borderBottom = borderBottom;
  }

  if (borderLeft) {
    cellStyle.borderLeft = borderLeft;
  }
}

function parseTextNodeContent(node: Node): string {
  if (node.nodeType === Node.TEXT_NODE) {
    return node.textContent ?? '';
  }

  if (!(node instanceof Element)) {
    return node.textContent ?? '';
  }

  if (node.localName === 'line-break') {
    return '\n';
  }

  if (node.localName === 'tab') {
    return '\t';
  }

  if (node.localName === 's') {
    return ' '.repeat(numberAttribute(node, 'c', 1));
  }

  return Array.from(node.childNodes).map(parseTextNodeContent).join('');
}

function directParagraphText(element: Element): string {
  const paragraphs = childElementsByLocalName(element, 'p');

  if (paragraphs.length === 0) {
    return '';
  }

  return paragraphs.map((paragraph) => Array.from(paragraph.childNodes).map(parseTextNodeContent).join('')).join('\n');
}

function parseOdsStyle(styleElement: Element): OdsStyle {
  const tableCellProperties = childElementByLocalName(styleElement, 'table-cell-properties');
  const paragraphProperties = childElementByLocalName(styleElement, 'paragraph-properties');
  const textProperties = childElementByLocalName(styleElement, 'text-properties');
  const tableColumnProperties = childElementByLocalName(styleElement, 'table-column-properties');
  const tableRowProperties = childElementByLocalName(styleElement, 'table-row-properties');
  const cellStyle: SpreadsheetCellStyle = {};
  const backgroundColor = localAttribute(tableCellProperties, 'background-color');
  const color = localAttribute(textProperties, 'color');
  const textAlign = mapTextAlign(localAttribute(paragraphProperties, 'text-align'));
  const verticalAlign = mapVerticalAlign(
    localAttribute(tableCellProperties, 'vertical-align') ?? localAttribute(paragraphProperties, 'vertical-align'),
  );
  const wrapOption = localAttribute(tableCellProperties, 'wrap-option') ?? localAttribute(paragraphProperties, 'wrap-option');
  const writingMode = localAttribute(tableCellProperties, 'writing-mode') ?? localAttribute(paragraphProperties, 'writing-mode');
  const fontSize = localAttribute(textProperties, 'font-size');

  applyOdsBorderStyles(cellStyle, tableCellProperties);

  if (backgroundColor && backgroundColor !== 'transparent') {
    cellStyle.backgroundColor = backgroundColor;
  }

  if (color) {
    cellStyle.color = color;
  }

  if (textAlign) {
    cellStyle.textAlign = textAlign;
  }

  if (verticalAlign) {
    cellStyle.verticalAlign = verticalAlign;
  }

  if (wrapOption === 'wrap') {
    cellStyle.whiteSpace = 'pre-wrap';
  } else if (wrapOption === 'no-wrap') {
    cellStyle.whiteSpace = 'nowrap';
  }

  if (writingMode?.includes('tb')) {
    cellStyle.writingMode = 'vertical-rl';
  }

  if (writingMode?.includes('rl')) {
    cellStyle.direction = 'rtl';
  }

  if (localAttribute(textProperties, 'font-weight') === 'bold') {
    cellStyle.fontWeight = '700';
  }

  if (localAttribute(textProperties, 'font-style') === 'italic') {
    cellStyle.fontStyle = 'italic';
  }

  if (localAttribute(textProperties, 'text-underline-style') && localAttribute(textProperties, 'text-underline-style') !== 'none') {
    cellStyle.textDecoration = 'underline';
  }

  if (fontSize) {
    cellStyle.fontSize = fontSize;
  }

  return {
    cellStyle: Object.keys(cellStyle).length > 0 ? cellStyle : undefined,
    columnWidth: measurementToPixels(localAttribute(tableColumnProperties, 'column-width')),
    rowHeight: measurementToPixels(localAttribute(tableRowProperties, 'row-height')),
  };
}

function parseStyleMap(...xmlDocuments: Array<Document | undefined>): Map<string, OdsStyle> {
  const styles = new Map<string, OdsStyle>();

  xmlDocuments.forEach((xmlDocument) => {
    if (!xmlDocument) {
      return;
    }

    elementsByLocalName(xmlDocument, 'style').forEach((styleElement) => {
      const name = localAttribute(styleElement, 'name');

      if (name) {
        styles.set(name, parseOdsStyle(styleElement));
      }
    });
  });

  return styles;
}

function parseCellValue(cellElement: Element): SpreadsheetCellValue {
  const valueType = localAttribute(cellElement, 'value-type');
  const text = directParagraphText(cellElement);

  if (valueType === 'float' || valueType === 'percentage' || valueType === 'currency') {
    const numericValue = Number(localAttribute(cellElement, 'value'));
    return Number.isFinite(numericValue) ? numericValue : text || null;
  }

  if (valueType === 'boolean') {
    return localAttribute(cellElement, 'boolean-value') === 'true';
  }

  if (valueType === 'date') {
    const dateValue = localAttribute(cellElement, 'date-value');
    return dateValue ? new Date(dateValue) : text || null;
  }

  if (valueType === 'time') {
    return localAttribute(cellElement, 'time-value') ?? text ?? null;
  }

  return text || null;
}

function parseComments(cellElement: Element): SpreadsheetComment[] {
  const comments: SpreadsheetComment[] = [];

  childElementsByLocalName(cellElement, 'annotation').forEach((annotation) => {
    const text = directParagraphText(annotation).trim();

    if (!text) {
      return;
    }

    comments.push({
      author: childElementByLocalName(annotation, 'creator')?.textContent?.trim() || undefined,
      date: childElementByLocalName(annotation, 'date')?.textContent?.trim() || undefined,
      kind: 'note',
      text,
    });
  });

  return comments;
}

function mergeStyles(...styles: Array<SpreadsheetCellStyle | undefined>): SpreadsheetCellStyle | undefined {
  const merged = Object.assign({}, ...styles.filter(Boolean));
  return Object.keys(merged).length > 0 ? merged : undefined;
}

function hasCellPayload(cell: ParsedSpreadsheetCell): boolean {
  return (
    cell.value !== null ||
    Boolean(cell.comments?.length) ||
    Boolean(cell.colSpan && cell.colSpan > 1) ||
    Boolean(cell.rowSpan && cell.rowSpan > 1)
  );
}

function hasVisibleCellData(cell: ParsedSpreadsheetCell): boolean {
  return hasCellPayload(cell) || Boolean(cell.style && Object.keys(cell.style).length > 0);
}

function cloneCell(cell: ParsedSpreadsheetCell): ParsedSpreadsheetCell {
  return {
    ...cell,
    comments: cell.comments ? [...cell.comments] : undefined,
    style: cell.style ? { ...cell.style } : undefined,
  };
}

function cloneSparseRow(row: ParsedSpreadsheetCell[]): ParsedSpreadsheetCell[] {
  const clone: ParsedSpreadsheetCell[] = [];

  row.forEach((cell, columnIndex) => {
    if (cell) {
      clone[columnIndex] = cloneCell(cell);
    }
  });

  return clone;
}

function parseColumns(tableElement: Element, styles: Map<string, OdsStyle>): ParsedOdsColumns {
  const widths: number[] = [];
  const defaultCellStyles: Array<SpreadsheetCellStyle | undefined> = [];
  let columnIndex = 0;

  for (const columnElement of childElementsByLocalName(tableElement, 'table-column')) {
    if (columnIndex >= maxParsedColumns) {
      break;
    }

    const repeatCount = Math.min(
      numberAttribute(columnElement, 'number-columns-repeated', 1),
      maxRepeatedCells,
      maxParsedColumns - columnIndex,
    );
    const columnStyle = styles.get(localAttribute(columnElement, 'style-name') ?? '');
    const defaultCellStyle = styles.get(localAttribute(columnElement, 'default-cell-style-name') ?? '')?.cellStyle;

    for (let repeatIndex = 0; repeatIndex < repeatCount; repeatIndex += 1) {
      if (columnStyle?.columnWidth) {
        widths[columnIndex] = columnStyle.columnWidth;
      }

      if (defaultCellStyle) {
        defaultCellStyles[columnIndex] = defaultCellStyle;
      }

      columnIndex += 1;
    }
  }

  return { defaultCellStyles, widths };
}

function parseRow(
  rowElement: Element,
  styles: Map<string, OdsStyle>,
  columnDefaultCellStyles: Array<SpreadsheetCellStyle | undefined>,
): ParsedSpreadsheetCell[] {
  const cells: ParsedSpreadsheetCell[] = [];
  const rowDefaultStyle = styles.get(localAttribute(rowElement, 'default-cell-style-name') ?? '')?.cellStyle;
  let columnIndex = 0;

  for (const cellElement of Array.from(rowElement.children)
    .filter((child) => child.localName === 'table-cell' || child.localName === 'covered-table-cell')
  ) {
    if (columnIndex >= maxParsedColumns) {
      break;
    }

    const repeatCount = Math.min(
      numberAttribute(cellElement, 'number-columns-repeated', 1),
      maxRepeatedCells,
      maxParsedColumns - columnIndex,
    );

    if (cellElement.localName === 'covered-table-cell') {
      columnIndex += repeatCount;
      continue;
    }

    const remainingColumns = maxParsedColumns - columnIndex;
    const colSpan = Math.min(numberAttribute(cellElement, 'number-columns-spanned', 1), remainingColumns);
    const rowSpan = Math.min(numberAttribute(cellElement, 'number-rows-spanned', 1), maxRowSpan);
    const comments = parseComments(cellElement);
    const parsedCell: ParsedSpreadsheetCell = {
      value: parseCellValue(cellElement),
      comments: comments.length > 0 ? comments : undefined,
      colSpan: colSpan > 1 ? colSpan : undefined,
      rowSpan: rowSpan > 1 ? rowSpan : undefined,
      style: mergeStyles(
        columnDefaultCellStyles[columnIndex],
        rowDefaultStyle,
        styles.get(localAttribute(cellElement, 'style-name') ?? '')?.cellStyle,
      ),
    };
    const hasPayload = hasCellPayload(parsedCell);
    const shouldMaterialize = hasPayload || Boolean(parsedCell.style && Object.keys(parsedCell.style).length > 0);
    const materializedRepeatCount = shouldMaterialize
      ? Math.min(repeatCount, hasPayload ? maxRepeatedCells : maxRepeatedStyledCells)
      : 0;

    for (let repeatIndex = 0; repeatIndex < materializedRepeatCount && columnIndex < maxParsedColumns; repeatIndex += 1) {
      cells[columnIndex] = cloneCell(parsedCell);
      columnIndex += colSpan;
    }

    if (shouldMaterialize) {
      columnIndex = Math.min(maxParsedColumns, columnIndex + Math.max(0, repeatCount - materializedRepeatCount) * colSpan);
    } else {
      columnIndex = Math.min(maxParsedColumns, columnIndex + repeatCount);
    }
  }

  return cells;
}

function applySpans(rows: ParsedSpreadsheetCell[][]) {
  rows.forEach((row, rowIndex) => {
    row.forEach((cell, columnIndex) => {
      if (!cell || cell.hiddenByMerge) {
        return;
      }

      const rowSpan = Math.min(cell.rowSpan ?? 1, maxParsedRows - rowIndex);
      const colSpan = Math.min(cell.colSpan ?? 1, maxParsedColumns - columnIndex);

      if (rowSpan <= 1 && colSpan <= 1) {
        return;
      }

      for (let offsetRow = 0; offsetRow < rowSpan; offsetRow += 1) {
        for (let offsetColumn = 0; offsetColumn < colSpan; offsetColumn += 1) {
          if (offsetRow === 0 && offsetColumn === 0) {
            continue;
          }

          const nextRowIndex = rowIndex + offsetRow;
          const nextColumnIndex = columnIndex + offsetColumn;

          if (nextRowIndex >= maxParsedRows || nextColumnIndex >= maxParsedColumns) {
            continue;
          }

          rows[nextRowIndex] ??= [];
          rows[nextRowIndex][nextColumnIndex] = {
            value: null,
            hiddenByMerge: true,
          };
        }
      }
    });
  });
}

function parseTable(tableElement: Element, styles: Map<string, OdsStyle>) {
  const columns = parseColumns(tableElement, styles);
  const rows: ParsedSpreadsheetCell[][] = [];
  const rowHeights: number[] = [];
  let rowIndex = 0;

  for (const rowElement of childElementsByLocalName(tableElement, 'table-row')) {
    if (rowIndex >= maxParsedRows) {
      break;
    }

    const parsedRow = parseRow(rowElement, styles, columns.defaultCellStyles);
    const rowStyle = styles.get(localAttribute(rowElement, 'style-name') ?? '');
    const hasVisibleRowData = parsedRow.some((cell) => cell && hasVisibleCellData(cell));
    const hasRowPayload = parsedRow.some((cell) => cell && hasCellPayload(cell));
    const repeatLimit = hasRowPayload
      ? maxRepeatedDataRows
      : hasVisibleRowData || rowStyle?.rowHeight
        ? maxRepeatedStyledRows
        : maxRepeatedBlankRows;
    const repeatCount = Math.min(numberAttribute(rowElement, 'number-rows-repeated', 1), repeatLimit, maxParsedRows - rowIndex);

    for (let repeatIndex = 0; repeatIndex < repeatCount; repeatIndex += 1) {
      if (hasVisibleRowData) {
        rows[rowIndex] = cloneSparseRow(parsedRow);
      }

      if (rowStyle?.rowHeight) {
        rowHeights[rowIndex] = rowStyle.rowHeight;
      }

      rowIndex += 1;
    }
  }

  applySpans(rows);

  return {
    columnWidths: columns.widths,
    name: localAttribute(tableElement, 'name') ?? 'Sheet',
    rowHeights,
    rows,
  };
}

export async function parseOdsWorkbook(arrayBuffer: ArrayBuffer): Promise<ParsedXlsxWorkbook> {
  const zip = await JSZip.loadAsync(arrayBuffer.slice(0));
  const contentXml = await zip.file('content.xml')?.async('text');

  if (!contentXml) {
    throw new Error('Unable to find content.xml in this ODS file.');
  }

  const stylesXml = await zip.file('styles.xml')?.async('text');
  const stylesDocument = stylesXml ? parseXml(stylesXml) : undefined;
  const contentDocument = parseXml(contentXml);
  const styles = parseStyleMap(stylesDocument, contentDocument);
  const tables = elementsByLocalName(contentDocument, 'table');
  const sheets = tables.map((table) => parseTable(table, styles));

  return {
    sheets: sheets.length > 0 ? sheets : [{ name: 'Sheet1', rows: [] }],
  };
}
