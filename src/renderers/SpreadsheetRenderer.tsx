import { useVirtualizer } from '@tanstack/react-virtual';
import { useEffect, useMemo, useRef, useState } from 'react';
import type { CSSProperties, RefObject } from 'react';
import type { DocumentRenderer, DocumentRendererProps } from '../types';
import { countMatches, highlightText } from '../utils/highlight';
import { parseOdsWorkbook } from '../utils/ods';
import { parseXlsWorkbook, type XlsCell } from '../utils/xls';
import {
  parseXlsxWorkbook,
  type ParsedSpreadsheetCell,
  type ParsedXlsxWorkbook,
  type SpreadsheetCellValue,
} from '../utils/xlsx';

type SpreadsheetCell = ParsedSpreadsheetCell;
type SpreadsheetRow = SpreadsheetCell[];
type SpreadsheetWorkbook = ParsedXlsxWorkbook;

const gridCellWidth = 96;
const gridRowHeight = 30;
const gridColumnHeaderHeight = 30;
const gridRowHeaderWidth = 48;
const sheetTabsHeight = 38;

interface Size {
  width: number;
  height: number;
}

function valueToText(value: SpreadsheetCellValue): string {
  if (value instanceof Date) {
    return value.toLocaleDateString();
  }

  if (value === null || typeof value === 'undefined') {
    return '';
  }

  return String(value);
}

function cellToText(cell?: SpreadsheetCell | null): string {
  return valueToText(cell?.value ?? null);
}

function xlsRowsToSpreadsheetRows(rows: XlsCell[][]): SpreadsheetRow[] {
  return rows.map((row) =>
    row.map((value) => ({
      value,
    })),
  );
}

function cellCommentLabel(cell?: SpreadsheetCell): string {
  return (
    cell?.comments
      ?.map((comment) => {
        const author = comment.author ? `${comment.author}: ` : '';
        return `${author}${comment.text}`;
      })
      .join('\n\n') ?? ''
  );
}

function columnLabel(index: number): string {
  let value = index + 1;
  let label = '';

  while (value > 0) {
    const remainder = (value - 1) % 26;
    label = String.fromCharCode(65 + remainder) + label;
    value = Math.floor((value - 1) / 26);
  }

  return label;
}

function scaledPixels(value: number): string {
  return `${Math.max(1, Math.round(value))}px`;
}

function columnWidthAt(columnIndex: number, columnWidths: number[], zoom: number): number {
  return (columnWidths[columnIndex] ?? gridCellWidth) * zoom;
}

function rowHeightAt(rowIndex: number, rowHeights: number[], zoom: number): number {
  return (rowHeights[rowIndex] ?? gridRowHeight) * zoom;
}

function columnSpanWidthAt(columnIndex: number, colSpan: number, columnWidths: number[], zoom: number): number {
  let width = 0;

  for (let offset = 0; offset < colSpan; offset += 1) {
    width += columnWidthAt(columnIndex + offset, columnWidths, zoom);
  }

  return width;
}

function rowSpanHeightAt(rowIndex: number, rowSpan: number, rowHeights: number[], zoom: number): number {
  let height = 0;

  for (let offset = 0; offset < rowSpan; offset += 1) {
    height += rowHeightAt(rowIndex + offset, rowHeights, zoom);
  }

  return height;
}

function defaultCellTextAlign(cell?: SpreadsheetCell): CSSProperties['textAlign'] {
  const value = cell?.value;

  if (typeof value === 'number' || value instanceof Date) {
    return 'right';
  }

  if (typeof value === 'boolean') {
    return 'center';
  }

  return 'left';
}

function effectiveCellWhiteSpace(cell?: SpreadsheetCell): CSSProperties['whiteSpace'] {
  if (cell?.style?.whiteSpace) {
    return cell.style.whiteSpace;
  }

  return cellToText(cell).includes('\n') ? 'pre-line' : 'nowrap';
}

function effectiveCellOverflowWrap(whiteSpace: CSSProperties['whiteSpace']): CSSProperties['overflowWrap'] {
  return whiteSpace === 'normal' || whiteSpace === 'pre-line' || whiteSpace === 'pre-wrap' ? 'break-word' : 'normal';
}

function estimateWrappedLineCount(text: string, width: number, zoom: number, whiteSpace: CSSProperties['whiteSpace']): number {
  if (!text) {
    return 1;
  }

  if (whiteSpace === 'nowrap') {
    return 1;
  }

  const availableWidth = Math.max(1, width - 6 * zoom);
  const averageCharacterWidth = Math.max(1, 7 * zoom);
  const charactersPerLine = Math.max(1, Math.floor(availableWidth / averageCharacterWidth));
  const lines = whiteSpace === 'normal' ? [text.replace(/\r\n|\r|\n/g, ' ')] : text.split(/\r\n|\r|\n/);

  return lines.reduce((count, line) => count + Math.max(1, Math.ceil(line.length / charactersPerLine)), 0);
}

function estimateRowHeight(row: SpreadsheetRow, rowIndex: number, rowHeights: number[], columnWidths: number[], zoom: number): number {
  let height = rowHeightAt(rowIndex, rowHeights, zoom);

  row.forEach((cell, columnIndex) => {
    if (!cell || cell.hiddenByMerge || (cell.rowSpan && cell.rowSpan > 1)) {
      return;
    }

    const whiteSpace = effectiveCellWhiteSpace(cell);
    const lineCount = estimateWrappedLineCount(
      cellToText(cell),
      columnSpanWidthAt(columnIndex, cell.colSpan ?? 1, columnWidths, zoom),
      zoom,
      whiteSpace,
    );

    height = Math.max(height, Math.ceil(lineCount * 17 * zoom + 6 * zoom));
  });

  return height;
}

function cellStyle(
  cell: SpreadsheetCell | undefined,
  columnIndex: number,
  rowIndex: number,
  columnWidths: number[],
  rowHeights: number[],
  zoom: number,
): CSSProperties {
  const colSpan = cell?.colSpan ?? 1;
  const rowSpan = cell?.rowSpan ?? 1;
  const customStyle = cell?.style ?? {};
  const whiteSpace = effectiveCellWhiteSpace(cell);
  const width = scaledPixels(columnSpanWidthAt(columnIndex, colSpan, columnWidths, zoom));
  const height = scaledPixels(rowSpanHeightAt(rowIndex, rowSpan, rowHeights, 1));
  const textAlign = customStyle.textAlign ?? defaultCellTextAlign(cell);
  const verticalAlign = customStyle.verticalAlign ?? 'bottom';

  return {
    alignItems: verticalAlign === 'middle' ? 'center' : verticalAlign === 'top' ? 'flex-start' : 'flex-end',
    display: 'flex',
    height,
    justifyContent: textAlign === 'right' ? 'flex-end' : textAlign === 'center' ? 'center' : 'flex-start',
    maxWidth: width,
    minWidth: width,
    overflowWrap: effectiveCellOverflowWrap(whiteSpace),
    textAlign,
    verticalAlign,
    whiteSpace,
    width,
    ...customStyle,
  };
}

function useElementSize(ref: RefObject<HTMLElement>): Size {
  const [size, setSize] = useState<Size>({ width: 0, height: 0 });

  useEffect(() => {
    const element = ref.current;
    if (!element) {
      return undefined;
    }

    const updateSize = () => {
      setSize({
        width: element.clientWidth,
        height: element.clientHeight,
      });
    };

    updateSize();
    const observer = new ResizeObserver(updateSize);
    observer.observe(element);

    return () => observer.disconnect();
  }, [ref]);

  return size;
}

function SpreadsheetRendererComponent({ file, state, actions, viewportRef }: DocumentRendererProps) {
  const [sheetNames, setSheetNames] = useState<string[]>([]);
  const [sheetName, setSheetName] = useState('');
  const [rows, setRows] = useState<SpreadsheetRow[]>([]);
  const [columnWidths, setColumnWidths] = useState<number[]>([]);
  const [rowHeights, setRowHeights] = useState<number[]>([]);
  const [workbook, setWorkbook] = useState<SpreadsheetWorkbook | null>(null);
  const viewportSize = useElementSize(viewportRef);
  const gridScrollRef = useRef<HTMLDivElement>(null);
  const isLegacyXls = file.extension === 'xls' || file.mimeType === 'application/vnd.ms-excel';
  const isOds = file.extension === 'ods' || file.mimeType === 'application/vnd.oasis.opendocument.spreadsheet';

  useEffect(() => {
    let cancelled = false;

    async function loadSheets() {
      actions.setLoading(true);
      let nextWorkbook: SpreadsheetWorkbook;

      if (isLegacyXls) {
        nextWorkbook = {
          sheets: parseXlsWorkbook(file.arrayBuffer).sheets.map((sheet) => ({
            name: sheet.name,
            rows: xlsRowsToSpreadsheetRows(sheet.rows),
          })),
        };
      } else if (isOds) {
        nextWorkbook = await parseOdsWorkbook(file.arrayBuffer);
      } else {
        nextWorkbook = await parseXlsxWorkbook(file.arrayBuffer);
      }

      if (cancelled) {
        return;
      }

      const nextSheetNames = nextWorkbook.sheets.map((sheet) => sheet.name);

      setWorkbook(nextWorkbook);
      setSheetNames(nextSheetNames);
      setSheetName(nextSheetNames[0] ?? '');
      actions.setPageCount(undefined);
      actions.setDocumentInfo({ title: file.fileName });
      actions.setLoading(false);
    }

    loadSheets().catch((error: unknown) => {
      if (!cancelled) {
        actions.setLoading(false);
        actions.reportError(error);
      }
    });

    return () => {
      cancelled = true;
      actions.setLoading(false);
    };
  }, [actions, file, isLegacyXls, isOds]);

  useEffect(() => {
    const sheet = workbook?.sheets.find((candidate) => candidate.name === sheetName);
    setColumnWidths(sheet?.columnWidths ?? []);
    setRowHeights(sheet?.rowHeights ?? []);
    setRows(sheet?.rows ?? []);
  }, [sheetName, workbook]);

  const activeSearchTerm = state.searchTerm.trim();
  const plainText = useMemo(() => {
    if (!activeSearchTerm) {
      return '';
    }

    const parts: string[] = [];

    rows.forEach((row) => {
      row.forEach((cell) => {
        const text = cellToText(cell);

        if (text) {
          parts.push(text);
        }
      });
    });

    return parts.join('\n');
  }, [activeSearchTerm, rows]);
  const matches = useMemo(() => (activeSearchTerm ? countMatches(plainText, activeSearchTerm) : 0), [activeSearchTerm, plainText]);
  const populatedColumnCount = useMemo(() => {
    let count = 0;

    rows.forEach((row) => {
      row.forEach((cell, columnIndex) => {
        if (cell && !cell.hiddenByMerge) {
          count = Math.max(count, columnIndex + (cell.colSpan ?? 1));
        }
      });
    });

    return count;
  }, [rows]);
  const showSheetTabs = sheetNames.length > 0;
  const displayRowHeights = useMemo(
    () => rows.map((row, rowIndex) => estimateRowHeight(row, rowIndex, rowHeights, columnWidths, state.zoom)),
    [columnWidths, rowHeights, rows, state.zoom],
  );
  const averageColumnWidth = useMemo(() => {
    const populatedWidths = columnWidths.filter((width) => Number.isFinite(width) && width > 0);
    const averageWidth =
      populatedWidths.length > 0
        ? populatedWidths.reduce((sum, width) => sum + width, 0) / populatedWidths.length
        : gridCellWidth;

    return averageWidth * state.zoom;
  }, [columnWidths, state.zoom]);
  const averageRowHeight = useMemo(() => {
    const populatedHeights = displayRowHeights.filter((height) => Number.isFinite(height) && height > 0);
    const averageHeight =
      populatedHeights.length > 0
        ? populatedHeights.reduce((sum, height) => sum + height, 0) / populatedHeights.length
        : gridRowHeight;

    return averageHeight;
  }, [displayRowHeights]);
  const visibleColumnCount = useMemo(() => {
    const viewportColumns = Math.ceil(
      Math.max(0, viewportSize.width - gridRowHeaderWidth * state.zoom) / Math.max(1, averageColumnWidth),
    );
    return Math.max(1, populatedColumnCount, viewportColumns);
  }, [averageColumnWidth, populatedColumnCount, state.zoom, viewportSize.width]);
  const visibleRowCount = useMemo(() => {
    const tabsHeight = showSheetTabs ? sheetTabsHeight : 0;
    const viewportRows = Math.ceil(
      Math.max(0, viewportSize.height - tabsHeight - gridColumnHeaderHeight * state.zoom) / Math.max(1, averageRowHeight),
    );
    return Math.max(1, rows.length, viewportRows);
  }, [averageRowHeight, rows.length, showSheetTabs, state.zoom, viewportSize.height]);
  const rowHeaderPixelWidth = Math.max(1, Math.round(gridRowHeaderWidth * state.zoom));
  const columnHeaderPixelHeight = Math.max(1, Math.round(gridColumnHeaderHeight * state.zoom));
  const rowHeaderStyle: CSSProperties = {
    height: scaledPixels(gridRowHeight * state.zoom),
    width: scaledPixels(rowHeaderPixelWidth),
  };
  const rowVirtualizer = useVirtualizer({
    count: visibleRowCount,
    estimateSize: (index) => rowHeightAt(index, displayRowHeights, 1),
    getScrollElement: () => gridScrollRef.current,
    overscan: 12,
    paddingStart: columnHeaderPixelHeight,
  });
  const columnVirtualizer = useVirtualizer({
    count: visibleColumnCount,
    estimateSize: (index) => columnWidthAt(index, columnWidths, state.zoom),
    getScrollElement: () => gridScrollRef.current,
    horizontal: true,
    overscan: 6,
    paddingStart: rowHeaderPixelWidth,
  });
  const virtualRows = rowVirtualizer.getVirtualItems();
  const virtualColumns = columnVirtualizer.getVirtualItems();
  const virtualGridHeight = rowVirtualizer.getTotalSize();
  const virtualGridWidth = columnVirtualizer.getTotalSize();

  useEffect(() => {
    actions.setSearchStats({
      totalMatches: matches,
      currentPageMatches: matches,
      message: activeSearchTerm ? `${matches} match${matches === 1 ? '' : 'es'}` : undefined,
    });
  }, [actions, activeSearchTerm, matches]);

  return (
    <div
      className="ldv-spreadsheet-document"
      style={{
        fontSize: `${13 * state.zoom}px`,
        transform: `rotate(${state.rotation}deg)`,
        transformOrigin: 'center top',
      }}
    >
      <div className="ldv-spreadsheet-grid-scroll" ref={gridScrollRef}>
        <div
          aria-colcount={visibleColumnCount}
          aria-rowcount={visibleRowCount}
          className="ldv-spreadsheet-virtual-grid"
          role="grid"
          style={{
            height: scaledPixels(virtualGridHeight),
            width: scaledPixels(virtualGridWidth),
          }}
        >
          <div
            className="ldv-spreadsheet-virtual-header"
            role="row"
            style={{
              height: scaledPixels(columnHeaderPixelHeight),
              width: scaledPixels(virtualGridWidth),
            }}
          >
            <div
              aria-label="Row and column headers"
              className="ldv-sheet-corner ldv-spreadsheet-virtual-corner"
              role="columnheader"
              style={{
                height: scaledPixels(columnHeaderPixelHeight),
                width: scaledPixels(rowHeaderPixelWidth),
              }}
            />
            {virtualColumns.map((virtualColumn) => (
              <div
                aria-colindex={virtualColumn.index + 1}
                className="ldv-column-header ldv-spreadsheet-virtual-column-header"
                key={virtualColumn.key}
                role="columnheader"
                style={{
                  height: scaledPixels(columnHeaderPixelHeight),
                  transform: `translateX(${virtualColumn.start}px)`,
                  width: scaledPixels(virtualColumn.size),
                }}
              >
                {columnLabel(virtualColumn.index)}
              </div>
            ))}
          </div>
          <div className="ldv-spreadsheet-virtual-body">
            {virtualRows.map((virtualRow) => {
              const rowIndex = virtualRow.index;
              const row = rows[rowIndex] ?? [];

              return (
                <div
                  aria-rowindex={rowIndex + 1}
                  className="ldv-spreadsheet-virtual-row"
                  key={virtualRow.key}
                  role="row"
                  style={{
                    height: scaledPixels(virtualRow.size),
                    transform: `translateY(${virtualRow.start}px)`,
                    width: scaledPixels(virtualGridWidth),
                  }}
                >
                  <div
                    className="ldv-row-header ldv-spreadsheet-virtual-row-header"
                    role="rowheader"
                    style={{ ...rowHeaderStyle, height: scaledPixels(virtualRow.size) }}
                  >
                    {rowIndex + 1}
                  </div>
                  {virtualColumns.map((virtualColumn) => {
                    const cellIndex = virtualColumn.index;
                    const cell = row[cellIndex];
                    const comments = cell?.comments ?? [];
                    const commentLabel = cellCommentLabel(cell);

                    if (cell?.hiddenByMerge) {
                      return null;
                    }

                    return (
                      <div
                        aria-colindex={cellIndex + 1}
                        aria-label={commentLabel ? `${cellToText(cell)} ${commentLabel}` : undefined}
                        className={comments.length > 0 ? 'ldv-cell-has-comment ldv-spreadsheet-virtual-cell' : 'ldv-spreadsheet-virtual-cell'}
                        key={`${rowIndex}-${cellIndex}`}
                        role="gridcell"
                        style={{
                          ...cellStyle(cell, cellIndex, rowIndex, columnWidths, displayRowHeights, state.zoom),
                          transform: `translateX(${virtualColumn.start}px)`,
                        }}
                        title={commentLabel || undefined}
                      >
                        {highlightText(cellToText(cell), state.searchTerm)}
                        {comments.length > 0 ? (
                          <span className="ldv-cell-comment" tabIndex={0}>
                            <span className="ldv-cell-comment-marker" aria-hidden="true" />
                            <span className="ldv-cell-comment-popover" role="note">
                              {comments.map((comment, commentIndex) => (
                                <span className="ldv-cell-comment-item" key={`${comment.kind}-${commentIndex}`}>
                                  {comment.author ? <strong>{comment.author}</strong> : null}
                                  {comment.date ? <small>{new Date(comment.date).toLocaleString()}</small> : null}
                                  <span>{comment.text}</span>
                                </span>
                              ))}
                            </span>
                          </span>
                        ) : null}
                      </div>
                    );
                  })}
                </div>
              );
            })}
          </div>
        </div>
      </div>
      {showSheetTabs ? (
        <div className="ldv-sheet-tabs" role="tablist">
          {sheetNames.map((name) => (
            <button
              aria-selected={name === sheetName}
              className="ldv-sheet-tab"
              key={name}
              onClick={() => setSheetName(name)}
              role="tab"
              type="button"
            >
              {name}
            </button>
          ))}
        </div>
      ) : null}
    </div>
  );
}

export const SpreadsheetRenderer: DocumentRenderer = {
  id: 'spreadsheet',
  label: 'Spreadsheet',
  priority: 20,
  canRender: (file) =>
    file.extension === 'xls' ||
    file.extension === 'xlsx' ||
    file.extension === 'ods' ||
    file.mimeType === 'application/vnd.ms-excel' ||
    file.mimeType === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' ||
    file.mimeType === 'application/vnd.oasis.opendocument.spreadsheet',
  Component: SpreadsheetRendererComponent,
};
