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
  type SpreadsheetChartDrawing,
  type ParsedXlsxWorkbook,
  type SpreadsheetDrawing,
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
  return cell?.formattedValue ?? valueToText(cell?.value ?? null);
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

function displayRowHeightAt(rowIndex: number, displayRowHeights: number[], zoom: number): number {
  return displayRowHeights[rowIndex] ?? gridRowHeight * zoom;
}

function displayRowSpanHeightAt(rowIndex: number, rowSpan: number, displayRowHeights: number[], zoom: number): number {
  let height = 0;

  for (let offset = 0; offset < rowSpan; offset += 1) {
    height += displayRowHeightAt(rowIndex + offset, displayRowHeights, zoom);
  }

  return height;
}

function columnOffsetAt(columnIndex: number, columnWidths: number[], zoom: number): number {
  let offset = 0;

  for (let index = 0; index < columnIndex; index += 1) {
    offset += columnWidthAt(index, columnWidths, zoom);
  }

  return offset;
}

function displayRowOffsetAt(rowIndex: number, displayRowHeights: number[], zoom: number): number {
  let offset = 0;

  for (let index = 0; index < rowIndex; index += 1) {
    offset += displayRowHeightAt(index, displayRowHeights, zoom);
  }

  return offset;
}

function drawingStyle(
  drawing: SpreadsheetDrawing,
  columnWidths: number[],
  displayRowHeights: number[],
  zoom: number,
  rowHeaderPixelWidth: number,
  columnHeaderPixelHeight: number,
): CSSProperties {
  const anchor = drawing.anchor;
  const left = rowHeaderPixelWidth + columnOffsetAt(anchor.fromColumn, columnWidths, zoom) + (anchor.columnOffsetPx ?? 0) * zoom;
  const top = columnHeaderPixelHeight + displayRowOffsetAt(anchor.fromRow, displayRowHeights, zoom) + (anchor.rowOffsetPx ?? 0) * zoom;
  const width =
    typeof anchor.toColumn === 'number'
      ? rowHeaderPixelWidth +
          columnOffsetAt(anchor.toColumn, columnWidths, zoom) +
          (anchor.toColumnOffsetPx ?? 0) * zoom -
          left
      : (anchor.widthPx ?? 220) * zoom;
  const height =
    typeof anchor.toRow === 'number'
      ? columnHeaderPixelHeight +
          displayRowOffsetAt(anchor.toRow, displayRowHeights, zoom) +
          (anchor.toRowOffsetPx ?? 0) * zoom -
          top
      : (anchor.heightPx ?? 140) * zoom;

  return {
    height: scaledPixels(height),
    left: scaledPixels(left),
    top: scaledPixels(top),
    width: scaledPixels(width),
  };
}

function drawingColumnCount(drawings: SpreadsheetDrawing[]): number {
  return drawings.reduce((count, drawing) => {
    const anchor = drawing.anchor;
    return Math.max(count, (anchor.toColumn ?? anchor.fromColumn + 1) + 1);
  }, 0);
}

function drawingRowCount(drawings: SpreadsheetDrawing[]): number {
  return drawings.reduce((count, drawing) => {
    const anchor = drawing.anchor;
    return Math.max(count, (anchor.toRow ?? anchor.fromRow + 1) + 1);
  }, 0);
}

function compactNumber(value: number): string {
  const absoluteValue = Math.abs(value);

  if (absoluteValue >= 1_000_000) {
    return `${(value / 1_000_000).toFixed(1)}M`;
  }

  if (absoluteValue >= 1_000) {
    return `${(value / 1_000).toFixed(0)}K`;
  }

  return Number.isInteger(value) ? String(value) : value.toFixed(1);
}

function stripSpreadsheetFormat(formatCode: string): string {
  return formatCode
    .replace(/\[[^\]]+\]/g, '')
    .replace(/"([^"]*)"/g, '$1')
    .replace(/\\(.)/g, '$1')
    .replace(/_./g, '')
    .replace(/\*./g, '');
}

function decimalPlacesFromSpreadsheetFormat(formatCode: string): { maximum: number; minimum: number } {
  const decimalPart = stripSpreadsheetFormat(formatCode).split('.')[1]?.match(/[0#?]+/)?.[0] ?? '';

  return {
    maximum: decimalPart.length,
    minimum: Array.from(decimalPart).filter((character) => character === '0').length,
  };
}

function currencySymbolFromSpreadsheetFormat(formatCode: string): string {
  const bracketCurrency = /\[\$([^\]-]+)(?:-[^\]]+)?\]/.exec(formatCode)?.[1];

  if (bracketCurrency) {
    return bracketCurrency;
  }

  return /[₹$€£¥]/.exec(stripSpreadsheetFormat(formatCode))?.[0] ?? '';
}

function formatChartNumber(value: number, formatCode: string | undefined): string {
  if (!formatCode || /^general$/i.test(formatCode.trim())) {
    return compactNumber(value);
  }

  const cleanFormat = stripSpreadsheetFormat(formatCode);
  const percent = cleanFormat.includes('%');
  const currency = currencySymbolFromSpreadsheetFormat(formatCode);
  const placeholderIndex = cleanFormat.search(/[0#?]/);
  const currencyIndex = currency ? cleanFormat.indexOf(currency) : -1;
  const prefix = currency && (currencyIndex < placeholderIndex || placeholderIndex === -1) ? currency : '';
  const suffix = `${percent ? '%' : ''}${currency && currencyIndex > placeholderIndex ? currency : ''}`;
  const decimalPlaces = decimalPlacesFromSpreadsheetFormat(formatCode);
  const adjustedValue = percent ? value * 100 : value;
  const sign = adjustedValue < 0 ? '-' : '';
  const formattedNumber = new Intl.NumberFormat('en-US', {
    maximumFractionDigits: decimalPlaces.maximum,
    minimumFractionDigits: decimalPlaces.minimum,
    useGrouping: /[0#?],[0#?]/.test(cleanFormat) || Math.abs(adjustedValue) >= 1000,
  }).format(Math.abs(adjustedValue));

  return `${sign}${prefix}${formattedNumber}${suffix}`;
}

function niceChartStep(value: number): number {
  if (!Number.isFinite(value) || value <= 0) {
    return 1;
  }

  const magnitude = 10 ** Math.floor(Math.log10(value));
  const normalized = value / magnitude;
  const niceNormalized = normalized <= 1 ? 1 : normalized <= 2 ? 2 : normalized <= 2.5 ? 2.5 : normalized <= 5 ? 5 : 10;

  return niceNormalized * magnitude;
}

function chartAxisScale(values: number[]): { max: number; min: number; ticks: number[] } {
  const finiteValues = values.filter((value) => Number.isFinite(value));

  if (finiteValues.length === 0) {
    return { max: 1, min: 0, ticks: [0, 0.5, 1] };
  }

  const rawMin = Math.min(0, ...finiteValues);
  const rawMax = Math.max(0, ...finiteValues);
  const step = niceChartStep((rawMax - rawMin || Math.abs(rawMax) || 1) / 4);
  const min = Math.floor(rawMin / step) * step;
  const max = Math.ceil(rawMax / step) * step || step;
  const tickCount = Math.max(1, Math.round((max - min) / step));
  const ticks = Array.from({ length: tickCount + 1 }, (_, index) => min + index * step);

  return { max, min, ticks };
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
  displayRowHeights: number[],
  zoom: number,
): CSSProperties {
  const colSpan = cell?.colSpan ?? 1;
  const rowSpan = cell?.rowSpan ?? 1;
  const customStyle = cell?.style ?? {};
  const whiteSpace = effectiveCellWhiteSpace(cell);
  const width = scaledPixels(columnSpanWidthAt(columnIndex, colSpan, columnWidths, zoom));
  const height = scaledPixels(displayRowSpanHeightAt(rowIndex, rowSpan, displayRowHeights, zoom));
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

function SpreadsheetChartPreview({ chart }: { chart: SpreadsheetChartDrawing }) {
  const series = chart.series.find((candidate) => candidate.values.length > 0);
  const values = series?.values ?? [];
  const categories = chart.categories.length > 0 ? chart.categories : values.map((_, index) => String(index + 1));
  const title = chart.title ?? series?.name ?? 'Chart';
  const isHorizontalBar = chart.chartType === 'bar';
  const isLine = chart.chartType === 'line' || chart.chartType === 'area';
  const isPie = chart.chartType === 'pie' || chart.chartType === 'doughnut';
  const axisScale = chartAxisScale(values);
  const axisRange = Math.max(1, axisScale.max - axisScale.min);
  const svgWidth = 340;
  const plotLeft = 68;
  const plotRight = 314;
  const plotTop = 24;
  const plotBottom = 142;
  const plotWidth = plotRight - plotLeft;
  const plotHeight = plotBottom - plotTop;
  const valueToY = (value: number) => plotBottom - ((value - axisScale.min) / axisRange) * plotHeight;
  const baselineY = valueToY(0);

  if (values.length === 0) {
    return (
      <div className="ldv-spreadsheet-chart-empty">
        <strong>{title}</strong>
        <span>Chart data unavailable</span>
      </div>
    );
  }

  if (isPie) {
    const total = values.reduce((sum, value) => sum + Math.max(0, value), 0) || 1;
    let accumulated = 0;

    return (
      <div className="ldv-spreadsheet-chart">
        <strong>{title}</strong>
        <svg aria-label={title} viewBox="0 0 300 180" role="img">
          {values.map((value, index) => {
            const start = accumulated / total;
            accumulated += Math.max(0, value);
            const end = accumulated / total;
            const largeArc = end - start > 0.5 ? 1 : 0;
            const startAngle = start * Math.PI * 2 - Math.PI / 2;
            const endAngle = end * Math.PI * 2 - Math.PI / 2;
            const x1 = 96 + Math.cos(startAngle) * 58;
            const y1 = 92 + Math.sin(startAngle) * 58;
            const x2 = 96 + Math.cos(endAngle) * 58;
            const y2 = 92 + Math.sin(endAngle) * 58;
            const color = `hsl(${(index * 58 + 190) % 360} 62% 40%)`;

            return (
              <path
                d={`M 96 92 L ${x1} ${y1} A 58 58 0 ${largeArc} 1 ${x2} ${y2} Z`}
                fill={color}
                key={`${categories[index]}-${index}`}
              />
            );
          })}
          {chart.chartType === 'doughnut' ? <circle cx="96" cy="92" fill="#ffffff" r="28" /> : null}
          {values.slice(0, 4).map((value, index) => (
            <g key={`legend-${categories[index]}-${index}`}>
              <rect fill={`hsl(${(index * 58 + 190) % 360} 62% 40%)`} height="8" width="8" x="178" y={52 + index * 22} />
              <text x="192" y={60 + index * 22}>
                {categories[index]} {formatChartNumber(value, chart.valueFormatCode)}
              </text>
            </g>
          ))}
        </svg>
      </div>
    );
  }

  const pointStep = values.length > 1 ? plotWidth / (values.length - 1) : plotWidth;
  const linePoints = values
    .map((value, index) => {
      const x = plotLeft + index * pointStep;
      const y = valueToY(value);
      return `${x},${y}`;
    })
    .join(' ');

  return (
    <div className="ldv-spreadsheet-chart">
      <strong>{title}</strong>
      <svg aria-label={title} viewBox={`0 0 ${svgWidth} 190`} role="img">
        <line className="ldv-spreadsheet-chart-axis" x1={plotLeft} x2={plotRight} y1={baselineY} y2={baselineY} />
        <line className="ldv-spreadsheet-chart-axis" x1={plotLeft} x2={plotLeft} y1={plotTop} y2={plotBottom} />
        {axisScale.ticks.map((tick) => {
          const y = valueToY(tick);

          return (
            <g key={tick}>
              <line className="ldv-spreadsheet-chart-grid" x1={plotLeft} x2={plotRight} y1={y} y2={y} />
              <text className="ldv-spreadsheet-chart-tick" x="4" y={y + 4}>
                {formatChartNumber(tick, chart.valueFormatCode)}
              </text>
            </g>
          );
        })}
        {isLine ? (
          <>
            {chart.chartType === 'area' ? (
              <polygon
                className="ldv-spreadsheet-chart-area"
                points={`${plotLeft},${baselineY} ${linePoints} ${plotLeft + pointStep * (values.length - 1)},${baselineY}`}
              />
            ) : null}
            <polyline className="ldv-spreadsheet-chart-line" points={linePoints} />
            {values.map((value, index) => (
              <circle cx={plotLeft + index * pointStep} cy={valueToY(value)} key={`${categories[index]}-${index}`} r="3" />
            ))}
          </>
        ) : (
          values.map((value, index) => {
            const bandSize = (isHorizontalBar ? plotHeight : plotWidth) / Math.max(1, values.length);
            const barSize = Math.max(4, bandSize * 0.62);
            const y = valueToY(value);
            const ratio = Math.abs(value) / Math.max(1, Math.max(Math.abs(axisScale.min), Math.abs(axisScale.max)));

            return isHorizontalBar ? (
              <g key={`${categories[index]}-${index}`}>
                <rect
                  className="ldv-spreadsheet-chart-bar"
                  height={barSize}
                  width={ratio * plotWidth}
                  x={plotLeft}
                  y={plotTop + index * bandSize + (bandSize - barSize) / 2}
                />
                <text className="ldv-spreadsheet-chart-label" x="4" y={plotTop + index * bandSize + bandSize / 2 + 3}>
                  {categories[index]}
                </text>
              </g>
            ) : (
              <g key={`${categories[index]}-${index}`}>
                <rect
                  className="ldv-spreadsheet-chart-bar"
                  height={Math.max(1, Math.abs(baselineY - y))}
                  width={barSize}
                  x={plotLeft + index * bandSize + (bandSize - barSize) / 2}
                  y={Math.min(y, baselineY)}
                />
                <text
                  className="ldv-spreadsheet-chart-label"
                  textAnchor="middle"
                  x={plotLeft + index * bandSize + bandSize / 2}
                  y="164"
                >
                  {categories[index]}
                </text>
              </g>
            );
          })
        )}
      </svg>
    </div>
  );
}

function SpreadsheetRendererComponent({ file, state, actions, viewportRef }: DocumentRendererProps) {
  const [sheetNames, setSheetNames] = useState<string[]>([]);
  const [sheetName, setSheetName] = useState('');
  const [rows, setRows] = useState<SpreadsheetRow[]>([]);
  const [drawings, setDrawings] = useState<SpreadsheetDrawing[]>([]);
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
    setDrawings(sheet?.drawings ?? []);
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
    drawings.forEach((drawing) => {
      if (drawing.kind === 'chart') {
        parts.push(drawing.title ?? '');
        parts.push(...drawing.categories);
        drawing.series.forEach((series) => {
          parts.push(series.name ?? '');
          parts.push(...series.values.map(String));
        });
      } else {
        parts.push(drawing.name ?? '');
      }
    });

    return parts.join('\n');
  }, [activeSearchTerm, drawings, rows]);
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
        : gridRowHeight * state.zoom;

    return averageHeight;
  }, [displayRowHeights, state.zoom]);
  const visibleColumnCount = useMemo(() => {
    const viewportColumns = Math.ceil(
      Math.max(0, viewportSize.width - gridRowHeaderWidth * state.zoom) / Math.max(1, averageColumnWidth),
    );
    return Math.max(1, populatedColumnCount, drawingColumnCount(drawings), viewportColumns);
  }, [averageColumnWidth, drawings, populatedColumnCount, state.zoom, viewportSize.width]);
  const visibleRowCount = useMemo(() => {
    const tabsHeight = showSheetTabs ? sheetTabsHeight : 0;
    const viewportRows = Math.ceil(
      Math.max(0, viewportSize.height - tabsHeight - gridColumnHeaderHeight * state.zoom) / Math.max(1, averageRowHeight),
    );
    return Math.max(1, rows.length, drawingRowCount(drawings), viewportRows);
  }, [averageRowHeight, drawings, rows.length, showSheetTabs, state.zoom, viewportSize.height]);
  const rowHeaderPixelWidth = Math.max(1, Math.round(gridRowHeaderWidth * state.zoom));
  const columnHeaderPixelHeight = Math.max(1, Math.round(gridColumnHeaderHeight * state.zoom));
  const rowHeaderStyle: CSSProperties = {
    height: scaledPixels(gridRowHeight * state.zoom),
    width: scaledPixels(rowHeaderPixelWidth),
  };
  const rowVirtualizer = useVirtualizer({
    count: visibleRowCount,
    estimateSize: (index) => displayRowHeightAt(index, displayRowHeights, state.zoom),
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

  useEffect(() => {
    rowVirtualizer.measure();
    columnVirtualizer.measure();
  }, [columnVirtualizer, columnWidths, displayRowHeights, rowVirtualizer, state.zoom]);

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
                    const cellText = cellToText(cell);
                    const comments = cell?.comments ?? [];
                    const commentLabel = cellCommentLabel(cell);
                    const cellClassName = [
                      'ldv-spreadsheet-virtual-cell',
                      comments.length > 0 ? 'ldv-cell-has-comment' : null,
                    ]
                      .filter(Boolean)
                      .join(' ');

                    if (cell?.hiddenByMerge) {
                      return null;
                    }

                    return (
                      <div
                        aria-colindex={cellIndex + 1}
                        aria-label={commentLabel ? `${cellText} ${commentLabel}` : undefined}
                        className={cellClassName}
                        key={`${rowIndex}-${cellIndex}`}
                        role="gridcell"
                        style={{
                          ...cellStyle(cell, cellIndex, rowIndex, columnWidths, displayRowHeights, state.zoom),
                          transform: `translateX(${virtualColumn.start}px)`,
                        }}
                        title={commentLabel || undefined}
                      >
                        {highlightText(cellText, state.searchTerm)}
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
          {drawings.length > 0 ? (
            <div className="ldv-spreadsheet-drawing-layer" aria-label="Spreadsheet drawings">
              {drawings.map((drawing, drawingIndex) => (
                <div
                  className={`ldv-spreadsheet-drawing ldv-spreadsheet-drawing-${drawing.kind}`}
                  key={`${drawing.kind}-${drawingIndex}`}
                  style={drawingStyle(
                    drawing,
                    columnWidths,
                    displayRowHeights,
                    state.zoom,
                    rowHeaderPixelWidth,
                    columnHeaderPixelHeight,
                  )}
                >
                  {drawing.kind === 'image' ? (
                    <img alt={drawing.name ?? 'Spreadsheet image'} src={drawing.dataUrl} />
                  ) : (
                    <SpreadsheetChartPreview chart={drawing} />
                  )}
                </div>
              ))}
            </div>
          ) : null}
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
