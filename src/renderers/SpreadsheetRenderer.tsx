import { useEffect, useMemo, useState } from 'react';
import type { RefObject } from 'react';
import type { DocumentRenderer, DocumentRendererProps } from '../types';
import { countMatches, highlightText } from '../utils/highlight';
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
  const [workbook, setWorkbook] = useState<SpreadsheetWorkbook | null>(null);
  const viewportSize = useElementSize(viewportRef);
  const isLegacyXls = file.extension === 'xls' || file.mimeType === 'application/vnd.ms-excel';

  useEffect(() => {
    let cancelled = false;

    async function loadSheets() {
      const nextWorkbook: SpreadsheetWorkbook = isLegacyXls
        ? {
            sheets: parseXlsWorkbook(file.arrayBuffer).sheets.map((sheet) => ({
              name: sheet.name,
              rows: xlsRowsToSpreadsheetRows(sheet.rows),
            })),
          }
        : await parseXlsxWorkbook(file.arrayBuffer);

      if (cancelled) {
        return;
      }

      const nextSheetNames = nextWorkbook.sheets.map((sheet) => sheet.name);

      setWorkbook(nextWorkbook);
      setSheetNames(nextSheetNames);
      setSheetName(nextSheetNames[0] ?? '');
      actions.setPageCount(undefined);
      actions.setDocumentInfo({ title: file.fileName });
    }

    loadSheets().catch(actions.reportError);

    return () => {
      cancelled = true;
    };
  }, [actions, file, isLegacyXls]);

  useEffect(() => {
    const sheet = workbook?.sheets.find((candidate) => candidate.name === sheetName);
    setRows(sheet?.rows ?? []);
  }, [sheetName, workbook]);

  const plainText = useMemo(() => rows.map((row) => row.map(cellToText).join(' ')).join('\n'), [rows]);
  const matches = useMemo(() => countMatches(plainText, state.searchTerm), [plainText, state.searchTerm]);
  const populatedColumnCount = useMemo(() => rows.reduce((max, row) => Math.max(max, row.length), 0), [rows]);
  const visibleColumnCount = useMemo(() => {
    const viewportColumns = Math.ceil(Math.max(0, viewportSize.width - gridRowHeaderWidth) / gridCellWidth);
    return Math.max(1, populatedColumnCount, viewportColumns);
  }, [populatedColumnCount, viewportSize.width]);
  const visibleRowCount = useMemo(() => {
    const tabsHeight = sheetNames.length > 1 ? sheetTabsHeight : 0;
    const viewportRows = Math.ceil(Math.max(0, viewportSize.height - tabsHeight - gridColumnHeaderHeight) / gridRowHeight);
    return Math.max(1, rows.length, viewportRows);
  }, [rows.length, sheetNames.length, viewportSize.height]);
  const columnIndexes = useMemo(
    () => Array.from({ length: visibleColumnCount }, (_, index) => index),
    [visibleColumnCount],
  );
  const rowIndexes = useMemo(() => Array.from({ length: visibleRowCount }, (_, index) => index), [visibleRowCount]);

  useEffect(() => {
    actions.setSearchStats({
      totalMatches: matches,
      currentPageMatches: matches,
      message: state.searchTerm ? `${matches} match${matches === 1 ? '' : 'es'}` : undefined,
    });
  }, [actions, matches, state.searchTerm]);

  return (
    <div
      className="ldv-spreadsheet-document"
      style={{
        fontSize: `${13 * state.zoom}px`,
        transform: `rotate(${state.rotation}deg)`,
        transformOrigin: 'center top',
      }}
    >
      {sheetNames.length > 1 ? (
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
      <div className="ldv-table-document ldv-grid-document">
        <table className="ldv-spreadsheet-grid">
          <thead>
            <tr>
              <th aria-label="Row and column headers" className="ldv-sheet-corner" scope="col" />
              {columnIndexes.map((columnIndex) => (
                <th className="ldv-column-header" key={columnIndex} scope="col">
                  {columnLabel(columnIndex)}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {rowIndexes.map((rowIndex) => {
              const row = rows[rowIndex] ?? [];

              return (
                <tr key={rowIndex}>
                  <th className="ldv-row-header" scope="row">
                    {rowIndex + 1}
                  </th>
                  {columnIndexes.map((cellIndex) => {
                    const cell = row[cellIndex];
                    const comments = cell?.comments ?? [];
                    const commentLabel = cellCommentLabel(cell);

                    return (
                      <td
                        aria-label={commentLabel ? `${cellToText(cell)} ${commentLabel}` : undefined}
                        className={comments.length > 0 ? 'ldv-cell-has-comment' : undefined}
                        key={`${rowIndex}-${cellIndex}`}
                        style={cell?.style}
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
                      </td>
                    );
                  })}
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
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
    file.mimeType === 'application/vnd.ms-excel' ||
    file.mimeType === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  Component: SpreadsheetRendererComponent,
};
