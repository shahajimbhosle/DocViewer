import { useEffect, useMemo, useState } from 'react';
import type { DocumentRenderer, DocumentRendererProps } from '../types';
import { countMatches, highlightText } from '../utils/highlight';

function parseDelimited(text: string, delimiter: ',' | '\t'): string[][] {
  const rows: string[][] = [];
  let row: string[] = [];
  let value = '';
  let inQuotes = false;

  for (let index = 0; index < text.length; index += 1) {
    const char = text[index];
    const next = text[index + 1];

    if (char === '"' && inQuotes && next === '"') {
      value += '"';
      index += 1;
    } else if (char === '"') {
      inQuotes = !inQuotes;
    } else if (char === delimiter && !inQuotes) {
      row.push(value);
      value = '';
    } else if ((char === '\n' || char === '\r') && !inQuotes) {
      if (char === '\r' && next === '\n') {
        index += 1;
      }
      row.push(value);
      rows.push(row);
      row = [];
      value = '';
    } else {
      value += char;
    }
  }

  row.push(value);
  rows.push(row);
  return rows.filter((cells) => cells.some((cell) => cell.length > 0));
}

function CsvRendererComponent({ file, state, actions }: DocumentRendererProps) {
  const [rows, setRows] = useState<string[][]>([]);
  const delimiter = file.extension === 'tsv' ? '\t' : ',';

  useEffect(() => {
    const text = new TextDecoder('utf-8').decode(file.arrayBuffer);
    setRows(parseDelimited(text, delimiter));
    actions.setPageCount(undefined);
    actions.setDocumentInfo({ title: file.fileName });
  }, [actions, delimiter, file]);

  const plainText = useMemo(() => rows.map((row) => row.join(' ')).join('\n'), [rows]);
  const matches = useMemo(() => countMatches(plainText, state.searchTerm), [plainText, state.searchTerm]);
  const [header, ...body] = rows;

  useEffect(() => {
    actions.setSearchStats({
      totalMatches: matches,
      currentPageMatches: matches,
      message: state.searchTerm ? `${matches} match${matches === 1 ? '' : 'es'}` : undefined,
    });
  }, [actions, matches, state.searchTerm]);

  return (
    <div
      className="ldv-table-document ldv-csv-document"
      style={{
        fontSize: `${13 * state.zoom}px`,
        transform: `rotate(${state.rotation}deg)`,
        transformOrigin: 'center top',
      }}
    >
      <table>
        {header ? (
          <thead>
            <tr>
              {header.map((cell, index) => (
                <th key={`${cell}-${index}`}>{highlightText(cell, state.searchTerm)}</th>
              ))}
            </tr>
          </thead>
        ) : null}
        <tbody>
          {body.map((row, rowIndex) => (
            <tr key={rowIndex}>
              {row.map((cell, cellIndex) => (
                <td key={`${rowIndex}-${cellIndex}`}>{highlightText(cell, state.searchTerm)}</td>
              ))}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

export const CsvRenderer: DocumentRenderer = {
  id: 'csv',
  label: 'CSV',
  priority: 35,
  canRender: (file) => file.extension === 'csv' || file.extension === 'tsv' || file.mimeType === 'text/csv',
  Component: CsvRendererComponent,
};
