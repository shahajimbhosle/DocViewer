import { useEffect, useMemo, useState } from 'react';
import type { DocumentRenderer, DocumentRendererProps } from '../types';
import { countMatches, highlightText } from '../utils/highlight';
import { isTextLikeMime } from '../utils/mime';

const textExtensions = new Set([
  'css',
  'html',
  'ini',
  'js',
  'jsx',
  'json',
  'log',
  'ts',
  'tsx',
  'txt',
  'xml',
  'yaml',
  'yml',
]);

function formatText(rawText: string, extension: string, mimeType: string): string {
  if (extension === 'json' || mimeType === 'application/json') {
    try {
      return JSON.stringify(JSON.parse(rawText), null, 2);
    } catch {
      return rawText;
    }
  }

  return rawText;
}

function TextRendererComponent({ file, state, actions }: DocumentRendererProps) {
  const [text, setText] = useState('');

  useEffect(() => {
    const decoder = new TextDecoder('utf-8');
    const decoded = decoder.decode(file.arrayBuffer);
    setText(formatText(decoded, file.extension, file.mimeType));
    actions.setPageCount(undefined);
    actions.setDocumentInfo({ title: file.fileName });
  }, [actions, file]);

  const matches = useMemo(() => countMatches(text, state.searchTerm), [state.searchTerm, text]);

  useEffect(() => {
    actions.setSearchStats({
      totalMatches: matches,
      currentPageMatches: matches,
      message: state.searchTerm ? `${matches} match${matches === 1 ? '' : 'es'}` : undefined,
    });
  }, [actions, matches, state.searchTerm]);

  return (
    <div
      className="ldv-text-shell"
      style={{
        transform: `rotate(${state.rotation}deg)`,
        transformOrigin: 'center top',
      }}
    >
      <pre
        className="ldv-text-document"
        style={{
          fontSize: `${14 * state.zoom}px`,
          lineHeight: 1.55,
        }}
      >
        {highlightText(text, state.searchTerm)}
      </pre>
    </div>
  );
}

export const TextRenderer: DocumentRenderer = {
  id: 'text',
  label: 'Text',
  priority: 10,
  canRender: (file) => textExtensions.has(file.extension) || isTextLikeMime(file.mimeType),
  Component: TextRendererComponent,
};
