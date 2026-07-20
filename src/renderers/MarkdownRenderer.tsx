import { marked } from 'marked';
import { useEffect, useMemo, useState } from 'react';
import type { DocumentRenderer, DocumentRendererProps } from '../types';
import { countMatches } from '../utils/highlight';
import { htmlToText, sanitizeHtml } from '../utils/sanitize';

function MarkdownRendererComponent({ file, state, actions, markdownOptions }: DocumentRendererProps) {
  const [html, setHtml] = useState('');
  const [plainText, setPlainText] = useState('');

  useEffect(() => {
    let cancelled = false;
    actions.setLoading(true);

    async function renderMarkdown() {
      const text = new TextDecoder('utf-8').decode(file.arrayBuffer);
      const rendered = await marked.parse(text, {
        async: false,
        gfm: true,
        breaks: false,
      });
      const cleanHtml = sanitizeHtml(rendered, {
        allowRemoteImages: markdownOptions.allowRemoteImages,
      });

      if (!cancelled) {
        setHtml(cleanHtml);
        setPlainText(htmlToText(cleanHtml));
        actions.setPageCount(undefined);
        actions.setDocumentInfo({ title: file.fileName });
        actions.setLoading(false);
      }
    }

    renderMarkdown().catch((error: unknown) => {
      if (!cancelled) {
        actions.setLoading(false);
        actions.reportError(error);
      }
    });

    return () => {
      cancelled = true;
      actions.setLoading(false);
    };
  }, [actions, file]);

  const matches = useMemo(() => countMatches(plainText, state.searchTerm), [plainText, state.searchTerm]);

  useEffect(() => {
    actions.setSearchStats({
      totalMatches: matches,
      currentPageMatches: matches,
      message: state.searchTerm ? `${matches} match${matches === 1 ? '' : 'es'}` : undefined,
    });
  }, [actions, matches, state.searchTerm]);

  return (
    <article
      className="ldv-rich-document ldv-markdown-document"
      style={{
        fontSize: `${16 * state.zoom}px`,
        transform: `rotate(${state.rotation}deg)`,
        transformOrigin: 'center top',
      }}
      dangerouslySetInnerHTML={{ __html: html }}
    />
  );
}

export const MarkdownRenderer: DocumentRenderer = {
  id: 'markdown',
  label: 'Markdown',
  priority: 30,
  canRender: (file) => file.extension === 'md' || file.mimeType === 'text/markdown',
  Component: MarkdownRendererComponent,
};
