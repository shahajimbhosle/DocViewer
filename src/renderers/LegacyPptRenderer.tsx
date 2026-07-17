import { AlertTriangle } from 'lucide-react';
import { useEffect } from 'react';
import type { DocumentRenderer, DocumentRendererProps } from '../types';

function LegacyPptRendererComponent({ file, actions, state }: DocumentRendererProps) {
  useEffect(() => {
    actions.setPageCount(undefined);
    actions.setDocumentInfo({
      title: file.fileName,
      warnings: ['Legacy binary PowerPoint files require a local/private converter before browser preview.'],
    });
  }, [actions, file.fileName]);

  useEffect(() => {
    actions.setSearchStats({
      totalMatches: 0,
      currentPageMatches: 0,
      message: state.searchTerm ? 'Search unavailable' : undefined,
    });
  }, [actions, state.searchTerm]);

  return (
    <div className="ldv-unsupported" role="status">
      <AlertTriangle aria-hidden="true" size={28} />
      <h2>Legacy PowerPoint preview unavailable</h2>
      <p>
        <strong>.ppt</strong> is the older binary PowerPoint format. The built-in browser renderer supports the newer{' '}
        <strong>.pptx</strong> Open XML format.
      </p>
      <p>
        To keep documents confidential, convert <strong>.ppt</strong> to <strong>.pptx</strong> or PDF with your own local/on-prem
        converter, then pass that converted file to this viewer or register it as a custom renderer.
      </p>
    </div>
  );
}

export const LegacyPptRenderer: DocumentRenderer = {
  id: 'legacy-ppt',
  label: 'Legacy PowerPoint',
  priority: 23,
  canRender: (file) => {
    if (file.extension === 'ppt') {
      return 2;
    }

    return file.extension !== 'pptx' && file.mimeType === 'application/vnd.ms-powerpoint';
  },
  Component: LegacyPptRendererComponent,
};
