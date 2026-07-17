import { AlertTriangle } from 'lucide-react';
import { useEffect } from 'react';
import type { DocumentRenderer, DocumentRendererProps } from '../types';
import { isOfficeExtension } from '../utils/mime';

function UnsupportedRendererComponent({ file, actions }: DocumentRendererProps) {
  const isOffice = isOfficeExtension(file.extension);

  useEffect(() => {
    actions.setPageCount(undefined);
    actions.setSearchStats({});
    actions.setDocumentInfo({
      title: file.fileName,
      warnings: ['No built-in renderer matched this document type.'],
    });
  }, [actions, file.fileName]);

  return (
    <div className="ldv-unsupported" role="status">
      <AlertTriangle aria-hidden="true" size={28} />
      <h2>Preview unavailable</h2>
      <p>
        This file is loaded locally, but no built-in browser renderer is available for <strong>.{file.extension || 'file'}</strong>.
      </p>
      {isOffice ? (
        <p>
          For strict fidelity across legacy Office, PowerPoint, OpenDocument, and encrypted files, register a custom renderer backed by
          your own local or private conversion service.
        </p>
      ) : null}
    </div>
  );
}

export const UnsupportedRenderer: DocumentRenderer = {
  id: 'unsupported',
  label: 'Unsupported',
  priority: -1000,
  canRender: () => true,
  Component: UnsupportedRendererComponent,
};
