import { AlertTriangle } from 'lucide-react';
import { useEffect } from 'react';
import type { DocumentRenderer, DocumentRendererProps } from '../types';
import { isOfficeExtension } from '../utils/mime';

const cadExtensions = new Set(['dwg', 'dxf', 'dgn', 'stl', 'step', 'stp', 'iges', 'igs']);

function UnsupportedRendererComponent({ file, actions }: DocumentRendererProps) {
  const isOffice = isOfficeExtension(file.extension);
  const isCad = cadExtensions.has(file.extension);

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
      {isCad ? (
        <p>
          CAD drawing formats such as DWG and DXF are not rendered by browsers natively. Convert them locally to PDF/SVG/PNG or
          register a custom CAD renderer backed by your own private service.
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
