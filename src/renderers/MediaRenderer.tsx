import { useEffect } from 'react';
import type { DocumentRenderer, DocumentRendererProps } from '../types';

function MediaRendererComponent({ file, actions }: DocumentRendererProps) {
  useEffect(() => {
    actions.setPageCount(undefined);
    actions.setSearchStats({});
    actions.setDocumentInfo({ title: file.fileName });
  }, [actions, file.fileName]);

  if (file.mimeType.startsWith('audio/')) {
    return (
      <div className="ldv-media-stage">
        <audio className="ldv-audio-document" controls src={file.objectUrl} />
      </div>
    );
  }

  return (
    <div className="ldv-media-stage">
      <video className="ldv-video-document" controls src={file.objectUrl} />
    </div>
  );
}

export const MediaRenderer: DocumentRenderer = {
  id: 'media',
  label: 'Media',
  priority: 5,
  canRender: (file) => file.mimeType.startsWith('audio/') || file.mimeType.startsWith('video/'),
  Component: MediaRendererComponent,
};
