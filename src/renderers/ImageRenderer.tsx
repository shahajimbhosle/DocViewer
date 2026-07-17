import { useEffect, useMemo, useState } from 'react';
import type { RefObject } from 'react';
import type { DocumentRenderer, DocumentRendererProps } from '../types';

interface Size {
  width: number;
  height: number;
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

function ImageRendererComponent({ file, state, actions, viewportRef }: DocumentRendererProps) {
  const viewportSize = useElementSize(viewportRef);

  useEffect(() => {
    actions.setPageCount(undefined);
    actions.setSearchStats({});
    actions.setDocumentInfo({ title: file.fileName });
  }, [actions, file.fileName]);

  const fitClass = state.fitMode === 'page' ? 'ldv-image-fit-page' : state.fitMode === 'width' ? 'ldv-image-fit-width' : '';
  const imageStyle = useMemo(
    () => ({
      ...(state.fitMode === 'page'
        ? {
            maxHeight: `${Math.max(1, viewportSize.height - 8)}px`,
            maxWidth: `${Math.max(1, viewportSize.width - 8)}px`,
          }
        : null),
      transform: `scale(${state.fitMode === 'manual' ? state.zoom : 1}) rotate(${state.rotation}deg)`,
    }),
    [state.fitMode, state.rotation, state.zoom, viewportSize.height, viewportSize.width],
  );

  return (
    <div className={`ldv-image-stage ${fitClass}`}>
      <img
        alt={file.fileName}
        className="ldv-image-document"
        src={file.objectUrl}
        style={imageStyle}
      />
    </div>
  );
}

export const ImageRenderer: DocumentRenderer = {
  id: 'image',
  label: 'Image',
  priority: 40,
  canRender: (file) => file.mimeType.startsWith('image/') && file.extension !== 'svg',
  Component: ImageRendererComponent,
};
