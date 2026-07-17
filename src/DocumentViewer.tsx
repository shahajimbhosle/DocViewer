import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import type { CSSProperties } from 'react';
import { Toolbar } from './components/Toolbar';
import { builtInRenderers } from './renderers';
import type {
  DocumentInfo,
  DocumentRenderer,
  DocumentViewerControls,
  DocumentViewerLabels,
  DocumentViewerProps,
  FitMode,
  ResolvedDocumentViewerControls,
  ResolvedDocument,
  SearchStats,
} from './types';
import { resolveDocumentSource, toError } from './utils/source';
import './styles.css';

const defaultControls: ResolvedDocumentViewerControls = {
  toolbar: true,
  fileName: true,
  pageNavigation: true,
  zoom: true,
  fit: true,
  rotate: true,
  search: true,
  print: true,
  download: true,
  fullscreen: true,
  thumbnails: true,
};

const defaultLabels: DocumentViewerLabels = {
  openFile: 'Open file',
  previousPage: 'Previous page',
  nextPage: 'Next page',
  page: 'Page',
  of: 'of',
  zoomOut: 'Zoom out',
  zoomIn: 'Zoom in',
  resetZoom: 'Reset zoom',
  fitWidth: 'Fit width',
  fitPage: 'Fit page',
  rotateLeft: 'Rotate left',
  rotateRight: 'Rotate right',
  search: 'Search',
  print: 'Print',
  download: 'Download',
  fullscreen: 'Fullscreen',
  exitFullscreen: 'Exit fullscreen',
  unsupportedTitle: 'Preview unavailable',
  noDocument: 'No document selected',
  loading: 'Loading document...',
};

type ControlName = keyof ResolvedDocumentViewerControls;

const disabledControlsByRenderer: Record<string, ControlName[]> = {
  csv: ['fit', 'pageNavigation', 'rotate'],
  docx: ['fit', 'rotate'],
  image: ['pageNavigation', 'rotate', 'search'],
  'legacy-ppt': ['fit', 'pageNavigation', 'print', 'rotate', 'search', 'zoom'],
  markdown: ['fit', 'pageNavigation', 'rotate'],
  media: ['fit', 'pageNavigation', 'print', 'rotate', 'search', 'zoom'],
  pptx: ['rotate'],
  spreadsheet: ['fit', 'pageNavigation', 'rotate'],
  text: ['fit', 'pageNavigation', 'rotate'],
  unsupported: ['fit', 'pageNavigation', 'print', 'rotate', 'search', 'zoom'],
};

function clamp(value: number, min: number, max: number): number {
  return Math.max(min, Math.min(value, max));
}

function mergeControls(controls?: DocumentViewerControls): ResolvedDocumentViewerControls {
  return { ...defaultControls, ...controls };
}

function mergeLabels(labels?: Partial<DocumentViewerLabels>): DocumentViewerLabels {
  return { ...defaultLabels, ...labels };
}

function controlsForRenderer(
  controls: ResolvedDocumentViewerControls,
  renderer?: DocumentRenderer | null,
): ResolvedDocumentViewerControls {
  if (!renderer) {
    return controls;
  }

  const nextControls = { ...controls };
  const disabledControls = disabledControlsByRenderer[renderer.id] ?? [];

  disabledControls.forEach((controlName) => {
    nextControls[controlName] = false;
  });

  return nextControls;
}

function rendererScore(renderer: DocumentRenderer, file: ResolvedDocument): number | null {
  const result = renderer.canRender(file);

  if (result === false) {
    return null;
  }

  if (typeof result === 'number') {
    return result;
  }

  return result ? 1 : null;
}

function selectRenderer(file: ResolvedDocument, customRenderers: DocumentRenderer[] = []): DocumentRenderer {
  const renderers = [...customRenderers, ...builtInRenderers];

  const [bestMatch] = renderers
    .map((renderer, index) => ({
      renderer,
      index,
      matchScore: rendererScore(renderer, file),
      priority: renderer.priority ?? 0,
    }))
    .filter((candidate): candidate is { renderer: DocumentRenderer; index: number; matchScore: number; priority: number } => {
      return candidate.matchScore !== null;
    })
    .sort((left, right) => {
      if (right.matchScore !== left.matchScore) {
        return right.matchScore - left.matchScore;
      }

      if (right.priority !== left.priority) {
        return right.priority - left.priority;
      }

      return left.index - right.index;
    });

  return bestMatch?.renderer ?? builtInRenderers[builtInRenderers.length - 1];
}

function escapeHtml(value: string): string {
  return value.replace(/[&<>"']/g, (char) => {
    const entities: Record<string, string> = {
      '&': '&amp;',
      '<': '&lt;',
      '>': '&gt;',
      '"': '&quot;',
      "'": '&#039;',
    };
    return entities[char];
  });
}

function downloadFile(file: ResolvedDocument) {
  const anchor = document.createElement('a');
  anchor.href = file.objectUrl;
  anchor.download = file.fileName;
  anchor.rel = 'noopener';
  document.body.append(anchor);
  anchor.click();
  anchor.remove();
}

function printObjectUrl(file: ResolvedDocument) {
  const frame = document.createElement('iframe');
  frame.style.position = 'fixed';
  frame.style.right = '0';
  frame.style.bottom = '0';
  frame.style.width = '0';
  frame.style.height = '0';
  frame.style.border = '0';
  frame.src = file.objectUrl;

  frame.onload = () => {
    frame.contentWindow?.focus();
    frame.contentWindow?.print();
    window.setTimeout(() => frame.remove(), 1000);
  };

  document.body.append(frame);
}

function printRenderedElement(element: HTMLElement, title: string) {
  const printWindow = window.open('', '_blank', 'width=960,height=720');

  if (!printWindow) {
    window.print();
    return;
  }

  const styles = Array.from(document.querySelectorAll('style,link[rel="stylesheet"]'))
    .map((node) => node.outerHTML)
    .join('\n');

  printWindow.document.write(`<!doctype html>
<html>
  <head>
    <title>${escapeHtml(title)}</title>
    ${styles}
  </head>
  <body>
    <main class="ldv-print-root">${element.innerHTML}</main>
  </body>
</html>`);
  printWindow.document.close();
  printWindow.focus();
  window.setTimeout(() => printWindow.print(), 150);
}

export function DocumentViewer({
  source,
  fileName,
  mimeType,
  className,
  style,
  height = 640,
  renderers,
  allowRemoteUrls = false,
  fetchCredentials = 'same-origin',
  minZoom = 0.25,
  maxZoom = 4,
  initialZoom = 1,
  initialPage = 1,
  controls,
  labels,
  emptyState,
  onLoad,
  onError,
}: DocumentViewerProps) {
  const viewerRef = useRef<HTMLDivElement>(null);
  const viewportRef = useRef<HTMLDivElement>(null);
  const pageCountRef = useRef<number | undefined>();
  const [resolvedFile, setResolvedFile] = useState<ResolvedDocument | null>(null);
  const [error, setError] = useState<Error | null>(null);
  const [isLoading, setIsLoading] = useState(false);
  const [page, setPageState] = useState(Math.max(1, initialPage));
  const [pageCount, setPageCountState] = useState<number | undefined>();
  const [zoom, setZoom] = useState(clamp(initialZoom, minZoom, maxZoom));
  const [rotation, setRotation] = useState(0);
  const [fitMode, setFitMode] = useState<FitMode>('manual');
  const [searchTerm, setSearchTerm] = useState('');
  const [searchStats, setSearchStats] = useState<SearchStats>({});
  const [documentInfo, setDocumentInfoState] = useState<DocumentInfo>({});
  const [isFullscreen, setIsFullscreen] = useState(false);
  const mergedControls = useMemo(() => mergeControls(controls), [controls]);
  const mergedLabels = useMemo(() => mergeLabels(labels), [labels]);

  const setPage = useCallback((nextPage: number) => {
    const maxPage = pageCountRef.current ?? nextPage;
    const safePage = clamp(Number.isFinite(nextPage) ? Math.round(nextPage) : 1, 1, Math.max(1, maxPage));
    setPageState(safePage);
  }, []);

  const setPageCount = useCallback((nextPageCount?: number) => {
    pageCountRef.current = nextPageCount;
    setPageCountState(nextPageCount);
    setPageState((currentPage) => {
      if (!nextPageCount) {
        return currentPage;
      }

      return clamp(currentPage, 1, nextPageCount);
    });
  }, []);

  const setDocumentInfo = useCallback((info: Partial<DocumentInfo>) => {
    setDocumentInfoState((current) => ({ ...current, ...info }));
  }, []);

  const reportError = useCallback(
    (unknownError: unknown) => {
      const nextError = toError(unknownError);
      setError(nextError);
      onError?.(nextError);
    },
    [onError],
  );

  const rendererActions = useMemo(
    () => ({
      setPage,
      setPageCount,
      setDocumentInfo,
      setSearchStats,
      reportError,
    }),
    [reportError, setDocumentInfo, setPage, setPageCount],
  );

  useEffect(() => {
    return () => {
      if (resolvedFile) {
        URL.revokeObjectURL(resolvedFile.objectUrl);
      }
    };
  }, [resolvedFile]);

  useEffect(() => {
    const controller = new AbortController();
    let cancelled = false;

    setError(null);
    setResolvedFile(null);
    setPageCount(undefined);
    setSearchStats({});
    setDocumentInfoState({});
    setPage(Math.max(1, initialPage));
    setZoom(clamp(initialZoom, minZoom, maxZoom));
    setRotation(0);
    setFitMode('manual');
    setSearchTerm('');

    if (!source) {
      setIsLoading(false);
      return () => controller.abort();
    }

    setIsLoading(true);

    resolveDocumentSource(source, {
      fileName,
      mimeType,
      allowRemoteUrls,
      fetchCredentials,
      signal: controller.signal,
    })
      .then((file) => {
        if (cancelled) {
          URL.revokeObjectURL(file.objectUrl);
          return;
        }

        setResolvedFile(file);
        setIsLoading(false);
      })
      .catch((unknownError) => {
        if (cancelled || controller.signal.aborted) {
          return;
        }

        setIsLoading(false);
        reportError(unknownError);
      });

    return () => {
      cancelled = true;
      controller.abort();
    };
  }, [
    allowRemoteUrls,
    fetchCredentials,
    fileName,
    initialPage,
    initialZoom,
    maxZoom,
    mimeType,
    minZoom,
    reportError,
    setPage,
    setPageCount,
    source,
  ]);

  const selectedRenderer = useMemo(() => {
    if (!resolvedFile) {
      return null;
    }

    return selectRenderer(resolvedFile, renderers);
  }, [renderers, resolvedFile]);

  const toolbarControls = useMemo(
    () => controlsForRenderer(mergedControls, selectedRenderer),
    [mergedControls, selectedRenderer],
  );

  useEffect(() => {
    if (!resolvedFile || !selectedRenderer) {
      return;
    }

    const baseInfo: DocumentInfo = {
      rendererId: selectedRenderer.id,
      rendererLabel: selectedRenderer.label,
      title: resolvedFile.fileName,
    };
    setDocumentInfoState(baseInfo);
    onLoad?.(baseInfo, resolvedFile);
  }, [onLoad, resolvedFile, selectedRenderer]);

  useEffect(() => {
    const handleFullscreenChange = () => {
      setIsFullscreen(document.fullscreenElement === viewerRef.current);
    };

    document.addEventListener('fullscreenchange', handleFullscreenChange);
    return () => document.removeEventListener('fullscreenchange', handleFullscreenChange);
  }, []);

  const runtimeState = useMemo(
    () => ({
      page,
      pageCount,
      zoom,
      rotation,
      fitMode,
      searchTerm,
    }),
    [fitMode, page, pageCount, rotation, searchTerm, zoom],
  );

  const rootStyle = useMemo<CSSProperties>(
    () => ({
      height: typeof height === 'number' ? `${height}px` : height,
      ...style,
    }),
    [height, style],
  );

  const handleZoomOut = useCallback(() => {
    setFitMode('manual');
    setZoom((current) => clamp(Number((current - 0.1).toFixed(2)), minZoom, maxZoom));
  }, [maxZoom, minZoom]);

  const handleZoomIn = useCallback(() => {
    setFitMode('manual');
    setZoom((current) => clamp(Number((current + 0.1).toFixed(2)), minZoom, maxZoom));
  }, [maxZoom, minZoom]);

  const handleZoomReset = useCallback(() => {
    setFitMode('manual');
    setZoom(clamp(initialZoom, minZoom, maxZoom));
  }, [initialZoom, maxZoom, minZoom]);

  const handleFullscreenToggle = useCallback(() => {
    if (!viewerRef.current) {
      return;
    }

    if (document.fullscreenElement) {
      document.exitFullscreen();
    } else {
      viewerRef.current.requestFullscreen();
    }
  }, []);

  const handlePrint = useCallback(() => {
    if (!resolvedFile) {
      return;
    }

    if (resolvedFile.mimeType === 'application/pdf' || resolvedFile.mimeType.startsWith('image/')) {
      printObjectUrl(resolvedFile);
      return;
    }

    if (viewportRef.current) {
      printRenderedElement(viewportRef.current, resolvedFile.fileName);
    }
  }, [resolvedFile]);

  const handleDownload = useCallback(() => {
    if (resolvedFile) {
      downloadFile(resolvedFile);
    }
  }, [resolvedFile]);

  const RendererComponent = selectedRenderer?.Component;
  const rootClassName = ['ldv-root', selectedRenderer ? `ldv-renderer-${selectedRenderer.id}` : null, className]
    .filter(Boolean)
    .join(' ');

  return (
    <section className={rootClassName} ref={viewerRef} style={rootStyle}>
      {mergedControls.toolbar ? (
        <Toolbar
          controls={toolbarControls}
          fileName={documentInfo.title || resolvedFile?.fileName}
          isFullscreen={isFullscreen}
          labels={mergedLabels}
          onDownload={handleDownload}
          onFitModeChange={setFitMode}
          onFullscreenToggle={handleFullscreenToggle}
          onPageChange={setPage}
          onPrint={handlePrint}
          onRotateLeft={() => setRotation((current) => (current - 90 + 360) % 360)}
          onRotateRight={() => setRotation((current) => (current + 90) % 360)}
          onSearchChange={setSearchTerm}
          onZoomIn={handleZoomIn}
          onZoomOut={handleZoomOut}
          onZoomReset={handleZoomReset}
          searchStats={searchStats}
          state={runtimeState}
        />
      ) : null}

      <div className="ldv-viewport" ref={viewportRef}>
        {!source ? (
          emptyState ?? <div className="ldv-empty-state">{mergedLabels.noDocument}</div>
        ) : isLoading ? (
          <div className="ldv-renderer-status">{mergedLabels.loading}</div>
        ) : error ? (
          <div className="ldv-error" role="alert">
            <strong>Unable to render document</strong>
            <span>{error.message}</span>
          </div>
        ) : resolvedFile && RendererComponent ? (
          <RendererComponent
            actions={rendererActions}
            controls={mergedControls}
            file={resolvedFile}
            labels={mergedLabels}
            state={runtimeState}
            viewportRef={viewportRef}
          />
        ) : null}
      </div>
    </section>
  );
}
