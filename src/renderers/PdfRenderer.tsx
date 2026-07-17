import * as pdfjs from 'pdfjs-dist';
import { useEffect, useMemo, useRef, useState } from 'react';
import type { RefObject } from 'react';
import type { DocumentRenderer, DocumentRendererProps, FitMode } from '../types';
import { countMatches } from '../utils/highlight';

const pdfWorkerFileName = 'pdf.worker.min.mjs';

function defaultPdfWorkerSource(): string | undefined {
  if (typeof window === 'undefined') {
    return undefined;
  }

  if (import.meta.env.DEV) {
    return `${window.location.origin}/node_modules/pdfjs-dist/build/pdf.worker.min.mjs`;
  }

  try {
    const moduleUrl = import.meta.url;
    const lastSlashIndex = moduleUrl.lastIndexOf('/');

    if (moduleUrl && lastSlashIndex >= 0) {
      return `${moduleUrl.slice(0, lastSlashIndex + 1)}${pdfWorkerFileName}`;
    }
  } catch {
    return pdfWorkerFileName;
  }

  return pdfWorkerFileName;
}

const initialPdfWorkerSource = defaultPdfWorkerSource();

if (initialPdfWorkerSource) {
  pdfjs.GlobalWorkerOptions.workerSrc = initialPdfWorkerSource;
}

export function configurePdfWorker(workerSrc: string): void {
  pdfjs.GlobalWorkerOptions.workerSrc = workerSrc;
}

type PdfLoadingTask = ReturnType<typeof pdfjs.getDocument>;
type PdfDocument = Awaited<PdfLoadingTask['promise']>;
type PdfPage = Awaited<ReturnType<PdfDocument['getPage']>>;

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

function clampScale(value: number): number {
  return Math.max(0.1, Math.min(value, 6));
}

function scaleForFitMode(baseSize: Size | null, containerSize: Size, fitMode: FitMode, zoom: number): number {
  if (!baseSize || fitMode === 'manual') {
    return zoom;
  }

  const availableWidth = Math.max(containerSize.width - 48, 160);
  const availableHeight = Math.max(containerSize.height - 48, 160);

  if (fitMode === 'width') {
    return clampScale(availableWidth / baseSize.width);
  }

  return clampScale(Math.min(availableWidth / baseSize.width, availableHeight / baseSize.height));
}

function textItemToString(item: unknown): string {
  if (item && typeof item === 'object' && 'str' in item) {
    return String((item as { str?: string }).str ?? '');
  }

  return '';
}

function isPdfCancellationError(error: unknown): boolean {
  const name = error && typeof error === 'object' && 'name' in error ? String((error as { name?: unknown }).name ?? '') : '';
  const message =
    error && typeof error === 'object' && 'message' in error
      ? String((error as { message?: unknown }).message ?? '')
      : String(error ?? '');

  return (
    name === 'RenderingCancelledException' ||
    name === 'AbortException' ||
    message.includes('Worker was destroyed') ||
    message.includes('Loading task destroyed') ||
    message.includes('worker has been destroyed')
  );
}

interface PdfPageCanvasProps {
  pdf: PdfDocument;
  pageNumber: number;
  zoom: number;
  rotation: number;
  fitMode: FitMode;
  viewportRef: RefObject<HTMLDivElement>;
  onError: (error: unknown) => void;
}

function PdfPageCanvas({ pdf, pageNumber, zoom, rotation, fitMode, viewportRef, onError }: PdfPageCanvasProps) {
  const canvasRef = useRef<HTMLCanvasElement>(null);
  const [page, setPage] = useState<PdfPage | null>(null);
  const [baseSize, setBaseSize] = useState<Size | null>(null);
  const containerSize = useElementSize(viewportRef);
  const effectiveScale = useMemo(
    () => scaleForFitMode(baseSize, containerSize, fitMode, zoom),
    [baseSize, containerSize, fitMode, zoom],
  );

  useEffect(() => {
    let cancelled = false;
    let loadedPage: PdfPage | null = null;

    pdf
      .getPage(pageNumber)
      .then((nextPage) => {
        if (cancelled) {
          nextPage.cleanup();
          return;
        }

        loadedPage = nextPage;
        const viewport = nextPage.getViewport({ scale: 1, rotation });
        setBaseSize({ width: viewport.width, height: viewport.height });
        setPage(nextPage);
      })
      .catch((error: unknown) => {
        if (!cancelled && !isPdfCancellationError(error)) {
          onError(error);
        }
      });

    return () => {
      cancelled = true;
      loadedPage?.cleanup();
    };
  }, [onError, pageNumber, pdf, rotation]);

  useEffect(() => {
    const canvas = canvasRef.current;
    if (!canvas || !page) {
      return undefined;
    }

    const context = canvas.getContext('2d');
    if (!context) {
      onError(new Error('Unable to create a 2D canvas context for the PDF page.'));
      return undefined;
    }

    const viewport = page.getViewport({ scale: effectiveScale, rotation });
    const outputScale = window.devicePixelRatio || 1;

    canvas.width = Math.floor(viewport.width * outputScale);
    canvas.height = Math.floor(viewport.height * outputScale);
    canvas.style.width = `${viewport.width}px`;
    canvas.style.height = `${viewport.height}px`;

    const renderTask = page.render({
      canvasContext: context,
      viewport,
      transform: outputScale !== 1 ? [outputScale, 0, 0, outputScale, 0, 0] : undefined,
    });

    let cancelledRender = false;

    renderTask.promise.catch((error: unknown) => {
      if (!cancelledRender && !isPdfCancellationError(error)) {
        onError(error);
      }
    });

    return () => {
      cancelledRender = true;
      renderTask.cancel();
    };
  }, [effectiveScale, onError, page, rotation]);

  return <canvas aria-label={`Page ${pageNumber}`} className="ldv-pdf-canvas" ref={canvasRef} />;
}

function PdfRendererComponent({ file, state, actions, viewportRef }: DocumentRendererProps) {
  const [pdf, setPdf] = useState<PdfDocument | null>(null);
  const [isLoading, setIsLoading] = useState(true);
  const [pageSearchCounts, setPageSearchCounts] = useState<number[]>([]);

  useEffect(() => {
    let cancelled = false;
    let loadedDocument: PdfDocument | null = null;
    const loadingTask = pdfjs.getDocument({
      data: file.arrayBuffer.slice(0),
      useWorkerFetch: false,
      isEvalSupported: false,
    });

    setIsLoading(true);
    setPdf(null);
    setPageSearchCounts([]);

    loadingTask.promise
      .then((document) => {
        if (cancelled) {
          document.destroy();
          return;
        }

        loadedDocument = document;
        setPdf(document);
        setIsLoading(false);
        actions.setPageCount(document.numPages);
        actions.setDocumentInfo({
          title: file.fileName,
          pageCount: document.numPages,
        });
      })
      .catch((error: unknown) => {
        if (!cancelled && !isPdfCancellationError(error)) {
          setIsLoading(false);
          actions.reportError(error);
        }
      });

    return () => {
      cancelled = true;
      loadingTask.destroy();
      loadedDocument?.destroy();
    };
  }, [actions, file]);

  useEffect(() => {
    if (!pdf || !state.searchTerm.trim()) {
      setPageSearchCounts([]);
      actions.setSearchStats({});
      return;
    }

    let cancelled = false;
    const activePdf = pdf;
    const query = state.searchTerm.trim();

    async function buildSearchIndex() {
      actions.setSearchStats({ isSearching: true, message: 'Searching document' });
      const counts: number[] = [];
      let firstMatchPage: number | undefined;

      for (let pageNumber = 1; pageNumber <= activePdf.numPages; pageNumber += 1) {
        if (cancelled) {
          return;
        }

        const page = await activePdf.getPage(pageNumber);
        const textContent = await page.getTextContent();
        const text = textContent.items.map(textItemToString).join(' ');
        const matches = countMatches(text, query);
        counts.push(matches);
        page.cleanup();

        if (matches > 0 && !firstMatchPage) {
          firstMatchPage = pageNumber;
        }
      }

      if (cancelled) {
        return;
      }

      setPageSearchCounts(counts);
      const totalMatches = counts.reduce((sum, count) => sum + count, 0);
      actions.setSearchStats({
        totalMatches,
        currentPageMatches: counts[state.page - 1] ?? 0,
        isSearching: false,
        message: `${totalMatches} match${totalMatches === 1 ? '' : 'es'}`,
      });

      if (firstMatchPage) {
        actions.setPage(firstMatchPage);
      }
    }

    buildSearchIndex().catch((error: unknown) => {
      if (!cancelled && !isPdfCancellationError(error)) {
        actions.reportError(error);
      }
    });

    return () => {
      cancelled = true;
    };
  }, [actions, pdf, state.searchTerm]);

  useEffect(() => {
    if (!state.searchTerm.trim() || pageSearchCounts.length === 0) {
      return;
    }

    const totalMatches = pageSearchCounts.reduce((sum, count) => sum + count, 0);
    actions.setSearchStats({
      totalMatches,
      currentPageMatches: pageSearchCounts[state.page - 1] ?? 0,
      message: `${totalMatches} match${totalMatches === 1 ? '' : 'es'}`,
    });
  }, [actions, pageSearchCounts, state.page, state.searchTerm]);

  if (isLoading || !pdf) {
    return <div className="ldv-renderer-status">Loading PDF...</div>;
  }

  return (
    <div className="ldv-pdf-stage">
      <PdfPageCanvas
        fitMode={state.fitMode}
        onError={actions.reportError}
        pageNumber={state.page}
        pdf={pdf}
        rotation={state.rotation}
        viewportRef={viewportRef}
        zoom={state.zoom}
      />
    </div>
  );
}

export const PdfRenderer: DocumentRenderer = {
  id: 'pdf',
  label: 'PDF',
  priority: 50,
  canRender: (file) => file.extension === 'pdf' || file.mimeType === 'application/pdf',
  Component: PdfRendererComponent,
};
