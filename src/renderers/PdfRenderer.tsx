import { useVirtualizer } from '@tanstack/react-virtual';
import { PanelLeftClose, PanelLeftOpen } from 'lucide-react';
import * as pdfjs from 'pdfjs-dist';
import { useEffect, useMemo, useRef, useState } from 'react';
import type { RefObject } from 'react';
import type { DocumentRenderer, DocumentRendererProps, FitMode } from '../types';
import { countMatches, findMatchRanges } from '../utils/highlight';

let pdfWorkerModuleReady: Promise<void> | null = null;
let hasConfiguredPdfWorker = false;

async function ensurePdfWorkerReady(): Promise<void> {
  if (typeof window === 'undefined' || hasConfiguredPdfWorker) {
    return;
  }

  // Register PDF.js' local worker module before getDocument so Vite consumers do
  // not fetch a worker from /node_modules/.vite/deps.
  pdfWorkerModuleReady ??= import('pdfjs-dist/build/pdf.worker.min.mjs').then(() => undefined);
  await pdfWorkerModuleReady;
}

export function configurePdfWorker(workerSrc: string): void {
  hasConfiguredPdfWorker = true;
  pdfjs.GlobalWorkerOptions.workerSrc = workerSrc;
}

type PdfLoadingTask = ReturnType<typeof pdfjs.getDocument>;
type PdfDocument = Awaited<PdfLoadingTask['promise']>;
type PdfPage = Awaited<ReturnType<PdfDocument['getPage']>>;

interface Size {
  width: number;
  height: number;
}

interface PdfHighlightRect {
  left: number;
  top: number;
  width: number;
  height: number;
}

interface PdfTextItem {
  str: string;
  transform: number[];
  width: number;
  height: number;
}

const pdfThumbnailWidth = 92;
const pdfThumbnailFallbackHeight = 130;
const pdfThumbnailEstimatedHeight = 148;

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

function isPdfTextItem(item: unknown): item is PdfTextItem {
  return (
    item !== null &&
    typeof item === 'object' &&
    'str' in item &&
    typeof (item as { str?: unknown }).str === 'string' &&
    Array.isArray((item as { transform?: unknown }).transform) &&
    typeof (item as { width?: unknown }).width === 'number'
  );
}

function textItemHighlightRects(item: PdfTextItem, query: string, viewport: ReturnType<PdfPage['getViewport']>): PdfHighlightRect[] {
  const ranges = findMatchRanges(item.str, query);

  if (ranges.length === 0 || item.str.length === 0) {
    return [];
  }

  const transform = pdfjs.Util.transform(viewport.transform, item.transform);
  const angle = Math.atan2(transform[1], transform[0]);
  const fontHeight = Math.max(6, Math.hypot(transform[2], transform[3]));
  const textWidth = Math.max(1, Math.abs(item.width * viewport.scale));
  const left = angle === 0 ? transform[4] : transform[4] + fontHeight * Math.sin(angle);
  const top = angle === 0 ? transform[5] - fontHeight : transform[5] - fontHeight * Math.cos(angle);

  return ranges.map((range) => {
    const rangeStart = range.start / item.str.length;
    const rangeWidth = (range.end - range.start) / item.str.length;

    return {
      left: left + textWidth * rangeStart,
      top,
      width: Math.max(2, textWidth * rangeWidth),
      height: fontHeight,
    };
  });
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
    message.includes('worker has been destroyed') ||
    message.includes("reading 'getPage'") ||
    message.includes('reading "getPage"')
  );
}

interface PdfPageCanvasProps {
  pdf: PdfDocument;
  pageNumber: number;
  zoom: number;
  rotation: number;
  fitMode: FitMode;
  searchTerm: string;
  viewportRef: RefObject<HTMLDivElement>;
  onError: (error: unknown) => void;
}

function PdfPageCanvas({ pdf, pageNumber, zoom, rotation, fitMode, searchTerm, viewportRef, onError }: PdfPageCanvasProps) {
  const canvasRef = useRef<HTMLCanvasElement>(null);
  const [page, setPage] = useState<PdfPage | null>(null);
  const [baseSize, setBaseSize] = useState<Size | null>(null);
  const [highlightRects, setHighlightRects] = useState<PdfHighlightRect[]>([]);
  const containerSize = useElementSize(viewportRef);
  const effectiveScale = useMemo(
    () => scaleForFitMode(baseSize, containerSize, fitMode, zoom),
    [baseSize, containerSize, fitMode, zoom],
  );
  const displaySize = useMemo(
    () =>
      baseSize
        ? {
            width: baseSize.width * effectiveScale,
            height: baseSize.height * effectiveScale,
          }
        : null,
    [baseSize, effectiveScale],
  );

  useEffect(() => {
    let cancelled = false;
    let loadedPage: PdfPage | null = null;
    let pagePromise: Promise<PdfPage>;

    setPage(null);
    setBaseSize(null);
    setHighlightRects([]);

    try {
      pagePromise = pdf.getPage(pageNumber);
    } catch (error: unknown) {
      if (!isPdfCancellationError(error)) {
        onError(error);
      }

      return () => {
        cancelled = true;
      };
    }

    pagePromise
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
    const query = searchTerm.trim();

    if (!page || !query) {
      setHighlightRects([]);
      return undefined;
    }

    let cancelled = false;
    const activePage = page;

    async function buildHighlights() {
      const textContent = await activePage.getTextContent();
      const viewport = activePage.getViewport({ scale: effectiveScale, rotation });
      const rects = textContent.items.flatMap((item) => {
        if (!isPdfTextItem(item)) {
          return [];
        }

        return textItemHighlightRects(item, query, viewport);
      });

      if (!cancelled) {
        setHighlightRects(rects);
      }
    }

    buildHighlights().catch((error: unknown) => {
      if (!cancelled && !isPdfCancellationError(error)) {
        onError(error);
      }
    });

    return () => {
      cancelled = true;
    };
  }, [effectiveScale, onError, page, rotation, searchTerm]);

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

    let cancelledRender = false;
    let renderTask: ReturnType<PdfPage['render']>;

    try {
      const viewport = page.getViewport({ scale: effectiveScale, rotation });
      const outputScale = window.devicePixelRatio || 1;

      canvas.width = Math.floor(viewport.width * outputScale);
      canvas.height = Math.floor(viewport.height * outputScale);
      canvas.style.width = `${viewport.width}px`;
      canvas.style.height = `${viewport.height}px`;

      renderTask = page.render({
        canvasContext: context,
        viewport,
        transform: outputScale !== 1 ? [outputScale, 0, 0, outputScale, 0, 0] : undefined,
      });
    } catch (error: unknown) {
      if (!isPdfCancellationError(error)) {
        onError(error);
      }

      return () => {
        cancelledRender = true;
      };
    }

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

  return (
    <div
      className="ldv-pdf-page-surface"
      style={{
        height: displaySize ? `${displaySize.height}px` : undefined,
        width: displaySize ? `${displaySize.width}px` : undefined,
      }}
    >
      <canvas aria-label={`Page ${pageNumber}`} className="ldv-pdf-canvas" ref={canvasRef} />
      {highlightRects.length > 0 ? (
        <div className="ldv-pdf-highlight-layer" aria-hidden="true">
          {highlightRects.map((rect, index) => (
            <span
              className="ldv-pdf-search-highlight"
              key={`${rect.left}-${rect.top}-${index}`}
              style={{
                height: `${rect.height}px`,
                left: `${rect.left}px`,
                top: `${rect.top}px`,
                width: `${rect.width}px`,
              }}
            />
          ))}
        </div>
      ) : null}
    </div>
  );
}

interface PdfThumbnailCanvasProps {
  pdf: PdfDocument;
  pageNumber: number;
  rotation: number;
}

function PdfThumbnailCanvas({ pdf, pageNumber, rotation }: PdfThumbnailCanvasProps) {
  const canvasRef = useRef<HTMLCanvasElement>(null);
  const [status, setStatus] = useState<'loading' | 'ready' | 'error'>('loading');
  const [displaySize, setDisplaySize] = useState<Size | null>(null);

  useEffect(() => {
    const canvas = canvasRef.current;

    if (!canvas) {
      return undefined;
    }

    const activeCanvas = canvas;
    let cancelled = false;
    let loadedPage: PdfPage | null = null;
    let renderTask: ReturnType<PdfPage['render']> | null = null;

    setStatus('loading');
    setDisplaySize(null);

    async function renderThumbnail() {
      const page = await pdf.getPage(pageNumber);

      if (cancelled) {
        page.cleanup();
        return;
      }

      loadedPage = page;

      const baseViewport = page.getViewport({ scale: 1, rotation });
      const scale = pdfThumbnailWidth / baseViewport.width;
      const viewport = page.getViewport({ scale, rotation });
      const outputScale = window.devicePixelRatio || 1;
      const context = activeCanvas.getContext('2d');

      if (!context) {
        throw new Error('Unable to create a 2D canvas context for the PDF thumbnail.');
      }

      activeCanvas.width = Math.floor(viewport.width * outputScale);
      activeCanvas.height = Math.floor(viewport.height * outputScale);
      activeCanvas.style.width = `${viewport.width}px`;
      activeCanvas.style.height = `${viewport.height}px`;
      setDisplaySize({ width: viewport.width, height: viewport.height });

      renderTask = page.render({
        canvasContext: context,
        viewport,
        transform: outputScale !== 1 ? [outputScale, 0, 0, outputScale, 0, 0] : undefined,
      });

      await renderTask.promise;

      if (!cancelled) {
        page.cleanup();
        loadedPage = null;
        setStatus('ready');
      }
    }

    renderThumbnail().catch((error: unknown) => {
      if (!cancelled && !isPdfCancellationError(error)) {
        setStatus('error');
      }
    });

    return () => {
      cancelled = true;
      renderTask?.cancel();
      loadedPage?.cleanup();
    };
  }, [pageNumber, pdf, rotation]);

  return (
    <span
      className="ldv-pdf-thumbnail-preview"
      style={{
        height: `${displaySize?.height ?? pdfThumbnailFallbackHeight}px`,
        width: `${pdfThumbnailWidth}px`,
      }}
    >
      <canvas
        aria-hidden="true"
        className="ldv-pdf-thumbnail-canvas"
        ref={canvasRef}
        style={{ visibility: status === 'ready' ? 'visible' : 'hidden' }}
      />
      {status === 'loading' ? (
        <span className="ldv-thumbnail-loader" aria-hidden="true">
          <span className="ldv-thumbnail-loader-spinner" />
        </span>
      ) : null}
      {status === 'error' ? <span className="ldv-thumbnail-error">Unable</span> : null}
    </span>
  );
}

function PdfRendererComponent({ file, state, actions, controls, pdfOptions }: DocumentRendererProps) {
  const [pdf, setPdf] = useState<PdfDocument | null>(null);
  const [isLoading, setIsLoading] = useState(true);
  const [isSidebarCollapsed, setIsSidebarCollapsed] = useState(false);
  const [pageSearchCounts, setPageSearchCounts] = useState<number[]>([]);
  const mainPaneRef = useRef<HTMLDivElement>(null);
  const thumbnailListRef = useRef<HTMLDivElement>(null);
  const pageRefs = useRef<Array<HTMLDivElement | null>>([]);
  const isUpdatingPageFromScrollRef = useRef(false);

  useEffect(() => {
    let cancelled = false;
    let loadedDocument: PdfDocument | null = null;
    let loadingTask: PdfLoadingTask | null = null;

    setIsLoading(true);
    actions.setLoading(true);
    setPdf(null);
    setPageSearchCounts([]);
    pageRefs.current = [];

    async function loadPdf() {
      await ensurePdfWorkerReady();

      if (cancelled) {
        return;
      }

      const nextLoadingTask = pdfjs.getDocument({
        data: file.arrayBuffer.slice(0),
        useWorkerFetch: false,
        isEvalSupported: false,
      });
      loadingTask = nextLoadingTask;

      nextLoadingTask.promise
        .then((document) => {
          if (cancelled) {
            document.destroy();
            return;
          }

          loadedDocument = document;
          setPdf(document);
          setIsLoading(false);
          actions.setLoading(false);
          actions.setPageCount(document.numPages);
          actions.setDocumentInfo({
            title: file.fileName,
            pageCount: document.numPages,
          });
        })
        .catch((error: unknown) => {
          if (!cancelled && !isPdfCancellationError(error)) {
            setIsLoading(false);
            actions.setLoading(false);
            actions.reportError(error);
          }
        });
    }

    loadPdf().catch((error: unknown) => {
      if (!cancelled && !isPdfCancellationError(error)) {
        setIsLoading(false);
        actions.setLoading(false);
        actions.reportError(error);
      }
    });

    return () => {
      cancelled = true;
      actions.setLoading(false);
      loadingTask?.destroy();
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

        let page: PdfPage;

        try {
          page = await activePdf.getPage(pageNumber);
        } catch (error: unknown) {
          if (cancelled || isPdfCancellationError(error)) {
            return;
          }

          throw error;
        }

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

  const pageNumbers = useMemo(() => {
    if (!pdf) {
      return [];
    }

    return Array.from({ length: pdf.numPages }, (_, index) => index + 1);
  }, [pdf]);
  const isContinuousMode = Boolean(pdf && !controls.pageNavigation && pdf.numPages > 1);
  const hasThumbnailSidebar = Boolean(pdf && pdfOptions.showThumbnails && controls.thumbnails && pdf.numPages > 0);
  const thumbnailVirtualizer = useVirtualizer({
    count: hasThumbnailSidebar && !isSidebarCollapsed ? pageNumbers.length : 0,
    estimateSize: () => pdfThumbnailEstimatedHeight,
    getScrollElement: () => thumbnailListRef.current,
    overscan: 6,
  });
  const virtualThumbnailPages = thumbnailVirtualizer.getVirtualItems();

  useEffect(() => {
    if (!isContinuousMode) {
      return;
    }

    if (isUpdatingPageFromScrollRef.current) {
      isUpdatingPageFromScrollRef.current = false;
      return;
    }

    pageRefs.current[state.page - 1]?.scrollIntoView({ block: 'start', inline: 'nearest' });
  }, [isContinuousMode, state.page]);

  useEffect(() => {
    const container = mainPaneRef.current;

    if (!isContinuousMode || !container) {
      return undefined;
    }

    let frame = 0;

    const updatePageFromScroll = () => {
      frame = 0;
      const containerRect = container.getBoundingClientRect();
      const marker = containerRect.top + Math.min(containerRect.height * 0.35, 240);
      let nextPage = state.page;
      let nearestDistance = Number.POSITIVE_INFINITY;

      pageRefs.current.forEach((element, index) => {
        if (!element) {
          return;
        }

        const rect = element.getBoundingClientRect();
        const distance =
          rect.top <= marker && rect.bottom >= marker
            ? 0
            : Math.min(Math.abs(rect.top - marker), Math.abs(rect.bottom - marker));

        if (distance < nearestDistance) {
          nearestDistance = distance;
          nextPage = index + 1;
        }
      });

      if (nextPage !== state.page) {
        isUpdatingPageFromScrollRef.current = true;
        actions.setPage(nextPage);
      }
    };

    const handleScroll = () => {
      if (frame) {
        return;
      }

      frame = window.requestAnimationFrame(updatePageFromScroll);
    };

    container.addEventListener('scroll', handleScroll, { passive: true });
    updatePageFromScroll();

    return () => {
      container.removeEventListener('scroll', handleScroll);
      if (frame) {
        window.cancelAnimationFrame(frame);
      }
    };
  }, [actions, isContinuousMode, pageNumbers, state.page]);

  useEffect(() => {
    if (!hasThumbnailSidebar || isSidebarCollapsed) {
      return;
    }

    thumbnailVirtualizer.scrollToIndex(state.page - 1, { align: 'auto' });
  }, [hasThumbnailSidebar, isSidebarCollapsed, state.page, thumbnailVirtualizer]);

  if (isLoading || !pdf) {
    return <div className="ldv-renderer-status">Loading PDF...</div>;
  }

  const shellClassName = [
    'ldv-pdf-shell',
    hasThumbnailSidebar ? null : 'ldv-pdf-sidebar-disabled',
    hasThumbnailSidebar && isSidebarCollapsed ? 'ldv-pdf-sidebar-collapsed' : null,
    isContinuousMode ? 'ldv-pdf-continuous' : null,
  ]
    .filter(Boolean)
    .join(' ');

  return (
    <div className={shellClassName}>
      {hasThumbnailSidebar ? (
        <aside aria-label="Pages" className="ldv-pdf-sidebar">
          <button
            aria-controls="ldv-pdf-thumbnails"
            aria-expanded={!isSidebarCollapsed}
            aria-label={isSidebarCollapsed ? 'Show pages' : 'Hide pages'}
            className="ldv-pdf-sidebar-toggle"
            onClick={() => setIsSidebarCollapsed((current) => !current)}
            title={isSidebarCollapsed ? 'Show pages' : 'Hide pages'}
            type="button"
          >
            {isSidebarCollapsed ? <PanelLeftOpen aria-hidden="true" size={18} /> : <PanelLeftClose aria-hidden="true" size={18} />}
          </button>
          {!isSidebarCollapsed ? (
            <div className="ldv-pdf-thumbnail-list" id="ldv-pdf-thumbnails" ref={thumbnailListRef}>
              <div
                className="ldv-pdf-thumbnail-virtual-space"
                style={{ height: `${thumbnailVirtualizer.getTotalSize()}px` }}
              >
                {virtualThumbnailPages.map((virtualPage) => {
                  const pageNumber = virtualPage.index + 1;

                  return (
                    <button
                      data-index={virtualPage.index}
                      aria-current={state.page === pageNumber ? 'page' : undefined}
                      aria-label={`Page ${pageNumber}`}
                      className="ldv-pdf-thumbnail"
                      key={virtualPage.key}
                      onClick={() => actions.setPage(pageNumber)}
                      ref={(element) => {
                        if (element) {
                          thumbnailVirtualizer.measureElement(element);
                        }
                      }}
                      style={{ transform: `translate(-50%, ${virtualPage.start}px)` }}
                      type="button"
                    >
                      <span className="ldv-pdf-thumbnail-number">{pageNumber}</span>
                      <PdfThumbnailCanvas pageNumber={pageNumber} pdf={pdf} rotation={state.rotation} />
                    </button>
                  );
                })}
              </div>
            </div>
          ) : null}
        </aside>
      ) : null}
      <div className="ldv-pdf-main" ref={mainPaneRef}>
        {isContinuousMode ? (
          <div className="ldv-pdf-continuous-stage">
            {pageNumbers.map((pageNumber) => (
              <div
                className="ldv-pdf-page-frame"
                key={pageNumber}
                ref={(element) => {
                  pageRefs.current[pageNumber - 1] = element;
                }}
              >
                <PdfPageCanvas
                  fitMode={state.fitMode}
                  onError={actions.reportError}
                  pageNumber={pageNumber}
                  pdf={pdf}
                  rotation={state.rotation}
                  searchTerm={state.searchTerm}
                  viewportRef={mainPaneRef}
                  zoom={state.zoom}
                />
              </div>
            ))}
          </div>
        ) : (
          <div className="ldv-pdf-stage">
            <PdfPageCanvas
              fitMode={state.fitMode}
              onError={actions.reportError}
              pageNumber={state.page}
              pdf={pdf}
              rotation={state.rotation}
              searchTerm={state.searchTerm}
              viewportRef={mainPaneRef}
              zoom={state.zoom}
            />
          </div>
        )}
      </div>
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
