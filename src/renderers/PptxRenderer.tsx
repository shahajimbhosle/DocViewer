import { useVirtualizer } from '@tanstack/react-virtual';
import { PanelLeftClose, PanelLeftOpen } from 'lucide-react';
import { PPTXViewer } from 'pptxviewjs';
import JSZip from 'jszip';
import { useEffect, useMemo, useRef, useState } from 'react';
import type { RefObject } from 'react';
import type { DocumentRenderer, DocumentRendererProps, FitMode } from '../types';
import { countMatches } from '../utils/highlight';

const emuPerInch = 914400;
const cssPixelsPerInch = 96;
const defaultSlideSize = {
  width: 960,
  height: 720,
};
const thumbnailWidth = 96;
const thumbnailEstimatedHeight = 96;

interface Size {
  width: number;
  height: number;
}

interface PptxSlideSize {
  cx: number;
  cy: number;
}

interface PreparedPptx {
  arrayBuffer: ArrayBuffer;
  slideTexts: string[];
  slideSearchBoxes: PptxSearchBox[][];
  warnings: string[];
}

interface PptxSearchBox {
  text: string;
  x: number;
  y: number;
  width: number;
  height: number;
}

type PptxViewerWithInternals = PPTXViewer & {
  processor?: {
    getSlideDimensions?: () => PptxSlideSize;
  };
  presentation?: {
    slideSize?: PptxSlideSize;
  };
};

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

function isNetworkReference(value: string): boolean {
  const clean = value.trim();
  return /^https?:\/\//i.test(clean) || /^\/\//.test(clean);
}

function serializeXml(document: Document): string {
  return new XMLSerializer().serializeToString(document);
}

function childByLocalName(element: Element, localName: string): Element | null {
  return Array.from(element.children).find((child) => child.localName === localName) ?? null;
}

function descendantsByLocalName(element: Element | Document, localName: string): Element[] {
  return Array.from(element.getElementsByTagName('*')).filter((node) => node.localName === localName);
}

function firstDescendantByLocalName(element: Element, localName: string): Element | null {
  return descendantsByLocalName(element, localName)[0] ?? null;
}

function emuToPixels(value: string | null): number {
  const emu = Number(value);

  if (!Number.isFinite(emu)) {
    return 0;
  }

  return (emu / emuPerInch) * cssPixelsPerInch;
}

async function hardenRelationshipFiles(zip: JSZip, parser: DOMParser): Promise<string[]> {
  const warnings: string[] = [];
  const relationshipFiles = Object.keys(zip.files).filter((path) => path.endsWith('.rels'));

  await Promise.all(
    relationshipFiles.map(async (path) => {
      const relationshipFile = zip.file(path);
      if (!relationshipFile) {
        return;
      }

      const xml = await relationshipFile.async('text');
      const document = parser.parseFromString(xml, 'application/xml');
      const relationships = Array.from(document.getElementsByTagName('Relationship'));
      let changed = false;
      let removedCount = 0;

      relationships.forEach((relationship) => {
        const targetMode = relationship.getAttribute('TargetMode') ?? '';
        const target = relationship.getAttribute('Target') ?? '';
        const isExternal = targetMode.toLowerCase() === 'external' || isNetworkReference(target);

        if (isExternal) {
          relationship.remove();
          changed = true;
          removedCount += 1;
        }
      });

      if (changed) {
        zip.file(path, serializeXml(document));
        warnings.push(`Removed ${removedCount} external PPTX relationship${removedCount === 1 ? '' : 's'} from ${path}.`);
      }
    }),
  );

  return warnings;
}

function slideNumberFromPath(path: string): number {
  const match = /\/slide(\d+)\.xml$/i.exec(path);
  return match ? Number(match[1]) : Number.MAX_SAFE_INTEGER;
}

function textFromSlideXml(xml: string, parser: DOMParser): string {
  const document = parser.parseFromString(xml, 'application/xml');

  return descendantsByLocalName(document, 't')
    .map((node) => node.textContent ?? '')
    .join(' ')
    .trim();
}

function searchBoxesFromSlideXml(xml: string, parser: DOMParser): PptxSearchBox[] {
  const document = parser.parseFromString(xml, 'application/xml');
  const shapes = Array.from(document.getElementsByTagName('*')).filter((node) => {
    return node.localName === 'sp' || node.localName === 'graphicFrame';
  });

  return shapes.flatMap((shape) => {
    const text = descendantsByLocalName(shape, 't')
      .map((node) => node.textContent ?? '')
      .join(' ')
      .trim();

    if (!text) {
      return [];
    }

    const shapeProperties = childByLocalName(shape, 'spPr') ?? firstDescendantByLocalName(shape, 'spPr');
    const transform = shapeProperties ? firstDescendantByLocalName(shapeProperties, 'xfrm') : firstDescendantByLocalName(shape, 'xfrm');
    const offset = transform ? childByLocalName(transform, 'off') : null;
    const extent = transform ? childByLocalName(transform, 'ext') : null;

    if (!offset || !extent) {
      return [];
    }

    const width = emuToPixels(extent.getAttribute('cx'));
    const height = emuToPixels(extent.getAttribute('cy'));

    if (width <= 0 || height <= 0) {
      return [];
    }

    return [
      {
        text,
        x: emuToPixels(offset.getAttribute('x')),
        y: emuToPixels(offset.getAttribute('y')),
        width,
        height,
      },
    ];
  });
}

function maxRunFontSize(paragraph: Element): number | null {
  const runProperties = Array.from(paragraph.getElementsByTagName('*')).filter((node) => node.localName === 'rPr');
  const sizes = runProperties
    .map((node) => Number(node.getAttribute('sz')))
    .filter((size) => Number.isFinite(size) && size > 0);

  if (sizes.length === 0) {
    return null;
  }

  return Math.max(...sizes);
}

function upsertParagraphFontDefault(document: Document, paragraph: Element, fontSize: number): boolean {
  let paragraphProperties = Array.from(paragraph.children).find((child) => child.localName === 'pPr');

  if (!paragraphProperties) {
    paragraphProperties = document.createElementNS('http://schemas.openxmlformats.org/drawingml/2006/main', 'a:pPr');
    paragraph.insertBefore(paragraphProperties, paragraph.firstChild);
  }

  let defaultRunProperties = Array.from(paragraphProperties.children).find((child) => child.localName === 'defRPr');

  if (!defaultRunProperties) {
    defaultRunProperties = document.createElementNS('http://schemas.openxmlformats.org/drawingml/2006/main', 'a:defRPr');
    paragraphProperties.append(defaultRunProperties);
  }

  const currentFontSize = Number(defaultRunProperties.getAttribute('sz'));

  if (Number.isFinite(currentFontSize) && currentFontSize >= fontSize) {
    return false;
  }

  defaultRunProperties.setAttribute('sz', String(fontSize));
  return true;
}

async function normalizeSlideParagraphFontDefaults(zip: JSZip, parser: DOMParser): Promise<boolean> {
  const slidePaths = Object.keys(zip.files).filter((path) => /^ppt\/slides\/slide\d+\.xml$/i.test(path));
  let changed = false;

  await Promise.all(
    slidePaths.map(async (path) => {
      const slideFile = zip.file(path);
      if (!slideFile) {
        return;
      }

      const xml = await slideFile.async('text');
      const document = parser.parseFromString(xml, 'application/xml');
      let slideChanged = false;

      Array.from(document.getElementsByTagName('*'))
        .filter((node) => node.localName === 'p')
        .forEach((paragraph) => {
          const fontSize = maxRunFontSize(paragraph);
          if (fontSize && upsertParagraphFontDefault(document, paragraph, fontSize)) {
            slideChanged = true;
          }
        });

      if (slideChanged) {
        zip.file(path, serializeXml(document));
        changed = true;
      }
    }),
  );

  return changed;
}

async function extractPptxSlideTexts(zip: JSZip, parser: DOMParser): Promise<string[]> {
  const slidePaths = Object.keys(zip.files)
    .filter((path) => /^ppt\/slides\/slide\d+\.xml$/i.test(path))
    .sort((left, right) => slideNumberFromPath(left) - slideNumberFromPath(right));

  return Promise.all(
    slidePaths.map(async (path) => {
      const slideFile = zip.file(path);
      if (!slideFile) {
        return '';
      }

      return textFromSlideXml(await slideFile.async('text'), parser);
    }),
  );
}

async function extractPptxSlideSearchBoxes(zip: JSZip, parser: DOMParser): Promise<PptxSearchBox[][]> {
  const slidePaths = Object.keys(zip.files)
    .filter((path) => /^ppt\/slides\/slide\d+\.xml$/i.test(path))
    .sort((left, right) => slideNumberFromPath(left) - slideNumberFromPath(right));

  return Promise.all(
    slidePaths.map(async (path) => {
      const slideFile = zip.file(path);
      if (!slideFile) {
        return [];
      }

      return searchBoxesFromSlideXml(await slideFile.async('text'), parser);
    }),
  );
}

async function preparePptx(arrayBuffer: ArrayBuffer): Promise<PreparedPptx> {
  const parser = new DOMParser();
  const zip = await JSZip.loadAsync(arrayBuffer.slice(0));
  const warnings = await hardenRelationshipFiles(zip, parser);
  const normalizedTextSizing = await normalizeSlideParagraphFontDefaults(zip, parser);
  const [slideTexts, slideSearchBoxes] = await Promise.all([
    extractPptxSlideTexts(zip, parser),
    extractPptxSlideSearchBoxes(zip, parser),
  ]);
  const safeArrayBuffer =
    warnings.length > 0 || normalizedTextSizing ? await zip.generateAsync({ type: 'arraybuffer' }) : arrayBuffer.slice(0);

  return {
    arrayBuffer: safeArrayBuffer,
    slideTexts,
    slideSearchBoxes,
    warnings,
  };
}

function slideSizeToPixels(size?: PptxSlideSize): Size {
  if (!size || !Number.isFinite(size.cx) || !Number.isFinite(size.cy) || size.cx <= 0 || size.cy <= 0) {
    return defaultSlideSize;
  }

  return {
    width: (size.cx / emuPerInch) * cssPixelsPerInch,
    height: (size.cy / emuPerInch) * cssPixelsPerInch,
  };
}

function slideSizeFromViewer(viewer: PPTXViewer): Size {
  const viewerWithInternals = viewer as PptxViewerWithInternals;
  const internalSize = viewerWithInternals.processor?.getSlideDimensions?.() ?? viewerWithInternals.presentation?.slideSize;

  return slideSizeToPixels(internalSize);
}

function scaleForFitMode(slideSize: Size, viewportSize: Size, fitMode: FitMode, zoom: number): number {
  if (fitMode === 'manual') {
    return zoom;
  }

  const availableWidth = Math.max(viewportSize.width - 48, 160);
  const availableHeight = Math.max(viewportSize.height - 48, 160);
  const widthScale = availableWidth / slideSize.width;
  const heightScale = availableHeight / slideSize.height;

  return Math.max(0.1, Math.min(fitMode === 'width' ? widthScale : Math.min(widthScale, heightScale), 6));
}

function displaySizeForSlide(slideSize: Size, viewportSize: Size, fitMode: FitMode, zoom: number): Size {
  const scale = scaleForFitMode(slideSize, viewportSize, fitMode, zoom);

  return {
    width: Math.max(1, Math.round(slideSize.width * scale)),
    height: Math.max(1, Math.round(slideSize.height * scale)),
  };
}

function prepareCanvas(canvas: HTMLCanvasElement, displaySize: Size) {
  const pixelRatio = window.devicePixelRatio || 1;

  canvas.style.width = `${displaySize.width}px`;
  canvas.style.height = `${displaySize.height}px`;
  canvas.width = Math.max(1, Math.round(displaySize.width * pixelRatio));
  canvas.height = Math.max(1, Math.round(displaySize.height * pixelRatio));
}

function setCanvasDisplaySize(canvas: HTMLCanvasElement, displaySize: Size) {
  canvas.style.width = `${displaySize.width}px`;
  canvas.style.height = `${displaySize.height}px`;
}

function thumbnailSizeForSlide(slideSize: Size): Size {
  return {
    width: thumbnailWidth,
    height: Math.max(1, Math.round((thumbnailWidth * slideSize.height) / slideSize.width)),
  };
}

function PptxRendererComponent({ file, state, actions, controls }: DocumentRendererProps) {
  const canvasRef = useRef<HTMLCanvasElement>(null);
  const mainPaneRef = useRef<HTMLDivElement>(null);
  const thumbnailListRef = useRef<HTMLDivElement>(null);
  const thumbnailCanvasRefs = useRef<Array<HTMLCanvasElement | null>>([]);
  const viewerRef = useRef<PPTXViewer | null>(null);
  const thumbnailViewerRef = useRef<PPTXViewer | null>(null);
  const thumbnailViewerReadyRef = useRef<Promise<PPTXViewer> | null>(null);
  const renderIdRef = useRef(0);
  const thumbnailRenderIdRef = useRef(0);
  const lastSearchNavigationTermRef = useRef('');
  const mainPaneSize = useElementSize(mainPaneRef);
  const [isLoading, setIsLoading] = useState(true);
  const [isRendering, setIsRendering] = useState(false);
  const [isSidebarCollapsed, setIsSidebarCollapsed] = useState(false);
  const [presentationBuffer, setPresentationBuffer] = useState<ArrayBuffer | null>(null);
  const [slideCount, setSlideCount] = useState(0);
  const [slideSize, setSlideSize] = useState<Size>(defaultSlideSize);
  const [slideSearchBoxes, setSlideSearchBoxes] = useState<PptxSearchBox[][]>([]);
  const [slideTexts, setSlideTexts] = useState<string[]>([]);
  const [thumbnailStatuses, setThumbnailStatuses] = useState<Record<number, 'loading' | 'ready' | 'error'>>({});
  const [warnings, setWarnings] = useState<string[]>([]);

  useEffect(() => {
    let cancelled = false;
    const viewer = new PPTXViewer({
      autoChartRerenderDelayMs: 200,
      backgroundColor: '#ffffff',
      debug: false,
      enableThumbnails: false,
    });

    viewerRef.current = viewer;
    setIsLoading(true);
    actions.setLoading(true);
    setIsRendering(false);
    setPresentationBuffer(null);
    setSlideCount(0);
    setSlideSearchBoxes([]);
    setSlideTexts([]);
    lastSearchNavigationTermRef.current = '';
    setThumbnailStatuses({});
    setWarnings([]);
    thumbnailCanvasRefs.current = [];
    thumbnailViewerRef.current?.destroy();
    thumbnailViewerRef.current = null;
    thumbnailViewerReadyRef.current = null;
    actions.setPageCount(undefined);
    actions.setSearchStats({});
    actions.setDocumentInfo({ title: file.fileName });

    async function loadPresentation() {
      const prepared = await preparePptx(file.arrayBuffer);
      await viewer.loadFile(prepared.arrayBuffer.slice(0));

      if (cancelled) {
        viewer.destroy();
        return;
      }

      const slideCount = viewer.getSlideCount();
      const nextWarnings = prepared.warnings.length > 0 ? ['Removed unsafe external PPTX references before preview.'] : [];

      setSlideSize(slideSizeFromViewer(viewer));
      setPresentationBuffer(prepared.arrayBuffer.slice(0));
      setSlideCount(slideCount);
      setSlideSearchBoxes(prepared.slideSearchBoxes);
      setSlideTexts(prepared.slideTexts);
      setWarnings(nextWarnings);
      setIsLoading(false);
      actions.setLoading(false);
      actions.setPageCount(slideCount || undefined);
      actions.setDocumentInfo({
        title: file.fileName,
        pageCount: slideCount || undefined,
        warnings: nextWarnings,
      });
    }

    loadPresentation().catch((error: unknown) => {
      if (!cancelled) {
        setIsLoading(false);
        actions.setLoading(false);
        actions.reportError(error);
      }
    });

    return () => {
      cancelled = true;
      actions.setLoading(false);
      viewer.destroy();
      thumbnailViewerRef.current?.destroy();
      if (viewerRef.current === viewer) {
        viewerRef.current = null;
      }
      thumbnailViewerRef.current = null;
      thumbnailViewerReadyRef.current = null;
    };
  }, [actions, file]);

  const displaySize = useMemo(
    () => displaySizeForSlide(slideSize, mainPaneSize, state.fitMode, state.zoom),
    [mainPaneSize, slideSize, state.fitMode, state.zoom],
  );
  const thumbnailSize = useMemo(() => thumbnailSizeForSlide(slideSize), [slideSize]);
  const hasThumbnailSidebar = controls.thumbnails && slideCount > 0;
  const thumbnailVirtualizer = useVirtualizer({
    count: hasThumbnailSidebar && !isSidebarCollapsed ? slideCount : 0,
    estimateSize: () => Math.max(thumbnailEstimatedHeight, thumbnailSize.height + 20),
    getScrollElement: () => thumbnailListRef.current,
    overscan: 6,
  });
  const virtualThumbnailSlides = thumbnailVirtualizer.getVirtualItems();
  const visibleThumbnailSlideIndexes = virtualThumbnailSlides.map((virtualSlide) => virtualSlide.index);
  const visibleThumbnailSlideKey = visibleThumbnailSlideIndexes.join(',');

  useEffect(() => {
    const viewer = viewerRef.current;
    const canvas = canvasRef.current;

    if (isLoading || !viewer || !canvas || viewer.getSlideCount() < 1) {
      return undefined;
    }

    const renderId = renderIdRef.current + 1;
    renderIdRef.current = renderId;
    const slideIndex = Math.min(Math.max(state.page - 1, 0), viewer.getSlideCount() - 1);

    prepareCanvas(canvas, displaySize);
    setIsRendering(true);

    viewer.renderSlide(slideIndex, canvas, { quality: 'high' }).then(
      () => {
        if (renderIdRef.current === renderId) {
          setIsRendering(false);
        }
      },
      (error: unknown) => {
        if (renderIdRef.current === renderId) {
          setIsRendering(false);
          actions.reportError(error);
        }
      },
    );

    return undefined;
  }, [actions, displaySize, isLoading, state.page]);

  useEffect(() => {
    const activePresentationBuffer = presentationBuffer;

    if (isLoading || !hasThumbnailSidebar || isSidebarCollapsed || !activePresentationBuffer) {
      thumbnailViewerRef.current?.destroy();
      thumbnailViewerRef.current = null;
      thumbnailViewerReadyRef.current = null;
      return undefined;
    }

    const thumbnailViewer = new PPTXViewer({
      autoChartRerenderDelayMs: 0,
      backgroundColor: '#ffffff',
      debug: false,
      enableThumbnails: false,
    });

    thumbnailViewerRef.current?.destroy();
    thumbnailViewerRef.current = thumbnailViewer;
    setThumbnailStatuses({});

    const readyPromise = thumbnailViewer.loadFile(activePresentationBuffer.slice(0)).then(() => thumbnailViewer);
    thumbnailViewerReadyRef.current = readyPromise;
    readyPromise.catch(() => undefined);

    return () => {
      thumbnailViewer.destroy();
      if (thumbnailViewerRef.current === thumbnailViewer) {
        thumbnailViewerRef.current = null;
      }
      if (thumbnailViewerReadyRef.current === readyPromise) {
        thumbnailViewerReadyRef.current = null;
      }
    };
  }, [hasThumbnailSidebar, isLoading, isSidebarCollapsed, presentationBuffer]);

  useEffect(() => {
    const readyPromise = thumbnailViewerReadyRef.current;

    if (
      isLoading ||
      !hasThumbnailSidebar ||
      isSidebarCollapsed ||
      !readyPromise ||
      visibleThumbnailSlideIndexes.length === 0
    ) {
      return undefined;
    }

    let cancelled = false;
    const renderId = thumbnailRenderIdRef.current + 1;
    thumbnailRenderIdRef.current = renderId;

    async function renderVisibleThumbnails() {
      const thumbnailViewer = await readyPromise;
      if (!thumbnailViewer) {
        return;
      }

      for (const slideIndex of visibleThumbnailSlideIndexes) {
        if (cancelled || thumbnailRenderIdRef.current !== renderId) {
          return;
        }

        const thumbnailCanvas = thumbnailCanvasRefs.current[slideIndex];
        if (!thumbnailCanvas) {
          continue;
        }

        if (
          thumbnailCanvas.dataset.ldvRenderedSlide === String(slideIndex) &&
          thumbnailCanvas.dataset.ldvRenderedWidth === String(thumbnailSize.width) &&
          thumbnailCanvas.dataset.ldvRenderedHeight === String(thumbnailSize.height)
        ) {
          setThumbnailStatuses((current) => ({ ...current, [slideIndex]: 'ready' }));
          continue;
        }

        setThumbnailStatuses((current) => ({ ...current, [slideIndex]: 'loading' }));

        try {
          prepareCanvas(thumbnailCanvas, slideSize);
          await thumbnailViewer.renderSlide(slideIndex, thumbnailCanvas, { quality: 'low' });

          if (cancelled || thumbnailRenderIdRef.current !== renderId) {
            return;
          }

          setCanvasDisplaySize(thumbnailCanvas, thumbnailSize);
          thumbnailCanvas.dataset.ldvRenderedSlide = String(slideIndex);
          thumbnailCanvas.dataset.ldvRenderedWidth = String(thumbnailSize.width);
          thumbnailCanvas.dataset.ldvRenderedHeight = String(thumbnailSize.height);
          setThumbnailStatuses((current) => ({ ...current, [slideIndex]: 'ready' }));
        } catch {
          if (!cancelled && thumbnailRenderIdRef.current === renderId) {
            setThumbnailStatuses((current) => ({ ...current, [slideIndex]: 'error' }));
          }
        }
      }
    }

    renderVisibleThumbnails().catch(() => {
      if (!cancelled && thumbnailRenderIdRef.current === renderId) {
        setThumbnailStatuses((current) => {
          const nextStatuses = { ...current };
          visibleThumbnailSlideIndexes.forEach((slideIndex) => {
            nextStatuses[slideIndex] = 'error';
          });
          return nextStatuses;
        });
      }
    });

    return () => {
      cancelled = true;
    };
  }, [
    hasThumbnailSidebar,
    isLoading,
    isSidebarCollapsed,
    slideSize,
    thumbnailSize,
    visibleThumbnailSlideKey,
  ]);

  useEffect(() => {
    if (!hasThumbnailSidebar || isSidebarCollapsed) {
      return;
    }

    thumbnailVirtualizer.scrollToIndex(state.page - 1, { align: 'auto' });
  }, [hasThumbnailSidebar, isSidebarCollapsed, state.page, thumbnailVirtualizer]);

  const totalMatches = useMemo(
    () => slideTexts.reduce((total, slideText) => total + countMatches(slideText, state.searchTerm), 0),
    [slideTexts, state.searchTerm],
  );
  const currentPageMatches = useMemo(
    () => countMatches(slideTexts[state.page - 1] ?? '', state.searchTerm),
    [slideTexts, state.page, state.searchTerm],
  );
  const currentSearchBoxes = useMemo(() => {
    const query = state.searchTerm.trim();

    if (!query) {
      return [];
    }

    return (slideSearchBoxes[state.page - 1] ?? []).filter((box) => countMatches(box.text, query) > 0);
  }, [slideSearchBoxes, state.page, state.searchTerm]);

  useEffect(() => {
    actions.setSearchStats({
      totalMatches,
      currentPageMatches,
      message: state.searchTerm ? `${totalMatches} match${totalMatches === 1 ? '' : 'es'}` : undefined,
    });
  }, [actions, currentPageMatches, state.searchTerm, totalMatches]);

  useEffect(() => {
    const query = state.searchTerm.trim();

    if (!query) {
      lastSearchNavigationTermRef.current = '';
      return;
    }

    if (lastSearchNavigationTermRef.current === query) {
      return;
    }

    lastSearchNavigationTermRef.current = query;
    const firstMatchSlideIndex = slideTexts.findIndex((slideText) => countMatches(slideText, query) > 0);

    if (firstMatchSlideIndex >= 0) {
      actions.setPage(firstMatchSlideIndex + 1);
    }
  }, [actions, slideTexts, state.searchTerm]);

  if (isLoading) {
    return <div className="ldv-renderer-status">Loading presentation...</div>;
  }

  const shellClassName = [
    'ldv-pptx-shell',
    hasThumbnailSidebar ? null : 'ldv-pptx-sidebar-disabled',
    hasThumbnailSidebar && isSidebarCollapsed ? 'ldv-pptx-sidebar-collapsed' : null,
  ]
    .filter(Boolean)
    .join(' ');

  return (
    <div className={shellClassName}>
      {hasThumbnailSidebar ? (
        <aside aria-label="Slides" className="ldv-pptx-sidebar">
          <button
            aria-controls="ldv-pptx-thumbnails"
            aria-expanded={!isSidebarCollapsed}
            aria-label={isSidebarCollapsed ? 'Show slides' : 'Hide slides'}
            className="ldv-pptx-sidebar-toggle"
            onClick={() => setIsSidebarCollapsed((current) => !current)}
            title={isSidebarCollapsed ? 'Show slides' : 'Hide slides'}
            type="button"
          >
            {isSidebarCollapsed ? <PanelLeftOpen aria-hidden="true" size={18} /> : <PanelLeftClose aria-hidden="true" size={18} />}
          </button>
          {!isSidebarCollapsed ? (
            <div className="ldv-pptx-thumbnail-list" id="ldv-pptx-thumbnails" ref={thumbnailListRef}>
              <div
                className="ldv-pptx-thumbnail-virtual-space"
                style={{ height: `${thumbnailVirtualizer.getTotalSize()}px` }}
              >
                {virtualThumbnailSlides.map((virtualSlide) => {
                  const slideIndex = virtualSlide.index;
                  const thumbnailStatus = thumbnailStatuses[slideIndex] ?? 'loading';

                  return (
                    <button
                      data-index={virtualSlide.index}
                      aria-current={state.page === slideIndex + 1 ? 'page' : undefined}
                      aria-label={`Slide ${slideIndex + 1}`}
                      className="ldv-pptx-thumbnail"
                      key={virtualSlide.key}
                      onClick={() => actions.setPage(slideIndex + 1)}
                      ref={(element) => {
                        if (element) {
                          thumbnailVirtualizer.measureElement(element);
                        }
                      }}
                      style={{ transform: `translate(-50%, ${virtualSlide.start}px)` }}
                      type="button"
                    >
                      <span className="ldv-pptx-thumbnail-number">{slideIndex + 1}</span>
                      <span
                        className="ldv-pptx-thumbnail-preview"
                        style={{
                          height: `${thumbnailSize.height}px`,
                          width: `${thumbnailSize.width}px`,
                        }}
                      >
                        <canvas
                          aria-hidden="true"
                          className="ldv-pptx-thumbnail-canvas"
                          ref={(element) => {
                            thumbnailCanvasRefs.current[slideIndex] = element;
                          }}
                          style={{ visibility: thumbnailStatus === 'ready' ? 'visible' : 'hidden' }}
                        />
                        {thumbnailStatus === 'loading' ? (
                          <span className="ldv-thumbnail-loader" aria-hidden="true">
                            <span className="ldv-thumbnail-loader-spinner" />
                          </span>
                        ) : null}
                        {thumbnailStatus === 'error' ? <span className="ldv-thumbnail-error">Unable</span> : null}
                      </span>
                    </button>
                  );
                })}
              </div>
            </div>
          ) : null}
        </aside>
      ) : null}
      <div className="ldv-pptx-main" ref={mainPaneRef}>
        <div className="ldv-pptx-stage">
          <div
            className="ldv-pptx-slide-frame"
            style={{
              height: `${displaySize.height}px`,
              transform: `rotate(${state.rotation}deg)`,
              transformOrigin: 'center top',
              width: `${displaySize.width}px`,
            }}
          >
            <canvas aria-label={`Slide ${state.page}`} className="ldv-pptx-canvas" ref={canvasRef} />
            {currentSearchBoxes.length > 0 ? (
              <div className="ldv-pptx-highlight-layer" aria-hidden="true">
                {currentSearchBoxes.map((box, index) => (
                  <span
                    className="ldv-pptx-search-highlight"
                    key={`${box.x}-${box.y}-${index}`}
                    style={{
                      height: `${(box.height / slideSize.height) * 100}%`,
                      left: `${(box.x / slideSize.width) * 100}%`,
                      top: `${(box.y / slideSize.height) * 100}%`,
                      width: `${(box.width / slideSize.width) * 100}%`,
                    }}
                  />
                ))}
              </div>
            ) : null}
            {isRendering ? <div className="ldv-pptx-render-status">Rendering slide...</div> : null}
          </div>
        </div>
        {warnings.length > 0 ? (
          <div className="ldv-warning" role="status">
            Unsafe embedded PPTX references were removed before preview.
          </div>
        ) : null}
      </div>
    </div>
  );
}

export const PptxRenderer: DocumentRenderer = {
  id: 'pptx',
  label: 'PPTX',
  priority: 24,
  canRender: (file) => {
    if (file.extension === 'pptx') {
      return 2;
    }

    return file.extension !== 'ppt' && file.mimeType === 'application/vnd.openxmlformats-officedocument.presentationml.presentation';
  },
  Component: PptxRendererComponent,
};
