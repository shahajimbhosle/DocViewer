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
const thumbnailWidth = 112;

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
  warnings: string[];
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

  return Array.from(document.getElementsByTagName('*'))
    .filter((node) => node.localName === 't')
    .map((node) => node.textContent ?? '')
    .join(' ')
    .trim();
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

async function preparePptx(arrayBuffer: ArrayBuffer): Promise<PreparedPptx> {
  const parser = new DOMParser();
  const zip = await JSZip.loadAsync(arrayBuffer.slice(0));
  const warnings = await hardenRelationshipFiles(zip, parser);
  const normalizedTextSizing = await normalizeSlideParagraphFontDefaults(zip, parser);
  const slideTexts = await extractPptxSlideTexts(zip, parser);
  const safeArrayBuffer =
    warnings.length > 0 || normalizedTextSizing ? await zip.generateAsync({ type: 'arraybuffer' }) : arrayBuffer.slice(0);

  return {
    arrayBuffer: safeArrayBuffer,
    slideTexts,
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
  const thumbnailCanvasRefs = useRef<Array<HTMLCanvasElement | null>>([]);
  const thumbnailButtonRefs = useRef<Array<HTMLButtonElement | null>>([]);
  const viewerRef = useRef<PPTXViewer | null>(null);
  const thumbnailViewerRef = useRef<PPTXViewer | null>(null);
  const renderIdRef = useRef(0);
  const thumbnailRenderIdRef = useRef(0);
  const mainPaneSize = useElementSize(mainPaneRef);
  const [isLoading, setIsLoading] = useState(true);
  const [isRendering, setIsRendering] = useState(false);
  const [isSidebarCollapsed, setIsSidebarCollapsed] = useState(false);
  const [presentationBuffer, setPresentationBuffer] = useState<ArrayBuffer | null>(null);
  const [slideCount, setSlideCount] = useState(0);
  const [slideSize, setSlideSize] = useState<Size>(defaultSlideSize);
  const [slideTexts, setSlideTexts] = useState<string[]>([]);
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
    setIsRendering(false);
    setPresentationBuffer(null);
    setSlideCount(0);
    setSlideTexts([]);
    setWarnings([]);
    thumbnailCanvasRefs.current = [];
    thumbnailButtonRefs.current = [];
    thumbnailViewerRef.current?.destroy();
    thumbnailViewerRef.current = null;
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
      setSlideTexts(prepared.slideTexts);
      setWarnings(nextWarnings);
      setIsLoading(false);
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
        actions.reportError(error);
      }
    });

    return () => {
      cancelled = true;
      viewer.destroy();
      thumbnailViewerRef.current?.destroy();
      if (viewerRef.current === viewer) {
        viewerRef.current = null;
      }
      thumbnailViewerRef.current = null;
    };
  }, [actions, file]);

  const displaySize = useMemo(
    () => displaySizeForSlide(slideSize, mainPaneSize, state.fitMode, state.zoom),
    [mainPaneSize, slideSize, state.fitMode, state.zoom],
  );
  const thumbnailSize = useMemo(() => thumbnailSizeForSlide(slideSize), [slideSize]);
  const slideIndexes = useMemo(
    () => Array.from({ length: slideCount }, (_, slideIndex) => slideIndex),
    [slideCount],
  );
  const hasThumbnailSidebar = controls.thumbnails && slideCount > 0;

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
      return undefined;
    }

    const thumbnailBuffer: ArrayBuffer = activePresentationBuffer;
    let cancelled = false;
    const renderId = thumbnailRenderIdRef.current + 1;
    thumbnailRenderIdRef.current = renderId;
    const thumbnailViewer = new PPTXViewer({
      autoChartRerenderDelayMs: 0,
      backgroundColor: '#ffffff',
      debug: false,
      enableThumbnails: false,
    });

    thumbnailViewerRef.current?.destroy();
    thumbnailViewerRef.current = thumbnailViewer;

    async function renderThumbnails() {
      await thumbnailViewer.loadFile(thumbnailBuffer.slice(0));

      for (const slideIndex of slideIndexes) {
        if (cancelled || thumbnailRenderIdRef.current !== renderId) {
          return;
        }

        const thumbnailCanvas = thumbnailCanvasRefs.current[slideIndex];
        if (!thumbnailCanvas) {
          continue;
        }

        prepareCanvas(thumbnailCanvas, slideSize);
        await thumbnailViewer.renderSlide(slideIndex, thumbnailCanvas, { quality: 'low' });
        setCanvasDisplaySize(thumbnailCanvas, thumbnailSize);
      }
    }

    renderThumbnails().catch(() => undefined);

    return () => {
      cancelled = true;
      thumbnailViewer.destroy();
      if (thumbnailViewerRef.current === thumbnailViewer) {
        thumbnailViewerRef.current = null;
      }
    };
  }, [hasThumbnailSidebar, isLoading, isSidebarCollapsed, presentationBuffer, slideIndexes, slideSize, thumbnailSize]);

  useEffect(() => {
    if (!hasThumbnailSidebar || isSidebarCollapsed) {
      return;
    }

    thumbnailButtonRefs.current[state.page - 1]?.scrollIntoView({ block: 'nearest' });
  }, [hasThumbnailSidebar, isSidebarCollapsed, state.page]);

  const totalMatches = useMemo(
    () => slideTexts.reduce((total, slideText) => total + countMatches(slideText, state.searchTerm), 0),
    [slideTexts, state.searchTerm],
  );
  const currentPageMatches = useMemo(
    () => countMatches(slideTexts[state.page - 1] ?? '', state.searchTerm),
    [slideTexts, state.page, state.searchTerm],
  );

  useEffect(() => {
    actions.setSearchStats({
      totalMatches,
      currentPageMatches,
      message: state.searchTerm ? `${totalMatches} match${totalMatches === 1 ? '' : 'es'}` : undefined,
    });
  }, [actions, currentPageMatches, state.searchTerm, totalMatches]);

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
            <div className="ldv-pptx-thumbnail-list" id="ldv-pptx-thumbnails">
              {slideIndexes.map((slideIndex) => (
                <button
                  aria-current={state.page === slideIndex + 1 ? 'page' : undefined}
                  aria-label={`Slide ${slideIndex + 1}`}
                  className="ldv-pptx-thumbnail"
                  key={slideIndex}
                  onClick={() => actions.setPage(slideIndex + 1)}
                  ref={(element) => {
                    thumbnailButtonRefs.current[slideIndex] = element;
                  }}
                  type="button"
                >
                  <span className="ldv-pptx-thumbnail-number">{slideIndex + 1}</span>
                  <canvas
                    aria-hidden="true"
                    className="ldv-pptx-thumbnail-canvas"
                    ref={(element) => {
                      thumbnailCanvasRefs.current[slideIndex] = element;
                    }}
                  />
                </button>
              ))}
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
