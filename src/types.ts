import type { ComponentType, CSSProperties, ReactNode, RefObject } from 'react';

export type BinaryDocumentData = ArrayBuffer | Uint8Array;
export type DocumentViewerUserSelect = CSSProperties['userSelect'] | boolean;

export type DocumentSource =
  | File
  | Blob
  | BinaryDocumentData
  | string
  | {
      blob?: Blob;
      data?: File | Blob | BinaryDocumentData;
      url?: string;
      fileName?: string;
      mimeType?: string;
    };

export type FitMode = 'manual' | 'width' | 'page';

export interface ResolvedDocument {
  blob: Blob;
  arrayBuffer: ArrayBuffer;
  objectUrl: string;
  fileName: string;
  mimeType: string;
  extension: string;
  byteLength: number;
  sourceKind: 'blob' | 'buffer' | 'url';
  originalUrl?: string;
}

export interface SearchStats {
  totalMatches?: number;
  currentPageMatches?: number;
  isSearching?: boolean;
  message?: string;
}

export interface DocumentInfo {
  title?: string;
  pageCount?: number;
  rendererId?: string;
  rendererLabel?: string;
  warnings?: string[];
}

export interface ViewerRuntimeState {
  page: number;
  pageCount?: number;
  zoom: number;
  rotation: number;
  fitMode: FitMode;
  searchTerm: string;
}

export interface RendererActions {
  setPage: (page: number) => void;
  setPageCount: (pageCount?: number) => void;
  setDocumentInfo: (info: Partial<DocumentInfo>) => void;
  setLoading: (isLoading: boolean) => void;
  setSearchStats: (stats: SearchStats) => void;
  reportError: (error: unknown) => void;
}

export interface DocumentRendererProps {
  file: ResolvedDocument;
  state: ViewerRuntimeState;
  actions: RendererActions;
  controls: ResolvedDocumentViewerControls;
  pdfOptions: ResolvedDocumentViewerPdfOptions;
  labels: DocumentViewerLabels;
  viewportRef: RefObject<HTMLDivElement>;
}

export interface DocumentRenderer {
  id: string;
  label: string;
  priority?: number;
  canRender: (file: ResolvedDocument) => boolean | number;
  Component: ComponentType<DocumentRendererProps>;
}

export interface DocumentViewerControls {
  toolbar?: boolean;
  fileName?: boolean;
  pageNavigation?: boolean;
  zoom?: boolean;
  fit?: boolean;
  rotate?: boolean;
  search?: boolean;
  print?: boolean;
  download?: boolean;
  fullscreen?: boolean;
  thumbnails?: boolean;
}

export type ResolvedDocumentViewerControls = Required<DocumentViewerControls>;

export interface DocumentViewerPdfOptions {
  showThumbnails?: boolean;
}

export type ResolvedDocumentViewerPdfOptions = Required<DocumentViewerPdfOptions>;

export interface DocumentViewerLabels {
  openFile: string;
  previousPage: string;
  nextPage: string;
  page: string;
  of: string;
  zoomOut: string;
  zoomIn: string;
  resetZoom: string;
  fitWidth: string;
  fitPage: string;
  rotateLeft: string;
  rotateRight: string;
  search: string;
  print: string;
  printing: string;
  download: string;
  fullscreen: string;
  exitFullscreen: string;
  unsupportedTitle: string;
  noDocument: string;
  loading: string;
}

export interface DocumentViewerProps {
  source?: DocumentSource | null;
  fileName?: string;
  mimeType?: string;
  className?: string;
  style?: CSSProperties;
  height?: number | string;
  renderers?: DocumentRenderer[];
  allowRemoteUrls?: boolean;
  fetchCredentials?: RequestCredentials;
  minZoom?: number;
  maxZoom?: number;
  initialZoom?: number;
  initialPage?: number;
  controls?: DocumentViewerControls;
  pdfOptions?: DocumentViewerPdfOptions;
  labels?: Partial<DocumentViewerLabels>;
  emptyState?: ReactNode;
  loader?: ReactNode;
  userSelect?: DocumentViewerUserSelect;
  onLoad?: (info: DocumentInfo, file: ResolvedDocument) => void;
  onError?: (error: Error) => void;
}
