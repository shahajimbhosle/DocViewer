import {
  ChevronLeft,
  ChevronRight,
  Download,
  FileText,
  Maximize2,
  Minimize2,
  MoveHorizontal,
  Printer,
  RefreshCcw,
  RotateCcw,
  RotateCw,
  Scan,
  Search,
  ZoomIn,
  ZoomOut,
} from 'lucide-react';
import type { ReactNode } from 'react';
import type {
  DocumentViewerLabels,
  FitMode,
  ResolvedDocumentViewerControls,
  SearchStats,
  ViewerRuntimeState,
} from '../types';

interface ToolbarProps {
  controls: ResolvedDocumentViewerControls;
  fileName?: string;
  isFullscreen: boolean;
  isPrinting: boolean;
  labels: DocumentViewerLabels;
  searchStats: SearchStats;
  state: ViewerRuntimeState;
  onDownload: () => void;
  onFitModeChange: (fitMode: FitMode) => void;
  onFullscreenToggle: () => void;
  onPageChange: (page: number) => void;
  onPrint: () => void;
  onRotateLeft: () => void;
  onRotateRight: () => void;
  onSearchChange: (value: string) => void;
  onZoomIn: () => void;
  onZoomOut: () => void;
  onZoomReset: () => void;
}

interface IconButtonProps {
  label: string;
  busy?: boolean;
  disabled?: boolean;
  pressed?: boolean;
  onClick: () => void;
  children: ReactNode;
}

function IconButton({ busy, children, disabled, label, onClick, pressed }: IconButtonProps) {
  return (
    <button
      aria-busy={busy || undefined}
      aria-label={label}
      aria-pressed={pressed}
      className="ldv-icon-button"
      disabled={disabled || busy}
      onClick={onClick}
      title={label}
      type="button"
    >
      {busy ? <span className="ldv-icon-button-spinner" aria-hidden="true" /> : children}
    </button>
  );
}

export function Toolbar({
  controls,
  fileName,
  isFullscreen,
  isPrinting,
  labels,
  searchStats,
  state,
  onDownload,
  onFitModeChange,
  onFullscreenToggle,
  onPageChange,
  onPrint,
  onRotateLeft,
  onRotateRight,
  onSearchChange,
  onZoomIn,
  onZoomOut,
  onZoomReset,
}: ToolbarProps) {
  const zoomPercent = Math.round(state.zoom * 100);
  const hasPages = typeof state.pageCount === 'number' && state.pageCount > 0;

  return (
    <div className="ldv-toolbar" role="toolbar">
      {controls.fileName ? (
        <div className="ldv-file-name" title={fileName}>
          <FileText aria-hidden="true" size={18} />
          <span>{fileName || labels.noDocument}</span>
        </div>
      ) : null}

      {controls.pageNavigation && hasPages ? (
        <div className="ldv-control-group">
          <IconButton disabled={state.page <= 1} label={labels.previousPage} onClick={() => onPageChange(state.page - 1)}>
            <ChevronLeft aria-hidden="true" size={18} />
          </IconButton>
          <label className="ldv-page-control">
            <span>{labels.page}</span>
            <input
              aria-label={labels.page}
              max={state.pageCount}
              min={1}
              onChange={(event) => onPageChange(Number(event.target.value))}
              type="number"
              value={state.page}
            />
            <span>
              {labels.of} {state.pageCount}
            </span>
          </label>
          <IconButton
            disabled={state.page >= (state.pageCount ?? state.page)}
            label={labels.nextPage}
            onClick={() => onPageChange(state.page + 1)}
          >
            <ChevronRight aria-hidden="true" size={18} />
          </IconButton>
        </div>
      ) : null}

      {controls.zoom ? (
        <div className="ldv-control-group">
          <IconButton label={labels.zoomOut} onClick={onZoomOut}>
            <ZoomOut aria-hidden="true" size={18} />
          </IconButton>
          <button className="ldv-zoom-reset" onClick={onZoomReset} title={labels.resetZoom} type="button">
            <RefreshCcw aria-hidden="true" size={15} />
            <span>{zoomPercent}%</span>
          </button>
          <IconButton label={labels.zoomIn} onClick={onZoomIn}>
            <ZoomIn aria-hidden="true" size={18} />
          </IconButton>
        </div>
      ) : null}

      {controls.fit ? (
        <div className="ldv-control-group">
          <IconButton
            label={labels.fitWidth}
            onClick={() => onFitModeChange(state.fitMode === 'width' ? 'manual' : 'width')}
            pressed={state.fitMode === 'width'}
          >
            <MoveHorizontal aria-hidden="true" size={18} />
          </IconButton>
          <IconButton
            label={labels.fitPage}
            onClick={() => onFitModeChange(state.fitMode === 'page' ? 'manual' : 'page')}
            pressed={state.fitMode === 'page'}
          >
            <Scan aria-hidden="true" size={18} />
          </IconButton>
        </div>
      ) : null}

      {controls.rotate ? (
        <div className="ldv-control-group">
          <IconButton label={labels.rotateLeft} onClick={onRotateLeft}>
            <RotateCcw aria-hidden="true" size={18} />
          </IconButton>
          <IconButton label={labels.rotateRight} onClick={onRotateRight}>
            <RotateCw aria-hidden="true" size={18} />
          </IconButton>
        </div>
      ) : null}

      {controls.search ? (
        <label className="ldv-search-control">
          <Search aria-hidden="true" size={17} />
          <input
            aria-label={labels.search}
            onChange={(event) => onSearchChange(event.target.value)}
            placeholder={labels.search}
            type="search"
            value={state.searchTerm}
          />
          {searchStats.message ? <span className="ldv-search-status">{searchStats.message}</span> : null}
        </label>
      ) : null}

      <div className="ldv-toolbar-spacer" />

      {controls.print ? (
        <IconButton busy={isPrinting} label={isPrinting ? labels.printing : labels.print} onClick={onPrint}>
          <Printer aria-hidden="true" size={18} />
        </IconButton>
      ) : null}

      {controls.download ? (
        <IconButton label={labels.download} onClick={onDownload}>
          <Download aria-hidden="true" size={18} />
        </IconButton>
      ) : null}

      {controls.fullscreen ? (
        <IconButton label={isFullscreen ? labels.exitFullscreen : labels.fullscreen} onClick={onFullscreenToggle}>
          {isFullscreen ? <Minimize2 aria-hidden="true" size={18} /> : <Maximize2 aria-hidden="true" size={18} />}
        </IconButton>
      ) : null}
    </div>
  );
}
