# Changelog

All notable changes to `@shahajimbhosle/local-doc-viewer` will be documented in this file.

This project follows semantic versioning. Dates use `YYYY-MM-DD`.

## [1.0.6] - 2026-07-18

### Added

- Added an inline password prompt for encrypted PDF files. Passwords are passed directly to PDF.js in the browser and are not uploaded.

## [1.0.5] - 2026-07-18

### Added

- Added `pdfOptions.showThumbnails` for PDF-specific thumbnail configuration.
- Added a virtualized, collapsible PDF page thumbnail drawer.
- Added continuous PDF scrolling when `controls.pageNavigation` is disabled, so users can still reach every page by scrolling.
- Added virtualized PPTX slide thumbnails with per-thumbnail loading states.
- Added search highlight overlays for PDF pages and PPTX slides.
- Added print preparation state so the print button can show progress while the browser print frame is prepared.

### Changed

- Moved document-type-specific thumbnail settings out of generic controls and into `pdfOptions` for PDF behavior.
- Kept `controls.thumbnails` as the shared switch for showing or hiding supported thumbnail drawers.
- Improved print handling by preparing a hidden local print frame instead of relying on an immediate popup flow.
- Centered PPTX slide rendering inside the viewer surface.

### Fixed

- Reflected spreadsheet cell borders from XLSX and ODS workbook styles.
- Prevented PDF thumbnail rendering from blocking the active PDF page render.
- Prevented PPTX thumbnail rendering from blocking the active slide render.
- Preserved the stable spreadsheet renderer by removing the experimental spreadsheet overflow mode path before release.

## [1.0.0] - 2026-07-17

### Added

- Initial public release of the privacy-first local React document viewer.
- Added local-only rendering for PDF, DOCX, XLSX, XLS, ODS, PPTX, images, video/audio, Markdown, CSV, JSON, and plain text where supported by browser-side renderers.
- Added built-in viewer controls for file name, page navigation, zoom, fit, rotation, search, print, download, fullscreen, and thumbnails where appropriate.
- Added Blob, File, ArrayBuffer, URL, and metadata object source support.
- Added TypeScript declarations, bundled CSS exports, and public npm package metadata.
