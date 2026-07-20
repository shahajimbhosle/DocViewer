# Changelog

All notable changes to `@shahajimbhosle/local-doc-viewer` will be documented in this file.

This project follows semantic versioning. Dates use `YYYY-MM-DD`.

## [1.0.8] - 2026-07-20

### Added

- Added `markdownOptions.allowRemoteImages` so apps can explicitly disable remote Markdown images when they need a stricter no-outbound-resource policy.

### Changed

- Aligned the default privacy model with user-provided content rendering: direct document URLs are allowed by default, while documents are still not uploaded to third-party preview services.
- Render remote Markdown images by default as safe user-provided display content, with lazy loading and `referrerpolicy="no-referrer"`.

### Fixed

- Kept safe remote image sources visible in the Markdown renderer while continuing to remove scripts, event handlers, unsafe URL schemes, and active embedded content.

## [1.0.7] - 2026-07-18

### Added

- Added local ODT rendering for OpenDocument Text files, including common paragraphs, headings, spans, lists, tables, links, styles, and embedded images.
- Added local RTF rendering for Rich Text Format files, including common inline styles, colors, links, bullets, and basic tables.
- Added basic example Blob samples for HTML, XML, YAML, TSV, PNG image, generated audio, generated video, and PPTX.
- Added embedded XLSX drawing previews for local worksheet images and basic charts.

### Changed

- Clarified that CAD drawing files such as `.dwg`/`.dxf` require a custom renderer or local/private conversion.
- Improved Markdown preview styling for headings, inline code, links, lists, blockquotes, code blocks, tables, and images.
- Kept sanitized Markdown links clickable in a new unlinked tab while continuing to remove script-capable links, inline event handlers, and automatic remote resource references.
- Aligned DOCX and ODT links with the same safe new-tab navigation behavior.
- Preserved document-provided page sizes and margins for local ODT and RTF previews.
- Moved basic example Blob generators out of `App.tsx` into separate files under `examples/basic/blobs`.

### Fixed

- Fixed OpenDocument `lr-tb` writing mode so normal ODT text and ODS cells do not render vertically.
- Kept selected PDF/PPTX thumbnail borders above adjacent hover states.
- Rendered basic XLSX charts/graphs that were previously ignored because they live in worksheet drawing parts outside the cell grid.
- Kept DOCX internal reference/bookmark links inside the viewer instead of opening the app domain in a new tab.
- Preserved visible DOCX VML shape geometry and added local fallback rendering for unresolved DOCX images and unsupported DrawingML Word shapes.
- Restored fit width and fit page controls for DOCX previews.
- Normalized common DOCX Symbol/Wingdings bullet glyphs so filled dot bullets render as bullets instead of square boxes.
- Normalized DOCX nested hollow circle bullets that Word stores as Courier New `o` markers.

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
