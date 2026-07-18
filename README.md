# Local Document Viewer for React

A privacy-first React document viewer that renders files in the browser without uploading them to Google Docs, Office Online, a SaaS preview API, or any third-party domain.

The package is designed for confidential documents. It accepts `File`, `Blob`, `ArrayBuffer`, `Uint8Array`, `blob:` URLs, and `data:` URLs. `http://` and `https://` document URLs are rejected by default.

## Links

- npm: [@shahajimbhosle/local-doc-viewer](https://www.npmjs.com/package/@shahajimbhosle/local-doc-viewer)
- GitHub: [shahajimbhosle/DocViewer](https://github.com/shahajimbhosle/DocViewer)
- Example: [GitHub Pages demo](https://shahajimbhosle.github.io/DocViewer/)
- Issues: [GitHub issues](https://github.com/shahajimbhosle/DocViewer/issues)

## Built-In Formats

- PDF, rendered with a bundled PDF.js worker
- Images: PNG, JPG, GIF, WebP, BMP, TIFF browser support permitting
- Text/code: TXT, JSON, XML, YAML, HTML as text/sanitized content, logs, JS/TS/CSS
- Markdown, rendered locally and sanitized
- CSV/TSV
- DOCX, rendered locally with Word-style layout and embedded styles where browser rendering permits
- XLS/XLSX/ODS, parsed locally in the browser
- PPTX, rendered locally to canvas with slide navigation and text search counts
- DOCX comments and XLSX/ODS notes/comments are shown locally when present
- Audio/video formats supported by the browser

Spreadsheet previews use row and column virtualization so larger XLS/XLSX/ODS files do not mount every grid cell into the DOM at once.

Legacy binary Word/PowerPoint documents, OpenDocument text/presentation files, strict Office fidelity edge cases, encrypted documents, CAD files, and uncommon enterprise formats need a custom renderer or a private/on-prem conversion service. The viewer exposes a renderer registry for that purpose.

### About `.ppt`

Legacy `.ppt` files are not the same format as `.pptx`. The built-in presentation renderer handles `.pptx` in-browser. For `.ppt`, use a local/private converter such as your own LibreOffice or Microsoft Office automation service to produce PDF or `.pptx`, then pass the converted file into this viewer through a custom renderer.

## Install

```bash
npm install @shahajimbhosle/local-doc-viewer
```

Import the component and CSS:

```tsx
import { DocumentViewer } from '@shahajimbhosle/local-doc-viewer';
import '@shahajimbhosle/local-doc-viewer/style.css';

export function SecurePreview({ file }: { file: File }) {
  return <DocumentViewer source={file} height={720} />;
}
```

Blob sources are supported directly. If the Blob has no reliable `type`, pass a filename or MIME type; otherwise the viewer will sniff common local formats such as PDF, DOCX, XLSX, ODS, and PPTX from the bytes.

```tsx
<DocumentViewer source={blob} />

<DocumentViewer
  source={{
    blob,
    fileName: 'mid-term-report.xlsx',
  }}
/>
```

## Loading State

The viewer shows a compact built-in loader while it is resolving the local file, Blob, buffer, or allowed private URL. Built-in async renderers also use it while preparing heavier formats such as PDF, DOCX, XLSX/ODS, Markdown, and PPTX.

Provide `loader` to replace the default loading UI:

```tsx
<DocumentViewer
  loader={<div className="my-loader">Opening secure preview...</div>}
  source={file}
/>
```

## Privacy Model

By default, the viewer:

- Does not upload files
- Does not call third-party document preview services
- Rejects `http://` and `https://` document URLs
- Uses local `Blob` URLs for preview/download/print flows
- Sanitizes generated HTML and removes remote resource references from rendered HTML
- Removes external PPTX relationships before rendering so linked resources are not fetched
- Uses a packaged PDF.js worker instead of a CDN worker

If your app serves confidential documents from its own trusted private endpoint, opt in explicitly:

```tsx
<DocumentViewer
  allowRemoteUrls
  fetchCredentials="include"
  source={{
    url: '/api/private-documents/123',
    fileName: 'board-pack.pdf',
    mimeType: 'application/pdf',
  }}
/>
```

For highly sensitive deployments, pair this with a strict Content Security Policy that blocks outbound connections except your own origin.

### PDF Worker

PDF rendering works locally without a CDN or third-party worker URL. By default, the viewer registers PDF.js' packaged worker module before opening a PDF, which avoids Vite resolving the worker from `/node_modules/.vite/deps`.

The worker is also published as a separate package asset for advanced setups:

```text
@shahajimbhosle/local-doc-viewer/pdf.worker.min.mjs
```

If your application wants to provide its own worker URL, configure it once before rendering PDFs:

```tsx
import { configurePdfWorker } from '@shahajimbhosle/local-doc-viewer';
import pdfWorkerUrl from '@shahajimbhosle/local-doc-viewer/pdf.worker.min.mjs?url';

configurePdfWorker(pdfWorkerUrl);
```

The worker is intentionally not embedded as an inline base64 JavaScript data URL in the library bundle, because large inline executable payloads can trigger package malware scanners even when the code comes from PDF.js.

## Controls

The default control surface includes:

- File name
- Page/slide navigation where available
- Zoom in/out/reset
- Fit width and fit page
- PDF rotate left/right
- Search where the renderer can extract text
- Print
- Download
- Fullscreen
- PPTX slide thumbnails
- Optional PDF page thumbnails

Controls are also filtered by renderer capability. For example, media files hide zoom, fit, search, print, page navigation, and rotate; images hide search, page navigation, and rotate; text/table renderers hide fit/page navigation/rotate; legacy or unsupported files keep only the generic actions such as file name, download, and fullscreen.

For PDFs, hiding page navigation switches the document surface to continuous scrolling so every page remains reachable without toolbar page buttons:

```tsx
<DocumentViewer
  controls={{ pageNavigation: false }}
  source={file}
/>
```

Show PDF page thumbnails in a collapsible left drawer:

```tsx
<DocumentViewer
  pdfOptions={{ showThumbnails: true }}
  source={file}
/>
```

PDF and PPTX thumbnail drawers are virtualized. Only visible thumbnails are mounted and rendered, and each thumbnail shows its own loader while it is prepared. This keeps the active PDF page or slide usable while thumbnails continue loading.

Set `controls.thumbnails` to `false` to hide thumbnail drawers for both PPTX and PDF.

Disable controls selectively:

```tsx
<DocumentViewer
  controls={{
    download: false,
    print: false,
    thumbnails: false,
  }}
  source={file}
/>
```

Every `controls` key is optional. Omitted controls use the default behavior.

## Document-Specific Options

Use document-specific option objects for behavior that only applies to one renderer type:

```tsx
<DocumentViewer
  pdfOptions={{ showThumbnails: true }}
  source={file}
/>
```

## Text Selection

Document text is selectable by default where the browser renderer supports it. Use `userSelect` to control selection inside the document viewport without affecting toolbar controls:

```tsx
<DocumentViewer source={file} userSelect="none" />
<DocumentViewer source={file} userSelect="text" />
```

Boolean shorthand is also supported:

```tsx
<DocumentViewer source={file} userSelect={false} />
```

## Custom Renderers

Register a renderer for regulated formats or for a local conversion flow:

```tsx
import type { DocumentRenderer } from '@shahajimbhosle/local-doc-viewer';

const PrivateLegacyPptRenderer: DocumentRenderer = {
  id: 'private-legacy-ppt',
  label: 'Private legacy PPT',
  priority: 100,
  canRender: (file) => file.extension === 'ppt',
  Component: ({ file }) => {
    return <div>Render {file.fileName} with your local renderer here.</div>;
  },
};

<DocumentViewer renderers={[PrivateLegacyPptRenderer]} source={file} />;
```

## Local Development

```bash
npm install
npm run dev
```

Open the printed local URL and choose a file from disk. The example passes the browser `File` object directly to the viewer.

Build the package:

```bash
npm run build
```

Build the GitHub Pages example:

```bash
npm run build:pages
npm run preview:pages
```

The example site is built from `examples/basic` into `dist-pages` and is deployed by `.github/workflows/pages.yml`. In GitHub, open the repository settings and set **Pages > Build and deployment > Source** to **GitHub Actions**. After that, every push to `main` deploys the example to:

```text
https://shahajimbhosle.github.io/DocViewer/
```

## License

This package is released under the MIT License.

MIT is a permissive public license. It allows users to use, copy, modify, merge, publish, distribute, sublicense, and sell copies of the software, as long as the copyright notice and license text are included in substantial copies of the software.

## Disclaimer

This software is provided "as is", without warranty of any kind. The authors, contributors, and package publisher are not responsible or liable for any direct, indirect, incidental, special, consequential, or other damages, data loss, security incidents, compliance issues, business interruption, or any other claims arising from use, misuse, inability to use, or distribution of this package.

You are responsible for validating whether this package is suitable for your security, privacy, legal, regulatory, browser-support, and document-fidelity requirements before using it in any environment.
