import type { BlobSample } from "./types";

const markdownContent = `# Markdown Blob

This markdown document is hardcoded in the example page.

- Rendered from a \`Blob\`
- Sanitized locally
- Searchable through the viewer toolbar

| Format | Source |
| --- | --- |
| Markdown | Blob |
| Storage | Browser memory |

[Open the package page](https://www.npmjs.com/package/@shahajimbhosle/local-doc-viewer)
`;

export const markdownBlobSample: BlobSample = {
  id: "markdown",
  label: "Markdown Blob",
  description:
    "Blob object with a filename so the markdown renderer is selected.",
  createSource: () => ({
    blob: new Blob([markdownContent], { type: "text/markdown" }),
    fileName: "hardcoded-notes.md",
  }),
};
