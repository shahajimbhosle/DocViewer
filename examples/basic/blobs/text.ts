import type { BlobSample } from "./types";

const textContent = `Local Document Viewer

This sample is a Blob, not a File. The bytes are created in the example page and passed directly to the React component.

Nothing is uploaded and no third-party preview service is used.`;

export const textBlobSample: BlobSample = {
  id: "text",
  label: "Text Blob",
  description: "Direct Blob with text/plain MIME type.",
  createSource: () => new Blob([textContent], { type: "text/plain" }),
};
