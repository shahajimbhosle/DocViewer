import type { BlobSample } from "./types";

const jsonContent = JSON.stringify(
  {
    title: "JSON Blob",
    confidential: true,
    renderer: "local",
    formats: [
      "pdf",
      "docx",
      "odt",
      "rtf",
      "xlsx",
      "ods",
      "pptx",
      "markdown",
      "csv",
      "tsv",
      "image",
      "audio",
      "video",
      "text",
    ],
    generatedAt: "browser-runtime",
  },
  null,
  2,
);

export const jsonBlobSample: BlobSample = {
  id: "json",
  label: "JSON Blob",
  description: "Direct Blob with application/json MIME type.",
  createSource: () => new Blob([jsonContent], { type: "application/json" }),
};
