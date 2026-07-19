import type { BlobSample } from "./types";

const xmlContent = `<?xml version="1.0" encoding="UTF-8"?>
<document-viewer>
  <title>XML Blob</title>
  <privacy local="true">No third-party preview service is used.</privacy>
  <format extension="xml" renderer="text" />
</document-viewer>
`;

export const xmlBlobSample: BlobSample = {
  id: "xml",
  label: "XML Blob",
  description: "XML content passed as a local application/xml Blob.",
  createSource: () => ({
    blob: new Blob([xmlContent], { type: "application/xml" }),
    fileName: "hardcoded-data.xml",
  }),
};
