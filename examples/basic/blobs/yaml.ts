import type { BlobSample } from "./types";

const yamlContent = `title: YAML Blob
renderer: local-text
confidential: true
formats:
  - yaml
  - text
security:
  uploads: false
  third_party_preview: false
`;

export const yamlBlobSample: BlobSample = {
  id: "yaml",
  label: "YAML Blob",
  description: "YAML content passed as a local application/yaml Blob.",
  createSource: () => ({
    blob: new Blob([yamlContent], { type: "application/yaml" }),
    fileName: "hardcoded-config.yaml",
  }),
};
