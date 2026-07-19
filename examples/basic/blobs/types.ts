import type { DocumentSource } from "../../../src";

export interface BlobSample {
  id: string;
  label: string;
  description: string;
  createSource: () => Promise<DocumentSource> | DocumentSource;
}
