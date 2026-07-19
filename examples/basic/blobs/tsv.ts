import type { BlobSample } from "./types";

const tsvContent = `format\tsource\tdetected_by\tstatus
TSV\tBlob with filename\tfileName prop\tRendered locally
CSV renderer\tTab delimiter\t.tsv extension\tParsed locally
`;

export const tsvBlobSample: BlobSample = {
  id: "tsv",
  label: "TSV Blob",
  description: "TSV content passed as a Blob plus .tsv filename metadata.",
  createSource: () => ({
    blob: new Blob([tsvContent], { type: "text/tab-separated-values" }),
    fileName: "hardcoded-table.tsv",
  }),
};
