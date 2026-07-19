import { audioBlobSample } from "./audio";
import { csvBlobSample } from "./csv";
import { docxBlobSample } from "./docx";
import { htmlBlobSample } from "./html";
import { imageBlobSample } from "./image";
import { jsonBlobSample } from "./json";
import { markdownBlobSample } from "./markdown";
import { odsBlobSample } from "./ods";
import { odtBlobSample } from "./odt";
import { pdfBlobSample } from "./pdf";
import { pptxBlobSample } from "./pptx";
import { rtfBlobSample } from "./rtf";
import { textBlobSample } from "./text";
import { tsvBlobSample } from "./tsv";
import { videoBlobSample } from "./video";
import { xlsxBlobSample } from "./xlsx";
import { xmlBlobSample } from "./xml";
import { yamlBlobSample } from "./yaml";
import type { BlobSample } from "./types";

export const blobSamples: BlobSample[] = [
  textBlobSample,
  markdownBlobSample,
  htmlBlobSample,
  jsonBlobSample,
  xmlBlobSample,
  yamlBlobSample,
  csvBlobSample,
  tsvBlobSample,
  rtfBlobSample,
  imageBlobSample,
  audioBlobSample,
  videoBlobSample,
  pdfBlobSample,
  pptxBlobSample,
  xlsxBlobSample,
  odsBlobSample,
  odtBlobSample,
  docxBlobSample,
];
