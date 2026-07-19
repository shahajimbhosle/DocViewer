import JSZip from "jszip";
import type { BlobSample } from "./types";

async function createOdtBlob(): Promise<Blob> {
  const zip = new JSZip();

  zip.file("mimetype", "application/vnd.oasis.opendocument.text", {
    compression: "STORE",
  });
  zip.file(
    "META-INF/manifest.xml",
    `<?xml version="1.0" encoding="UTF-8"?>
<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.2">
  <manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.text"/>
  <manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>
</manifest:manifest>`,
  );
  zip.file(
    "content.xml",
    `<?xml version="1.0" encoding="UTF-8"?>
<office:document-content
  xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
  xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
  xmlns:xlink="http://www.w3.org/1999/xlink"
  xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"
  office:version="1.2">
  <office:automatic-styles>
    <style:style style:name="Heading1" style:family="paragraph">
      <style:text-properties fo:font-size="24pt" fo:font-weight="bold" fo:color="#111827"/>
    </style:style>
    <style:style style:name="Strong" style:family="text">
      <style:text-properties fo:font-weight="bold" fo:color="#0f766e"/>
    </style:style>
    <style:style style:name="Center" style:family="paragraph">
      <style:paragraph-properties fo:text-align="center"/>
    </style:style>
    <style:style style:name="CellHeader" style:family="table-cell">
      <style:table-cell-properties fo:background-color="#edf2f7" fo:border="1px solid #cbd5e1"/>
      <style:text-properties fo:font-weight="bold"/>
    </style:style>
    <style:style style:name="Cell" style:family="table-cell">
      <style:table-cell-properties fo:border="1px solid #cbd5e1"/>
    </style:style>
    <style:page-layout style:name="LetterPage">
      <style:page-layout-properties fo:page-width="8.5in" fo:page-height="11in" fo:margin-top="0.75in" fo:margin-right="0.8in" fo:margin-bottom="0.75in" fo:margin-left="0.8in"/>
    </style:page-layout>
  </office:automatic-styles>
  <office:master-styles>
    <style:master-page style:name="Standard" style:page-layout-name="LetterPage"/>
  </office:master-styles>
  <office:body>
    <office:text>
      <text:h text:outline-level="1" text:style-name="Heading1">Local ODT Blob</text:h>
      <text:p>This OpenDocument Text file is assembled from hardcoded XML and rendered locally from a Blob.</text:p>
      <text:p text:style-name="Center">Centered paragraph with <text:span text:style-name="Strong">styled inline text</text:span>.</text:p>
      <text:p>External links stay clickable in a new unlinked tab: <text:a xlink:href="https://www.npmjs.com/package/@shahajimbhosle/local-doc-viewer">npm package</text:a>.</text:p>
      <text:list>
        <text:list-item><text:p>Headings, paragraphs, spans, lists, and tables are parsed from content.xml.</text:p></text:list-item>
        <text:list-item><text:p>No document bytes are uploaded to a cloud preview service.</text:p></text:list-item>
      </text:list>
      <table:table table:name="Local formats">
        <table:table-row>
          <table:table-cell table:style-name="CellHeader"><text:p>Format</text:p></table:table-cell>
          <table:table-cell table:style-name="CellHeader"><text:p>Status</text:p></table:table-cell>
        </table:table-row>
        <table:table-row>
          <table:table-cell table:style-name="Cell"><text:p>ODT</text:p></table:table-cell>
          <table:table-cell table:style-name="Cell"><text:p>Rendered locally</text:p></table:table-cell>
        </table:table-row>
      </table:table>
    </office:text>
  </office:body>
</office:document-content>`,
  );

  const arrayBuffer = await zip.generateAsync({ type: "arraybuffer" });
  return new Blob([arrayBuffer]);
}

export const odtBlobSample: BlobSample = {
  id: "odt",
  label: "Anonymous ODT Blob",
  description:
    "No filename and no MIME type; the viewer should sniff the ODT package.",
  createSource: () => createOdtBlob(),
};
