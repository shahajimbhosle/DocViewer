import JSZip from "jszip";
import type { BlobSample } from "./types";

async function createOdsBlob(): Promise<Blob> {
  const zip = new JSZip();

  zip.file("mimetype", "application/vnd.oasis.opendocument.spreadsheet", {
    compression: "STORE",
  });
  zip.file(
    "META-INF/manifest.xml",
    `<?xml version="1.0" encoding="UTF-8"?>
<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.2">
  <manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.spreadsheet"/>
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
  xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"
  office:version="1.2">
  <office:automatic-styles>
    <style:style style:name="co1" style:family="table-column"><style:table-column-properties style:column-width="1in"/></style:style>
    <style:style style:name="co2" style:family="table-column"><style:table-column-properties style:column-width="1.35in"/></style:style>
    <style:style style:name="co3" style:family="table-column"><style:table-column-properties style:column-width="2in"/></style:style>
    <style:style style:name="co4" style:family="table-column"><style:table-column-properties style:column-width="3.6in"/></style:style>
    <style:style style:name="roTall" style:family="table-row"><style:table-row-properties style:row-height="0.9in"/></style:style>
    <style:style style:name="header" style:family="table-cell">
      <style:table-cell-properties fo:background-color="#EDF2F7" style:vertical-align="bottom"/>
      <style:paragraph-properties fo:text-align="center"/>
      <style:text-properties fo:font-weight="bold"/>
    </style:style>
    <style:style style:name="wrap" style:family="table-cell">
      <style:table-cell-properties fo:wrap-option="wrap" style:vertical-align="top"/>
    </style:style>
    <style:style style:name="red" style:family="table-cell">
      <style:text-properties fo:color="#FF0000"/>
    </style:style>
  </office:automatic-styles>
  <office:body>
    <office:spreadsheet>
      <table:table table:name="ODS Sheet">
        <table:table-column table:style-name="co1"/>
        <table:table-column table:style-name="co2"/>
        <table:table-column table:style-name="co3"/>
        <table:table-column table:style-name="co4"/>
        <table:table-row>
          <table:table-cell table:style-name="header" office:value-type="string"><text:p>Id</text:p></table:table-cell>
          <table:table-cell table:style-name="header" office:value-type="string"><text:p>Date</text:p></table:table-cell>
          <table:table-cell table:style-name="header" office:value-type="string"><text:p>Title</text:p></table:table-cell>
          <table:table-cell table:style-name="header" office:value-type="string"><text:p>Note</text:p></table:table-cell>
        </table:table-row>
        <table:table-row>
          <table:table-cell office:value-type="float" office:value="1"><text:p>1</text:p></table:table-cell>
          <table:table-cell office:value-type="date" office:date-value="2026-07-18"><text:p>7/18/2026</text:p></table:table-cell>
          <table:table-cell office:value-type="string"><text:p>ODS support</text:p></table:table-cell>
          <table:table-cell office:value-type="string"><text:p>Parsed locally from content.xml</text:p></table:table-cell>
        </table:table-row>
        <table:table-row table:style-name="roTall">
          <table:table-cell office:value-type="float" office:value="2"><text:p>2</text:p></table:table-cell>
          <table:table-cell table:style-name="red" office:value-type="string"><text:p>Styled</text:p></table:table-cell>
          <table:table-cell office:value-type="string"><text:p>Wrapped text</text:p></table:table-cell>
          <table:table-cell table:style-name="wrap" office:value-type="string"><text:p>This ODS blob includes wrapped multiline text.</text:p><text:p>It should render inside the spreadsheet grid without using any cloud service.</text:p></table:table-cell>
        </table:table-row>
      </table:table>
      <table:table table:name="Second Sheet">
        <table:table-row>
          <table:table-cell office:value-type="string"><text:p>Second sheet content</text:p></table:table-cell>
        </table:table-row>
      </table:table>
    </office:spreadsheet>
  </office:body>
</office:document-content>`,
  );

  const arrayBuffer = await zip.generateAsync({ type: "arraybuffer" });
  return new Blob([arrayBuffer]);
}

export const odsBlobSample: BlobSample = {
  id: "ods",
  label: "Anonymous ODS Blob",
  description:
    "No filename and no MIME type; the viewer should sniff the ODS package.",
  createSource: () => createOdsBlob(),
};
