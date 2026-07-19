import JSZip from "jszip";
import type { BlobSample } from "./types";

async function createXlsxBlob(): Promise<Blob> {
  const zip = new JSZip();

  zip.file(
    "[Content_Types].xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`,
  );
  zip.file(
    "_rels/.rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
  );
  zip.file(
    "xl/workbook.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Blob Sheet" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>`,
  );
  zip.file(
    "xl/_rels/workbook.xml.rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`,
  );
  zip.file(
    "xl/worksheets/sheet1.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1:D5"/>
  <sheetData>
    <row r="1">
      <c r="A1" t="str"><v>format</v></c>
      <c r="B1" t="str"><v>source</v></c>
      <c r="C1" t="str"><v>detected_by</v></c>
      <c r="D1" t="str"><v>status</v></c>
    </row>
    <row r="2">
      <c r="A2" t="str"><v>XLSX</v></c>
      <c r="B2" t="str"><v>anonymous Blob</v></c>
      <c r="C2" t="str"><v>zip package sniffing</v></c>
      <c r="D2" t="str"><v>Rendered locally</v></c>
    </row>
    <row r="3">
      <c r="A3" t="str"><v>PDF</v></c>
      <c r="B3" t="str"><v>typed Blob</v></c>
      <c r="C3" t="str"><v>MIME type</v></c>
      <c r="D3" t="str"><v>Rendered locally</v></c>
    </row>
    <row r="4">
      <c r="A4" t="str"><v>CSV</v></c>
      <c r="B4" t="str"><v>Blob with filename</v></c>
      <c r="C4" t="str"><v>fileName prop</v></c>
      <c r="D4" t="str"><v>Rendered locally</v></c>
    </row>
  </sheetData>
</worksheet>`,
  );

  const arrayBuffer = await zip.generateAsync({ type: "arraybuffer" });
  return new Blob([arrayBuffer]);
}

export const xlsxBlobSample: BlobSample = {
  id: "xlsx",
  label: "Anonymous XLSX Blob",
  description:
    "No filename and no MIME type; the viewer should sniff the XLSX package.",
  createSource: () => createXlsxBlob(),
};
