import JSZip from "jszip";
import { useEffect, useMemo, useState } from "react";
import { DocumentViewer } from "../../src";
import type { DocumentSource } from "../../src";

interface BlobSample {
  id: string;
  label: string;
  description: string;
  createSource: () => Promise<DocumentSource> | DocumentSource;
}

const textContent = `Local Document Viewer

This sample is a Blob, not a File. The bytes are created in the example page and passed directly to the React component.

Nothing is uploaded and no third-party preview service is used.`;

const markdownContent = `# Markdown Blob

This markdown document is hardcoded in the example page.

- Rendered from a \`Blob\`
- Sanitized locally
- Searchable through the viewer toolbar

| Format | Source |
| --- | --- |
| Markdown | Blob |
| Storage | Browser memory |
`;

const csvContent = `question,selected_answer,correct_answer,time,result,explanation
"<p>Testv 2</p>",B,D,0s,Wrong,ddddd
"True or False: The sum of probabilities of all possible outcomes of a random experiment is always less than 1.",False,False,0s,Correct,"The sum of probabilities of all possible outcomes of a random experiment is always equal to 1."
"A box contains 10 defective and 90 non-defective items. If one item is selected at random, what is the probability that it is non-defective?",1/9,9/10,0s,Wrong,"The probability of selecting a non-defective item is 90/100 = 9/10."
`;

const jsonContent = JSON.stringify(
  {
    title: "JSON Blob",
    confidential: true,
    renderer: "local",
    formats: ["pdf", "docx", "xlsx", "markdown", "csv", "text"],
    generatedAt: "browser-runtime",
  },
  null,
  2,
);

function createPdfBlob(): Blob {
  const stream = [
    "BT",
    "/F1 24 Tf",
    "72 720 Td",
    "(Local PDF Blob) Tj",
    "/F1 12 Tf",
    "0 -32 Td",
    "(This PDF was generated from hardcoded strings in the example page.) Tj",
    "ET",
  ].join("\n");
  const objects = [
    "<< /Type /Catalog /Pages 2 0 R >>",
    "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
    "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources << /Font << /F1 4 0 R >> >> /Contents 5 0 R >>",
    "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
    `<< /Length ${stream.length} >>\nstream\n${stream}\nendstream`,
  ];
  const offsets: number[] = [];
  let pdf = "%PDF-1.4\n";

  objects.forEach((object, index) => {
    offsets.push(pdf.length);
    pdf += `${index + 1} 0 obj\n${object}\nendobj\n`;
  });

  const xrefStart = pdf.length;
  pdf += `xref\n0 ${objects.length + 1}\n`;
  pdf += "0000000000 65535 f \n";
  offsets.forEach((offset) => {
    pdf += `${String(offset).padStart(10, "0")} 00000 n \n`;
  });
  pdf += `trailer\n<< /Size ${objects.length + 1} /Root 1 0 R >>\nstartxref\n${xrefStart}\n%%EOF\n`;

  return new Blob([pdf], { type: "application/pdf" });
}

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

async function createDocxBlob(): Promise<Blob> {
  const zip = new JSZip();

  zip.file(
    "[Content_Types].xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>`,
  );
  zip.file(
    "_rels/.rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`,
  );
  zip.file(
    "word/styles.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/>
    <w:qFormat/>
  </w:style>
</w:styles>`,
  );
  zip.file(
    "word/document.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r>
        <w:rPr><w:b/><w:sz w:val="32"/></w:rPr>
        <w:t>Local DOCX Blob</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r><w:t>This Word document package is assembled from hardcoded XML strings in the example page.</w:t></w:r>
    </w:p>
    <w:p>
      <w:r><w:t>The viewer receives it as a Blob and renders it without uploading the document.</w:t></w:r>
    </w:p>
    <w:sectPr>
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>
    </w:sectPr>
  </w:body>
</w:document>`,
  );

  const arrayBuffer = await zip.generateAsync({ type: "arraybuffer" });
  return new Blob([arrayBuffer]);
}

const blobSamples: BlobSample[] = [
  {
    id: "text",
    label: "Text Blob",
    description: "Direct Blob with text/plain MIME type.",
    createSource: () => new Blob([textContent], { type: "text/plain" }),
  },
  {
    id: "markdown",
    label: "Markdown Blob",
    description:
      "Blob object with a filename so the markdown renderer is selected.",
    createSource: () => ({
      blob: new Blob([markdownContent], { type: "text/markdown" }),
      fileName: "hardcoded-notes.md",
    }),
  },
  {
    id: "csv",
    label: "CSV Blob",
    description: "CSV content passed as a Blob plus filename metadata.",
    createSource: () => ({
      blob: new Blob([csvContent], { type: "text/csv" }),
      fileName: "mid-term-report.csv",
    }),
  },
  {
    id: "json",
    label: "JSON Blob",
    description: "Direct Blob with application/json MIME type.",
    createSource: () => new Blob([jsonContent], { type: "application/json" }),
  },
  {
    id: "pdf",
    label: "PDF Blob",
    description: "PDF bytes generated from hardcoded PDF objects.",
    createSource: () => createPdfBlob(),
  },
  {
    id: "xlsx",
    label: "Anonymous XLSX Blob",
    description:
      "No filename and no MIME type; the viewer should sniff the XLSX package.",
    createSource: () => createXlsxBlob(),
  },
  {
    id: "docx",
    label: "DOCX Blob",
    description:
      "Word package generated from hardcoded XML and passed as a Blob.",
    createSource: async () => ({
      blob: await createDocxBlob(),
      fileName: "hardcoded-blob.docx",
    }),
  },
];

export function App() {
  const [file, setFile] = useState<File | null>(null);
  const [selectedSampleId, setSelectedSampleId] = useState(blobSamples[0].id);
  const [sampleSource, setSampleSource] = useState<DocumentSource | null>(null);
  const [sampleError, setSampleError] = useState<Error | null>(null);
  const selectedSample = useMemo(
    () =>
      blobSamples.find((sample) => sample.id === selectedSampleId) ??
      blobSamples[0],
    [selectedSampleId],
  );

  useEffect(() => {
    let cancelled = false;

    setSampleSource(null);
    setSampleError(null);

    Promise.resolve(selectedSample.createSource()).then(
      (source) => {
        if (!cancelled) {
          setSampleSource(source);
        }
      },
      (error: unknown) => {
        if (!cancelled) {
          setSampleError(
            error instanceof Error
              ? error
              : new Error("Unable to create Blob sample."),
          );
        }
      },
    );

    return () => {
      cancelled = true;
    };
  }, [selectedSample]);

  const source = file ?? sampleSource;

  return (
    <main className="example-shell">
      <header className="example-header">
        <div>
          <h1>Local Document Viewer</h1>
          <p>
            Test hardcoded Blob documents or pick a file from this machine. The
            browser passes the data directly to the React component.
          </p>
        </div>
        <label className="example-file-button">
          <input
            onChange={(event) => {
              setFile(event.target.files?.[0] ?? null);
            }}
            type="file"
          />
          Select file
        </label>
      </header>

      <section className="example-samples" aria-label="Hardcoded Blob samples">
        <div className="example-sample-buttons">
          {blobSamples.map((sample) => (
            <button
              aria-pressed={!file && selectedSampleId === sample.id}
              className="example-sample-button"
              key={sample.id}
              onClick={() => {
                setFile(null);
                setSelectedSampleId(sample.id);
              }}
              type="button"
            >
              {sample.label}
            </button>
          ))}
        </div>
        <p>
          {file
            ? `Viewing uploaded file: ${file.name}`
            : selectedSample.description}
        </p>
      </section>

      {sampleError ? (
        <div className="example-error" role="alert">
          {sampleError.message}
        </div>
      ) : (
        <DocumentViewer
          height="calc(100vh - 204px)"
          onError={(error) => {
            console.error(error);
          }}
          source={source}
        />
      )}
    </main>
  );
}
