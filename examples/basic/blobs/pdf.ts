import type { BlobSample } from "./types";

function createPdfBlob(): Blob {
  const stream = [
    "BT",
    "/F1 24 Tf",
    "72 720 Td",
    "(Local PDF Blob) Tj",
    "/F1 12 Tf",
    "0 -32 Td",
    "(This PDF was generated from hardcoded PDF objects in the example page.) Tj",
    "0 -24 Td",
    "(The viewer receives the bytes directly as a Blob.) Tj",
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

export const pdfBlobSample: BlobSample = {
  id: "pdf",
  label: "PDF Blob",
  description: "PDF bytes generated from hardcoded PDF objects.",
  createSource: () => createPdfBlob(),
};
