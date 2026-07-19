import type { BlobSample } from "./types";

const htmlContent = `<!doctype html>
<html>
  <head>
    <title>Local HTML Blob</title>
  </head>
  <body>
    <h1>Local HTML Blob</h1>
    <p>This HTML source is displayed as local text content.</p>
    <script>console.log("This script is text only and is not executed.");</script>
  </body>
</html>
`;

export const htmlBlobSample: BlobSample = {
  id: "html",
  label: "HTML Blob",
  description: "HTML content rendered as local text, without executing scripts.",
  createSource: () => ({
    blob: new Blob([htmlContent], { type: "text/html" }),
    fileName: "hardcoded-page.html",
  }),
};
