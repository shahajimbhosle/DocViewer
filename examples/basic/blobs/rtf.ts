import type { BlobSample } from "./types";

const rtfContent = String.raw`{\rtf1\ansi\deff0\paperw12240\paperh15840\margl1440\margr1440\margt1080\margb1080
{\fonttbl{\f0 Arial;}}
{\colortbl;\red15\green118\blue110;\red220\green38\blue38;}
\f0\fs32 Local RTF Blob\par
\fs24 This RTF file is rendered locally from a Blob.\par
\b Bold text\b0, \i italic text\i0, \ul underlined text\ul0, and \cf1 teal text\cf0.\par
Open package: {\field{\*\fldinst{HYPERLINK "https://www.npmjs.com/package/@shahajimbhosle/local-doc-viewer"}}{\fldrslt{npm package}}}\par
\par
\u8226? Rendered locally\par
\u8226? Links open in a new unlinked tab\par
\u8226? No third-party document preview service\par
\par
\trowd\trgaph108\trleft0
\cellx2200\cellx5600
\b Format\b0\cell\b Status\b0\cell\row
\trowd\trgaph108\trleft0
\cellx2200\cellx5600
RTF\cell Rendered locally\cell\row
}`;

export const rtfBlobSample: BlobSample = {
  id: "rtf",
  label: "RTF Blob",
  description:
    "Rich Text Format content passed as a Blob plus filename metadata.",
  createSource: () => ({
    blob: new Blob([rtfContent], { type: "application/rtf" }),
    fileName: "hardcoded-rich-text.rtf",
  }),
};
