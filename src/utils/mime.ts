export const mimeByExtension: Record<string, string> = {
  avi: 'video/x-msvideo',
  bmp: 'image/bmp',
  csv: 'text/csv',
  doc: 'application/msword',
  docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  gif: 'image/gif',
  htm: 'text/html',
  html: 'text/html',
  jpeg: 'image/jpeg',
  jpg: 'image/jpeg',
  json: 'application/json',
  log: 'text/plain',
  md: 'text/markdown',
  mov: 'video/quicktime',
  mp3: 'audio/mpeg',
  mp4: 'video/mp4',
  odp: 'application/vnd.oasis.opendocument.presentation',
  ods: 'application/vnd.oasis.opendocument.spreadsheet',
  odt: 'application/vnd.oasis.opendocument.text',
  pdf: 'application/pdf',
  png: 'image/png',
  ppt: 'application/vnd.ms-powerpoint',
  pptx: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
  rtf: 'application/rtf',
  svg: 'image/svg+xml',
  tif: 'image/tiff',
  tiff: 'image/tiff',
  txt: 'text/plain',
  webm: 'video/webm',
  webp: 'image/webp',
  xls: 'application/vnd.ms-excel',
  xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  xml: 'application/xml',
  yaml: 'application/yaml',
  yml: 'application/yaml',
};

export function extensionFromName(fileName = ''): string {
  const cleanName = fileName.split(/[?#]/)[0] ?? '';
  const lastDot = cleanName.lastIndexOf('.');
  if (lastDot < 0 || lastDot === cleanName.length - 1) {
    return '';
  }

  return cleanName.slice(lastDot + 1).toLowerCase();
}

export function fileNameFromUrl(url: string): string {
  try {
    const parsed = new URL(url, globalThis.location?.href);
    const name = decodeURIComponent(parsed.pathname.split('/').pop() ?? '');
    return name || 'document';
  } catch {
    return url.split('/').pop()?.split(/[?#]/)[0] || 'document';
  }
}

export function inferMimeType(fileName?: string, fallback = 'application/octet-stream'): string {
  const extension = extensionFromName(fileName);
  return (extension && mimeByExtension[extension]) || fallback;
}

export function extensionFromMimeType(mimeType = ''): string {
  const cleanMimeType = mimeType.split(';')[0].trim().toLowerCase();
  const entry = Object.entries(mimeByExtension).find(([, candidateMimeType]) => candidateMimeType === cleanMimeType);

  return entry?.[0] ?? '';
}

export function isTextLikeMime(mimeType: string): boolean {
  return (
    mimeType.startsWith('text/') ||
    mimeType === 'application/json' ||
    mimeType === 'application/xml' ||
    mimeType === 'application/yaml' ||
    mimeType === 'image/svg+xml'
  );
}

export function isOfficeExtension(extension: string): boolean {
  return ['doc', 'docx', 'xls', 'xlsx', 'ppt', 'pptx', 'odt', 'ods', 'odp', 'rtf'].includes(extension);
}
