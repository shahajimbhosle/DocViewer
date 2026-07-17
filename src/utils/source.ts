import JSZip from 'jszip';
import type { DocumentSource, ResolvedDocument } from '../types';
import { extensionFromMimeType, extensionFromName, fileNameFromUrl, inferMimeType } from './mime';

export interface ResolveDocumentSourceOptions {
  fileName?: string;
  mimeType?: string;
  allowRemoteUrls?: boolean;
  fetchCredentials?: RequestCredentials;
  signal?: AbortSignal;
}

export class DocumentSourceError extends Error {
  constructor(message: string) {
    super(message);
    this.name = 'DocumentSourceError';
  }
}

interface DetectedDocumentType {
  extension: string;
  mimeType: string;
}

function isBlobLike(value: unknown): value is Blob {
  if (typeof Blob !== 'undefined' && value instanceof Blob) {
    return true;
  }

  return (
    typeof value === 'object' &&
    value !== null &&
    typeof (value as Blob).arrayBuffer === 'function' &&
    typeof (value as Blob).slice === 'function' &&
    typeof (value as Blob).size === 'number' &&
    typeof (value as Blob).type === 'string'
  );
}

function isBlob(value: unknown): value is Blob {
  return isBlobLike(value);
}

function isFile(value: unknown): value is File {
  return typeof File !== 'undefined' && value instanceof File;
}

function isArrayBuffer(value: unknown): value is ArrayBuffer {
  return value instanceof ArrayBuffer;
}

function isUint8Array(value: unknown): value is Uint8Array {
  return value instanceof Uint8Array;
}

function copyUint8Array(value: Uint8Array): ArrayBuffer {
  const copy = new Uint8Array(value.byteLength);
  copy.set(value);
  return copy.buffer;
}

function bytesStartWith(bytes: Uint8Array, signature: number[]): boolean {
  return signature.every((byte, index) => bytes[index] === byte);
}

function asciiAt(bytes: Uint8Array, start: number, length: number): string {
  return String.fromCharCode(...bytes.subarray(start, start + length));
}

async function detectZipDocumentType(arrayBuffer: ArrayBuffer): Promise<DetectedDocumentType | undefined> {
  try {
    const zip = await JSZip.loadAsync(arrayBuffer.slice(0));

    if (zip.file('xl/workbook.xml')) {
      return {
        extension: 'xlsx',
        mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      };
    }

    if (zip.file('word/document.xml')) {
      return {
        extension: 'docx',
        mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      };
    }

    if (zip.file('ppt/presentation.xml')) {
      return {
        extension: 'pptx',
        mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
      };
    }

    const odfMimeType = await zip.file('mimetype')?.async('text');
    if (odfMimeType === 'application/vnd.oasis.opendocument.spreadsheet') {
      return { extension: 'ods', mimeType: odfMimeType };
    }
    if (odfMimeType === 'application/vnd.oasis.opendocument.text') {
      return { extension: 'odt', mimeType: odfMimeType };
    }
    if (odfMimeType === 'application/vnd.oasis.opendocument.presentation') {
      return { extension: 'odp', mimeType: odfMimeType };
    }
  } catch {
    return undefined;
  }

  return undefined;
}

async function detectDocumentType(arrayBuffer: ArrayBuffer): Promise<DetectedDocumentType | undefined> {
  const bytes = new Uint8Array(arrayBuffer, 0, Math.min(arrayBuffer.byteLength, 16));

  if (bytesStartWith(bytes, [0x25, 0x50, 0x44, 0x46])) {
    return { extension: 'pdf', mimeType: 'application/pdf' };
  }

  if (bytesStartWith(bytes, [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a])) {
    return { extension: 'png', mimeType: 'image/png' };
  }

  if (bytesStartWith(bytes, [0xff, 0xd8, 0xff])) {
    return { extension: 'jpg', mimeType: 'image/jpeg' };
  }

  if (asciiAt(bytes, 0, 6) === 'GIF87a' || asciiAt(bytes, 0, 6) === 'GIF89a') {
    return { extension: 'gif', mimeType: 'image/gif' };
  }

  if (asciiAt(bytes, 0, 4) === 'RIFF' && asciiAt(bytes, 8, 4) === 'WEBP') {
    return { extension: 'webp', mimeType: 'image/webp' };
  }

  if (bytesStartWith(bytes, [0x49, 0x49, 0x2a, 0x00]) || bytesStartWith(bytes, [0x4d, 0x4d, 0x00, 0x2a])) {
    return { extension: 'tiff', mimeType: 'image/tiff' };
  }

  if (bytesStartWith(bytes, [0x50, 0x4b])) {
    return detectZipDocumentType(arrayBuffer);
  }

  return undefined;
}

function shouldUseDetectedMimeType(mimeType: string): boolean {
  return !mimeType || ['application/octet-stream', 'application/zip', 'application/x-zip-compressed'].includes(mimeType.toLowerCase());
}

function fileNameWithExtension(fileName: string, extension: string): string {
  if (!extension || extensionFromName(fileName)) {
    return fileName;
  }

  const baseName = fileName && fileName !== 'document' ? fileName : 'document';
  return `${baseName}.${extension}`;
}

function isRemoteUrl(url: string): boolean {
  return /^https?:\/\//i.test(url);
}

async function blobFromUrl(
  url: string,
  allowRemoteUrls: boolean,
  credentials: RequestCredentials,
  signal?: AbortSignal,
): Promise<Blob> {
  if (isRemoteUrl(url) && !allowRemoteUrls) {
    throw new DocumentSourceError(
      'Remote document URLs are disabled by default. Pass a File, Blob, ArrayBuffer, blob: URL, data: URL, or set allowRemoteUrls only for trusted same-origin/private endpoints.',
    );
  }

  const response = await fetch(url, {
    method: 'GET',
    credentials,
    signal,
  });

  if (!response.ok) {
    throw new DocumentSourceError(`Unable to load document URL: ${response.status} ${response.statusText}`);
  }

  return response.blob();
}

function normalizeSourceInput(source: DocumentSource): {
  data?: File | Blob | ArrayBuffer | Uint8Array;
  url?: string;
  fileName?: string;
  mimeType?: string;
} {
  if (typeof source === 'string') {
    return { url: source };
  }

  if (isFile(source) || isBlob(source) || isArrayBuffer(source) || isUint8Array(source)) {
    return { data: source };
  }

  return {
    ...source,
    data: source.data ?? source.blob,
  };
}

export async function resolveDocumentSource(
  source: DocumentSource,
  options: ResolveDocumentSourceOptions = {},
): Promise<ResolvedDocument> {
  const input = normalizeSourceInput(source);
  const allowRemoteUrls = options.allowRemoteUrls ?? false;
  const fetchCredentials = options.fetchCredentials ?? 'same-origin';
  let blob: Blob;
  let sourceKind: ResolvedDocument['sourceKind'] = 'blob';
  let originalUrl: string | undefined;
  let fileName = input.fileName || options.fileName || 'document';
  let mimeType = input.mimeType || options.mimeType || '';

  if (input.data) {
    if (isFile(input.data)) {
      fileName = input.fileName || options.fileName || input.data.name || fileName;
      mimeType = input.mimeType || options.mimeType || input.data.type || '';
      blob = input.data;
    } else if (isBlob(input.data)) {
      mimeType = input.mimeType || options.mimeType || input.data.type || '';
      blob = input.data;
    } else {
      sourceKind = 'buffer';
      const data = isUint8Array(input.data) ? copyUint8Array(input.data) : input.data;
      mimeType = mimeType || inferMimeType(fileName);
      blob = new Blob([data], { type: mimeType });
    }
  } else if (input.url) {
    sourceKind = 'url';
    originalUrl = input.url;
    fileName = input.fileName || options.fileName || fileNameFromUrl(input.url);
    blob = await blobFromUrl(input.url, allowRemoteUrls, fetchCredentials, options.signal);
    mimeType = input.mimeType || options.mimeType || blob.type || inferMimeType(fileName);
  } else {
    throw new DocumentSourceError('Document source must contain data or a URL.');
  }

  const arrayBuffer = await blob.arrayBuffer();
  const detectedType = await detectDocumentType(arrayBuffer);

  if (shouldUseDetectedMimeType(mimeType) && detectedType) {
    mimeType = detectedType.mimeType;
  }

  if (!mimeType) {
    mimeType = inferMimeType(fileName, blob.type || detectedType?.mimeType || 'application/octet-stream');
  }

  const inferredExtension = extensionFromName(fileName) || extensionFromMimeType(mimeType) || detectedType?.extension || '';
  fileName = fileNameWithExtension(fileName, inferredExtension);
  const objectUrl = URL.createObjectURL(blob);

  return {
    blob,
    arrayBuffer,
    objectUrl,
    fileName,
    mimeType,
    extension: extensionFromName(fileName) || inferredExtension,
    byteLength: blob.size,
    sourceKind,
    originalUrl,
  };
}

export function toError(error: unknown): Error {
  if (error instanceof Error) {
    return error;
  }

  return new Error(typeof error === 'string' ? error : 'Unknown document viewer error');
}
