export type XlsCell = string | number | boolean | Date | null;

export interface ParsedXlsSheet {
  name: string;
  rows: XlsCell[][];
}

export interface ParsedXlsWorkbook {
  sheets: ParsedXlsSheet[];
}

const cfbSignature = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];
const endOfChain = 0xfffffffe;
const freeSector = 0xffffffff;
const maxRegularChainLength = 100_000;

function uint16(view: DataView, offset: number): number {
  return view.getUint16(offset, true);
}

function uint32(view: DataView, offset: number): number {
  return view.getUint32(offset, true);
}

function decodeUtf16(bytes: Uint8Array): string {
  return new TextDecoder('utf-16le').decode(bytes);
}

function decodeCompressedString(bytes: Uint8Array): string {
  let result = '';

  for (const byte of bytes) {
    result += String.fromCharCode(byte);
  }

  return result;
}

function concatParts(parts: Uint8Array[], size?: number): Uint8Array {
  const totalLength = parts.reduce((sum, part) => sum + part.length, 0);
  const targetLength = typeof size === 'number' ? Math.min(size, totalLength) : totalLength;
  const output = new Uint8Array(targetLength);
  let offset = 0;

  for (const part of parts) {
    if (offset >= targetLength) {
      break;
    }

    output.set(part.subarray(0, Math.min(part.length, targetLength - offset)), offset);
    offset += part.length;
  }

  return output;
}

function isCfbFile(bytes: Uint8Array): boolean {
  return cfbSignature.every((byte, index) => bytes[index] === byte);
}

interface DirectoryEntry {
  name: string;
  type: number;
  startSector: number;
  size: number;
}

function readCfbWorkbookStream(bytes: Uint8Array): Uint8Array {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const sectorSize = 2 ** uint16(view, 0x1e);
  const miniSectorSize = 2 ** uint16(view, 0x20);
  const fatSectorCount = uint32(view, 0x2c);
  const firstDirectorySector = uint32(view, 0x30);
  const miniStreamCutoff = uint32(view, 0x38);
  const firstMiniFatSector = uint32(view, 0x3c);
  const miniFatSectorCount = uint32(view, 0x40);
  const firstDifatSector = uint32(view, 0x44);
  const difatSectorCount = uint32(view, 0x48);
  const difat: number[] = [];

  for (let offset = 0x4c; offset < 0x200; offset += 4) {
    const sector = uint32(view, offset);
    if (sector !== freeSector) {
      difat.push(sector);
    }
  }

  let nextDifatSector = firstDifatSector;
  for (let index = 0; index < difatSectorCount && nextDifatSector !== endOfChain && nextDifatSector !== freeSector; index += 1) {
    const sectorOffset = sectorByteOffset(nextDifatSector, sectorSize);
    const entriesPerDifatSector = sectorSize / 4 - 1;

    for (let entryIndex = 0; entryIndex < entriesPerDifatSector; entryIndex += 1) {
      const sector = uint32(view, sectorOffset + entryIndex * 4);
      if (sector !== freeSector) {
        difat.push(sector);
      }
    }

    nextDifatSector = uint32(view, sectorOffset + entriesPerDifatSector * 4);
  }

  const fat: number[] = [];
  difat.slice(0, fatSectorCount).forEach((fatSector) => {
    const sectorOffset = sectorByteOffset(fatSector, sectorSize);
    for (let entryOffset = 0; entryOffset < sectorSize; entryOffset += 4) {
      fat.push(uint32(view, sectorOffset + entryOffset));
    }
  });

  function readRegularChain(startSector: number, size?: number): Uint8Array {
    if (startSector === endOfChain || startSector === freeSector) {
      return new Uint8Array();
    }

    const parts: Uint8Array[] = [];
    const visited = new Set<number>();
    let sector = startSector;

    while (sector !== endOfChain && sector !== freeSector && !visited.has(sector) && visited.size < maxRegularChainLength) {
      visited.add(sector);
      const offset = sectorByteOffset(sector, sectorSize);
      parts.push(bytes.subarray(offset, offset + sectorSize));
      sector = fat[sector] ?? endOfChain;
    }

    return concatParts(parts, size);
  }

  const directoryData = readRegularChain(firstDirectorySector);
  const entries: DirectoryEntry[] = [];

  for (let offset = 0; offset + 128 <= directoryData.length; offset += 128) {
    const entryView = new DataView(directoryData.buffer, directoryData.byteOffset + offset, 128);
    const nameLength = uint16(entryView, 0x40);
    const type = entryView.getUint8(0x42);

    if (nameLength < 2 || type === 0) {
      continue;
    }

    const nameBytes = directoryData.subarray(offset, offset + nameLength - 2);
    entries.push({
      name: decodeUtf16(nameBytes),
      type,
      startSector: uint32(entryView, 0x74),
      size: uint32(entryView, 0x78),
    });
  }

  const rootEntry = entries.find((entry) => entry.type === 5);
  const workbookEntry = entries.find((entry) => {
    const name = entry.name.toLowerCase();
    return entry.type === 2 && (name === 'workbook' || name === 'book');
  });

  if (!workbookEntry) {
    throw new Error('Unable to find a Workbook stream in this XLS file.');
  }

  if (workbookEntry.size >= miniStreamCutoff || !rootEntry) {
    return readRegularChain(workbookEntry.startSector, workbookEntry.size);
  }

  const rootStream = readRegularChain(rootEntry.startSector, rootEntry.size);
  const miniFatBytes = readRegularChain(firstMiniFatSector, miniFatSectorCount * sectorSize);
  const miniFatView = new DataView(miniFatBytes.buffer, miniFatBytes.byteOffset, miniFatBytes.byteLength);
  const miniFat: number[] = [];

  for (let offset = 0; offset + 4 <= miniFatBytes.length; offset += 4) {
    miniFat.push(uint32(miniFatView, offset));
  }

  const parts: Uint8Array[] = [];
  const visited = new Set<number>();
  let sector = workbookEntry.startSector;

  while (sector !== endOfChain && sector !== freeSector && !visited.has(sector) && visited.size < maxRegularChainLength) {
    visited.add(sector);
    const offset = sector * miniSectorSize;
    parts.push(rootStream.subarray(offset, offset + miniSectorSize));
    sector = miniFat[sector] ?? endOfChain;
  }

  return concatParts(parts, workbookEntry.size);
}

function sectorByteOffset(sector: number, sectorSize: number): number {
  return (sector + 1) * sectorSize;
}

function parseSheetName(payload: Uint8Array): { name: string; offset: number } {
  const view = new DataView(payload.buffer, payload.byteOffset, payload.byteLength);
  const offset = uint32(view, 0);
  const nameLength = payload[6] ?? 0;
  const flags = payload[7] ?? 0;
  const nameBytes = payload.subarray(8, 8 + nameLength * (flags & 1 ? 2 : 1));
  const name = flags & 1 ? decodeUtf16(nameBytes) : decodeCompressedString(nameBytes);

  return {
    name: name || 'Sheet',
    offset,
  };
}

function parseUnicodeString(bytes: Uint8Array, offset: number): { text: string; nextOffset: number } {
  if (offset + 3 > bytes.length) {
    return { text: '', nextOffset: bytes.length };
  }

  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const charCount = uint16(view, offset);
  const flags = bytes[offset + 2] ?? 0;
  const isUtf16 = Boolean(flags & 0x01);
  const hasRichText = Boolean(flags & 0x08);
  const hasExtendedText = Boolean(flags & 0x04);
  let cursor = offset + 3;
  let richRunCount = 0;
  let extendedByteCount = 0;

  if (hasRichText && cursor + 2 <= bytes.length) {
    richRunCount = uint16(view, cursor);
    cursor += 2;
  }

  if (hasExtendedText && cursor + 4 <= bytes.length) {
    extendedByteCount = uint32(view, cursor);
    cursor += 4;
  }

  const stringByteLength = charCount * (isUtf16 ? 2 : 1);
  const textBytes = bytes.subarray(cursor, Math.min(bytes.length, cursor + stringByteLength));
  const text = isUtf16 ? decodeUtf16(textBytes) : decodeCompressedString(textBytes);
  cursor += stringByteLength + richRunCount * 4 + extendedByteCount;

  return {
    text,
    nextOffset: Math.min(cursor, bytes.length),
  };
}

function parseSst(payloads: Uint8Array[]): string[] {
  const bytes = concatParts(payloads);
  if (bytes.length < 8) {
    return [];
  }

  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const uniqueCount = uint32(view, 4);
  const strings: string[] = [];
  let offset = 8;

  for (let index = 0; index < uniqueCount && offset < bytes.length; index += 1) {
    const parsed = parseUnicodeString(bytes, offset);
    strings.push(parsed.text);
    offset = parsed.nextOffset;
  }

  return strings;
}

function setCell(rows: XlsCell[][], rowIndex: number, columnIndex: number, value: XlsCell) {
  if (!rows[rowIndex]) {
    rows[rowIndex] = [];
  }

  rows[rowIndex][columnIndex] = value;
}

function decodeRkNumber(raw: number): number {
  const shouldDivide = Boolean(raw & 0x01);
  const isInteger = Boolean(raw & 0x02);
  let value: number;

  if (isInteger) {
    value = raw >> 2;
  } else {
    const buffer = new ArrayBuffer(8);
    const view = new DataView(buffer);
    view.setUint32(4, raw & 0xfffffffc, true);
    value = view.getFloat64(0, true);
  }

  return shouldDivide ? value / 100 : value;
}

function parseBiffWorkbook(workbookStream: Uint8Array): ParsedXlsWorkbook {
  const view = new DataView(workbookStream.buffer, workbookStream.byteOffset, workbookStream.byteLength);
  const sheets: Array<{ name: string; offset: number }> = [];
  let sharedStrings: string[] = [];
  let cursor = 0;

  while (cursor + 4 <= workbookStream.length) {
    const type = uint16(view, cursor);
    const length = uint16(view, cursor + 2);
    const payloadStart = cursor + 4;
    const payload = workbookStream.subarray(payloadStart, payloadStart + length);
    cursor = payloadStart + length;

    if (type === 0x0085) {
      sheets.push(parseSheetName(payload));
    } else if (type === 0x00fc) {
      const sstPayloads = [payload];

      while (cursor + 4 <= workbookStream.length && uint16(view, cursor) === 0x003c) {
        const continueLength = uint16(view, cursor + 2);
        const continueStart = cursor + 4;
        sstPayloads.push(workbookStream.subarray(continueStart, continueStart + continueLength));
        cursor = continueStart + continueLength;
      }

      sharedStrings = parseSst(sstPayloads);
    }
  }

  const sheetDescriptors = sheets.length > 0 ? sheets : [{ name: 'Sheet1', offset: 0 }];
  const parsedSheets = sheetDescriptors.map((sheet, index) => ({
    name: sheet.name || `Sheet${index + 1}`,
    rows: parseWorksheet(workbookStream, sheet.offset, sharedStrings),
  }));

  return {
    sheets: parsedSheets.length > 0 ? parsedSheets : [{ name: 'Sheet1', rows: [] }],
  };
}

function parseWorksheet(workbookStream: Uint8Array, startOffset: number, sharedStrings: string[]): XlsCell[][] {
  const view = new DataView(workbookStream.buffer, workbookStream.byteOffset, workbookStream.byteLength);
  const rows: XlsCell[][] = [];
  let cursor = startOffset;

  while (cursor + 4 <= workbookStream.length) {
    const type = uint16(view, cursor);
    const length = uint16(view, cursor + 2);
    const payloadStart = cursor + 4;
    const payload = workbookStream.subarray(payloadStart, payloadStart + length);
    const payloadView = new DataView(payload.buffer, payload.byteOffset, payload.byteLength);
    cursor = payloadStart + length;

    if (type === 0x000a) {
      break;
    }

    if (payload.length < 6) {
      continue;
    }

    if (type === 0x0203 && payload.length >= 14) {
      setCell(rows, uint16(payloadView, 0), uint16(payloadView, 2), payloadView.getFloat64(6, true));
    } else if (type === 0x00fd && payload.length >= 10) {
      const sharedStringIndex = uint32(payloadView, 6);
      setCell(rows, uint16(payloadView, 0), uint16(payloadView, 2), sharedStrings[sharedStringIndex] ?? '');
    } else if (type === 0x0204 && payload.length >= 9) {
      const parsed = parseUnicodeString(payload, 6);
      setCell(rows, uint16(payloadView, 0), uint16(payloadView, 2), parsed.text);
    } else if (type === 0x027e && payload.length >= 10) {
      setCell(rows, uint16(payloadView, 0), uint16(payloadView, 2), decodeRkNumber(uint32(payloadView, 6)));
    } else if (type === 0x00bd && payload.length >= 10) {
      const rowIndex = uint16(payloadView, 0);
      const firstColumn = uint16(payloadView, 2);
      const lastColumn = uint16(payloadView, payload.length - 2);
      let valueOffset = 4;

      for (let columnIndex = firstColumn; columnIndex <= lastColumn && valueOffset + 6 <= payload.length - 2; columnIndex += 1) {
        setCell(rows, rowIndex, columnIndex, decodeRkNumber(uint32(payloadView, valueOffset + 2)));
        valueOffset += 6;
      }
    } else if (type === 0x0205 && payload.length >= 8) {
      const value = payload[6] ?? 0;
      const isError = Boolean(payload[7]);
      setCell(rows, uint16(payloadView, 0), uint16(payloadView, 2), isError ? `#ERR${value}` : Boolean(value));
    } else if (type === 0x0006 && payload.length >= 14) {
      const result = payloadView.getFloat64(6, true);
      if (Number.isFinite(result)) {
        setCell(rows, uint16(payloadView, 0), uint16(payloadView, 2), result);
      }
    }
  }

  return rows;
}

export function parseXlsWorkbook(arrayBuffer: ArrayBuffer): ParsedXlsWorkbook {
  const bytes = new Uint8Array(arrayBuffer);
  const workbookStream = isCfbFile(bytes) ? readCfbWorkbookStream(bytes) : bytes;
  return parseBiffWorkbook(workbookStream);
}
