/**
 * Some XLSX writers emit ZIP entry names using the Windows path separator
 * (`xl\worksheets\sheet1.xml`) instead of the forward slash required by the ZIP
 * specification (APPNOTE 4.4.17.1) and OPC. Readers that look parts up by their
 * canonical name cannot find those parts and report the workbook as corrupt.
 *
 * ZIP entry names are stored in both the central directory and each local file
 * header. Since `\` and `/` have the same byte length, both copies can be
 * patched without decompressing or recompressing any workbook content.
 */

const BACKSLASH_BYTE = 0x5c;
const FORWARD_SLASH_BYTE = 0x2f;
const END_OF_CENTRAL_DIRECTORY_SIGNATURE = 0x06054b50;
const CENTRAL_FILE_HEADER_SIGNATURE = 0x02014b50;
const LOCAL_FILE_HEADER_SIGNATURE = 0x04034b50;
const END_OF_CENTRAL_DIRECTORY_LENGTH = 22;
const CENTRAL_FILE_HEADER_LENGTH = 46;
const LOCAL_FILE_HEADER_LENGTH = 30;
const MAX_ZIP_COMMENT_LENGTH = 0xffff;
const ZIP16_SENTINEL = 0xffff;
const ZIP32_SENTINEL = 0xffffffff;

type CentralDirectoryInfo = {
  endOffset: number;
  entryCount: number;
  offset: number;
};

type CentralDirectoryEntry = {
  diskNumberStart: number;
  localHeaderOffset: number;
  nameEnd: number;
  nameStart: number;
  nextOffset: number;
};

function findEndOfCentralDirectoryOffset(view: DataView) {
  const firstPossibleOffset = Math.max(
    0,
    view.byteLength - END_OF_CENTRAL_DIRECTORY_LENGTH - MAX_ZIP_COMMENT_LENGTH
  );

  for (let offset = view.byteLength - END_OF_CENTRAL_DIRECTORY_LENGTH; offset >= firstPossibleOffset; offset -= 1) {
    if (view.getUint32(offset, true) !== END_OF_CENTRAL_DIRECTORY_SIGNATURE) {
      continue;
    }

    const commentLength = view.getUint16(offset + 20, true);
    if (offset + END_OF_CENTRAL_DIRECTORY_LENGTH + commentLength === view.byteLength) {
      return offset;
    }
  }

  return null;
}

function readCentralDirectoryInfo(view: DataView): CentralDirectoryInfo | null {
  if (view.byteLength < END_OF_CENTRAL_DIRECTORY_LENGTH) {
    return null;
  }

  const endOfCentralDirectoryOffset = findEndOfCentralDirectoryOffset(view);
  if (endOfCentralDirectoryOffset === null) {
    return null;
  }

  const diskNumber = view.getUint16(endOfCentralDirectoryOffset + 4, true);
  const centralDirectoryDiskNumber = view.getUint16(endOfCentralDirectoryOffset + 6, true);
  const entriesOnDisk = view.getUint16(endOfCentralDirectoryOffset + 8, true);
  const entryCount = view.getUint16(endOfCentralDirectoryOffset + 10, true);
  const centralDirectorySize = view.getUint32(endOfCentralDirectoryOffset + 12, true);
  const centralDirectoryOffset = view.getUint32(endOfCentralDirectoryOffset + 16, true);

  if (
    diskNumber !== 0 ||
    centralDirectoryDiskNumber !== 0 ||
    entriesOnDisk !== entryCount ||
    entryCount === ZIP16_SENTINEL ||
    centralDirectorySize === ZIP32_SENTINEL ||
    centralDirectoryOffset === ZIP32_SENTINEL
  ) {
    return null;
  }

  const centralDirectoryEndOffset = centralDirectoryOffset + centralDirectorySize;
  if (
    centralDirectoryOffset > endOfCentralDirectoryOffset ||
    centralDirectoryEndOffset > endOfCentralDirectoryOffset ||
    centralDirectoryEndOffset > view.byteLength
  ) {
    return null;
  }

  return {
    endOffset: centralDirectoryEndOffset,
    entryCount,
    offset: centralDirectoryOffset
  };
}

function readCentralDirectoryEntry(
  view: DataView,
  centralDirectoryEndOffset: number,
  offset: number
): CentralDirectoryEntry | null {
  if (
    offset < 0 ||
    offset + CENTRAL_FILE_HEADER_LENGTH > centralDirectoryEndOffset ||
    offset + CENTRAL_FILE_HEADER_LENGTH > view.byteLength ||
    view.getUint32(offset, true) !== CENTRAL_FILE_HEADER_SIGNATURE
  ) {
    return null;
  }

  const nameLength = view.getUint16(offset + 28, true);
  const extraLength = view.getUint16(offset + 30, true);
  const commentLength = view.getUint16(offset + 32, true);
  const nameStart = offset + CENTRAL_FILE_HEADER_LENGTH;
  const nameEnd = nameStart + nameLength;
  const nextOffset = nameEnd + extraLength + commentLength;

  if (nameEnd > centralDirectoryEndOffset || nextOffset > centralDirectoryEndOffset) {
    return null;
  }

  return {
    diskNumberStart: view.getUint16(offset + 34, true),
    localHeaderOffset: view.getUint32(offset + 42, true),
    nameEnd,
    nameStart,
    nextOffset
  };
}

function entryNameUsesBackslash(bytes: Uint8Array, entry: CentralDirectoryEntry) {
  for (let cursor = entry.nameStart; cursor < entry.nameEnd; cursor += 1) {
    if (bytes[cursor] === BACKSLASH_BYTE) {
      return true;
    }
  }

  return false;
}

function centralDirectoryUsesBackslashSeparators(
  bytes: Uint8Array,
  view: DataView,
  centralDirectory: CentralDirectoryInfo
): boolean | null {
  let offset = centralDirectory.offset;

  for (let index = 0; index < centralDirectory.entryCount; index += 1) {
    const entry = readCentralDirectoryEntry(view, centralDirectory.endOffset, offset);
    if (!entry) {
      return null;
    }
    if (entryNameUsesBackslash(bytes, entry)) {
      return true;
    }
    offset = entry.nextOffset;
  }

  return false;
}

function normalizedEntryNameKey(bytes: Uint8Array, entry: CentralDirectoryEntry) {
  let key = "";
  for (let cursor = entry.nameStart; cursor < entry.nameEnd; cursor += 1) {
    const byte = bytes[cursor] === BACKSLASH_BYTE ? FORWARD_SLASH_BYTE : bytes[cursor];
    key += String.fromCharCode(byte);
  }
  return key;
}

function validateEntryNamePatches(
  bytes: Uint8Array,
  view: DataView,
  centralDirectory: CentralDirectoryInfo
): boolean {
  const normalizedNames = new Set<string>();
  const localHeaderOffsets = new Set<number>();
  let offset = centralDirectory.offset;

  for (let index = 0; index < centralDirectory.entryCount; index += 1) {
    const entry = readCentralDirectoryEntry(view, centralDirectory.endOffset, offset);
    if (
      !entry ||
      entry.diskNumberStart !== 0 ||
      entry.localHeaderOffset === ZIP32_SENTINEL ||
      entry.localHeaderOffset + LOCAL_FILE_HEADER_LENGTH > centralDirectory.offset ||
      localHeaderOffsets.has(entry.localHeaderOffset)
    ) {
      return false;
    }

    localHeaderOffsets.add(entry.localHeaderOffset);

    const normalizedName = normalizedEntryNameKey(bytes, entry);
    if (normalizedNames.has(normalizedName)) {
      return false;
    }
    normalizedNames.add(normalizedName);

    if (!entryNameUsesBackslash(bytes, entry)) {
      offset = entry.nextOffset;
      continue;
    }

    if (
      view.getUint32(entry.localHeaderOffset, true) !== LOCAL_FILE_HEADER_SIGNATURE
    ) {
      return false;
    }

    const localNameLength = view.getUint16(entry.localHeaderOffset + 26, true);
    const localNameStart = entry.localHeaderOffset + LOCAL_FILE_HEADER_LENGTH;
    const localNameEnd = localNameStart + localNameLength;
    if (
      localNameLength !== entry.nameEnd - entry.nameStart ||
      localNameEnd > centralDirectory.offset
    ) {
      return false;
    }

    for (let nameIndex = 0; nameIndex < localNameLength; nameIndex += 1) {
      const centralNameOffset = entry.nameStart + nameIndex;
      const localNameOffset = localNameStart + nameIndex;
      if (bytes[centralNameOffset] !== bytes[localNameOffset]) {
        return false;
      }
    }

    offset = entry.nextOffset;
  }

  return true;
}

function patchEntryNameSeparators(
  bytes: Uint8Array,
  view: DataView,
  centralDirectory: CentralDirectoryInfo
) {
  let offset = centralDirectory.offset;

  for (let index = 0; index < centralDirectory.entryCount; index += 1) {
    const entry = readCentralDirectoryEntry(view, centralDirectory.endOffset, offset);
    if (!entry) {
      return;
    }

    if (entryNameUsesBackslash(bytes, entry)) {
      const localNameStart = entry.localHeaderOffset + LOCAL_FILE_HEADER_LENGTH;
      for (let nameIndex = 0; nameIndex < entry.nameEnd - entry.nameStart; nameIndex += 1) {
        const centralNameOffset = entry.nameStart + nameIndex;
        if (bytes[centralNameOffset] === BACKSLASH_BYTE) {
          bytes[centralNameOffset] = FORWARD_SLASH_BYTE;
          bytes[localNameStart + nameIndex] = FORWARD_SLASH_BYTE;
        }
      }
    }

    offset = entry.nextOffset;
  }
}

/**
 * Detects backslash-separated entry names by walking only the central
 * directory. Invalid, multi-disk, and ZIP64 archives return `false`.
 */
export function zipEntryNamesUseBackslashSeparators(bytes: Uint8Array) {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const centralDirectory = readCentralDirectoryInfo(view);
  if (!centralDirectory) {
    return false;
  }

  return centralDirectoryUsesBackslashSeparators(bytes, view, centralDirectory) === true;
}

/**
 * Returns a copy with `/`-separated entry names, or the original bytes when no
 * rewrite is needed or the archive cannot be patched without ambiguity.
 */
export function normalizeZipEntryNames(bytes: Uint8Array) {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const centralDirectory = readCentralDirectoryInfo(view);
  if (
    !centralDirectory ||
    centralDirectoryUsesBackslashSeparators(bytes, view, centralDirectory) !== true
  ) {
    return bytes;
  }

  if (!validateEntryNamePatches(bytes, view, centralDirectory)) {
    return bytes;
  }

  const normalizedBytes = new Uint8Array(bytes.byteLength);
  normalizedBytes.set(bytes);
  const normalizedView = new DataView(
    normalizedBytes.buffer,
    normalizedBytes.byteOffset,
    normalizedBytes.byteLength
  );
  patchEntryNameSeparators(normalizedBytes, normalizedView, centralDirectory);
  return normalizedBytes;
}

export function normalizeWorkbookArrayBuffer(buffer: ArrayBuffer) {
  const bytes = new Uint8Array(buffer);
  const normalizedBytes = normalizeZipEntryNames(bytes);
  if (normalizedBytes === bytes) {
    return buffer;
  }

  if (
    normalizedBytes.byteOffset === 0 &&
    normalizedBytes.byteLength === normalizedBytes.buffer.byteLength &&
    normalizedBytes.buffer instanceof ArrayBuffer
  ) {
    return normalizedBytes.buffer;
  }

  const normalizedBuffer = new ArrayBuffer(normalizedBytes.byteLength);
  new Uint8Array(normalizedBuffer).set(normalizedBytes);
  return normalizedBuffer;
}
