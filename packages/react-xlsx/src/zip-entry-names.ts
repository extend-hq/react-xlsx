import { unzipSync, zipSync } from "fflate";

/**
 * Some XLSX writers emit ZIP entry names using the Windows path separator
 * (`xl\worksheets\sheet1.xml`) instead of the forward slash required by the ZIP
 * specification (APPNOTE 4.4.17.1) and OPC. Readers that look parts up by their
 * canonical name — including the Duke WASM core and this package's own XML
 * helpers — cannot find those parts and report the workbook as corrupt.
 *
 * Rewriting the entry names once at ingestion keeps every downstream consumer
 * on canonical `/`-separated paths.
 */

const BACKSLASH_BYTE = 0x5c;
const END_OF_CENTRAL_DIRECTORY_SIGNATURE = 0x06054b50;
const CENTRAL_FILE_HEADER_SIGNATURE = 0x02014b50;
const END_OF_CENTRAL_DIRECTORY_LENGTH = 22;
const CENTRAL_FILE_HEADER_LENGTH = 46;
const MAX_ZIP_COMMENT_LENGTH = 0xffff;

function findEndOfCentralDirectoryOffset(view: DataView) {
  const firstPossibleOffset = Math.max(
    0,
    view.byteLength - END_OF_CENTRAL_DIRECTORY_LENGTH - MAX_ZIP_COMMENT_LENGTH
  );

  for (let offset = view.byteLength - END_OF_CENTRAL_DIRECTORY_LENGTH; offset >= firstPossibleOffset; offset -= 1) {
    if (view.getUint32(offset, true) === END_OF_CENTRAL_DIRECTORY_SIGNATURE) {
      return offset;
    }
  }

  return null;
}

/**
 * Detects backslash-separated entry names by walking the central directory, so
 * a conformant workbook is never decompressed just to be checked. Returns
 * `false` for archives whose central directory cannot be read (for example
 * ZIP64), leaving those files to the existing code path.
 */
export function zipEntryNamesUseBackslashSeparators(bytes: Uint8Array) {
  if (bytes.byteLength < END_OF_CENTRAL_DIRECTORY_LENGTH) {
    return false;
  }

  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const endOfCentralDirectoryOffset = findEndOfCentralDirectoryOffset(view);
  if (endOfCentralDirectoryOffset === null) {
    return false;
  }

  const entryCount = view.getUint16(endOfCentralDirectoryOffset + 10, true);
  let offset = view.getUint32(endOfCentralDirectoryOffset + 16, true);

  for (let index = 0; index < entryCount; index += 1) {
    if (offset < 0 || offset + CENTRAL_FILE_HEADER_LENGTH > bytes.byteLength) {
      return false;
    }
    if (view.getUint32(offset, true) !== CENTRAL_FILE_HEADER_SIGNATURE) {
      return false;
    }

    const nameLength = view.getUint16(offset + 28, true);
    const extraLength = view.getUint16(offset + 30, true);
    const commentLength = view.getUint16(offset + 32, true);
    const nameStart = offset + CENTRAL_FILE_HEADER_LENGTH;
    const nameEnd = nameStart + nameLength;
    if (nameEnd > bytes.byteLength) {
      return false;
    }

    for (let cursor = nameStart; cursor < nameEnd; cursor += 1) {
      if (bytes[cursor] === BACKSLASH_BYTE) {
        return true;
      }
    }

    offset = nameEnd + extraLength + commentLength;
  }

  return false;
}

/**
 * Returns the workbook rebuilt with `/`-separated entry names, or the original
 * bytes when no rewrite is needed or the archive cannot be repacked.
 */
export function normalizeZipEntryNames(bytes: Uint8Array) {
  if (!zipEntryNamesUseBackslashSeparators(bytes)) {
    return bytes;
  }

  try {
    const archive = unzipSync(bytes);
    const normalizedArchive: Record<string, Uint8Array> = {};
    let didRename = false;

    for (const [name, content] of Object.entries(archive)) {
      const normalizedName = name.replace(/\\/g, "/");
      if (normalizedName !== name) {
        didRename = true;
      }
      normalizedArchive[normalizedName] = content;
    }

    return didRename ? zipSync(normalizedArchive) : bytes;
  } catch {
    return bytes;
  }
}

export function normalizeWorkbookArrayBuffer(buffer: ArrayBuffer) {
  const bytes = new Uint8Array(buffer);
  const normalizedBytes = normalizeZipEntryNames(bytes);
  if (normalizedBytes === bytes) {
    return buffer;
  }

  const normalizedBuffer = new ArrayBuffer(normalizedBytes.byteLength);
  new Uint8Array(normalizedBuffer).set(normalizedBytes);
  return normalizedBuffer;
}
