import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { describe, it } from "node:test";
import initSheetsWasm, { Workbook } from "@dukelib/sheets-wasm";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import {
  normalizeWorkbookArrayBuffer,
  normalizeZipEntryNames,
  zipEntryNamesUseBackslashSeparators
} from "./zip-entry-names.ts";

const END_OF_CENTRAL_DIRECTORY_SIGNATURE = 0x06054b50;
const END_OF_CENTRAL_DIRECTORY_LENGTH = 22;

function findEndOfCentralDirectoryOffset(bytes: Uint8Array) {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  for (let offset = bytes.byteLength - END_OF_CENTRAL_DIRECTORY_LENGTH; offset >= 0; offset -= 1) {
    if (view.getUint32(offset, true) === END_OF_CENTRAL_DIRECTORY_SIGNATURE) {
      return offset;
    }
  }
  throw new Error("Missing end of central directory");
}

function readFirstEntryNames(bytes: Uint8Array) {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const endOfCentralDirectoryOffset = findEndOfCentralDirectoryOffset(bytes);
  const centralHeaderOffset = view.getUint32(endOfCentralDirectoryOffset + 16, true);
  const centralNameLength = view.getUint16(centralHeaderOffset + 28, true);
  const centralNameStart = centralHeaderOffset + 46;
  const localHeaderOffset = view.getUint32(centralHeaderOffset + 42, true);
  const localNameLength = view.getUint16(localHeaderOffset + 26, true);
  const localNameStart = localHeaderOffset + 30;
  const decoder = new TextDecoder();

  return {
    central: decoder.decode(bytes.subarray(centralNameStart, centralNameStart + centralNameLength)),
    local: decoder.decode(bytes.subarray(localNameStart, localNameStart + localNameLength))
  };
}

function exactArrayBuffer(bytes: Uint8Array) {
  const buffer = new ArrayBuffer(bytes.byteLength);
  new Uint8Array(buffer).set(bytes);
  return buffer;
}

function fixture(name: string) {
  const file = readFileSync(new URL(`../test-fixtures/${name}`, import.meta.url));
  return new Uint8Array(file.buffer, file.byteOffset, file.byteLength);
}

describe("ZIP entry-name normalization", () => {
  it("returns conformant bytes and buffers by identity", () => {
    const bytes = zipSync({ "xl/workbook.xml": strToU8("workbook") });
    const buffer = exactArrayBuffer(bytes);

    assert.equal(zipEntryNamesUseBackslashSeparators(bytes), false);
    assert.strictEqual(normalizeZipEntryNames(bytes), bytes);
    assert.strictEqual(normalizeWorkbookArrayBuffer(buffer), buffer);
  });

  it("patches both local and central filenames without rebuilding the archive", () => {
    const bytes = zipSync({ "xl\\workbook.xml": strToU8("workbook") });
    const normalized = normalizeZipEntryNames(bytes);

    assert.equal(zipEntryNamesUseBackslashSeparators(bytes), true);
    assert.notStrictEqual(normalized, bytes);
    assert.equal(normalized.byteLength, bytes.byteLength);
    const changedOffsets = Array.from(bytes.keys()).filter((index) => bytes[index] !== normalized[index]);
    assert.equal(changedOffsets.length, 2);
    for (const index of changedOffsets) {
      assert.equal(bytes[index], "\\".charCodeAt(0));
      assert.equal(normalized[index], "/".charCodeAt(0));
    }
    assert.deepEqual(readFirstEntryNames(normalized), {
      central: "xl/workbook.xml",
      local: "xl/workbook.xml"
    });
    assert.equal(strFromU8(unzipSync(normalized)["xl/workbook.xml"]), "workbook");
  });

  it("normalizes both reported regression fixtures", () => {
    for (const name of [
      "backslash-entry-names-minimal.xlsx",
      "backslash-entry-names-reported.xlsx"
    ]) {
      const bytes = fixture(name);
      const normalized = normalizeZipEntryNames(bytes);
      const entryNames = Object.keys(unzipSync(normalized));

      assert.equal(normalized.byteLength, bytes.byteLength);
      assert.equal(entryNames.some((entryName) => entryName.includes("\\")), false);
      assert.equal(entryNames.includes("xl/workbook.xml"), true);
    }
  });

  it("opens both normalized fixtures with the production workbook parser", async () => {
    const wasmFile = readFileSync(
      new URL(import.meta.resolve("@dukelib/sheets-wasm/duke_sheets_wasm_bg.wasm"))
    );
    await initSheetsWasm({ module_or_path: wasmFile });

    const expectations = [
      ["backslash-entry-names-minimal.xlsx", ["Sheet1"]],
      ["backslash-entry-names-reported.xlsx", ["项目概览", "实施计划", "预算概算"]]
    ] as const;

    for (const [name, expectedSheetNames] of expectations) {
      const workbook = Workbook.fromBytes(normalizeZipEntryNames(fixture(name)));
      const sheetNames = Array.from(
        { length: workbook.sheetCount },
        (_, index) => workbook.getSheet(index).name
      );
      assert.deepEqual(sheetNames, expectedSheetNames);
      workbook.free();
    }
  });

  it("rejects canonical-name collisions regardless of entry order", () => {
    for (const archive of [
      {
        "xl/workbook.xml": strToU8("canonical"),
        "xl\\workbook.xml": strToU8("backslash")
      },
      {
        "xl\\workbook.xml": strToU8("backslash"),
        "xl/workbook.xml": strToU8("canonical")
      }
    ]) {
      const bytes = zipSync(archive);
      const normalized = normalizeZipEntryNames(bytes);

      assert.strictEqual(normalized, bytes);
      const contents = unzipSync(normalized);
      assert.equal(strFromU8(contents["xl/workbook.xml"]), "canonical");
      assert.equal(strFromU8(contents["xl\\workbook.xml"]), "backslash");
    }
  });

  it("returns malformed local-header references unchanged", () => {
    const bytes = zipSync({ "xl\\workbook.xml": strToU8("workbook") });
    const corrupted = bytes.slice();
    const view = new DataView(corrupted.buffer, corrupted.byteOffset, corrupted.byteLength);
    const endOfCentralDirectoryOffset = findEndOfCentralDirectoryOffset(corrupted);
    const centralHeaderOffset = view.getUint32(endOfCentralDirectoryOffset + 16, true);
    view.setUint32(centralHeaderOffset + 42, corrupted.byteLength, true);

    assert.strictEqual(normalizeZipEntryNames(corrupted), corrupted);
  });

  it("returns ZIP64 archives unchanged", () => {
    const bytes = zipSync({ "xl\\workbook.xml": strToU8("workbook") });
    const zip64Marker = bytes.slice();
    const view = new DataView(zip64Marker.buffer, zip64Marker.byteOffset, zip64Marker.byteLength);
    const endOfCentralDirectoryOffset = findEndOfCentralDirectoryOffset(zip64Marker);
    view.setUint16(endOfCentralDirectoryOffset + 10, 0xffff, true);

    assert.equal(zipEntryNamesUseBackslashSeparators(zip64Marker), false);
    assert.strictEqual(normalizeZipEntryNames(zip64Marker), zip64Marker);
  });
});
