import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import test from "node:test";
import { initSync, Workbook } from "./duke-sheets/generated/duke_sheets_wasm.js";
import { readWorksheetAutofitCells, readWorksheetAutoHeightCells } from "./worksheet-autofit.ts";

initSync({
  module: readFileSync(new URL("./duke-sheets/generated/duke_sheets_wasm_bg.wasm", import.meta.url))
});

test("projects the selected column, including cells outside the viewport", () => {
  const workbook = new Workbook();
  const sheet = workbook.getSheet(0);
  try {
    sheet.setCell("A1", "Header");
    sheet.setCell("A130", "Long value below the viewport");
    sheet.setCell("B130", "Neighboring column");

    const rows = readWorksheetAutofitCells(sheet, "column", 0, "en-US");
    assert.deepEqual(rows.map(({ index, cells }) => [index, cells.map(({ col, value }) => [col, value])]), [
      [0, [[0, "Header"]]],
      [129, [[0, "Long value below the viewport"]]]
    ]);
  } finally {
    sheet.free();
    workbook.free();
  }
});

test("projects the selected row and localizes numbers from the batch", () => {
  const workbook = new Workbook();
  const sheet = workbook.getSheet(0);
  try {
    sheet.setCell("A1", "Neighboring row");
    sheet.setCell("A130", 1234.5);
    sheet.setCell("B130", "1234.5");

    const rows = readWorksheetAutofitCells(sheet, "row", 129, "de-DE");
    assert.equal(rows.length, 1);
    assert.equal(rows[0].index, 129);
    assert.deepEqual(rows[0].cells.map(({ col, value, numericValue }) => ({ col, value, numericValue })), [
      { col: 0, value: "1234,5", numericValue: 1234.5 },
      { col: 1, value: "1234.5", numericValue: undefined }
    ]);
  } finally {
    sheet.free();
    workbook.free();
  }
});

test("includes calculated values in spill targets", () => {
  const workbook = new Workbook();
  const sheet = workbook.getSheet(0);
  try {
    sheet.setFormula("A1", "=SEQUENCE(1,2,1234567890)");
    workbook.calculate();
    const rows = readWorksheetAutofitCells(sheet, "column", 1, "en-US");
    assert.equal(rows[0].cells[0].value, "1234567891");
    assert.equal(rows[0].cells[0].numericValue, 1234567891);
  } finally {
    sheet.free();
    workbook.free();
  }
});

test("projects only cells that can affect automatic row height", () => {
  const workbook = new Workbook();
  const sheet = workbook.getSheet(0);
  try {
    sheet.setCell("A1", "One line");
    sheet.setCell("B1", "A long value with wrapping");
    sheet.setCellStyle("B1", { alignment: { wrapText: true } });
    sheet.setCell("C1", "Line one\nLine two");
    sheet.setCell("D1", "Line one\nOutside the column range");
    sheet.setCell("A2", "Line one\nFixed height");
    sheet.setRowHeight(1, 20);
    sheet.setCell("A3", "Line one\nOutside the range");

    const rows = readWorksheetAutoHeightCells(sheet, 2, 3, "en-US");
    assert.deepEqual(rows.map(({ index, cells }) => [index, cells.map(({ col }) => col)]), [[0, [1, 2]]]);
  } finally {
    sheet.free();
    workbook.free();
  }
});

test("preserves general alignment and numeric operands with intentionally blank display text", () => {
  const workbook = new Workbook();
  const sheet = workbook.getSheet(0);
  try {
    sheet.setCell("A1", 1543);
    sheet.setCellStyle("A1", { numberFormat: { formatString: '""' } });
    sheet.setCell("B1", 42);
    sheet.setCell("C1", true);
    sheet.setCell("D1", "42");

    const rows = readWorksheetAutofitCells(sheet, "row", 0, "en-US");
    assert.equal(rows[0].cells[0].value, "");
    assert.equal(rows[0].cells[0].numericValue, 1543);
    assert.deepEqual(rows[0].cells.map(({ col, generalAlignment }) => [col, generalAlignment]), [
      [0, "right"],
      [1, "right"],
      [2, "center"],
      [3, "left"]
    ]);
  } finally {
    sheet.free();
    workbook.free();
  }
});
