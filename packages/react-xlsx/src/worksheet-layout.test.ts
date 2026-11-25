import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import test from "node:test";
import { strFromU8, unzipSync } from "fflate";
import { initSync, Workbook } from "./duke-sheets/generated/duke_sheets_wasm.js";

import { applyReadOnlyLayout, changedVisibilityRanges } from "./read-only-layout.ts";
import type { XlsxSheetData } from "./types";

initSync({ module: readFileSync(new URL("./duke-sheets/generated/duke_sheets_wasm_bg.wasm", import.meta.url)) });

test("visibility and pane edits survive snapshots and XLSX export on an empty worksheet", (t) => {
  const workbook = new Workbook();
  const sheet = workbook.getSheet(0);
  t.after(() => {
    sheet.free();
    workbook.free();
  });
  sheet.setRowHeight(5, 31);
  sheet.setColumnWidth(2, 17.25);
  sheet.setRowsHidden(4, 6, true);
  sheet.setColumnsHidden(1, 3, true);
  sheet.setFreezePanes(2, 1);
  const restored = Workbook.fromSnapshot(workbook.exportSnapshot());
  t.after(() => restored.free());
  const bytes = restored.saveXlsxBytes();
  const xml = strFromU8(unzipSync(bytes)["xl/worksheets/sheet1.xml"]);
  assert.match(xml, /<pane\b[^>]*\btopLeftCell="B3"/);
  const reopened = Workbook.fromBytesForViewer(bytes);
  const layout = reopened.getSheet(0);
  t.after(() => {
    layout.free();
    reopened.free();
  });
  assert.deepEqual(Array.from(layout.hiddenRows), [4, 5, 6]);
  assert.deepEqual(Array.from(layout.hiddenColumns), [1, 2, 3]);
  assert.deepEqual(layout.freezePanes, { row: 2, col: 1 });

  layout.setRowsHidden(4, 5, false);
  layout.setColumnsHidden(1, 2, false);
  layout.setFreezePanes(0, 1);
  assert.deepEqual(layout.freezePanes, { row: 0, col: 1 });
  layout.setFreezePanes(1, 0);
  assert.deepEqual(layout.freezePanes, { row: 1, col: 0 });
  layout.setFreezePanes(0, 0);
  const cleared = reopened.saveXlsxBytes();
  const clearedXml = strFromU8(unzipSync(cleared)["xl/worksheets/sheet1.xml"]);
  assert.doesNotMatch(clearedXml, /<pane\b|\bpane="/);
  const finalWorkbook = Workbook.fromBytes(cleared);
  const finalSheet = finalWorkbook.getSheet(0);
  t.after(() => {
    finalSheet.free();
    finalWorkbook.free();
  });
  assert.equal(finalSheet.freezePanes, null);
  assert.deepEqual(Array.from(finalSheet.hiddenRows), [6]);
  assert.deepEqual(Array.from(finalSheet.hiddenColumns), [3]);
  assert.equal(finalSheet.getRowHeight(5), 31);
  assert.equal(finalSheet.getColumnWidth(1), undefined);
  assert.equal(finalSheet.getColumnWidth(2), 17.25);
});

test("read-only layouts preserve source dimensions, hidden headers, and freeze panes", () => {
  const sheet: XlsxSheetData = {
    name: "Sheet", workbookSheetIndex: 0, visibility: "visible", showGridLines: true,
    minUsedCol: 0, minUsedRow: 0, hasHorizontalMerges: false, hasVerticalMerges: false,
    maxHorizontalMergeEndCol: -1, maxVerticalMergeEndRow: -1,
    autoFilterRanges: [], dataValidations: [], conditionalFormatRules: [], sparklines: [],
    cachedFormulaValues: {}, colStyleIds: {}, rowStyleIds: {}, namedCellStyleByName: {},
    styleById: {}, tableStyleByName: {}, themePalette: { colorsByIndex: {} },
    freezePanes: { row: 2, col: 1 },
    maxUsedRow: 3, visibleRows: [1, 2, 3], hiddenRows: [0], rowCount: 3,
    rowHeights: [21, 32, 43], rowHeightOverridesPx: { 0: 56 }, defaultRowHeightPx: 18,
    maxUsedCol: 1, visibleCols: [0], hiddenCols: [1], colCount: 1,
    colWidths: [64], colWidthOverridesPx: { 1: 88 }, defaultColWidthPx: 60
  };
  const source = structuredClone(sheet);
  const view = applyReadOnlyLayout(sheet, {
    freezePanes: null, hiddenRows: new Set([8, 1]), hiddenCols: new Set()
  });
  assert.deepEqual(view.hiddenRows, [1, 8]);
  assert.deepEqual(view.visibleRows, [0, 2, 3]);
  assert.deepEqual(view.rowHeights, [56, 32, 43]);
  assert.equal(view.rowCount, 3);
  assert.deepEqual(view.visibleCols, [0, 1]);
  assert.deepEqual(view.colWidths, [64, 88]);
  assert.equal(view.colCount, 2);
  assert.equal(view.freezePanes, null);
  assert.deepEqual(sheet, source);
});

test("selection-based unhide includes leading hidden headers while explicit intervals stay exact", () => {
  const existing = new Set([0, 1, 4]);
  const ranges = [{ start: 2, end: 3 }, { start: 4, end: 5 }];
  assert.deepEqual(changedVisibilityRanges(existing, ranges, false, true), [{ start: 0, end: 3 }, ranges[1]]);
  assert.deepEqual(changedVisibilityRanges(existing, ranges, false, false), [ranges[1]]);
});
