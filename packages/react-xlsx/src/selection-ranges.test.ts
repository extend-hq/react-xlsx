import assert from "node:assert/strict";
import test from "node:test";
import { addSelectionRange, bandIndices, clipboardBands, resolveFreezePanesCell, selectionCoversRange, subtractRange } from "./selection-ranges.ts";

const rows = (start: number, end = start) => ({ start: { row: start, col: 0 }, end: { row: end, col: 9 } });
const cols = (start: number, end = start) => ({ start: { row: 0, col: start }, end: { row: 9, col: end } });

test("freezing at a header inside a merge uses the selected row or column", () => {
  const anchor = { row: 0, col: 0 };
  assert.deepEqual(resolveFreezePanesCell(anchor, rows(2)), { row: 2, col: 0 });
  assert.deepEqual(resolveFreezePanesCell(anchor, cols(2)), { row: 0, col: 2 });
  const cell = { row: 3, col: 2 };
  assert.deepEqual(resolveFreezePanesCell(cell, { start: anchor, end: cell }), cell);
  assert.deepEqual(resolveFreezePanesCell(anchor, { start: anchor, end: cell }), anchor);
  assert.equal(resolveFreezePanesCell(null, null), null);
});

test("removing a header splits the selected band and preserves other areas", () => {
  const areas = [rows(1), rows(3, 7)];
  assert.equal(selectionCoversRange(areas, rows(4)), true);
  assert.equal(selectionCoversRange(areas, rows(1, 4)), false);
  assert.deepEqual(areas.flatMap(area => subtractRange(area, rows(4))), [rows(1), rows(3), rows(5, 7)]);
});

test("crossing a row and a column selects their intersection once", () => {
  const areas = addSelectionRange([rows(2)], cols(3));
  assert.equal(selectionCoversRange(areas, rows(2)), true);
  assert.equal(selectionCoversRange(areas, cols(3)), true);
  assert.equal(selectionCoversRange(areas, rows(3)), false);
  assert.equal(areas.reduce((size, { start, end }) => size + (end.row - start.row + 1) * (end.col - start.col + 1), 0), 19);
  assert.equal(clipboardBands(areas), null);
});

test("copying separated headers compacts selected indices in worksheet order", () => {
  const columnBands = clipboardBands([cols(5, 6), cols(2)])!;
  assert.deepEqual([...bandIndices(columnBands.cols)].map(({ index, offset }) => [index, offset]), [[2, 0], [5, 1], [6, 2]]);
  const rowBands = clipboardBands([rows(8), rows(1, 2)])!;
  assert.deepEqual([...bandIndices(rowBands.rows)].map(({ index }) => index), [1, 2, 8]);
});
