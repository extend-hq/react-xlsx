import assert from "node:assert/strict";
import test from "node:test";
import { createMergedRangeLookup } from "./merged-range-lookup.ts";

const titleRange = { start: { row: 0, col: 0 }, end: { row: 2, col: 2 } };
const sideRange = { start: { row: 1, col: 4 }, end: { row: 1, col: 6 } };

test("returns the merged range for its anchor and every secondary cell", () => {
  const lookup = createMergedRangeLookup([titleRange, sideRange]);

  assert.deepEqual(lookup({ row: 0, col: 0 }), titleRange);
  assert.deepEqual(lookup({ row: 1, col: 1 }), titleRange);
  assert.deepEqual(lookup({ row: 2, col: 2 }), titleRange);
  assert.deepEqual(lookup({ row: 1, col: 5 }), sideRange);
});

test("returns null for cells outside every merge", () => {
  const lookup = createMergedRangeLookup([titleRange, sideRange]);

  assert.equal(lookup({ row: 3, col: 0 }), null);
  assert.equal(lookup({ row: 1, col: 3 }), null);
  assert.equal(lookup({ row: 0, col: 4 }), null);
});

test("normalizes reversed ranges and ignores single-cell ranges", () => {
  const lookup = createMergedRangeLookup([
    { start: { row: 5, col: 3 }, end: { row: 4, col: 1 } },
    { start: { row: 8, col: 8 }, end: { row: 8, col: 8 } }
  ]);

  assert.deepEqual(lookup({ row: 5, col: 2 }), { start: { row: 4, col: 1 }, end: { row: 5, col: 3 } });
  assert.equal(lookup({ row: 8, col: 8 }), null);
});

test("indexes tall merges per row instead of per cell", () => {
  const tallRange = { start: { row: 0, col: 0 }, end: { row: 99_999, col: 50 } };
  const lookup = createMergedRangeLookup([tallRange]);

  assert.deepEqual(lookup({ row: 99_999, col: 50 }), tallRange);
  assert.equal(lookup({ row: 100_000, col: 0 }), null);
});
