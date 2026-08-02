import assert from "node:assert/strict";
import test from "node:test";
import { resolveCellTextClipOverscan } from "./cell-text-clip.ts";

test("wrapped text cannot overscan into adjacent rows", () => {
  assert.deepEqual(resolveCellTextClipOverscan(2.5, true), {
    horizontal: 2.5,
    vertical: 0
  });
});

test("single-line text retains glyph overscan in both directions", () => {
  assert.deepEqual(resolveCellTextClipOverscan(2.5, false), {
    horizontal: 2.5,
    vertical: 2.5
  });
});
