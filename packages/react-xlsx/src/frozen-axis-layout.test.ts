import assert from "node:assert/strict";
import test from "node:test";
import { resolveFrozenAxisLayout } from "./frozen-axis-layout.ts";

test("a frozen axis scrolls when its bands and header fill or exceed the viewport", () => {
  const visibleCols = [0, 1, 2, 3];
  const widths = [230, 170, 100, 100];

  for (const viewportWidth of [350, 450]) {
    assert.deepEqual(resolveFrozenAxisLayout(visibleCols, widths, 2, 50, viewportWidth), {
      frozenIndices: [],
      stickyOffsets: new Map(),
      endOffset: 50
    });
  }
});

test("a wider viewport or zooming out restores the configured freeze", () => {
  const visibleCols = [0, 1, 2, 3];
  const widths = [230, 170, 100, 100];

  assert.deepEqual(resolveFrozenAxisLayout(visibleCols, widths, 2, 50, 550), {
    frozenIndices: [0, 1],
    stickyOffsets: new Map([[0, 50], [1, 280]]),
    endOffset: 450
  });
  assert.deepEqual(resolveFrozenAxisLayout(visibleCols, widths.map((width) => width / 2), 2, 25, 350), {
    frozenIndices: [0, 1],
    stickyOffsets: new Map([[0, 25], [1, 140]]),
    endOffset: 225
  });
});

test("hidden bands contribute no size and the freeze boundary uses sheet indices", () => {
  assert.deepEqual(resolveFrozenAxisLayout([0, 2, 3, 4], [90, 100, 250, 100], 3, 50, 260), {
    frozenIndices: [0, 2],
    stickyOffsets: new Map([[0, 50], [2, 140]]),
    endOffset: 240
  });
});
