import assert from "node:assert/strict";
import test from "node:test";
import { growScrollDisplayLimit } from "./scroll-display-limit.ts";

for (const { axis, initial, growth, maximum } of [
  { axis: "rows", initial: 200, growth: 200, maximum: 450 },
  { axis: "columns", initial: 26, growth: 26, maximum: 65 }
]) {
  test(`read-only ${axis} stop growing after repeated scrolling reaches the limit`, () => {
    let current = initial;
    assert.equal(growScrollDisplayLimit(current, growth, maximum, true), initial + growth);
    for (let scroll = 0; scroll < 20; scroll += 1) {
      current = growScrollDisplayLimit(current, growth, maximum, true);
      assert.ok(current <= maximum);
    }
    assert.equal(current, maximum);
  });

  test(`editable ${axis} continue growing past the populated range`, () => {
    assert.equal(growScrollDisplayLimit(maximum, growth, maximum, false), maximum + growth);
  });
}

test("an expanded read-only grid holds its limit instead of shrinking", () => {
  assert.equal(growScrollDisplayLimit(1000, 200, 450, true), 1000);
});

test("read-only rows keep a drawing anchored below the populated range in view", () => {
  assert.equal(growScrollDisplayLimit(546, 200, 221, true), 546);
});
