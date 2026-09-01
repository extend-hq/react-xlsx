import assert from "node:assert/strict";
import test from "node:test";
import { BUILTIN_TABLE_STYLES, resolveBuiltinTableStyle } from "./builtin-table-styles.ts";

test("ships all 60 Excel built-in table styles", () => {
  const names = Object.keys(BUILTIN_TABLE_STYLES);
  assert.equal(names.length, 60);
  for (let index = 1; index <= 21; index += 1) {
    assert.ok(resolveBuiltinTableStyle(`TableStyleLight${index}`), `TableStyleLight${index}`);
  }
  for (let index = 1; index <= 28; index += 1) {
    assert.ok(resolveBuiltinTableStyle(`TableStyleMedium${index}`), `TableStyleMedium${index}`);
  }
  for (let index = 1; index <= 11; index += 1) {
    assert.ok(resolveBuiltinTableStyle(`TableStyleDark${index}`), `TableStyleDark${index}`);
  }
});

test("TableStyleMedium2 matches Excel's default table look", () => {
  const style = resolveBuiltinTableStyle("TableStyleMedium2");
  assert.ok(style);
  assert.deepEqual(style.headerRow, {
    fill: { color: { theme: 4 }, fillType: "solid" },
    font: { bold: true, color: { theme: 0 } }
  });
  assert.deepEqual(style.firstRowStripe, {
    fill: { color: { theme: 4, tint: 0.8 }, fillType: "solid" }
  });
  assert.equal(style.secondRowStripe, undefined);
  assert.equal(style.wholeTable?.border?.horizontal?.style, "thin");
  assert.equal(style.wholeTable?.border?.vertical, undefined);
  assert.equal(style.totalRow?.border?.top?.style, "double");
});

test("neutral styles use the theme text colour rather than an accent", () => {
  assert.deepEqual(resolveBuiltinTableStyle("TableStyleLight1")?.wholeTable?.border?.top?.color, { theme: 1 });
  assert.deepEqual(resolveBuiltinTableStyle("TableStyleMedium1")?.headerRow?.fill, {
    color: { theme: 1 },
    fillType: "solid"
  });
  assert.deepEqual(resolveBuiltinTableStyle("TableStyleDark1")?.headerRow?.fill?.color, { theme: 1 });
});

test("TableStyleLight14 resolves every element despite the source's off-by-one dxf count", () => {
  const style = resolveBuiltinTableStyle("TableStyleLight14");
  assert.ok(style);
  assert.ok(style.secondRowStripe);
  assert.ok(style.secondColumnStripe);
  assert.ok(style.wholeTable);
});

test("unknown and prototype names resolve to null", () => {
  assert.equal(resolveBuiltinTableStyle(undefined), null);
  assert.equal(resolveBuiltinTableStyle("MyCustomStyle"), null);
  assert.equal(resolveBuiltinTableStyle("constructor"), null);
  assert.equal(resolveBuiltinTableStyle("__proto__"), null);
});
