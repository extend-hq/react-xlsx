import assert from "node:assert/strict";
import test from "node:test";
import { resolveWorkbookColor, resolveWorkbookFillColor } from "./colors.ts";

const themePalette = { colorsByIndex: { 0: "#ffffff", 1: "#000000", 4: "#4f81bd" } };

test("resolves theme colors in the shape getCellStyleAt returns", () => {
  assert.equal(resolveWorkbookColor({ colorType: "theme", themeIndex: 4, tint: 0 }, themePalette), "#4f81bd");
  assert.equal(
    resolveWorkbookFillColor({ fillType: "solid", color: { colorType: "theme", themeIndex: 4, tint: 0 } }, themePalette),
    "#4f81bd"
  );
});

test("applies the tint the same way for both theme color shapes", () => {
  assert.equal(resolveWorkbookColor({ colorType: "theme", themeIndex: 4, tint: 0.8 }, themePalette), "#dce6f2");
  assert.equal(resolveWorkbookColor({ theme: 4, tint: 0.8 }, themePalette), "#dce6f2");
});
