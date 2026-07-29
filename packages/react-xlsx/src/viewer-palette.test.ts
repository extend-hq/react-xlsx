import assert from "node:assert/strict";
import test from "node:test";
import { resolveViewerPalette } from "./viewer-palette.ts";

test("uses the built-in light and dark header colors by default", () => {
  const lightPalette = resolveViewerPalette();
  const darkPalette = resolveViewerPalette(true);

  assert.equal(lightPalette.headerSurface, "#f4f4f5");
  assert.equal(lightPalette.rowHeaderSurface, "#f4f4f5");
  assert.equal(lightPalette.headerText, "#71717a");
  assert.equal(darkPalette.headerSurface, "#18181b");
  assert.equal(darkPalette.rowHeaderSurface, "#18181b");
  assert.equal(darkPalette.headerText, "#a1a1aa");
});

test("applies header background and text overrides without changing the rest of the palette", () => {
  const palette = resolveViewerPalette(false, "#dbeafe", "#1e3a8a");

  assert.equal(palette.headerSurface, "#dbeafe");
  assert.equal(palette.rowHeaderSurface, "#dbeafe");
  assert.equal(palette.headerText, "#1e3a8a");
  assert.equal(palette.canvas, "#fafafa");
  assert.equal(palette.mutedText, "#71717a");
});

test("allows either header color to be overridden independently", () => {
  const backgroundOnly = resolveViewerPalette(false, "#fef3c7");
  const textOnly = resolveViewerPalette(true, undefined, "#ffffff");

  assert.equal(backgroundOnly.headerSurface, "#fef3c7");
  assert.equal(backgroundOnly.headerText, "#71717a");
  assert.equal(textOnly.headerSurface, "#18181b");
  assert.equal(textOnly.headerText, "#ffffff");
});
