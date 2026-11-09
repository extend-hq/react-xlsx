import assert from "node:assert/strict";
import { mkdtemp, readFile, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { after, before, test } from "node:test";
import { fileURLToPath, pathToFileURL } from "node:url";
import { build } from "tsup";

let output;

before(async () => {
  output = await mkdtemp(join(tmpdir(), "react-xlsx-office-art-"));
  await build({
    config: false,
    entry: { loader: fileURLToPath(new URL("../packages/react-xlsx/src/office-art/loader.ts", import.meta.url)) },
    format: ["esm", "cjs"],
    noExternal: ["fflate"],
    outDir: output,
    outExtension: ({ format }) => ({ js: format === "esm" ? ".mjs" : ".cjs" }),
    platform: "browser",
    silent: true
  });
});

after(async () => {
  if (output) await rm(output, { recursive: true, force: true });
});

const source = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+jxmcAAAAASUVORK5CYII=";
const picture = `<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <pic:blipFill><a:blip r:embed="rId1"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>
  <pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="762000" cy="381000"/></a:xfrm><a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom></pic:spPr>
</pic:pic>`;

for (const extension of ["mjs", "cjs"]) {
  test(`${extension} bundle does not reference an unpublished WASM asset`, async () => {
    const code = await readFile(join(output, `loader.${extension}`), "utf8");
    // Turbopack resolves literal WASM URLs at build time, even when a caller
    // always supplies bytes. Protect the emitted asset contract, not source layout.
    assert.equal(code.includes("office_art_bg.wasm"), false, "Bundle references the unpublished office_art_bg.wasm asset");
  });

  test(`${extension} bundle renders a picture using only embedded WASM`, async (t) => {
    t.mock.method(globalThis, "fetch", () => {
      throw new Error("Office Art must not fetch a separate WASM asset");
    });
    const entry = await import(pathToFileURL(join(output, `loader.${extension}`)).href);
    const { loadOfficeArt } = extension === "cjs" ? entry.default : entry;
    const renderer = await loadOfficeArt();
    const result = renderer.renderPicture(picture, source, 80, 40, () => "");
    assert.ok(result, "The initialized renderer should render a rounded picture");
    assert.match(result.svg, /<svg[^>]*width="80"[^>]*height="40"/);
    assert.ok(result.svg.includes(source), "The rendered picture should retain its image source");
  });
}
