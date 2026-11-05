import { copyFileSync, mkdirSync } from "node:fs";
import { createRequire } from "node:module";
import { dirname, join } from "node:path";

const require = createRequire(new URL("../packages/react-xlsx/package.json", import.meta.url));
const dukeEntrypoint = require.resolve("@dukelib/sheets-wasm");
const wasmSource = join(dirname(dukeEntrypoint), "duke_sheets_wasm_bg.wasm");
const distDir = new URL("../packages/react-xlsx/dist/", import.meta.url);

mkdirSync(distDir, { recursive: true });
copyFileSync(wasmSource, new URL("duke_sheets_wasm_bg.wasm", distDir));

copyFileSync(new URL("../packages/react-xlsx/src/fonts/OFL.txt", import.meta.url), new URL("Carlito-OFL.txt", distDir));

for (const name of ["preact"]) {
  const directory = dirname(require.resolve(`${name}/package.json`));
  copyFileSync(join(directory, "LICENSE"), new URL(`${name}-LICENSE.txt`, distDir));
}

for (const name of ["LICENSE", "MAP-DATA-LICENSE", "provenance.json"]) {
  copyFileSync(new URL(`../packages/react-xlsx/src/office-maps/${name}`, import.meta.url), new URL(`office-maps-${name}`, distDir));
}

copyFileSync(new URL("../packages/react-xlsx/src/office-maps/office_charts_bg.wasm", import.meta.url), new URL("office_charts_bg.wasm", distDir));
