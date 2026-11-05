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

const chartsDirectory = dirname(require.resolve("@extend-ai/rust-office-charts/package.json"));
for (const name of ["LICENSE", "MAP-DATA-LICENSE"]) {
  copyFileSync(join(chartsDirectory, name), new URL(`rust-office-charts-${name}.txt`, distDir));
}
copyFileSync(require.resolve("@extend-ai/rust-office-charts/office_charts_bg.wasm"), new URL("office_charts_bg.wasm", distDir));
copyFileSync(require.resolve("@extend-ai/rust-office-charts/full/office_charts_bg.wasm"), new URL("office_charts_full_bg.wasm", distDir));
