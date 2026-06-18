import { copyFileSync, mkdirSync } from "node:fs";
import { createRequire } from "node:module";
import { dirname, join } from "node:path";

const require = createRequire(new URL("../packages/react-xlsx/package.json", import.meta.url));
const dukeEntrypoint = require.resolve("@dukelib/sheets-wasm");
const wasmSource = join(dirname(dukeEntrypoint), "duke_sheets_wasm_bg.wasm");
const distDir = new URL("../packages/react-xlsx/dist/", import.meta.url);

mkdirSync(distDir, { recursive: true });
copyFileSync(wasmSource, new URL("duke_sheets_wasm_bg.wasm", distDir));
