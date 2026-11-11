import { copyFileSync, mkdirSync } from "node:fs";
const wasmSource = new URL("../../../packages/react-xlsx/src/duke-sheets/generated/duke_sheets_wasm_bg.wasm", import.meta.url);
const publicDir = new URL("../public/", import.meta.url);
mkdirSync(publicDir, { recursive: true });
copyFileSync(wasmSource, new URL("duke_sheets_wasm_bg.wasm", publicDir));
