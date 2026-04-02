let wasmModulePromise: Promise<typeof import("@dukelib/sheets-wasm")> | null = null;

export function getSheetsWasmModule() {
  if (!wasmModulePromise) {
    wasmModulePromise = import("@dukelib/sheets-wasm").then(async (mod) => {
      try {
        const wasmAsset = await import("@dukelib/sheets-wasm/duke_sheets_wasm_bg.wasm?url");
        await mod.default(wasmAsset.default);
      } catch {
        await mod.default();
      }
      return mod;
    });
  }

  return wasmModulePromise;
}
