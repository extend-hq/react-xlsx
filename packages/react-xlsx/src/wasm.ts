let wasmModulePromise: Promise<typeof import("./duke-sheets/generated/duke_sheets_wasm.js")> | null = null;

export type XlsxWasmSource =
  | string
  | URL
  | Request
  | Response
  | BufferSource
  | WebAssembly.Module;

export type WorkerWasmSource = string | ArrayBuffer | WebAssembly.Module;

let hasConfiguredWasmSource = false;
let configuredWasmSource: XlsxWasmSource | undefined;
let configuredWorkerWasmSource: WorkerWasmSource | undefined;
let configuredWorkerWasmPromise: Promise<WorkerWasmSource> | undefined;
let compiledWorkerWasmPromise: Promise<WebAssembly.Module> | undefined;

function bufferSourceToArrayBuffer(source: ArrayBuffer | ArrayBufferView<ArrayBufferLike>): ArrayBuffer {
  if (source instanceof ArrayBuffer) {
    return source.slice(0);
  }

  const bytes = new Uint8Array(source.buffer, source.byteOffset, source.byteLength);
  const copy = new Uint8Array(bytes);
  return copy.buffer;
}

function sourceToWorkerSource(source: XlsxWasmSource): WorkerWasmSource | undefined {
  if (typeof source === "string") {
    return source;
  }
  if (typeof URL !== "undefined" && source instanceof URL) {
    return source.href;
  }
  if (typeof Request !== "undefined" && source instanceof Request) {
    return source.url;
  }
  if (source instanceof ArrayBuffer || ArrayBuffer.isView(source)) {
    return bufferSourceToArrayBuffer(source);
  }
  if (typeof WebAssembly !== "undefined" && source instanceof WebAssembly.Module) {
    return source;
  }

  return undefined;
}

export function setWasmSource(source: XlsxWasmSource): void {
  hasConfiguredWasmSource = true;
  configuredWasmSource = source;
  configuredWorkerWasmSource = sourceToWorkerSource(source);
  compiledWorkerWasmPromise = undefined;
  configuredWorkerWasmPromise = typeof Response !== "undefined" && source instanceof Response
    ? source.clone().arrayBuffer()
    : typeof Request !== "undefined" && source instanceof Request
      ? fetch(source.clone()).then(response => { if (!response.ok) throw new Error("Unable to load spreadsheet runtime."); return response.arrayBuffer(); })
      : undefined;
}

export function initWasm(source?: XlsxWasmSource) {
  if (source !== undefined) {
    setWasmSource(source);
  }

  return getSheetsWasmModule();
}

export function canUseConfiguredWasmSourceInWorker(): boolean {
  return !hasConfiguredWasmSource || configuredWorkerWasmSource !== undefined || configuredWorkerWasmPromise !== undefined;
}

export async function getConfiguredWorkerWasmSource(): Promise<WorkerWasmSource | undefined> {
  if (!compiledWorkerWasmPromise) {
    compiledWorkerWasmPromise = (async () => {
      const source = await (configuredWorkerWasmPromise ?? configuredWorkerWasmSource);
      if (source instanceof ArrayBuffer) return WebAssembly.compile(source);
      if (source instanceof WebAssembly.Module) return source;
      const response = await fetch(source ?? new URL("./duke_sheets_wasm_bg.wasm", import.meta.url));
      if (!response.ok) throw new Error("Unable to load spreadsheet runtime.");
      if (typeof WebAssembly.compileStreaming === "function") {
        try { return await WebAssembly.compileStreaming(response.clone()); } catch { /* Retry servers without the WASM MIME type. */ }
      }
      return WebAssembly.compile(await response.arrayBuffer());
    })().catch(error => { compiledWorkerWasmPromise = undefined; throw error; });
  }
  return compiledWorkerWasmPromise;
}

export function getSheetsWasmModule() {
  if (!wasmModulePromise) {
    wasmModulePromise = import("./duke-sheets/generated/duke_sheets_wasm.js").then(async (mod) => {
      const exports = await mod.default(configuredWasmSource !== undefined ? { module_or_path: configuredWasmSource } : undefined);
      if (typeof exports.workbook_fromBytesForViewer !== "function" || typeof exports.workbook_fromSnapshot !== "function") {
        throw new Error("The spreadsheet WASM asset is incompatible. Use duke_sheets_wasm_bg.wasm from the same react-xlsx package version.");
      }
      return mod;
    });
  }

  return wasmModulePromise;
}
