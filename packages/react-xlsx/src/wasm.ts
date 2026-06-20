let wasmModulePromise: Promise<typeof import("@dukelib/sheets-wasm")> | null = null;

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
}

export function initWasm(source?: XlsxWasmSource) {
  if (source !== undefined) {
    setWasmSource(source);
  }

  return getSheetsWasmModule();
}

export function canUseConfiguredWasmSourceInWorker(): boolean {
  return !hasConfiguredWasmSource || configuredWorkerWasmSource !== undefined;
}

export function getConfiguredWorkerWasmSource(): WorkerWasmSource | undefined {
  return configuredWorkerWasmSource;
}

export function getSheetsWasmModule() {
  if (!wasmModulePromise) {
    wasmModulePromise = import("@dukelib/sheets-wasm").then(async (mod) => {
      if (configuredWasmSource !== undefined) {
        await mod.default({ module_or_path: configuredWasmSource });
      } else {
        await mod.default();
      }
      return mod;
    });
  }

  return wasmModulePromise;
}
