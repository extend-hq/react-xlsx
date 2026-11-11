import type { ChartWorkerResult, RenderChartAsync } from "./chart-renderer-types";
import type { WorkbookImageAssets, WorkbookImageSheetOrigin } from "./images";
import type { WorkbookChartOrigin } from "./charts";
import type { takeWorkbookXmlRecords } from "./workbook-xml";
import type { XlsxPivotTable } from "./pivot-types";
import { restoreWorkerImageBlobs } from "./workbook-object-urls";

const columnFontMeasurements = new Map<string, number>();

export type XlsxWorkerLoadResult = {
  pivotTables: XlsxPivotTable[];
  columnFontMeasurements: Array<[string, number]>;
  chartsByWorkbookSheetIndex: XlsxChart[][];
  chartsheets: XlsxChartsheet[];
  formControlsByWorkbookSheetIndex: XlsxFormControl[][];
  sheets: XlsxSheetData[];
  tablesByWorkbookSheetIndex: XlsxTable[][];
  tabs: XlsxWorkbookTab[];
  chartOriginsById: Map<string, WorkbookChartOrigin>;
  imageAssets: WorkbookImageAssets | null;
  imageBlobs?: Array<[string, Blob]>;
  xmlRecords: ReturnType<typeof takeWorkbookXmlRecords>;
  fontFamilies: string[];
  workbookSnapshot?: Uint8Array;
  shouldAutoCalculate: boolean;
};

import type { XlsxChart, XlsxChartsheet, XlsxFormControl, XlsxSheetData, XlsxTable, XlsxWorkbookTab } from "./types";
import { getConfiguredWorkerWasmSource, type WorkerWasmSource } from "./wasm";

export type WorkerMessage =
  | { id: number; type: "renderChart"; payload: Parameters<RenderChartAsync> }
  | { id: number; type: "export"; payload: { snapshot: Uint8Array; assets: WorkbookImageAssets | null; origins: Array<WorkbookImageSheetOrigin | null> } }

  | {
      id: number;
      type: "load";
      payload: {
        buffer: ArrayBuffer;
        includeSnapshot?: boolean;
        columnFontMeasurements?: Array<[string, number]>;
        showHiddenSheets?: boolean;
        skipXmlParsing?: boolean;
        wasmSource?: WorkerWasmSource;
      };
    }
  | {
      id: number;
      type: "getCellSnapshot";
      payload: {
        workbookSheetIndex: number;
        row: number;
        col: number;
        locale?: string;
      };
    }
  | {
      id: number;
      type: "getRowsBatch";
      payload: {
        workbookSheetIndex: number;
        startRow: number;
        rowCount: number;
        locale?: string;
      };
    };

type WorkerSuccessMessage =
  | {
      id: number;
      success: true;
      result: {
        chartsByWorkbookSheetIndex: XlsxChart[][];
        chartsheets: XlsxChartsheet[];
        tabs: XlsxWorkbookTab[];
      };
    }
  | {
      id: number;
      success: true;
      result: {
        chartsByWorkbookSheetIndex: XlsxChart[][];
        chartsheets: XlsxChartsheet[];
        formControlsByWorkbookSheetIndex: XlsxFormControl[][];
        sheets: XlsxSheetData[];
        tablesByWorkbookSheetIndex: XlsxTable[][];
        tabs: XlsxWorkbookTab[];
      };
    }
  | {
      id: number;
      success: true;
      result: {
        displayValue: string;
        formula: string;
      };
    }
  | {
      id: number;
      success: true;
      result: unknown[] | null;
    };

type WorkerErrorMessage = {
  id: number;
  success: false;
  error: string;
};

type WorkerResponse = WorkerSuccessMessage | WorkerErrorMessage | { type: "ready" };

type PendingRequest = {
  reject: (error: Error) => void;
  resolve: (value: unknown) => void;
};

function createAbortError() {
  if (typeof DOMException !== "undefined") {
    return new DOMException("Aborted", "AbortError");
  }

  const error = new Error("Aborted");
  error.name = "AbortError";
  return error;
}

export class XlsxWorkerUnavailableError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "XlsxWorkerUnavailableError";
  }
}

export class XlsxWorkerClient {
  private worker: Worker | null = null;

  private mainRuntime: Promise<ReturnType<typeof import("./workbook-runtime")["createWorkbookRuntime"]>> | null = null;

  private mainLoad: Promise<XlsxWorkerLoadResult> | null = null;

  private loadPayload: Extract<WorkerMessage, { type: "load" }>["payload"] | null = null;

  private nextRequestId = 1;

  private readonly pendingRequests = new Map<number, PendingRequest>();

  private disposed = false;

  private failure: XlsxWorkerUnavailableError | null = null;

  private startupTimer: ReturnType<typeof setTimeout> | null = null;

  constructor(useWorker = true) {
    if (!useWorker || typeof Worker === "undefined") return;
    try {
      this.worker = new Worker(new URL("./xlsx-worker.js", import.meta.url), { type: "module" });
    } catch {
      return;
    }
    this.worker.addEventListener("message", this.handleMessage);
    this.worker.addEventListener("error", this.handleError);
    this.worker.addEventListener("messageerror", this.handleError);
    this.startupTimer = setTimeout(() => {
      this.fail(new XlsxWorkerUnavailableError("The workbook worker did not start in time."));
    }, 15000);
  }

  dispose() {
    if (this.disposed) {
      return;
    }

    this.disposed = true;
    this.stopWorker();
    const abortError = createAbortError();
    for (const request of this.pendingRequests.values()) {
      request.reject(abortError);
    }
    this.pendingRequests.clear();
    void this.mainRuntime?.then(runtime => runtime.dispose(), () => undefined);
    this.loadPayload = null;
  }

  private stopWorker() {
    if (this.startupTimer !== null) clearTimeout(this.startupTimer);
    this.startupTimer = null;
    this.worker?.removeEventListener("message", this.handleMessage);
    this.worker?.removeEventListener("error", this.handleError);
    this.worker?.removeEventListener("messageerror", this.handleError);
    this.worker?.terminate();
  }

  async loadWorkbook(buffer: ArrayBuffer, skipXmlParsing = false, showHiddenSheets = false, includeSnapshot = false) {
    this.loadPayload = { buffer, columnFontMeasurements: Array.from(columnFontMeasurements), includeSnapshot, showHiddenSheets, skipXmlParsing };
    const result = await this.withFallback(async () => {
      const workerBuffer = cloneArrayBufferForTransfer(buffer);
      return this.request<XlsxWorkerLoadResult>({
        id: 0,
        payload: { ...this.loadPayload!, buffer: workerBuffer, wasmSource: await getConfiguredWorkerWasmSource() },
        type: "load"
      }, [workerBuffer]);
    }, () => this.loadOnMainThread());
    restoreWorkerImageBlobs(result.imageAssets, result.imageBlobs);
    delete result.imageBlobs;
    for (const [key, value] of result.columnFontMeasurements) columnFontMeasurements.set(key, value);
    while (columnFontMeasurements.size > 256) columnFontMeasurements.delete(columnFontMeasurements.keys().next().value!);
    return result;
  }

  getCellSnapshot(workbookSheetIndex: number, row: number, col: number, locale?: string) {
    return this.requestWithFallback<{
      displayValue: string;
      formula: string;
    }>({
      id: 0,
      payload: { col, row, workbookSheetIndex, locale },
      type: "getCellSnapshot"
    });
  }

  renderChart: RenderChartAsync = (...payload) => this.requestWithFallback<ChartWorkerResult>({ id: 0, type: "renderChart", payload });

  exportWorkbook(snapshot: Uint8Array, assets: WorkbookImageAssets | null, origins: Array<WorkbookImageSheetOrigin | null>) {
    return this.requestWithFallback<Uint8Array>({ id: 0, type: "export", payload: { snapshot, assets, origins } });
  }

  getRowsBatch(workbookSheetIndex: number, startRow: number, rowCount: number, locale?: string) {
    return this.requestWithFallback<unknown[] | null>({
      id: 0,
      payload: { rowCount, startRow, workbookSheetIndex, locale },
      type: "getRowsBatch"
    });
  }

  private async withFallback<TResult>(runWorker: () => Promise<TResult>, runMain: () => Promise<TResult>): Promise<TResult> {
    if (this.disposed) throw createAbortError();
    if (this.worker && !this.failure) {
      try {
        return await runWorker();
      } catch (error) {
        if (this.disposed || (error instanceof Error && error.name === "AbortError")) throw error;
        this.fail(new XlsxWorkerUnavailableError("The workbook worker is unavailable; using the main thread."));
      }
    }
    if (this.disposed) throw createAbortError();
    return runMain();
  }

  private async getMainRuntime() {
    if (this.disposed) throw createAbortError();
    this.mainRuntime ??= import("./workbook-runtime").then(({ createWorkbookRuntime }) => createWorkbookRuntime());
    const runtime = await this.mainRuntime;
    if (this.disposed) {
      runtime.dispose();
      throw createAbortError();
    }
    return runtime;
  }

  private loadOnMainThread(): Promise<XlsxWorkerLoadResult> {
    this.mainLoad ??= (async () => {
      const runtime = await this.getMainRuntime();
      if (!this.loadPayload) throw createAbortError();
      const result = await runtime.handleMessage({ id: 0, type: "load", payload: { ...this.loadPayload, includeSnapshot: true } }) as XlsxWorkerLoadResult;
      if (this.disposed) {
        runtime.dispose();
        throw createAbortError();
      }
      return result;
    })();
    return this.mainLoad;
  }

  private requestWithFallback<TResult>(message: WorkerMessage): Promise<TResult> {
    return this.withFallback(() => this.request<TResult>(message), async () => {
      const runtime = await this.getMainRuntime();
      if (this.loadPayload) await this.loadOnMainThread();
      if (this.disposed) throw createAbortError();
      return await runtime.handleMessage(message) as TResult;
    });
  }

  private request<TResult>(message: WorkerMessage, transfer: Transferable[] = []) {
    return new Promise<TResult>((resolve, reject) => {
      if (this.disposed) {
        reject(createAbortError());
        return;
      }
      if (this.failure) {
        reject(this.failure);
        return;
      }

      const id = this.nextRequestId;
      this.nextRequestId += 1;
      this.pendingRequests.set(id, { reject, resolve: resolve as (value: unknown) => void });
      try {
        this.worker!.postMessage({ ...message, id }, transfer);
      } catch {
        this.fail(new XlsxWorkerUnavailableError("The workbook worker could not receive a request."));
      }
    });
  }

  private readonly handleError = () => {
    this.fail(new XlsxWorkerUnavailableError("The workbook worker failed."));
  };

  private fail(error: XlsxWorkerUnavailableError) {
    if (this.disposed || this.failure) return;
    this.failure = error;
    this.stopWorker();
    for (const request of this.pendingRequests.values()) {
      request.reject(error);
    }
    this.pendingRequests.clear();
  }

  private readonly handleMessage = (event: MessageEvent<WorkerResponse>) => {
    const message = event.data;
    if ("type" in message) {
      if (this.startupTimer !== null) clearTimeout(this.startupTimer);
      this.startupTimer = null;
      return;
    }
    const request = this.pendingRequests.get(message.id);
    if (!request) {
      return;
    }

    this.pendingRequests.delete(message.id);
    if (!message.success) {
      request.reject(new Error(message.error));
      return;
    }

    request.resolve(message.result);
  };
}

function cloneArrayBufferForTransfer(buffer: ArrayBuffer) {
  return buffer.slice(0);
}
