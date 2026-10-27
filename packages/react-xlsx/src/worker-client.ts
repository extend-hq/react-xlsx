import type { XlsxChart, XlsxChartsheet, XlsxFormControl, XlsxSheetData, XlsxTable, XlsxWorkbookTab } from "./types";
import { getConfiguredWorkerWasmSource, type WorkerWasmSource } from "./wasm";

type WorkerMessage =
  | {
      id: number;
      type: "load";
      payload: {
        buffer: ArrayBuffer;
        showHiddenSheets?: boolean;
        skipXmlParsing?: boolean;
        wasmSource?: WorkerWasmSource;
      };
    }
  | {
      id: number;
      type: "parseCharts";
      payload: {
        buffer: ArrayBuffer;
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
      };
    }
  | {
      id: number;
      type: "getRowsBatch";
      payload: {
        workbookSheetIndex: number;
        startRow: number;
        rowCount: number;
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
  private readonly worker: Worker;

  private nextRequestId = 1;

  private readonly pendingRequests = new Map<number, PendingRequest>();

  private disposed = false;

  private failure: XlsxWorkerUnavailableError | null = null;

  private startupTimer: ReturnType<typeof setTimeout> | null = null;

  constructor() {
    try {
      this.worker = new Worker(new URL("./xlsx-worker.js", import.meta.url), { type: "module" });
    } catch {
      throw new XlsxWorkerUnavailableError("The workbook worker could not be started.");
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
  }

  private stopWorker() {
    if (this.startupTimer !== null) clearTimeout(this.startupTimer);
    this.startupTimer = null;
    this.worker.removeEventListener("message", this.handleMessage);
    this.worker.removeEventListener("error", this.handleError);
    this.worker.removeEventListener("messageerror", this.handleError);
    this.worker.terminate();
  }

  loadWorkbook(buffer: ArrayBuffer, skipXmlParsing = false, showHiddenSheets = false) {
    const workerBuffer = cloneArrayBufferForTransfer(buffer);
    return this.request<{
      chartsByWorkbookSheetIndex: XlsxChart[][];
      chartsheets: XlsxChartsheet[];
      formControlsByWorkbookSheetIndex: XlsxFormControl[][];
      sheets: XlsxSheetData[];
      tablesByWorkbookSheetIndex: XlsxTable[][];
      tabs: XlsxWorkbookTab[];
    }>({
      id: 0,
      payload: {
        buffer: workerBuffer,
        showHiddenSheets,
        skipXmlParsing,
        wasmSource: getConfiguredWorkerWasmSource()
      },
      type: "load"
    }, [workerBuffer]);
  }

  getCellSnapshot(workbookSheetIndex: number, row: number, col: number) {
    return this.request<{
      displayValue: string;
      formula: string;
    }>({
      id: 0,
      payload: { col, row, workbookSheetIndex },
      type: "getCellSnapshot"
    });
  }

  parseCharts(buffer: ArrayBuffer, skipXmlParsing = false, showHiddenSheets = false) {
    const workerBuffer = cloneArrayBufferForTransfer(buffer);
    return this.request<{
      chartsByWorkbookSheetIndex: XlsxChart[][];
      chartsheets: XlsxChartsheet[];
      tabs: XlsxWorkbookTab[];
    }>({
      id: 0,
      payload: {
        buffer: workerBuffer,
        showHiddenSheets,
        skipXmlParsing,
        wasmSource: getConfiguredWorkerWasmSource()
      },
      type: "parseCharts"
    }, [workerBuffer]);
  }

  getRowsBatch(workbookSheetIndex: number, startRow: number, rowCount: number) {
    return this.request<unknown[] | null>({
      id: 0,
      payload: { rowCount, startRow, workbookSheetIndex },
      type: "getRowsBatch"
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
        this.worker.postMessage({ ...message, id }, transfer);
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
