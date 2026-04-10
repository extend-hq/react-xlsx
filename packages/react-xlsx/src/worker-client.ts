import type { XlsxChart, XlsxChartsheet, XlsxSheetData, XlsxTable, XlsxWorkbookTab } from "./types";

type WorkerMessage =
  | {
      id: number;
      type: "load";
      payload: {
        buffer: ArrayBuffer;
        showHiddenSheets?: boolean;
        skipXmlParsing?: boolean;
      };
    }
  | {
      id: number;
      type: "parseCharts";
      payload: {
        buffer: ArrayBuffer;
        showHiddenSheets?: boolean;
        skipXmlParsing?: boolean;
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

type WorkerResponse = WorkerSuccessMessage | WorkerErrorMessage;

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

export class XlsxWorkerClient {
  private readonly worker: Worker;

  private nextRequestId = 1;

  private readonly pendingRequests = new Map<number, PendingRequest>();

  private disposed = false;

  constructor() {
    this.worker = new Worker(new URL("./xlsx-worker.js", import.meta.url), { type: "module" });
    this.worker.addEventListener("message", this.handleMessage);
    this.worker.addEventListener("error", this.handleError);
  }

  dispose() {
    if (this.disposed) {
      return;
    }

    this.disposed = true;
    this.worker.removeEventListener("message", this.handleMessage);
    this.worker.removeEventListener("error", this.handleError);
    this.worker.terminate();
    const abortError = createAbortError();
    for (const request of this.pendingRequests.values()) {
      request.reject(abortError);
    }
    this.pendingRequests.clear();
  }

  loadWorkbook(buffer: ArrayBuffer, skipXmlParsing = false, showHiddenSheets = false) {
    const workerBuffer = cloneArrayBufferForTransfer(buffer);
    return this.request<{
      chartsByWorkbookSheetIndex: XlsxChart[][];
      chartsheets: XlsxChartsheet[];
      sheets: XlsxSheetData[];
      tablesByWorkbookSheetIndex: XlsxTable[][];
      tabs: XlsxWorkbookTab[];
    }>({
      id: 0,
      payload: {
        buffer: workerBuffer,
        showHiddenSheets,
        skipXmlParsing
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
        skipXmlParsing
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

      const id = this.nextRequestId;
      this.nextRequestId += 1;
      this.pendingRequests.set(id, { reject, resolve: resolve as (value: unknown) => void });
      this.worker.postMessage({ ...message, id }, transfer);
    });
  }

  private readonly handleError = () => {
    for (const request of this.pendingRequests.values()) {
      request.reject(new Error("Worker request failed."));
    }
    this.pendingRequests.clear();
  };

  private readonly handleMessage = (event: MessageEvent<WorkerResponse>) => {
    const message = event.data;
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
