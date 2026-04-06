import type { XlsxSheetData, XlsxTable } from "./types";

type WorkerMessage =
  | {
      id: number;
      type: "load";
      payload: {
        buffer: ArrayBuffer;
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
        sheets: XlsxSheetData[];
        tablesByWorkbookSheetIndex: XlsxTable[][];
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

export class XlsxWorkerClient {
  private readonly worker: Worker;

  private nextRequestId = 1;

  private readonly pendingRequests = new Map<number, PendingRequest>();

  constructor() {
    this.worker = new Worker(new URL("./xlsx-worker.js", import.meta.url), { type: "module" });
    this.worker.addEventListener("message", this.handleMessage);
    this.worker.addEventListener("error", this.handleError);
  }

  dispose() {
    this.worker.removeEventListener("message", this.handleMessage);
    this.worker.removeEventListener("error", this.handleError);
    this.worker.terminate();
    for (const request of this.pendingRequests.values()) {
      request.reject(new Error("Worker was disposed."));
    }
    this.pendingRequests.clear();
  }

  loadWorkbook(buffer: ArrayBuffer) {
    return this.request<{
      sheets: XlsxSheetData[];
      tablesByWorkbookSheetIndex: XlsxTable[][];
    }>({
      id: 0,
      payload: { buffer },
      type: "load"
    }, [buffer]);
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

  getRowsBatch(workbookSheetIndex: number, startRow: number, rowCount: number) {
    return this.request<unknown[] | null>({
      id: 0,
      payload: { rowCount, startRow, workbookSheetIndex },
      type: "getRowsBatch"
    });
  }

  private request<TResult>(message: WorkerMessage, transfer: Transferable[] = []) {
    return new Promise<TResult>((resolve, reject) => {
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
