import { createWorkbookRuntime } from "./workbook-runtime";
import type { WorkerMessage } from "./worker-client";

type WorkerResponse =
  | { id: number; success: true; result: unknown }
  | { id: number; success: false; error: string };

const runtime = createWorkbookRuntime();

function respond(message: WorkerResponse) {
  const buffers = new Set<ArrayBuffer>();
  const seen = new Set<object>();
  const visit = (value: unknown) => {
    if (!value || typeof value !== "object" || seen.has(value)) return;
    seen.add(value);
    if (ArrayBuffer.isView(value) && value.buffer instanceof ArrayBuffer) buffers.add(value.buffer);
    else if (value instanceof ArrayBuffer) buffers.add(value);
    else if (value instanceof Map || value instanceof Set) for (const entry of value.values()) visit(entry);
    else for (const child of Object.values(value)) visit(child);
  };
  visit(message);
  (self as unknown as { postMessage: (message: WorkerResponse, transfer: Transferable[]) => void }).postMessage(message, [...buffers]);
}

self.addEventListener("message", (event: MessageEvent<WorkerMessage>) => {
  const message = event.data;
  void runtime.handleMessage(message)
    .then((result) => {
      respond({
        id: message.id,
        result,
        success: true
      });
    })
    .catch((error: unknown) => {
      respond({
        error: error instanceof Error ? error.message : "Worker request failed.",
        id: message.id,
        success: false
      });
    });
});

self.postMessage({ type: "ready" });
