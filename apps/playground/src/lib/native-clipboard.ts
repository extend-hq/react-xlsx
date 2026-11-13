import type { XlsxDrawingClipboardAdapter } from "@extend-ai/react-xlsx";

declare const __XLSX_NATIVE_CLIPBOARD_TOKEN__: string;

async function base64(blob: Blob) {
  const bytes = new Uint8Array(await blob.arrayBuffer());
  let binary = "";
  for (let offset = 0; offset < bytes.length; offset += 8192) {
    binary += String.fromCharCode(...bytes.subarray(offset, offset + 8192));
  }
  return btoa(binary);
}

export const nativeDrawingClipboardAdapter: XlsxDrawingClipboardAdapter | undefined = __XLSX_NATIVE_CLIPBOARD_TOKEN__
  ? {
    async read() {
      const response = await request({ operation: "read" });
      const result = await response.json() as { data: string | null };
      return result.data ? new Blob([Uint8Array.from(atob(result.data), char => char.charCodeAt(0))]) : null;
    },
    async write({ kind, officeDrawing, previewImage, viewerDrawing }) {
      const [office, preview, viewer] = await Promise.all([base64(officeDrawing), base64(previewImage), base64(viewerDrawing)]);
      await request({ operation: "write", kind, officeDrawing: office, previewImage: preview, viewerDrawing: viewer });
    },
  }
  : undefined;

async function request(body: object) {
  const response = await fetch("/__xlsx_native_clipboard", {
    method: "POST",
    headers: { "Content-Type": "application/json", "X-Xlsx-Clipboard-Token": __XLSX_NATIVE_CLIPBOARD_TOKEN__ },
    body: JSON.stringify(body),
  });
  if (!response.ok) {
    const result = await response.json().catch(() => null);
    throw new Error(result?.error ?? "The native clipboard bridge is unavailable.");
  }
  return response;
}
