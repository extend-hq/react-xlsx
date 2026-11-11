import type { WorkbookImageAssets } from "./images";

const workerBlobs = new Map<string, Blob>();

export function createWorkbookObjectUrl(blob: Blob) {
  const url = URL.createObjectURL(blob);
  if (typeof document === "undefined") workerBlobs.set(url, blob);
  return url;
}

export function revokeWorkbookObjectUrl(url: string) {
  workerBlobs.delete(url);
  URL.revokeObjectURL(url);
}

export function takeWorkerImageBlobs(urls: string[]): Array<[string, Blob]> {
  return urls.flatMap(url => {
    const blob = workerBlobs.get(url);
    workerBlobs.delete(url);
    return blob ? [[url, blob] as [string, Blob]] : [];
  });
}

export function restoreWorkerImageBlobs(assets: WorkbookImageAssets | null, blobs: Array<[string, Blob]> | undefined) {
  if (!assets || !blobs?.length) return;
  const urls = new Map(blobs.map(([url, blob]) => [url, URL.createObjectURL(blob)]));
  assets.objectUrls = assets.objectUrls.map(url => urls.get(url) ?? url);
  for (const images of assets.imagesByWorkbookSheetIndex) {
    for (const image of images) image.src = urls.get(image.src) ?? image.src;
  }
}
