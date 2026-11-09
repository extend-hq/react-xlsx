import { gunzipSync } from "fflate";
import { initSync, normalizeRaster, renderPicture } from "./generated/office_art.js";
import payload from "./generated/payload.js";

export type PicturePaintBounds = {
  xPx: number;
  yPx: number;
  widthPx: number;
  heightPx: number;
  layoutWidthPx: number;
  layoutHeightPx: number;
};

export type RenderedPicture = { svg: string; paintBounds: PicturePaintBounds | null };

let pending: Promise<void> | undefined;

export async function loadOfficeArt() {
  pending ??= Promise.resolve().then(async () => {
    const compressed = Uint8Array.from(atob(payload), (character) => character.charCodeAt(0));
    // Keep compilation asynchronous without retaining wasm-bindgen's default
    // URL loader: the standalone office_art_bg.wasm is not shipped in the package.
    const module = await WebAssembly.compile(gunzipSync(compressed));
    initSync({ module });
  }).catch((error: unknown) => {
    pending = undefined;
    throw error;
  });
  await pending;
  return {
    normalizeRaster,
    renderPicture(xml: string, source: string, width: number, height: number, resolveColor: (xml: string) => string): RenderedPicture | null {
      const output = renderPicture(xml, source, width, height, resolveColor);
      return output ? JSON.parse(output) as RenderedPicture : null;
    }
  };
}
