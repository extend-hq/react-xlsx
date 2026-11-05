import { gunzipSync } from "fflate";
import init, { normalizeRaster, renderPicture } from "./generated/office_art.js";
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
    await init({ module_or_path: gunzipSync(compressed) });
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
