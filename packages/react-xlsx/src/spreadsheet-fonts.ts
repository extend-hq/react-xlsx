let fontLoadPromise: Promise<void> | undefined;

export const SPREADSHEET_FALLBACK_FONT = "React XLSX Carlito";

export function loadSpreadsheetFonts(): Promise<void> {
  if (typeof document === "undefined" || typeof FontFace === "undefined") {
    return Promise.resolve();
  }
  return fontLoadPromise ??= import("./spreadsheet-font-assets").then(async ({ spreadsheetFontFaces }) => {
    await Promise.allSettled(spreadsheetFontFaces.map(async ({ source, ...descriptors }) => {
      const face = new FontFace(SPREADSHEET_FALLBACK_FONT, `url(${JSON.stringify(source)})`, descriptors);
      await face.load();
      document.fonts.add(face);
    }));
  }).catch(() => {});
}
