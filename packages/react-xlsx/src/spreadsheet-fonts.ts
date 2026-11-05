import { loadOfficeFonts, OFFICE_FALLBACK_FONT } from "@extend-ai/office-fonts";
import { strFromU8, unzipSync } from "fflate";

export const SPREADSHEET_FALLBACK_FONT = OFFICE_FALLBACK_FONT;

export async function loadSpreadsheetFonts(buffer: ArrayBuffer): Promise<void> {
  const families = new Set(["Calibri", "Arial"]);
  try {
    let remaining = 32 * 1024 * 1024;
    const parts = unzipSync(new Uint8Array(buffer), {
      filter: ({ name, originalSize }) => {
        if (!/^xl\/(?:styles\.xml|sharedStrings\.xml|(?:theme|charts|drawings)\/(?:[^/]+\/)*[^/]+\.xml)$/.test(name)
          || originalSize > 4 * 1024 * 1024 || originalSize > remaining) return false;
        remaining -= originalSize;
        return true;
      },
    });
    for (const bytes of Object.values(parts)) {
      const xml = strFromU8(bytes);
      for (const match of xml.matchAll(/<(?:[\w-]+:)?(?:name|rFont|latin|ea|cs)\b[^>]*?\b(?:val|typeface)=(?:"([^"]*)"|'([^']*)')/g)) {
        const family = (match[1] ?? match[2]).replace(/&(amp|quot|apos|lt|gt);/g, (_, entity: string) =>
          ({ amp: "&", quot: '"', apos: "'", lt: "<", gt: ">" })[entity] ?? "");
        if (family && !family.startsWith("+")) families.add(family);
      }
    }
  } catch {
    // The workbook parser reports malformed packages; font loading remains optional.
  }
  await loadOfficeFonts(undefined, families).catch(() => undefined);
}
