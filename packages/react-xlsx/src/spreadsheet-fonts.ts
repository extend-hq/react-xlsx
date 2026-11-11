import { loadOfficeFonts, OFFICE_FALLBACK_FONT } from "@extend-ai/office-fonts";

export const SPREADSHEET_FALLBACK_FONT = OFFICE_FALLBACK_FONT;

export async function loadSpreadsheetFonts(families: Iterable<string>, target?: Document): Promise<void> {
  await loadOfficeFonts(target, families).catch(() => undefined);
}
