import type { JsRow, Worksheet } from "./duke-sheets/generated/duke_sheets_wasm.js";
import { localizeFormattedNumber } from "./number-locale.ts";

export type AutofitAxis = "column" | "row";

export function readWorksheetAutofitCells(
  worksheet: Worksheet,
  axis: AutofitAxis,
  index: number,
  locale: string
): JsRow[] {
  return localizeSizingCells(worksheet.getAutofitCells(axis, index), locale);
}

export function readWorksheetAutoHeightCells(
  worksheet: Worksheet,
  rowCount: number,
  colCount: number,
  locale: string
): JsRow[] {
  return localizeSizingCells(worksheet.getAutoHeightCells(rowCount, colCount), locale);
}

function localizeSizingCells(rows: JsRow[], locale: string): JsRow[] {
  if (locale === "en-US") return rows;

  for (const row of rows) {
    for (const cell of row.cells) {
      if (cell.numericValue !== undefined) {
        cell.value = localizeFormattedNumber(cell.value, cell.style?.numberFormat, locale);
      }
    }
  }
  return rows;
}
