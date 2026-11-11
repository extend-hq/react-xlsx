import type { Worksheet } from "./duke-sheets/generated/duke_sheets_wasm.js";

export type CellGeneralAlignment = "left" | "center" | "right";

export function getCellGeneralAlignment(worksheet: Worksheet, row: number, col: number): CellGeneralAlignment {
  const value = worksheet.getCalculatedValueAt(row, col);
  try {
    if (value.is_number) return "right";
    if (value.is_boolean || value.is_error) return "center";
    return "left";
  } finally {
    value.free();
  }
}
