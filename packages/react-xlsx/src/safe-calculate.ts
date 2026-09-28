import type { Workbook } from "@dukelib/sheets-wasm";

const SHEET_REF_REGEX = /'((?:[^']|'')+)'!|([A-Za-z_\u0080-\uFFFF][\w.\u0080-\uFFFF]*)!/g;

type FormulaCell = { formula?: string | null; row?: number; col?: number };

type CellPosition = { sheetIdx: number; row: number; col: number };

function collectReferencedSheetNames(workbook: Workbook): Set<string> {
  const referenced = new Set<string>();
  for (let sheetIdx = 0; sheetIdx < workbook.sheetCount; sheetIdx += 1) {
    let sheet;
    try {
      sheet = workbook.getSheet(sheetIdx);
    } catch {
      continue;
    }
    const cells = sheet.formulaCells as FormulaCell[] | null | undefined;
    if (!Array.isArray(cells)) {
      continue;
    }
    for (const cell of cells) {
      const formula = cell?.formula;
      if (!formula) {
        continue;
      }
      SHEET_REF_REGEX.lastIndex = 0;
      let match: RegExpExecArray | null;
      while ((match = SHEET_REF_REGEX.exec(formula)) !== null) {
        const raw = match[1] ?? match[2];
        if (!raw) {
          continue;
        }
        referenced.add(raw.replace(/''/g, "'"));
      }
    }
  }
  return referenced;
}

function hasUnresolvedSheetReferences(workbook: Workbook): boolean {
  let names: string[];
  try {
    names = workbook.sheetNames;
  } catch {
    return false;
  }
  const known = new Set(names);
  const referenced = collectReferencedSheetNames(workbook);
  for (const name of referenced) {
    if (!known.has(name)) {
      return true;
    }
  }
  return false;
}

function collectFormulaCellsWithSavedValues(workbook: Workbook): CellPosition[] {
  const positions: CellPosition[] = [];
  for (let sheetIdx = 0; sheetIdx < workbook.sheetCount; sheetIdx += 1) {
    let sheet;
    try {
      sheet = workbook.getSheet(sheetIdx);
    } catch {
      continue;
    }
    const cells = sheet.formulaCells as FormulaCell[] | null | undefined;
    if (!Array.isArray(cells)) {
      continue;
    }
    for (const cell of cells) {
      if (typeof cell?.row !== "number" || typeof cell?.col !== "number") {
        continue;
      }
      const saved = sheet.getCalculatedValueAt(cell.row, cell.col);
      if (!saved.is_empty && !saved.is_error) {
        positions.push({ sheetIdx, row: cell.row, col: cell.col });
      }
    }
  }
  return positions;
}

function anyBecameError(workbook: Workbook, positions: CellPosition[]): boolean {
  return positions.some(({ sheetIdx, row, col }) => {
    try {
      return workbook.getSheet(sheetIdx).getCalculatedValueAt(row, col).is_error;
    } catch {
      return false;
    }
  });
}

export type SafeCalculateSkipReason = "unresolved-sheet-refs" | "calculate-trapped" | "calculate-regressed";

export type SafeCalculateResult = {
  workbook: Workbook;
  calculated: boolean;
  skipReason: SafeCalculateSkipReason | null;
};

export type SafeCalculateOptions = {
  reparse?: () => Workbook;
};

// Pre-scans for formulas referencing missing sheets (which cause the Rust
// engine to panic into a wasm `unreachable` trap that poisons the Workbook
// instance). On trap, `reparse` is used to return a fresh usable instance.
//
// Also falls back to the saved values when calculating turns a formula that
// had a saved value into an error — e.g. a reference to a name defined as a
// formula (`Total = SUM(A1:A3)`), which the engine evaluates to #VALUE! —
// instead of showing errors the file never had.
export function safeCalculate(workbook: Workbook, options: SafeCalculateOptions = {}): SafeCalculateResult {
  if (hasUnresolvedSheetReferences(workbook)) {
    return { workbook, calculated: false, skipReason: "unresolved-sheet-refs" };
  }
  const withSavedValues = options.reparse ? collectFormulaCellsWithSavedValues(workbook) : [];
  try {
    workbook.calculate();
  } catch (err) {
    console.warn("[react-xlsx] workbook.calculate() trapped; falling back to cached formula values", err);
    if (options.reparse) {
      try {
        return { workbook: options.reparse(), calculated: false, skipReason: "calculate-trapped" };
      } catch (reparseErr) {
        console.warn("[react-xlsx] workbook reparse after calculate trap failed", reparseErr);
      }
    }
    return { workbook, calculated: false, skipReason: "calculate-trapped" };
  }
  if (options.reparse && anyBecameError(workbook, withSavedValues)) {
    try {
      return { workbook: options.reparse(), calculated: false, skipReason: "calculate-regressed" };
    } catch (reparseErr) {
      console.warn("[react-xlsx] workbook reparse after calculate regression failed", reparseErr);
    }
  }
  return { workbook, calculated: true, skipReason: null };
}

export function tryRecalculate(workbook: Workbook): { calculated: boolean; error: unknown } {
  try {
    workbook.calculate();
    return { calculated: true, error: null };
  } catch (err) {
    console.warn("[react-xlsx] workbook.calculate() trapped during recalculation", err);
    return { calculated: false, error: err };
  }
}
