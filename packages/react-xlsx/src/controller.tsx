import * as React from "react";
import type { Workbook } from "@dukelib/sheets-wasm";
import { getSheetsWasmModule } from "./wasm";
import type {
  UseXlsxViewerControllerOptions,
  XlsxCellAddress,
  XlsxCellRange,
  XlsxClipboardData,
  XlsxSheetData,
  XlsxViewerController
} from "./types";

const FORMULA_COUNT_THRESHOLD = 1000;
const DEFAULT_ROW_HEIGHT = 24;
const DEFAULT_COL_WIDTH = 80;
const XLSX_MIME_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
const CSV_MIME_TYPE = "text/csv;charset=utf-8";
const MIN_COL_WIDTH_PX = 30;
const MIN_ROW_HEIGHT_PX = 16;
const HISTORY_LIMIT = 100;
const INTERNAL_CLIPBOARD_MIME = "application/x-react-xlsx-range+json";

type SnapshotHistoryEntry = {
  kind: "snapshot";
  activeCell: XlsxCellAddress | null;
  activeSheetIndex: number;
  bytes: Uint8Array;
  selection: XlsxCellRange | null;
};

type CellMutationState = {
  formula: string | null;
  value: unknown;
};

type CellEditHistoryEntry = {
  kind: "cell-edit";
  activeCellAfter: XlsxCellAddress | null;
  activeCellBefore: XlsxCellAddress | null;
  after: CellMutationState;
  before: CellMutationState;
  cell: XlsxCellAddress;
  selectionAfter: XlsxCellRange | null;
  selectionBefore: XlsxCellRange | null;
  sheetIndex: number;
};

type RangeCellMutation = {
  after: CellMutationState;
  before: CellMutationState;
  cell: XlsxCellAddress;
};

type RangeEditHistoryEntry = {
  kind: "range-edit";
  activeCellAfter: XlsxCellAddress | null;
  activeCellBefore: XlsxCellAddress | null;
  mutations: RangeCellMutation[];
  selectionAfter: XlsxCellRange | null;
  selectionBefore: XlsxCellRange | null;
  sheetIndex: number;
};

type HistoryEntry = SnapshotHistoryEntry | CellEditHistoryEntry | RangeEditHistoryEntry;

type ClipboardMatrixCell = {
  colOffset: number;
  formula: string | null;
  rowOffset: number;
  value: string;
};

type ClipboardMerge = {
  colSpan: number;
  colOffset: number;
  rowOffset: number;
  rowSpan: number;
};

type ClipboardPayload = {
  cells: ClipboardMatrixCell[];
  cols: number;
  merges: ClipboardMerge[];
  rows: number;
};

function resolveDisplayFileName(src?: string, fileName?: string): string {
  if (typeof fileName === "string" && fileName.trim().length > 0) {
    return fileName.trim();
  }

  if (!src) {
    return "Workbook.xlsx";
  }

  const pathWithoutQuery = src.split("?")[0] ?? "";
  const pathSegments = pathWithoutQuery.split("/");
  const lastSegment = pathSegments[pathSegments.length - 1] ?? "";

  if (!lastSegment) {
    return "Workbook.xlsx";
  }

  try {
    return decodeURIComponent(lastSegment);
  } catch {
    return lastSegment;
  }
}

function buildSheetList(workbook: Workbook): XlsxSheetData[] {
  const sheets: XlsxSheetData[] = [];

  for (let index = 0; index < workbook.sheetCount; index += 1) {
    const worksheet = workbook.getSheet(index);
    if (worksheet.visibility !== "visible") {
      continue;
    }

    const usedRange = worksheet.usedRange() as [number, number, number, number] | null;
    if (!usedRange) {
      sheets.push({
        maxUsedCol: -1,
        maxUsedRow: -1,
        name: worksheet.name,
        rowCount: 0,
        colCount: 0,
        visibleRows: [],
        visibleCols: [],
        colWidths: [],
        rowHeights: [],
        workbookSheetIndex: index
      });
      continue;
    }

    const [, , maxRow, maxCol] = usedRange;
    const visibleRows: number[] = [];
    const visibleCols: number[] = [];

    for (let row = 0; row <= maxRow; row += 1) {
      if (!worksheet.isRowHidden(row)) {
        visibleRows.push(row);
      }
    }

    for (let col = 0; col <= maxCol; col += 1) {
      if (!worksheet.isColumnHidden(col)) {
        visibleCols.push(col);
      }
    }

    sheets.push({
      maxUsedCol: maxCol,
      maxUsedRow: maxRow,
      name: worksheet.name,
      rowCount: visibleRows.length,
      colCount: visibleCols.length,
      visibleRows,
      visibleCols,
      colWidths: visibleCols.map((col) => {
        const width = worksheet.getColumnWidth(col);
        return width !== undefined && width !== null ? Math.max(Math.round(width * 7.5), 30) : DEFAULT_COL_WIDTH;
      }),
      rowHeights: visibleRows.map((row) => {
        const height = worksheet.getRowHeight(row);
        return height !== undefined && height !== null ? Math.max(Math.round(height * 1.33), 16) : DEFAULT_ROW_HEIGHT;
      }),
      workbookSheetIndex: index
    });
  }

  return sheets;
}

function columnLabel(col: number): string {
  let label = "";
  let nextValue = col;

  while (nextValue >= 0) {
    label = String.fromCharCode(65 + (nextValue % 26)) + label;
    nextValue = Math.floor(nextValue / 26) - 1;
  }

  return label;
}

function cellAddressToA1(cell: XlsxCellAddress): string {
  return `${columnLabel(cell.col)}${cell.row + 1}`;
}

function normalizeRange(range: XlsxCellRange): XlsxCellRange {
  return {
    start: {
      row: Math.min(range.start.row, range.end.row),
      col: Math.min(range.start.col, range.end.col)
    },
    end: {
      row: Math.max(range.start.row, range.end.row),
      col: Math.max(range.start.col, range.end.col)
    }
  };
}

function rangeToA1(range: XlsxCellRange): string {
  const normalized = normalizeRange(range);
  const start = cellAddressToA1(normalized.start);
  const end = cellAddressToA1(normalized.end);
  return start === end ? start : `${start}:${end}`;
}

function rangeContainsCell(range: XlsxCellRange, cell: XlsxCellAddress): boolean {
  const normalized = normalizeRange(range);
  return (
    cell.row >= normalized.start.row &&
    cell.row <= normalized.end.row &&
    cell.col >= normalized.start.col &&
    cell.col <= normalized.end.col
  );
}

function fileStem(fileName: string): string {
  const normalized = fileName.trim();
  const lastDot = normalized.lastIndexOf(".");
  return lastDot > 0 ? normalized.slice(0, lastDot) : normalized;
}

function pxToSheetColumnWidth(widthPx: number): number {
  return Math.max(widthPx, MIN_COL_WIDTH_PX) / 7.5;
}

function pxToSheetRowHeight(heightPx: number): number {
  return Math.max(heightPx, MIN_ROW_HEIGHT_PX) / 1.33;
}

function cssColor(color: Record<string, unknown> | undefined): string | null {
  if (!color?.hex) {
    return null;
  }

  const hex = String(color.hex);
  const rgb = hex.length === 8 ? hex.slice(2) : hex;
  return `#${rgb}`;
}

function mapBorder(edge: { style: string; color?: { hex?: string } }): string {
  const color = cssColor(edge.color as Record<string, unknown> | undefined) ?? "#000000";
  const widthMap: Record<string, string> = {
    dashed: "1px",
    dotted: "1px",
    double: "3px",
    hair: "1px",
    medium: "2px",
    thick: "3px",
    thin: "1px"
  };
  const styleMap: Record<string, string> = {
    dashDot: "dashed",
    dashDotDot: "dotted",
    dashed: "dashed",
    dotted: "dotted",
    double: "double",
    hair: "solid",
    medium: "solid",
    mediumDashDot: "dashed",
    mediumDashDotDot: "dotted",
    mediumDashed: "dashed",
    slantDashDot: "dashed",
    thick: "solid",
    thin: "solid"
  };

  return `${widthMap[edge.style] ?? "1px"} ${styleMap[edge.style] ?? "solid"} ${color}`;
}

function decodeHtmlEntities(value: string): string {
  return value
    .replace(/&quot;/g, "\"")
    .replace(/&#34;/g, "\"")
    .replace(/&apos;/g, "'")
    .replace(/&#39;/g, "'")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&amp;/g, "&");
}

function cloneBytes(bytes: Uint8Array): Uint8Array {
  const nextBytes = new Uint8Array(bytes.byteLength);
  nextBytes.set(bytes);
  return nextBytes;
}

function pushHistoryEntry(stack: HistoryEntry[], entry: HistoryEntry) {
  stack.push(entry);
  if (stack.length > HISTORY_LIMIT) {
    stack.shift();
  }
}

function normalizeCellValue(value: unknown) {
  return value ?? "";
}

function applyCellMutationState(
  worksheet: ReturnType<Workbook["getSheet"]>,
  cell: XlsxCellAddress,
  state: CellMutationState
) {
  if (state.formula) {
    worksheet.setFormula(cellAddressToA1(cell), state.formula);
    return;
  }

  worksheet.setCell(cellAddressToA1(cell), normalizeCellValue(state.value));
}

function escapeHtml(value: string) {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function parseClipboardText(text: string): string[][] {
  const normalized = text.replace(/\r\n/g, "\n").replace(/\r/g, "\n");
  const rows = normalized.split("\n");

  if (rows.length > 1 && rows[rows.length - 1] === "") {
    rows.pop();
  }

  return rows.map((row) => row.split("\t"));
}

async function loadWorkbook({ file, src }: UseXlsxViewerControllerOptions): Promise<Workbook> {
  const wasmModule = await getSheetsWasmModule();
  let buffer: ArrayBuffer;

  if (file) {
    buffer = file;
  } else if (src) {
    const response = await fetch(src);
    if (!response.ok) {
      throw new Error(`Failed to fetch workbook (status ${response.status})`);
    }
    buffer = await response.arrayBuffer();
  } else {
    throw new Error("Either `file` or `src` must be provided.");
  }

  const workbook = wasmModule.Workbook.fromBytes(new Uint8Array(buffer));
  let totalFormulas = 0;

  for (let index = 0; index < workbook.sheetCount; index += 1) {
    totalFormulas += workbook.getSheet(index).formulaCount;
  }

  if (totalFormulas <= FORMULA_COUNT_THRESHOLD) {
    workbook.calculate();
  }

  return workbook;
}

function downloadArrayBuffer(file: ArrayBuffer, fileName: string) {
  const blob = new Blob([file], { type: XLSX_MIME_TYPE });
  const objectUrl = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = objectUrl;
  anchor.download = fileName;
  document.body.append(anchor);
  anchor.click();
  anchor.remove();
  setTimeout(() => URL.revokeObjectURL(objectUrl), 0);
}

function downloadBytes(bytes: Uint8Array, fileName: string, mimeType: string) {
  const normalizedBytes = new Uint8Array(bytes.byteLength);
  normalizedBytes.set(bytes);
  const blob = new Blob([normalizedBytes], { type: mimeType });
  const objectUrl = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = objectUrl;
  anchor.download = fileName;
  document.body.append(anchor);
  anchor.click();
  anchor.remove();
  setTimeout(() => URL.revokeObjectURL(objectUrl), 0);
}

function downloadText(text: string, fileName: string, mimeType: string) {
  const blob = new Blob([text], { type: mimeType });
  const objectUrl = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = objectUrl;
  anchor.download = fileName;
  document.body.append(anchor);
  anchor.click();
  anchor.remove();
  setTimeout(() => URL.revokeObjectURL(objectUrl), 0);
}

function downloadUrl(src: string, fileName: string) {
  const anchor = document.createElement("a");
  anchor.href = src;
  anchor.download = fileName;
  anchor.rel = "noreferrer";
  document.body.append(anchor);
  anchor.click();
  anchor.remove();
}

export function useXlsxViewerController(options: UseXlsxViewerControllerOptions): XlsxViewerController {
  const { file, fileName, readOnly = false, src } = options;
  const [isLoading, setIsLoading] = React.useState(Boolean(file ?? src));
  const [error, setError] = React.useState<Error | null>(null);
  const [workbook, setWorkbook] = React.useState<Workbook | null>(null);
  const [sheets, setSheets] = React.useState<XlsxSheetData[]>([]);
  const [activeSheetIndex, setActiveSheetIndexState] = React.useState(0);
  const [activeCell, setActiveCell] = React.useState<XlsxCellAddress | null>(null);
  const [selection, setSelection] = React.useState<XlsxCellRange | null>(null);
  const [revision, setRevision] = React.useState(0);
  const selectionAnchorRef = React.useRef<XlsxCellAddress | null>(null);
  const undoStackRef = React.useRef<HistoryEntry[]>([]);
  const redoStackRef = React.useRef<HistoryEntry[]>([]);
  const isApplyingHistoryRef = React.useRef(false);
  const [historyRevision, setHistoryRevision] = React.useState(0);
  const displayFileName = React.useMemo(() => resolveDisplayFileName(src, fileName), [fileName, src]);

  const refreshWorkbookState = React.useCallback((targetWorkbook: Workbook) => {
    setSheets(buildSheetList(targetWorkbook));
    setRevision((current) => current + 1);
  }, []);

  React.useEffect(() => {
    if (!file && !src) {
      setWorkbook(null);
      setSheets([]);
      setError(null);
      setIsLoading(false);
      setActiveSheetIndexState(0);
      setActiveCell(null);
      setSelection(null);
      selectionAnchorRef.current = null;
      undoStackRef.current = [];
      redoStackRef.current = [];
      setHistoryRevision(0);
      setRevision(0);
      return;
    }

    let isCurrent = true;
    setIsLoading(true);
    setError(null);
    setActiveSheetIndexState(0);
    setActiveCell(null);
    setSelection(null);
    selectionAnchorRef.current = null;
    undoStackRef.current = [];
    redoStackRef.current = [];
    setHistoryRevision(0);
    setRevision(0);

    void loadWorkbook({ file, src })
      .then((nextWorkbook) => {
        if (!isCurrent) {
          return;
        }

        setWorkbook(nextWorkbook);
        setSheets(buildSheetList(nextWorkbook));
        setIsLoading(false);
      })
      .catch((nextError: unknown) => {
        if (!isCurrent) {
          return;
        }

        setWorkbook(null);
        setSheets([]);
        setError(nextError instanceof Error ? nextError : new Error("Could not load workbook."));
        setIsLoading(false);
      });

    return () => {
      isCurrent = false;
    };
  }, [file, src]);

  const activeSheet = sheets[activeSheetIndex] ?? null;

  React.useEffect(() => {
    setActiveCell(null);
    setSelection(null);
    selectionAnchorRef.current = null;
  }, [activeSheetIndex]);

  const setActiveSheetIndex = React.useCallback((index: number) => {
    setActiveSheetIndexState((currentIndex) => {
      if (index < 0 || index >= sheets.length) {
        return currentIndex;
      }
      return index;
    });
  }, [sheets.length]);

  const getActiveWorksheet = React.useCallback(() => {
    if (!workbook || !activeSheet) {
      return null;
    }

    return workbook.getSheet(activeSheet.workbookSheetIndex);
  }, [activeSheet, workbook]);

  const getCellDisplayValue = React.useCallback((cell?: XlsxCellAddress | null) => {
    const worksheet = getActiveWorksheet();
    if (!worksheet || !cell) {
      return "";
    }

    const formatted = worksheet.getFormattedValueAt(cell.row, cell.col);
    if (formatted) {
      return decodeHtmlEntities(formatted);
    }

    const calculated = worksheet.getCalculatedValueAt(cell.row, cell.col);
    if (calculated.is_error) {
      return calculated.asError() ?? "";
    }
    if (calculated.is_empty) {
      return "";
    }

    return calculated.toString();
  }, [getActiveWorksheet]);

  const getCellFormula = React.useCallback((cell?: XlsxCellAddress | null) => {
    const worksheet = getActiveWorksheet();
    if (!worksheet || !cell) {
      return "";
    }

    return worksheet.getFormulaAt(cell.row, cell.col) ?? "";
  }, [getActiveWorksheet]);

  const getClipboardData = React.useCallback((): XlsxClipboardData | null => {
    const worksheet = getActiveWorksheet();
    const targetRange = selection ?? (activeCell ? { start: activeCell, end: activeCell } : null);
    if (!worksheet || !targetRange) {
      return null;
    }

    const normalized = normalizeRange(targetRange);
    const rows: string[] = [];
    const htmlRows: string[] = [];
    const payload: ClipboardPayload = {
      cells: [],
      cols: normalized.end.col - normalized.start.col + 1,
      merges: [],
      rows: normalized.end.row - normalized.start.row + 1
    };

    for (let row = normalized.start.row; row <= normalized.end.row; row += 1) {
      const textCells: string[] = [];
      const htmlCells: string[] = [];

      for (let col = normalized.start.col; col <= normalized.end.col; col += 1) {
        if (worksheet.isMergedSecondary(row, col)) {
          textCells.push("");
          continue;
        }

        const formula = worksheet.getFormulaAt(row, col) ?? null;
        const value = getCellDisplayValue({ row, col });
        const merge = worksheet.getMergeSpan(row, col) as
          | { colSpan?: number; rowSpan?: number }
          | null
          | undefined;
        const rawStyle = worksheet.getCellStyleAt(row, col) as Record<string, unknown> | null | undefined;
        const cellStyles: string[] = [
          "padding:2px 4px",
          "white-space:pre-wrap",
          "vertical-align:top"
        ];

        const fill = rawStyle?.fill as Record<string, unknown> | undefined;
        if (fill) {
          const fillColor =
            fill.fillType === "solid"
              ? cssColor(fill.color as Record<string, unknown> | undefined)
              : fill.fillType === "pattern"
                ? cssColor(fill.foreground as Record<string, unknown> | undefined)
                : null;

          if (fillColor && fillColor.toLowerCase() !== "#ffffff") {
            cellStyles.push(`background-color:${fillColor}`);
          }
        }

        const font = rawStyle?.font as Record<string, unknown> | undefined;
        if (font) {
          if (font.bold) {
            cellStyles.push("font-weight:700");
          }
          if (font.italic) {
            cellStyles.push("font-style:italic");
          }
          if (font.underline && font.underline !== "none") {
            cellStyles.push("text-decoration:underline");
          }
          if (font.strikethrough) {
            cellStyles.push("text-decoration:line-through");
          }
          const fontColor = cssColor(font.color as Record<string, unknown> | undefined);
          if (fontColor) {
            cellStyles.push(`color:${fontColor}`);
          }
          if (typeof font.size === "number") {
            cellStyles.push(`font-size:${font.size}pt`);
          }
        }

        const alignment = rawStyle?.alignment as Record<string, unknown> | undefined;
        if (alignment?.horizontal && alignment.horizontal !== "general") {
          cellStyles.push(`text-align:${String(alignment.horizontal)}`);
        }
        if (alignment?.wrapText) {
          cellStyles.push("white-space:pre-wrap");
          cellStyles.push("word-break:break-word");
        }

        const border = rawStyle?.border as Record<string, Record<string, unknown>> | undefined;
        if (border?.top?.style && border.top.style !== "none") {
          cellStyles.push(`border-top:${mapBorder(border.top as { color?: { hex?: string }; style: string })}`);
        }
        if (border?.right?.style && border.right.style !== "none") {
          cellStyles.push(`border-right:${mapBorder(border.right as { color?: { hex?: string }; style: string })}`);
        }
        if (border?.bottom?.style && border.bottom.style !== "none") {
          cellStyles.push(`border-bottom:${mapBorder(border.bottom as { color?: { hex?: string }; style: string })}`);
        }
        if (border?.left?.style && border.left.style !== "none") {
          cellStyles.push(`border-left:${mapBorder(border.left as { color?: { hex?: string }; style: string })}`);
        }

        const rowSpan = Math.min(merge?.rowSpan ?? 1, normalized.end.row - row + 1);
        const colSpan = Math.min(merge?.colSpan ?? 1, normalized.end.col - col + 1);

        payload.cells.push({
          colOffset: col - normalized.start.col,
          formula,
          rowOffset: row - normalized.start.row,
          value
        });

        if (rowSpan > 1 || colSpan > 1) {
          payload.merges.push({
            colOffset: col - normalized.start.col,
            colSpan,
            rowOffset: row - normalized.start.row,
            rowSpan
          });
        }

        textCells.push(value);
        htmlCells.push(
          `<td${rowSpan > 1 ? ` rowspan="${rowSpan}"` : ""}${colSpan > 1 ? ` colspan="${colSpan}"` : ""} style="${escapeHtml(cellStyles.join(";"))}">${escapeHtml(value)}</td>`
        );
      }

      rows.push(textCells.join("\t"));
      htmlRows.push(`<tr>${htmlCells.join("")}</tr>`);
    }

    return {
      html: `<table style="border-collapse:collapse">${htmlRows.join("")}</table>`,
      structured: JSON.stringify(payload),
      text: rows.join("\n")
    };
  }, [activeCell, getActiveWorksheet, getCellDisplayValue, selection]);

  const activeCellAddress = React.useMemo(() => (activeCell ? cellAddressToA1(activeCell) : null), [activeCell]);
  const selectedRangeAddress = React.useMemo(() => (selection ? rangeToA1(selection) : null), [selection]);
  const selectedValue = React.useMemo(() => getCellDisplayValue(activeCell), [activeCell, getCellDisplayValue, revision]);
  const selectedFormula = React.useMemo(() => getCellFormula(activeCell), [activeCell, getCellFormula, revision]);
  const canUndo = !readOnly && undoStackRef.current.length > 0;
  const canRedo = !readOnly && redoStackRef.current.length > 0;

  const createHistoryEntry = React.useCallback((): SnapshotHistoryEntry | null => {
    if (!workbook) {
      return null;
    }

    return {
      kind: "snapshot",
      activeCell,
      activeSheetIndex,
      bytes: cloneBytes(workbook.saveXlsxBytes()),
      selection
    };
  }, [activeCell, activeSheetIndex, selection, workbook]);

  const captureCellMutationState = React.useCallback((cell: XlsxCellAddress): CellMutationState | null => {
    const worksheet = getActiveWorksheet();
    if (!worksheet) {
      return null;
    }

    return {
      formula: worksheet.getFormulaAt(cell.row, cell.col) ?? null,
      value: worksheet.getCellAt(cell.row, cell.col).toJs()
    };
  }, [getActiveWorksheet]);

  const restoreHistoryEntry = React.useCallback(async (entry: SnapshotHistoryEntry) => {
    const wasmModule = await getSheetsWasmModule();
    const nextWorkbook = wasmModule.Workbook.fromBytes(cloneBytes(entry.bytes));
    const nextSheets = buildSheetList(nextWorkbook);
    const nextSheetIndex = Math.max(0, Math.min(entry.activeSheetIndex, Math.max(0, nextSheets.length - 1)));

    setError(null);
    setIsLoading(false);
    setWorkbook(nextWorkbook);
    setSheets(nextSheets);
    setActiveSheetIndexState(nextSheetIndex);
    setActiveCell(entry.activeCell);
    setSelection(entry.selection);
    selectionAnchorRef.current = entry.selection ? normalizeRange(entry.selection).start : entry.activeCell;
    setRevision((current) => current + 1);
  }, []);

  const applyCellEditHistoryEntry = React.useCallback((
    entry: CellEditHistoryEntry,
    direction: "undo" | "redo"
  ) => {
    if (!workbook) {
      return;
    }

    const worksheet = workbook.getSheet(entry.sheetIndex);
    const visibleSheetIndex = sheets.findIndex((sheet) => sheet.workbookSheetIndex === entry.sheetIndex);
    const targetState = direction === "undo" ? entry.before : entry.after;

    isApplyingHistoryRef.current = true;
    applyCellMutationState(worksheet, entry.cell, targetState);
    workbook.calculate();
    refreshWorkbookState(workbook);

    const nextActiveCell = direction === "undo" ? entry.activeCellBefore : entry.activeCellAfter;
    const nextSelection = direction === "undo" ? entry.selectionBefore : entry.selectionAfter;
    if (visibleSheetIndex >= 0) {
      setActiveSheetIndexState(visibleSheetIndex);
    }
    setActiveCell(nextActiveCell);
    setSelection(nextSelection);
    selectionAnchorRef.current = nextSelection ? normalizeRange(nextSelection).start : nextActiveCell;
    isApplyingHistoryRef.current = false;
  }, [refreshWorkbookState, sheets, workbook]);

  const applyRangeEditHistoryEntry = React.useCallback((
    entry: RangeEditHistoryEntry,
    direction: "undo" | "redo"
  ) => {
    if (!workbook) {
      return;
    }

    const worksheet = workbook.getSheet(entry.sheetIndex);
    const visibleSheetIndex = sheets.findIndex((sheet) => sheet.workbookSheetIndex === entry.sheetIndex);

    isApplyingHistoryRef.current = true;
    for (const mutation of entry.mutations) {
      applyCellMutationState(worksheet, mutation.cell, direction === "undo" ? mutation.before : mutation.after);
    }
    workbook.calculate();
    refreshWorkbookState(workbook);

    const nextActiveCell = direction === "undo" ? entry.activeCellBefore : entry.activeCellAfter;
    const nextSelection = direction === "undo" ? entry.selectionBefore : entry.selectionAfter;
    if (visibleSheetIndex >= 0) {
      setActiveSheetIndexState(visibleSheetIndex);
    }
    setActiveCell(nextActiveCell);
    setSelection(nextSelection);
    selectionAnchorRef.current = nextSelection ? normalizeRange(nextSelection).start : nextActiveCell;
    isApplyingHistoryRef.current = false;
  }, [refreshWorkbookState, sheets, workbook]);

  const recordHistoryBeforeMutation = React.useCallback(() => {
    if (isApplyingHistoryRef.current) {
      return;
    }

    const snapshot = createHistoryEntry();
    if (!snapshot) {
      return;
    }

    pushHistoryEntry(undoStackRef.current, snapshot);
    redoStackRef.current = [];
    setHistoryRevision((current) => current + 1);
  }, [createHistoryEntry]);

  const recordCellEditHistory = React.useCallback((
    cell: XlsxCellAddress,
    before: CellMutationState,
    after: CellMutationState
  ) => {
    if (!activeSheet || isApplyingHistoryRef.current) {
      return;
    }

    pushHistoryEntry(undoStackRef.current, {
      kind: "cell-edit",
      activeCellAfter: cell,
      activeCellBefore: activeCell,
      after,
      before,
      cell,
      selectionAfter: { start: cell, end: cell },
      selectionBefore: selection,
      sheetIndex: activeSheet.workbookSheetIndex
    });
    redoStackRef.current = [];
    setHistoryRevision((current) => current + 1);
  }, [activeCell, activeSheet, selection]);

  const recordRangeEditHistory = React.useCallback((
    mutations: RangeCellMutation[],
    selectionAfter: XlsxCellRange | null,
    activeCellAfter: XlsxCellAddress | null
  ) => {
    if (!activeSheet || isApplyingHistoryRef.current || mutations.length === 0) {
      return;
    }

    pushHistoryEntry(undoStackRef.current, {
      kind: "range-edit",
      activeCellAfter,
      activeCellBefore: activeCell,
      mutations,
      selectionAfter,
      selectionBefore: selection,
      sheetIndex: activeSheet.workbookSheetIndex
    });
    redoStackRef.current = [];
    setHistoryRevision((current) => current + 1);
  }, [activeCell, activeSheet, selection]);

  const download = React.useCallback(() => {
    if (file) {
      downloadArrayBuffer(file, displayFileName);
      return;
    }

    if (src) {
      downloadUrl(src, displayFileName);
    }
  }, [displayFileName, file, src]);

  const exportXlsx = React.useCallback(() => {
    if (!workbook) {
      return;
    }

    downloadBytes(workbook.saveXlsxBytes(), `${fileStem(displayFileName)}.xlsx`, XLSX_MIME_TYPE);
  }, [displayFileName, workbook]);

  const exportCsv = React.useCallback(() => {
    if (!workbook) {
      return;
    }

    const activeSheetName = activeSheet?.name ?? "sheet";
    downloadText(workbook.saveCsvString(), `${fileStem(displayFileName)}-${activeSheetName}.csv`, CSV_MIME_TYPE);
  }, [activeSheet?.name, displayFileName, workbook]);

  const recalculate = React.useCallback(() => {
    if (!workbook) {
      return;
    }

    workbook.calculate();
    refreshWorkbookState(workbook);
  }, [refreshWorkbookState, workbook]);

  const resizeColumn = React.useCallback((col: number, widthPx: number) => {
    if (readOnly || !workbook || !activeSheet) {
      return;
    }

    recordHistoryBeforeMutation();
    const worksheet = workbook.getSheet(activeSheet.workbookSheetIndex);
    worksheet.setColumnWidth(col, pxToSheetColumnWidth(widthPx));
    refreshWorkbookState(workbook);
  }, [activeSheet, readOnly, recordHistoryBeforeMutation, refreshWorkbookState, workbook]);

  const resizeRow = React.useCallback((row: number, heightPx: number) => {
    if (readOnly || !workbook || !activeSheet) {
      return;
    }

    recordHistoryBeforeMutation();
    const worksheet = workbook.getSheet(activeSheet.workbookSheetIndex);
    worksheet.setRowHeight(row, pxToSheetRowHeight(heightPx));
    refreshWorkbookState(workbook);
  }, [activeSheet, readOnly, recordHistoryBeforeMutation, refreshWorkbookState, workbook]);

  const selectCell = React.useCallback((cell: XlsxCellAddress, options?: { extend?: boolean }) => {
    setActiveCell(cell);
    if (options?.extend && selectionAnchorRef.current) {
      setSelection(normalizeRange({ start: selectionAnchorRef.current, end: cell }));
      return;
    }

    selectionAnchorRef.current = cell;
    setSelection({ start: cell, end: cell });
  }, []);

  const selectRange = React.useCallback((range: XlsxCellRange) => {
    const normalized = normalizeRange(range);
    selectionAnchorRef.current = normalized.start;
    setActiveCell(normalized.end);
    setSelection(normalized);
  }, []);

  const clearSelection = React.useCallback(() => {
    selectionAnchorRef.current = null;
    setActiveCell(null);
    setSelection(null);
  }, []);

  const clearSelectedCells = React.useCallback(() => {
    const worksheet = getActiveWorksheet();
    const targetRange = selection ?? (activeCell ? { start: activeCell, end: activeCell } : null);
    if (readOnly || !worksheet || !workbook || !targetRange) {
      return;
    }

    const normalized = normalizeRange(targetRange);
    const mutations: RangeCellMutation[] = [];
    for (let row = normalized.start.row; row <= normalized.end.row; row += 1) {
      for (let col = normalized.start.col; col <= normalized.end.col; col += 1) {
        if (worksheet.isMergedSecondary(row, col)) {
          continue;
        }

        const cell = { row, col };
        const before = captureCellMutationState(cell);
        if (!before) {
          continue;
        }

        worksheet.setCell(cellAddressToA1({ row, col }), "");
        mutations.push({
          after: { formula: null, value: "" },
          before,
          cell
        });
      }
    }

    workbook.calculate();
    refreshWorkbookState(workbook);
    recordRangeEditHistory(mutations, normalized, activeCell ?? normalized.start);
  }, [
    activeCell,
    captureCellMutationState,
    getActiveWorksheet,
    readOnly,
    recordRangeEditHistory,
    refreshWorkbookState,
    selection,
    workbook
  ]);

  const setCellValue = React.useCallback((cell: XlsxCellAddress, value: string) => {
    const worksheet = getActiveWorksheet();
    if (readOnly || !worksheet || !workbook) {
      return;
    }

    const before = captureCellMutationState(cell);
    if (!before) {
      return;
    }

    worksheet.setCell(cellAddressToA1(cell), value);
    workbook.calculate();
    refreshWorkbookState(workbook);
    recordCellEditHistory(cell, before, { formula: null, value });
  }, [captureCellMutationState, getActiveWorksheet, readOnly, recordCellEditHistory, refreshWorkbookState, workbook]);

  const setCellFormula = React.useCallback((cell: XlsxCellAddress, formula: string) => {
    const worksheet = getActiveWorksheet();
    if (readOnly || !worksheet || !workbook) {
      return;
    }

    const before = captureCellMutationState(cell);
    if (!before) {
      return;
    }

    const trimmedFormula = formula.trim();
    if (!formula.trim()) {
      worksheet.setCell(cellAddressToA1(cell), "");
    } else {
      worksheet.setFormula(cellAddressToA1(cell), formula);
    }
    workbook.calculate();
    refreshWorkbookState(workbook);
    recordCellEditHistory(cell, before, {
      formula: trimmedFormula || null,
      value: trimmedFormula ? null : ""
    });
  }, [captureCellMutationState, getActiveWorksheet, readOnly, recordCellEditHistory, refreshWorkbookState, workbook]);

  const setSelectedCellValue = React.useCallback((value: string) => {
    if (!activeCell) {
      return;
    }

    setCellValue(activeCell, value);
  }, [activeCell, setCellValue]);

  const setSelectedCellFormula = React.useCallback((formula: string) => {
    if (!activeCell) {
      return;
    }

    setCellFormula(activeCell, formula);
  }, [activeCell, setCellFormula]);

  const fillSelection = React.useCallback((targetRange: XlsxCellRange) => {
    const worksheet = getActiveWorksheet();
    if (readOnly || !worksheet || !workbook || !selection) {
      return;
    }

    const sourceRange = normalizeRange(selection);
    const nextRange = normalizeRange(targetRange);
    const sourceHeight = sourceRange.end.row - sourceRange.start.row + 1;
    const sourceWidth = sourceRange.end.col - sourceRange.start.col + 1;

    if (sourceHeight <= 0 || sourceWidth <= 0) {
      return;
    }

    const mutations: RangeCellMutation[] = [];
    for (let row = nextRange.start.row; row <= nextRange.end.row; row += 1) {
      for (let col = nextRange.start.col; col <= nextRange.end.col; col += 1) {
        if (rangeContainsCell(sourceRange, { row, col })) {
          continue;
        }

        const targetCell = { row, col };
        const before = captureCellMutationState(targetCell);
        if (!before) {
          continue;
        }

        const sourceRow = sourceRange.start.row + ((row - nextRange.start.row) % sourceHeight);
        const sourceCol = sourceRange.start.col + ((col - nextRange.start.col) % sourceWidth);
        const sourceFormula = worksheet.getFormulaAt(sourceRow, sourceCol);

        if (sourceFormula) {
          worksheet.setFormula(cellAddressToA1(targetCell), sourceFormula);
          mutations.push({
            after: { formula: sourceFormula, value: null },
            before,
            cell: targetCell
          });
          continue;
        }

        const sourceValue = normalizeCellValue(worksheet.getCellAt(sourceRow, sourceCol).toJs());
        worksheet.setCell(cellAddressToA1(targetCell), sourceValue);
        mutations.push({
          after: { formula: null, value: sourceValue },
          before,
          cell: targetCell
        });
      }
    }

    workbook.calculate();
    refreshWorkbookState(workbook);
    setSelection(nextRange);
    setActiveCell(nextRange.end);
    selectionAnchorRef.current = nextRange.start;
    recordRangeEditHistory(mutations, nextRange, nextRange.end);
  }, [captureCellMutationState, getActiveWorksheet, readOnly, recordRangeEditHistory, refreshWorkbookState, selection, workbook]);

  const mergeSelection = React.useCallback(() => {
    const worksheet = getActiveWorksheet();
    if (readOnly || !worksheet || !selection || !workbook) {
      return;
    }

    recordHistoryBeforeMutation();
    worksheet.mergeCells(rangeToA1(selection));
    refreshWorkbookState(workbook);
  }, [getActiveWorksheet, readOnly, recordHistoryBeforeMutation, refreshWorkbookState, selection, workbook]);

  const unmergeSelection = React.useCallback(() => {
    const worksheet = getActiveWorksheet();
    if (readOnly || !worksheet || !selection || !workbook) {
      return;
    }

    recordHistoryBeforeMutation();
    worksheet.unmergeCells(rangeToA1(selection));
    refreshWorkbookState(workbook);
  }, [getActiveWorksheet, readOnly, recordHistoryBeforeMutation, refreshWorkbookState, selection, workbook]);

  const addSheet = React.useCallback((name?: string) => {
    if (readOnly || !workbook) {
      return;
    }

    recordHistoryBeforeMutation();
    const baseName = name?.trim() || "Sheet";
    let candidate = baseName;
    let counter = 2;
    while (workbook.sheetIndex(candidate) !== undefined) {
      candidate = `${baseName} ${counter}`;
      counter += 1;
    }

    workbook.addSheet(candidate);
    const nextSheets = buildSheetList(workbook);
    setSheets(nextSheets);
    const nextIndex = nextSheets.findIndex((sheet) => sheet.name === candidate);
    setActiveSheetIndexState(nextIndex >= 0 ? nextIndex : 0);
    setRevision((current) => current + 1);
  }, [readOnly, recordHistoryBeforeMutation, workbook]);

  const removeActiveSheet = React.useCallback(() => {
    if (readOnly || !workbook || !activeSheet) {
      return;
    }

    recordHistoryBeforeMutation();
    workbook.removeSheet(activeSheet.workbookSheetIndex);
    const nextSheets = buildSheetList(workbook);
    setSheets(nextSheets);
    setActiveSheetIndexState((current) => Math.max(0, Math.min(current, nextSheets.length - 1)));
    setRevision((current) => current + 1);
  }, [activeSheet, readOnly, recordHistoryBeforeMutation, workbook]);

  const defineNamedRange = React.useCallback((name: string, range?: XlsxCellRange | null) => {
    if (readOnly || !workbook) {
      return;
    }

    const trimmed = name.trim();
    if (!trimmed) {
      return;
    }

    const targetRange = range ?? selection;
    if (!targetRange) {
      return;
    }

    recordHistoryBeforeMutation();
    workbook.defineName(trimmed, rangeToA1(targetRange));
    setRevision((current) => current + 1);
  }, [readOnly, recordHistoryBeforeMutation, selection, workbook]);

  const pasteText = React.useCallback((text: string) => {
    const worksheet = getActiveWorksheet();
    const targetCell = activeCell ?? selection?.start ?? null;
    if (readOnly || !worksheet || !workbook || !targetCell || !text) {
      return false;
    }

    const grid = parseClipboardText(text);
    if (grid.length === 0 || grid.every((row) => row.length === 0)) {
      return false;
    }

    const mutations: RangeCellMutation[] = [];
    for (let rowIndex = 0; rowIndex < grid.length; rowIndex += 1) {
      const row = grid[rowIndex] ?? [];
      for (let colIndex = 0; colIndex < row.length; colIndex += 1) {
        const rawValue = row[colIndex] ?? "";
        const nextCell = {
          col: targetCell.col + colIndex,
          row: targetCell.row + rowIndex
        };
        const before = captureCellMutationState(nextCell);
        if (!before) {
          continue;
        }
        if (rawValue.startsWith("=") && rawValue.length > 1) {
          worksheet.setFormula(cellAddressToA1(nextCell), rawValue);
          mutations.push({
            after: { formula: rawValue, value: null },
            before,
            cell: nextCell
          });
        } else {
          worksheet.setCell(cellAddressToA1(nextCell), rawValue);
          mutations.push({
            after: { formula: null, value: rawValue },
            before,
            cell: nextCell
          });
        }
      }
    }

    workbook.calculate();
    refreshWorkbookState(workbook);
    const nextRange = normalizeRange({
      start: targetCell,
      end: {
        col: targetCell.col + Math.max(0, Math.max(...grid.map((row) => row.length), 1) - 1),
        row: targetCell.row + grid.length - 1
      }
    });
    setActiveCell(targetCell);
    setSelection(nextRange);
    selectionAnchorRef.current = targetCell;
    recordRangeEditHistory(mutations, nextRange, targetCell);
    return true;
  }, [activeCell, captureCellMutationState, getActiveWorksheet, readOnly, recordRangeEditHistory, refreshWorkbookState, selection, workbook]);

  const pasteStructuredClipboardData = React.useCallback((serializedPayload: string) => {
    const worksheet = getActiveWorksheet();
    const targetCell = activeCell ?? selection?.start ?? null;
    if (readOnly || !worksheet || !workbook || !targetCell || !serializedPayload) {
      return false;
    }

    let payload: ClipboardPayload;
    try {
      payload = JSON.parse(serializedPayload) as ClipboardPayload;
    } catch {
      return false;
    }

    if (!Array.isArray(payload.cells) || payload.cells.length === 0) {
      return false;
    }

    const hasMergeOperations = Array.isArray(payload.merges) && payload.merges.some((merge) => (merge.rowSpan ?? 1) > 1 || (merge.colSpan ?? 1) > 1);
    const mutations: RangeCellMutation[] = [];
    if (hasMergeOperations) {
      recordHistoryBeforeMutation();
    }
    for (const cell of payload.cells) {
      const nextCell = {
        col: targetCell.col + cell.colOffset,
        row: targetCell.row + cell.rowOffset
      };
      const before = hasMergeOperations ? null : captureCellMutationState(nextCell);

      if (cell.formula) {
        worksheet.setFormula(cellAddressToA1(nextCell), cell.formula);
        if (before) {
          mutations.push({
            after: { formula: cell.formula, value: null },
            before,
            cell: nextCell
          });
        }
      } else {
        worksheet.setCell(cellAddressToA1(nextCell), cell.value);
        if (before) {
          mutations.push({
            after: { formula: null, value: cell.value },
            before,
            cell: nextCell
          });
        }
      }
    }

    if (Array.isArray(payload.merges)) {
      for (const merge of payload.merges) {
        if ((merge.rowSpan ?? 1) <= 1 && (merge.colSpan ?? 1) <= 1) {
          continue;
        }

        const mergeRange = normalizeRange({
          start: {
            col: targetCell.col + merge.colOffset,
            row: targetCell.row + merge.rowOffset
          },
          end: {
            col: targetCell.col + merge.colOffset + merge.colSpan - 1,
            row: targetCell.row + merge.rowOffset + merge.rowSpan - 1
          }
        });
        worksheet.mergeCells(rangeToA1(mergeRange));
      }
    }

    workbook.calculate();
    refreshWorkbookState(workbook);
    const nextRange = normalizeRange({
      start: targetCell,
      end: {
        col: targetCell.col + Math.max((payload.cols ?? 1) - 1, 0),
        row: targetCell.row + Math.max((payload.rows ?? 1) - 1, 0)
      }
    });
    setActiveCell(targetCell);
    setSelection(nextRange);
    selectionAnchorRef.current = targetCell;
    if (!hasMergeOperations) {
      recordRangeEditHistory(mutations, nextRange, targetCell);
    }
    return true;
  }, [
    activeCell,
    captureCellMutationState,
    getActiveWorksheet,
    readOnly,
    recordHistoryBeforeMutation,
    recordRangeEditHistory,
    refreshWorkbookState,
    selection,
    workbook
  ]);

  const copySelectionToClipboard = React.useCallback(async () => {
    const clipboardData = getClipboardData();
    if (!clipboardData || typeof navigator === "undefined" || !navigator.clipboard) {
      return false;
    }

    if (typeof ClipboardItem === "function" && navigator.clipboard.write) {
      const item = new ClipboardItem({
        [INTERNAL_CLIPBOARD_MIME]: new Blob([clipboardData.structured], { type: INTERNAL_CLIPBOARD_MIME }),
        "text/html": new Blob([clipboardData.html], { type: "text/html" }),
        "text/plain": new Blob([clipboardData.text], { type: "text/plain" })
      });
      await navigator.clipboard.write([item]);
      return true;
    }

    await navigator.clipboard.writeText(clipboardData.text);
    return true;
  }, [getClipboardData]);

  const pasteFromClipboard = React.useCallback(async () => {
    if (readOnly || typeof navigator === "undefined" || !navigator.clipboard) {
      return false;
    }

    if (navigator.clipboard.read) {
      const items = await navigator.clipboard.read();
      for (const item of items) {
        if (item.types.includes(INTERNAL_CLIPBOARD_MIME)) {
          const blob = await item.getType(INTERNAL_CLIPBOARD_MIME);
          return pasteStructuredClipboardData(await blob.text());
        }
      }

      for (const item of items) {
        if (item.types.includes("text/plain")) {
          const blob = await item.getType("text/plain");
          return pasteText(await blob.text());
        }
      }
    }

    return pasteText(await navigator.clipboard.readText());
  }, [pasteStructuredClipboardData, pasteText, readOnly]);

  const undo = React.useCallback(() => {
    if (readOnly || !workbook || undoStackRef.current.length === 0) {
      return;
    }

    const entry = undoStackRef.current.pop();
    if (!entry) {
      return;
    }

    if (entry.kind === "cell-edit") {
      pushHistoryEntry(redoStackRef.current, entry);
      setHistoryRevision((current) => current + 1);
      applyCellEditHistoryEntry(entry, "undo");
      return;
    }

    if (entry.kind === "range-edit") {
      pushHistoryEntry(redoStackRef.current, entry);
      setHistoryRevision((current) => current + 1);
      applyRangeEditHistoryEntry(entry, "undo");
      return;
    }

    const currentSnapshot = createHistoryEntry();
    if (currentSnapshot) {
      pushHistoryEntry(redoStackRef.current, currentSnapshot);
    }
    setHistoryRevision((current) => current + 1);
    void restoreHistoryEntry(entry);
  }, [applyCellEditHistoryEntry, applyRangeEditHistoryEntry, createHistoryEntry, readOnly, restoreHistoryEntry, workbook]);

  const redo = React.useCallback(() => {
    if (readOnly || !workbook || redoStackRef.current.length === 0) {
      return;
    }

    const entry = redoStackRef.current.pop();
    if (!entry) {
      return;
    }

    if (entry.kind === "cell-edit") {
      pushHistoryEntry(undoStackRef.current, entry);
      setHistoryRevision((current) => current + 1);
      applyCellEditHistoryEntry(entry, "redo");
      return;
    }

    if (entry.kind === "range-edit") {
      pushHistoryEntry(undoStackRef.current, entry);
      setHistoryRevision((current) => current + 1);
      applyRangeEditHistoryEntry(entry, "redo");
      return;
    }

    const currentSnapshot = createHistoryEntry();
    if (currentSnapshot) {
      pushHistoryEntry(undoStackRef.current, currentSnapshot);
    }
    setHistoryRevision((current) => current + 1);
    void restoreHistoryEntry(entry);
  }, [applyCellEditHistoryEntry, applyRangeEditHistoryEntry, createHistoryEntry, readOnly, restoreHistoryEntry, workbook]);

  return React.useMemo(
    () => ({
      activeCell,
      activeCellAddress,
      activeSheet,
      activeSheetIndex,
      addSheet,
      canRedo,
      canDownload: Boolean(file ?? src),
      canExport: Boolean(workbook),
      canUndo,
      clearSelectedCells,
      clearSelection,
      copySelectionToClipboard,
      defineNamedRange,
      displayFileName,
      download,
      exportCsv,
      exportXlsx,
      error,
      fillSelection,
      file,
      getClipboardData,
      getCellDisplayValue,
      getCellFormula,
      getActiveWorksheet,
      isLoading,
      mergeSelection,
      pasteFromClipboard,
      pasteStructuredClipboardData,
      pasteText,
      removeActiveSheet,
      readOnly,
      recalculate,
      redo,
      revision,
      resizeColumn,
      resizeRow,
      setCellFormula,
      setCellValue,
      selectedFormula,
      selectedRangeAddress,
      selectedValue,
      selectCell,
      selectRange,
      selection,
      setActiveSheetIndex,
      setSelectedCellFormula,
      setSelectedCellValue,
      sheets,
      src,
      undo,
      unmergeSelection,
      workbook
    }),
    [
      activeCell,
      activeCellAddress,
      activeSheet,
      activeSheetIndex,
      addSheet,
      canRedo,
      canUndo,
      clearSelectedCells,
      copySelectionToClipboard,
      defineNamedRange,
      displayFileName,
      download,
      error,
      exportCsv,
      exportXlsx,
      file,
      fillSelection,
      getClipboardData,
      getCellDisplayValue,
      getCellFormula,
      getActiveWorksheet,
      historyRevision,
      isLoading,
      mergeSelection,
      pasteFromClipboard,
      pasteStructuredClipboardData,
      pasteText,
      removeActiveSheet,
      readOnly,
      recalculate,
      redo,
      revision,
      resizeColumn,
      resizeRow,
      setCellFormula,
      setCellValue,
      selectedFormula,
      selectedRangeAddress,
      selectedValue,
      selectCell,
      selectRange,
      selection,
      setActiveSheetIndex,
      setSelectedCellFormula,
      setSelectedCellValue,
      sheets,
      src,
      undo,
      unmergeSelection,
      workbook
    ]
  );
}
