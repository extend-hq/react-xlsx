import * as React from "react";
import type { Workbook } from "@dukelib/sheets-wasm";
import { getSheetsWasmModule } from "./wasm";
import type { UseXlsxViewerControllerOptions, XlsxSheetData, XlsxViewerController } from "./types";

const FORMULA_COUNT_THRESHOLD = 1000;
const DEFAULT_ROW_HEIGHT = 24;
const DEFAULT_COL_WIDTH = 80;
const XLSX_MIME_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
const CSV_MIME_TYPE = "text/csv;charset=utf-8";
const MIN_COL_WIDTH_PX = 30;
const MIN_ROW_HEIGHT_PX = 16;

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

    const [minRow, minCol, maxRow, maxCol] = usedRange;
    const visibleRows: number[] = [];
    const visibleCols: number[] = [];

    for (let row = minRow; row <= maxRow; row += 1) {
      if (!worksheet.isRowHidden(row)) {
        visibleRows.push(row);
      }
    }

    for (let col = minCol; col <= maxCol; col += 1) {
      if (!worksheet.isColumnHidden(col)) {
        visibleCols.push(col);
      }
    }

    sheets.push({
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
  const { file, fileName, src } = options;
  const [isLoading, setIsLoading] = React.useState(Boolean(file ?? src));
  const [error, setError] = React.useState<Error | null>(null);
  const [workbook, setWorkbook] = React.useState<Workbook | null>(null);
  const [sheets, setSheets] = React.useState<XlsxSheetData[]>([]);
  const [activeSheetIndex, setActiveSheetIndexState] = React.useState(0);
  const [revision, setRevision] = React.useState(0);
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
      setRevision(0);
      return;
    }

    let isCurrent = true;
    setIsLoading(true);
    setError(null);
    setActiveSheetIndexState(0);
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
    if (!workbook || !activeSheet) {
      return;
    }

    const worksheet = workbook.getSheet(activeSheet.workbookSheetIndex);
    worksheet.setColumnWidth(col, pxToSheetColumnWidth(widthPx));
    refreshWorkbookState(workbook);
  }, [activeSheet, refreshWorkbookState, workbook]);

  const resizeRow = React.useCallback((row: number, heightPx: number) => {
    if (!workbook || !activeSheet) {
      return;
    }

    const worksheet = workbook.getSheet(activeSheet.workbookSheetIndex);
    worksheet.setRowHeight(row, pxToSheetRowHeight(heightPx));
    refreshWorkbookState(workbook);
  }, [activeSheet, refreshWorkbookState, workbook]);

  return React.useMemo(
    () => ({
      activeSheet,
      activeSheetIndex,
      canDownload: Boolean(file ?? src),
      canExport: Boolean(workbook),
      displayFileName,
      download,
      exportCsv,
      exportXlsx,
      error,
      file,
      getActiveWorksheet,
      isLoading,
      recalculate,
      revision,
      resizeColumn,
      resizeRow,
      setActiveSheetIndex,
      sheets,
      src,
      workbook
    }),
    [
      activeSheet,
      activeSheetIndex,
      displayFileName,
      download,
      error,
      exportCsv,
      exportXlsx,
      file,
      getActiveWorksheet,
      isLoading,
      recalculate,
      revision,
      resizeColumn,
      resizeRow,
      setActiveSheetIndex,
      sheets,
      src,
      workbook
    ]
  );
}
