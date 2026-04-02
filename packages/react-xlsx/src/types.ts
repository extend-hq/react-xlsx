import type * as React from "react";
import type { Workbook, Worksheet } from "@dukelib/sheets-wasm";

export interface XlsxSheetData {
  name: string;
  rowCount: number;
  colCount: number;
  visibleRows: number[];
  visibleCols: number[];
  colWidths: number[];
  rowHeights: number[];
  workbookSheetIndex: number;
}

export interface UseXlsxViewerControllerOptions {
  file?: ArrayBuffer;
  fileName?: string;
  src?: string;
}

export interface XlsxViewerController {
  activeSheet: XlsxSheetData | null;
  activeSheetIndex: number;
  canDownload: boolean;
  canExport: boolean;
  displayFileName: string;
  download: () => void;
  exportCsv: () => void;
  exportXlsx: () => void;
  error: Error | null;
  file?: ArrayBuffer;
  isLoading: boolean;
  recalculate: () => void;
  revision: number;
  resizeColumn: (col: number, widthPx: number) => void;
  resizeRow: (row: number, heightPx: number) => void;
  setActiveSheetIndex: (index: number) => void;
  sheets: XlsxSheetData[];
  src?: string;
  workbook: Workbook | null;
  getActiveWorksheet: () => Worksheet | null;
}

export interface XlsxViewerProviderProps extends UseXlsxViewerControllerOptions {
  children: React.ReactNode;
  controller?: XlsxViewerController;
}

export interface XlsxViewerProps extends UseXlsxViewerControllerOptions {
  className?: string;
  controller?: XlsxViewerController;
  emptyState?: React.ReactNode;
  errorState?: React.ReactNode | ((error: Error) => React.ReactNode);
  height?: React.CSSProperties["height"];
  loadingState?: React.ReactNode;
  rounded?: boolean;
  showDefaultToolbar?: boolean;
  toolbar?: React.ReactNode | ((controller: XlsxViewerController) => React.ReactNode);
}
