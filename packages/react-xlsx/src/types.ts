import type * as React from "react";
import type { Workbook, Worksheet } from "@dukelib/sheets-wasm";

export interface XlsxSheetData {
  name: string;
  maxUsedCol: number;
  maxUsedRow: number;
  rowCount: number;
  colCount: number;
  visibleRows: number[];
  visibleCols: number[];
  colWidths: number[];
  rowHeights: number[];
  workbookSheetIndex: number;
}

export interface XlsxCellAddress {
  col: number;
  row: number;
}

export interface XlsxCellRange {
  end: XlsxCellAddress;
  start: XlsxCellAddress;
}

export interface XlsxClipboardData {
  html: string;
  structured: string;
  text: string;
}

export interface UseXlsxViewerControllerOptions {
  deferLoadingAboveBytes?: number;
  file?: ArrayBuffer;
  fileName?: string;
  readOnly?: boolean;
  src?: string;
}

export interface XlsxViewerController {
  activeCell: XlsxCellAddress | null;
  activeCellAddress: string | null;
  activeSheet: XlsxSheetData | null;
  activeSheetIndex: number;
  canDownload: boolean;
  canExport: boolean;
  canLoadDeferred: boolean;
  canRedo: boolean;
  canUndo: boolean;
  clearSelectedCells: () => void;
  clearSelection: () => void;
  continueDeferredLoad: () => void;
  copySelectionToClipboard: () => Promise<boolean>;
  deferredLoadFileSize: number | null;
  defineNamedRange: (name: string, range?: XlsxCellRange | null) => void;
  displayFileName: string;
  download: () => void;
  exportCsv: () => void;
  exportXlsx: () => void;
  error: Error | null;
  file?: ArrayBuffer;
  fillSelection: (targetRange: XlsxCellRange) => void;
  getClipboardData: () => XlsxClipboardData | null;
  getCellDisplayValue: (cell?: XlsxCellAddress | null) => string;
  getCellFormula: (cell?: XlsxCellAddress | null) => string;
  isLoadDeferred: boolean;
  isLoading: boolean;
  mergeSelection: () => void;
  removeActiveSheet: () => void;
  readOnly: boolean;
  recalculate: () => void;
  revision: number;
  resizeColumn: (col: number, widthPx: number) => void;
  resizeRow: (row: number, heightPx: number) => void;
  redo: () => void;
  pasteFromClipboard: () => Promise<boolean>;
  pasteStructuredClipboardData: (payload: string) => boolean;
  pasteText: (text: string) => boolean;
  selectedRangeAddress: string | null;
  selectedValue: string;
  selectedFormula: string;
  setCellFormula: (cell: XlsxCellAddress, formula: string) => void;
  setCellValue: (cell: XlsxCellAddress, value: string) => void;
  selectCell: (cell: XlsxCellAddress, options?: { extend?: boolean }) => void;
  selectRange: (range: XlsxCellRange) => void;
  selection: XlsxCellRange | null;
  setActiveSheetIndex: (index: number) => void;
  setSelectedCellFormula: (formula: string) => void;
  setSelectedCellValue: (value: string) => void;
  sheets: XlsxSheetData[];
  src?: string;
  undo: () => void;
  unmergeSelection: () => void;
  workbook: Workbook | null;
  getActiveWorksheet: () => Worksheet | null;
  addSheet: (name?: string) => void;
}

export interface XlsxViewerSelection {
  activeCell: XlsxCellAddress | null;
  activeCellAddress: string | null;
  clearSelection: () => void;
  selectedRangeAddress: string | null;
  selectCell: (cell: XlsxCellAddress, options?: { extend?: boolean }) => void;
  selectRange: (range: XlsxCellRange) => void;
  selection: XlsxCellRange | null;
}

export interface XlsxViewerEditing {
  addSheet: (name?: string) => void;
  canRedo: boolean;
  canUndo: boolean;
  clearSelectedCells: () => void;
  copySelectionToClipboard: () => Promise<boolean>;
  defineNamedRange: (name: string, range?: XlsxCellRange | null) => void;
  fillSelection: (targetRange: XlsxCellRange) => void;
  getClipboardData: () => XlsxClipboardData | null;
  getCellDisplayValue: (cell?: XlsxCellAddress | null) => string;
  getCellFormula: (cell?: XlsxCellAddress | null) => string;
  mergeSelection: () => void;
  pasteFromClipboard: () => Promise<boolean>;
  pasteStructuredClipboardData: (payload: string) => boolean;
  pasteText: (text: string) => boolean;
  removeActiveSheet: () => void;
  readOnly: boolean;
  redo: () => void;
  selectedFormula: string;
  selectedValue: string;
  setCellFormula: (cell: XlsxCellAddress, formula: string) => void;
  setCellValue: (cell: XlsxCellAddress, value: string) => void;
  setSelectedCellFormula: (formula: string) => void;
  setSelectedCellValue: (value: string) => void;
  undo: () => void;
  unmergeSelection: () => void;
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
  loadingComponent?: React.ReactElement;
  loadingState?: React.ReactNode;
  rounded?: boolean;
  readOnly?: boolean;
  selectionColor?: string;
  selectionFillColor?: string;
  selectionHeaderColor?: string;
  showDefaultToolbar?: boolean;
  toolbar?: React.ReactNode | ((controller: XlsxViewerController) => React.ReactNode);
}
