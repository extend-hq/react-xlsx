import type * as React from "react";
import type { Workbook, Worksheet } from "@dukelib/sheets-wasm";

export interface XlsxThemePalette {
  colorsByIndex: Record<number, string>;
  majorLatinFont?: string;
  minorLatinFont?: string;
}

export interface XlsxResolvedCellStyle {
  [key: string]: unknown;
  alignment?: Record<string, unknown>;
  border?: Record<string, Record<string, unknown>>;
  fill?: Record<string, unknown>;
  font?: Record<string, unknown>;
}

export interface XlsxTableStyleDefinition {
  [elementType: string]: XlsxResolvedCellStyle;
}

export interface XlsxConditionalFormatValueObject {
  type: string;
  value?: number;
}

export interface XlsxConditionalDataBarRule {
  color?: Record<string, unknown>;
  borderColor?: Record<string, unknown>;
  cfvos: XlsxConditionalFormatValueObject[];
  gradient?: boolean;
  kind: "dataBar";
  maxLength?: number;
  minLength?: number;
  priority: number;
  ranges: XlsxCellRange[];
  showValue?: boolean;
}

export interface XlsxConditionalFormatIcon {
  iconId: number;
  iconSet: string;
}

export interface XlsxConditionalIconSetRule {
  cfvos: XlsxConditionalFormatValueObject[];
  icons: XlsxConditionalFormatIcon[];
  kind: "iconSet";
  priority: number;
  ranges: XlsxCellRange[];
  reverse?: boolean;
  showValue?: boolean;
}

export type XlsxConditionalFormatRule = XlsxConditionalDataBarRule | XlsxConditionalIconSetRule;

export interface XlsxDataValidation {
  allowBlank?: boolean;
  errorMessage?: string;
  errorStyle?: string;
  inputMessage?: string;
  listSource?: string;
  ranges: XlsxCellRange[];
  showDropdown?: boolean;
  showErrorAlert?: boolean;
  showInputMessage?: boolean;
  validationType: string;
}

export interface XlsxFreezePanes {
  col: number;
  row: number;
}

export interface XlsxSheetData {
  cachedFormulaValues: Record<string, string>;
  colWidthOverridesPx: Record<number, number>;
  colStyleIds: Record<number, number>;
  hiddenCols?: number[];
  hiddenRows?: number[];
  conditionalFormatRules: XlsxConditionalFormatRule[];
  dataValidations: XlsxDataValidation[];
  name: string;
  defaultColWidthPx: number;
  defaultRowHeightPx: number;
  freezePanes: XlsxFreezePanes | null;
  hasHorizontalMerges: boolean;
  hasVerticalMerges: boolean;
  maxUsedCol: number;
  maxUsedRow: number;
  rowCount: number;
  colCount: number;
  rowHeightOverridesPx: Record<number, number>;
  rowStyleIds: Record<number, number>;
  namedCellStyleByName: Record<string, XlsxResolvedCellStyle>;
  styleById: Record<number, XlsxResolvedCellStyle>;
  tableStyleByName: Record<string, XlsxTableStyleDefinition>;
  visibleRows: number[];
  visibleCols: number[];
  colWidths: number[];
  rowHeights: number[];
  showGridLines: boolean;
  themePalette: XlsxThemePalette;
  workbookSheetIndex: number;
  zoomScale?: number;
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

export interface XlsxTableColumn {
  id: number;
  index: number;
  name: string;
}

export interface XlsxTableStyleInfo {
  name?: string;
  showColumnStripes?: boolean;
  showFirstColumn?: boolean;
  showLastColumn?: boolean;
  showRowStripes?: boolean;
}

export interface XlsxTable {
  columns: XlsxTableColumn[];
  displayName: string;
  end: XlsxCellAddress;
  headerRowCount: number;
  headerRowCellStyle?: string;
  name: string;
  reference: string;
  start: XlsxCellAddress;
  styleInfo?: XlsxTableStyleInfo;
  totalsRowCount: number;
  totalsRowShown: boolean;
}

export type XlsxTableSortDirection = "ascending" | "descending";

export interface XlsxTableSortState {
  columnIndex: number;
  direction: XlsxTableSortDirection;
  tableName: string;
}

export interface XlsxImageMarker {
  col: number;
  colOffsetEmu: number;
  row: number;
  rowOffsetEmu: number;
}

export type XlsxImageAnchor =
  | {
      from: XlsxImageMarker;
      kind: "one-cell";
      sizeEmu: {
        cx: number;
        cy: number;
      };
    }
  | {
      kind: "absolute";
      positionEmu: {
        x: number;
        y: number;
      };
      sizeEmu: {
        cx: number;
        cy: number;
      };
    }
  | {
      from: XlsxImageMarker;
      kind: "two-cell";
      to: XlsxImageMarker;
    };

export interface XlsxImage {
  anchor: XlsxImageAnchor;
  description?: string;
  editable?: boolean;
  hyperlink?: string;
  id: string;
  mimeType: string;
  name?: string;
  sheetIndex: number;
  src: string;
  workbookSheetIndex: number;
  zIndex: number;
}

export interface XlsxShapeFill {
  color?: string;
  none?: boolean;
  opacity?: number;
}

export interface XlsxShapeStroke {
  color?: string;
  dash?: string;
  headEndType?: string;
  none?: boolean;
  opacity?: number;
  tailEndType?: string;
  widthPx?: number;
}

export interface XlsxShapeTextRun {
  bold?: boolean;
  color?: string;
  fontFamily?: string;
  fontSizePt?: number;
  italic?: boolean;
  text: string;
  underline?: boolean;
}

export interface XlsxShapeParagraph {
  align?: "center" | "justify" | "left" | "right";
  runs: XlsxShapeTextRun[];
}

export interface XlsxShapeTextBox {
  insetPx?: {
    bottom: number;
    left: number;
    right: number;
    top: number;
  };
  verticalAlign?: "bottom" | "middle" | "top";
}

export interface XlsxShape {
  anchor: XlsxImageAnchor;
  description?: string;
  fill?: XlsxShapeFill;
  flipH?: boolean;
  flipV?: boolean;
  geometry: string;
  geometryAdjustments?: Record<string, number>;
  hyperlink?: string;
  id: string;
  name?: string;
  paragraphs: XlsxShapeParagraph[];
  rotationDeg?: number;
  scaleX?: number;
  scaleY?: number;
  sheetIndex: number;
  svgPath?: string;
  svgViewBox?: {
    height: number;
    width: number;
  };
  stroke?: XlsxShapeStroke;
  textBox?: XlsxShapeTextBox;
  workbookSheetIndex: number;
  zIndex: number;
}

export interface XlsxImageRect {
  height: number;
  left: number;
  top: number;
  width: number;
}

export type XlsxImageResizeHandlePosition = "n" | "s" | "e" | "w" | "ne" | "nw" | "se" | "sw";

export interface XlsxImageRenderProps {
  defaultNode: React.ReactNode;
  image: XlsxImage;
  rect: XlsxImageRect;
  style: React.CSSProperties;
}

export interface XlsxImageSelectionRenderProps {
  defaultNode: React.ReactNode;
  getHandleProps: (
    position: XlsxImageResizeHandlePosition
  ) => {
    onPointerDown: (event: React.PointerEvent<HTMLElement>) => void;
    style: React.CSSProperties;
  };
  image: XlsxImage;
  rect: XlsxImageRect;
}

export interface UseXlsxViewerControllerOptions {
  deferLoadingAboveBytes?: number;
  file?: ArrayBuffer;
  fileName?: string;
  readOnly?: boolean;
  readOnlyAboveBytes?: number;
  src?: string;
  useWorker?: boolean;
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
  clearSelectedImage: () => void;
  getImageById: (id: string) => XlsxImage | null;
  getSheetImages: (sheetIndex?: number) => XlsxImage[];
  getSheetShapes: (sheetIndex?: number) => XlsxShape[];
  getClipboardData: () => XlsxClipboardData | null;
  getCellDisplayValue: (cell?: XlsxCellAddress | null) => string;
  getCellFormula: (cell?: XlsxCellAddress | null) => string;
  getCellSnapshotAsync?: (workbookSheetIndex: number, row: number, col: number) => Promise<{
    displayValue: string;
    formula: string;
  }>;
  isLoadDeferred: boolean;
  isLoading: boolean;
  isWorkerBacked?: boolean;
  images: XlsxImage[];
  shapes: XlsxShape[];
  mergeSelection: () => void;
  moveImageBy: (id: string, deltaX: number, deltaY: number) => void;
  removeActiveSheet: () => void;
  readOnly: boolean;
  recalculate: () => void;
  revision: number;
  resizeImageBy: (
    id: string,
    handle: XlsxImageResizeHandlePosition,
    deltaX: number,
    deltaY: number
  ) => void;
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
  selectedImage: XlsxImage | null;
  selectedImageId: string | null;
  setSelectedCellFormula: (formula: string) => void;
  setSelectedCellValue: (value: string) => void;
  sheets: XlsxSheetData[];
  src?: string;
  sortState: XlsxTableSortState | null;
  sortTable: (tableName: string, columnIndex: number, direction: XlsxTableSortDirection) => void;
  selectImage: (id: string | null) => void;
  setImageRect: (id: string, rect: XlsxImageRect) => void;
  getRowsBatchAsync?: (workbookSheetIndex: number, startRow: number, rowCount: number) => Promise<unknown[] | null>;
  tables: XlsxTable[];
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

export interface XlsxViewerTables {
  sortState: XlsxTableSortState | null;
  sortTable: (tableName: string, columnIndex: number, direction: XlsxTableSortDirection) => void;
  tables: XlsxTable[];
}

export interface XlsxViewerImages {
  clearSelectedImage: () => void;
  getImageById: (id: string) => XlsxImage | null;
  getSheetImages: (sheetIndex?: number) => XlsxImage[];
  images: XlsxImage[];
  moveImageBy: (id: string, deltaX: number, deltaY: number) => void;
  readOnly: boolean;
  resizeImageBy: (
    id: string,
    handle: XlsxImageResizeHandlePosition,
    deltaX: number,
    deltaY: number
  ) => void;
  selectedImage: XlsxImage | null;
  selectedImageId: string | null;
  selectImage: (id: string | null) => void;
  setImageRect: (id: string, rect: XlsxImageRect) => void;
}

export interface XlsxTableHeaderMenuRenderProps {
  close: () => void;
  column: XlsxTableColumn;
  direction: XlsxTableSortDirection | null;
  sortAscending: () => void;
  sortDescending: () => void;
  table: XlsxTable;
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
  renderImage?: (props: XlsxImageRenderProps) => React.ReactNode;
  renderImageSelection?: (props: XlsxImageSelectionRenderProps) => React.ReactNode;
  rounded?: boolean;
  readOnly?: boolean;
  selectionColor?: string;
  selectionFillColor?: string;
  selectionHeaderColor?: string;
  renderTableHeaderMenu?: (props: XlsxTableHeaderMenuRenderProps) => React.ReactNode;
  showImages?: boolean;
  showDefaultToolbar?: boolean;
  toolbar?: React.ReactNode | ((controller: XlsxViewerController) => React.ReactNode);
}
