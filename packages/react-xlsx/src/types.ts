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
  cellControl?: {
    kind: "checkbox";
  };
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
  axisColor?: Record<string, unknown>;
  border?: boolean;
  color?: Record<string, unknown>;
  borderColor?: Record<string, unknown>;
  cfvos: XlsxConditionalFormatValueObject[];
  gradient?: boolean;
  kind: "dataBar";
  maxLength?: number;
  minLength?: number;
  negativeBarBorderColorSameAsPositive?: boolean;
  negativeBorderColor?: Record<string, unknown>;
  negativeFillColor?: Record<string, unknown>;
  priority: number;
  ranges: XlsxCellRange[];
  showValue?: boolean;
}

export interface XlsxConditionalColorScaleRule {
  cfvos: XlsxConditionalFormatValueObject[];
  colors: Record<string, unknown>[];
  kind: "colorScale";
  priority: number;
  ranges: XlsxCellRange[];
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

export type XlsxConditionalFormatRule =
  | XlsxConditionalColorScaleRule
  | XlsxConditionalDataBarRule
  | XlsxConditionalIconSetRule;

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

export type XlsxSheetVisibility = "hidden" | "veryHidden" | "visible";

export interface XlsxSheetData {
  cachedFormulaValues: Record<string, string>;
  colWidthOverridesPx: Record<number, number>;
  colStyleIds: Record<number, number>;
  hiddenCols?: number[];
  hiddenRows?: number[];
  conditionalFormatRules: XlsxConditionalFormatRule[];
  dataValidations: XlsxDataValidation[];
  name: string;
  visibility: XlsxSheetVisibility;
  columnWidthCharacterWidthPx?: number;
  defaultColWidthPx: number;
  defaultRowHeightPx: number;
  freezePanes: XlsxFreezePanes | null;
  hasHorizontalMerges: boolean;
  hasVerticalMerges: boolean;
  maxHorizontalMergeEndCol: number;
  maxVerticalMergeEndRow: number;
  minUsedCol: number;
  minUsedRow: number;
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
  sparklines: XlsxSparkline[];
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

export interface XlsxSparkline {
  color?: string;
  firstColor?: string;
  highColor?: string;
  lastColor?: string;
  lowColor?: string;
  markerColor?: string;
  markers?: boolean;
  negative?: boolean;
  negativeColor?: string;
  range: XlsxCellRange;
  sheetName?: string;
  target: XlsxCellAddress;
  type: "column" | "line" | "winLoss";
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
  mediaPath?: string;
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
  horizontalAlign?: "center" | "left";
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
  hidden?: boolean;
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

export type XlsxRowsBatchResult =
  | {
      rows: unknown[];
      stylesById?: Record<number, Record<string, unknown>>;
    }
  | unknown[]
  | null;

export type XlsxFormControlKind =
  | "button"
  | "checkbox"
  | "dropdown"
  | "editbox"
  | "group-box"
  | "label"
  | "listbox"
  | "radio"
  | "scrollbar"
  | "spinner"
  | "unknown";

export interface XlsxFormControl {
  anchor: XlsxImageAnchor;
  checked?: boolean;
  fontFamily?: string;
  fontSizePt?: number;
  hidden?: boolean;
  id: string;
  kind: XlsxFormControlKind;
  label?: string;
  linkedCell?: string;
  name?: string;
  sheetIndex: number;
  textAlign?: "center" | "left" | "right";
  textColor?: string;
  workbookSheetIndex: number;
  zIndex: number;
}

export interface XlsxChartReference {
  formula?: string;
  refType?: string;
  values?: Array<number | string | null>;
}

export interface XlsxChartDataLabels {
  pointLabels?: XlsxChartPointDataLabel[];
  raw?: Record<string, unknown>;
  showBubbleSize?: boolean;
  showCategoryName?: boolean;
  showLegendKey?: boolean;
  showPercent?: boolean;
  showSeriesName?: boolean;
  showValue?: boolean;
}

export interface XlsxChartPointDataLabel {
  deleted?: boolean;
  fontSizePt?: number;
  index: number;
  showBubbleSize?: boolean;
  showCategoryName?: boolean;
  showPercent?: boolean;
  showSeriesName?: boolean;
  showValue?: boolean;
  x?: number;
  y?: number;
}

export interface XlsxChartLegend {
  overlay?: boolean;
  position?: string;
  raw?: Record<string, unknown>;
}

export interface XlsxChartAxis {
  crossId?: number;
  crosses?: string;
  crossBetween?: string;
  delete?: boolean;
  id?: number;
  labelPosition?: string;
  logBase?: number;
  orientation?: string;
  majorUnit?: number;
  max?: number;
  min?: number;
  majorGridlines?: boolean;
  majorTickMark?: string;
  minorUnit?: number;
  minorGridlines?: boolean;
  minorTickMark?: string;
  numberFormat?: {
    formatCode?: string;
    sourceLinked?: boolean;
  };
  position?: string;
  raw?: Record<string, unknown>;
  shapeProperties?: Record<string, unknown>;
  tickLabelSkip?: number;
  tickMarkSkip?: number;
}

export interface XlsxChartPointStyle {
  color?: string;
  explosion?: number;
  index: number;
  lineColor?: string;
}

export interface XlsxChartSeries {
  bubbleSizeRef?: XlsxChartReference | null;
  bubbleSizes?: Array<number | null>;
  categories: Array<number | string | null>;
  categoriesRef?: XlsxChartReference | null;
  color?: string;
  dataPoints: unknown[];
  dataPointStyles?: XlsxChartPointStyle[];
  formatIdx?: number;
  hidden?: boolean;
  id: string;
  invertIfNegative?: boolean;
  lineColor?: string;
  lineWidthPx?: number;
  marker?: Record<string, unknown>;
  markerColor?: string;
  markerLineColor?: string;
  markerSize?: number;
  markerSymbol?: string;
  name?: string;
  negativeColor?: string;
  negativeLineColor?: string;
  raw?: Record<string, unknown>;
  shapeProperties?: Record<string, unknown>;
  smooth?: boolean;
  values: Array<number | null>;
  valuesRef?: XlsxChartReference | null;
}

export interface XlsxChartTypeGroup {
  axisIds?: number[];
  chartType: string;
  dataLabels?: XlsxChartDataLabels | null;
  gapWidth?: number;
  is3d?: boolean;
  overlap?: number;
  raw?: Record<string, unknown>;
  series: XlsxChartSeries[];
  varyColors?: boolean;
}

export interface XlsxChartWall {
  fillColor?: string;
  hidden?: boolean;
  lineColor?: string;
  thickness?: number;
}

export interface XlsxChart {
  anchor: XlsxImageAnchor;
  autoTitleDeleted?: boolean;
  axes: XlsxChartAxis[];
  axisLabelColor?: string;
  axisLineColor?: string;
  categoryAxis?: XlsxChartAxis | null;
  chartExLayout?: string;
  chartAreaBorderColor?: string;
  chartAreaFillColor?: string;
  chartColorPalette?: string[];
  chartColorPaletteOffset?: number;
  chartPath?: string;
  chartStyleId?: number;
  chartType: string;
  dataLabels?: XlsxChartDataLabels | null;
  displayBlanksAs?: string;
  editable?: boolean;
  firstSliceAngle?: number;
  fontFamily?: string;
  gapWidth?: number;
  holeSize?: number;
  id: string;
  is3d?: boolean;
  legend?: XlsxChartLegend | null;
  name?: string;
  overlap?: number;
  plotVisibleOnly?: boolean;
  raw?: Record<string, unknown>;
  radarStyle?: string;
  scatterStyle?: string;
  roundedCorners?: boolean;
  shape3d?: string;
  seriesAxis?: XlsxChartAxis | null;
  series: XlsxChartSeries[];
  sheetIndex: number;
  showDlblsOverMax?: boolean;
  sideWall?: XlsxChartWall | null;
  backWall?: XlsxChartWall | null;
  bubbleScale?: number;
  bubble3d?: boolean;
  floor?: XlsxChartWall | null;
  surfaceMaterial?: string;
  textColor?: string;
  title?: string;
  titleColor?: string;
  titleFontFamily?: string;
  typeGroups?: XlsxChartTypeGroup[];
  valueAxis?: XlsxChartAxis | null;
  varyColors?: boolean;
  view3d?: {
    depthPercent?: number;
    perspective?: number;
    rAngAx?: boolean;
    rotX?: number;
    rotY?: number;
  };
  wireframe?: boolean;
  workbookSheetIndex: number;
  zIndex: number;
}

export interface XlsxChartsheet {
  chartIds: string[];
  chartPath?: string;
  id: string;
  index: number;
  name: string;
  raw?: Record<string, unknown>;
  workbookSheetIndex?: number;
}

export interface XlsxWorkbookTab {
  chartsheetIndex?: number;
  id: string;
  index: number;
  kind: "chartsheet" | "sheet";
  name: string;
  sheetIndex?: number;
  visibility?: XlsxSheetVisibility;
  workbookSheetIndex?: number;
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

export interface XlsxChartLoadingRenderProps {
  chart: XlsxChart;
  defaultNode: React.ReactNode;
  rect: XlsxImageRect;
}

export interface XlsxFileTooLargeRenderProps {
  defaultNode: React.ReactNode;
  displayFileName: string;
  fileSizeBytes: number;
  maxFileSizeBytes: number;
}

export interface UseXlsxViewerControllerOptions {
  deferLoadingAboveBytes?: number;
  file?: ArrayBuffer;
  fileName?: string;
  maxFileSizeBytes?: number;
  readOnly?: boolean;
  readOnlyAboveBytes?: number;
  showHiddenSheets?: boolean;
  skipXmlParsing?: boolean;
  src?: string;
  useWorker?: boolean;
}

export interface XlsxViewerController {
  activeCell: XlsxCellAddress | null;
  activeCellAddress: string | null;
  activeSheet: XlsxSheetData | null;
  activeSheetIndex: number;
  activeTab: XlsxWorkbookTab | null;
  activeTabIndex: number;
  canDownload: boolean;
  canExport: boolean;
  canLoadDeferred: boolean;
  canRedo: boolean;
  canUndo: boolean;
  canZoomIn: boolean;
  canZoomOut: boolean;
  clearSelectedCells: () => void;
  clearSelection: () => void;
  continueDeferredLoad: () => void;
  copySelectionToClipboard: () => Promise<boolean>;
  defaultZoomScale: number;
  deferredLoadFileSize: number | null;
  defineNamedRange: (name: string, range?: XlsxCellRange | null) => void;
  displayFileName: string;
  download: () => void;
  charts: XlsxChart[];
  chartsheets: XlsxChartsheet[];
  exportCsv: () => void;
  exportXlsx: () => void;
  error: Error | null;
  file?: ArrayBuffer;
  fillSelection: (targetRange: XlsxCellRange) => void;
  clearSelectedChart: () => void;
  clearSelectedImage: () => void;
  getChartById: (id: string) => XlsxChart | null;
  getChartsheetById: (id: string) => XlsxChartsheet | null;
  formControls: XlsxFormControl[];
  getSheetCharts: (sheetIndex?: number) => XlsxChart[];
  getSheetFormControls: (sheetIndex?: number) => XlsxFormControl[];
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
  isChartsLoading: boolean;
  isWorkerBacked?: boolean;
  images: XlsxImage[];
  moveChartBy: (id: string, deltaX: number, deltaY: number) => void;
  shapes: XlsxShape[];
  mergeSelection: () => void;
  maxZoomScale: number;
  minZoomScale: number;
  moveImageBy: (id: string, deltaX: number, deltaY: number) => void;
  removeActiveSheet: () => void;
  readOnly: boolean;
  recalculate: () => void;
  revision: number;
  resetZoom: () => void;
  resizeChartBy: (
    id: string,
    handle: XlsxImageResizeHandlePosition,
    deltaX: number,
    deltaY: number
  ) => void;
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
  setZoomScale: (zoomScale: number) => void;
  selectCell: (cell: XlsxCellAddress, options?: { extend?: boolean }) => void;
  selectChart: (id: string | null) => void;
  selectRange: (range: XlsxCellRange) => void;
  selection: XlsxCellRange | null;
  setActiveSheetIndex: (index: number) => void;
  setActiveTabIndex: (index: number) => void;
  selectedChart: XlsxChart | null;
  selectedChartId: string | null;
  selectedImage: XlsxImage | null;
  selectedImageId: string | null;
  setSelectedCellFormula: (formula: string) => void;
  setSelectedCellValue: (value: string) => void;
  sheets: XlsxSheetData[];
  src?: string;
  sortState: XlsxTableSortState | null;
  sortTable: (tableName: string, columnIndex: number, direction: XlsxTableSortDirection) => void;
  selectImage: (id: string | null) => void;
  setChartRect: (id: string, rect: XlsxImageRect) => void;
  setImageRect: (id: string, rect: XlsxImageRect) => void;
  getRowsBatchAsync?: (workbookSheetIndex: number, startRow: number, rowCount: number) => Promise<XlsxRowsBatchResult>;
  tables: XlsxTable[];
  tabs: XlsxWorkbookTab[];
  undo: () => void;
  unmergeSelection: () => void;
  updateChart: (id: string, patch: Partial<XlsxChart>) => void;
  workbook: Workbook | null;
  zoomIn: () => void;
  zoomOut: () => void;
  zoomScale: number;
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

export interface XlsxViewerZoom {
  canZoomIn: boolean;
  canZoomOut: boolean;
  defaultZoomScale: number;
  maxZoomScale: number;
  minZoomScale: number;
  resetZoom: () => void;
  setZoomScale: (zoomScale: number) => void;
  zoomIn: () => void;
  zoomOut: () => void;
  zoomScale: number;
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
  charts: XlsxChart[];
  clearSelectedChart: () => void;
  clearSelectedImage: () => void;
  getChartById: (id: string) => XlsxChart | null;
  getSheetCharts: (sheetIndex?: number) => XlsxChart[];
  getImageById: (id: string) => XlsxImage | null;
  getSheetImages: (sheetIndex?: number) => XlsxImage[];
  images: XlsxImage[];
  moveChartBy: (id: string, deltaX: number, deltaY: number) => void;
  moveImageBy: (id: string, deltaX: number, deltaY: number) => void;
  isChartsLoading: boolean;
  readOnly: boolean;
  resizeChartBy: (
    id: string,
    handle: XlsxImageResizeHandlePosition,
    deltaX: number,
    deltaY: number
  ) => void;
  resizeImageBy: (
    id: string,
    handle: XlsxImageResizeHandlePosition,
    deltaX: number,
    deltaY: number
  ) => void;
  selectedChart: XlsxChart | null;
  selectedChartId: string | null;
  selectedImage: XlsxImage | null;
  selectedImageId: string | null;
  selectChart: (id: string | null) => void;
  selectImage: (id: string | null) => void;
  setChartRect: (id: string, rect: XlsxImageRect) => void;
  setImageRect: (id: string, rect: XlsxImageRect) => void;
  updateChart: (id: string, patch: Partial<XlsxChart>) => void;
}

export interface XlsxViewerCharts {
  activeTab: XlsxWorkbookTab | null;
  activeTabIndex: number;
  charts: XlsxChart[];
  chartsheets: XlsxChartsheet[];
  clearSelectedChart: () => void;
  getChartById: (id: string) => XlsxChart | null;
  getChartsheetById: (id: string) => XlsxChartsheet | null;
  getSheetCharts: (sheetIndex?: number) => XlsxChart[];
  isChartsLoading: boolean;
  moveChartBy: (id: string, deltaX: number, deltaY: number) => void;
  readOnly: boolean;
  resizeChartBy: (
    id: string,
    handle: XlsxImageResizeHandlePosition,
    deltaX: number,
    deltaY: number
  ) => void;
  selectChart: (id: string | null) => void;
  selectedChart: XlsxChart | null;
  selectedChartId: string | null;
  setActiveTabIndex: (index: number) => void;
  setChartRect: (id: string, rect: XlsxImageRect) => void;
  tabs: XlsxWorkbookTab[];
  updateChart: (id: string, patch: Partial<XlsxChart>) => void;
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
  isDark?: boolean;
}

export interface XlsxViewerProps extends UseXlsxViewerControllerOptions {
  allowResizeInReadOnly?: boolean;
  className?: string;
  controller?: XlsxViewerController;
  emptyState?: React.ReactNode;
  enableCanvasSelectionAnimation?: boolean;
  enableGestureZoom?: boolean;
  experimentalCanvas?: boolean;
  errorState?: React.ReactNode | ((error: Error) => React.ReactNode);
  fileTooLargeState?: React.ReactNode | ((props: XlsxFileTooLargeRenderProps) => React.ReactNode);
  height?: React.CSSProperties["height"];
  isDark?: boolean;
  loadingComponent?: React.ReactElement;
  loadingState?: React.ReactNode;
  renderChartLoading?: (props: XlsxChartLoadingRenderProps) => React.ReactNode;
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
