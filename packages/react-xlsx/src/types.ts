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
  /** The built-in image element that react-xlsx would render without customization. */
  defaultNode: React.ReactNode;
  /** Workbook image metadata, including source, anchor, name, alt text, and editability. */
  image: XlsxImage;
  /** The image rectangle in viewer pixels, including the current zoom level. */
  rect: XlsxImageRect;
  /** Absolute positioning styles to apply when rendering a replacement image node. */
  style: React.CSSProperties;
}

export interface XlsxImageSelectionRenderProps {
  /** The built-in selected-image outline and resize handles. */
  defaultNode: React.ReactNode;
  /** Returns pointer handlers and styles for a custom resize handle. */
  getHandleProps: (
    position: XlsxImageResizeHandlePosition
  ) => {
    onPointerDown: (event: React.PointerEvent<HTMLElement>) => void;
    style: React.CSSProperties;
  };
  /** The currently selected image. */
  image: XlsxImage;
  /** The selected image rectangle in viewer pixels, including the current zoom level. */
  rect: XlsxImageRect;
}

export interface XlsxChartLoadingRenderProps {
  /** The chart that is waiting for its renderer or data. */
  chart: XlsxChart;
  /** The built-in chart loading placeholder. */
  defaultNode: React.ReactNode;
  /** The chart rectangle in viewer pixels, including the current zoom level. */
  rect: XlsxImageRect;
}

export interface XlsxFileTooLargeRenderProps {
  /** The built-in file-too-large message. */
  defaultNode: React.ReactNode;
  /** File name displayed in the viewer UI. */
  displayFileName: string;
  /** Actual file size in bytes. */
  fileSizeBytes: number;
  /** Configured maximum file size in bytes. */
  maxFileSizeBytes: number;
}

export interface XlsxScrollerRenderProps {
  /** Workbook grid content that should be rendered inside the scrollable viewport. */
  children: React.ReactNode;
  /** Props that must be applied to the actual scrollable viewport element. */
  viewportProps: React.HTMLAttributes<HTMLDivElement> & {
    key: React.Key;
    ref: React.Ref<HTMLDivElement>;
    style: React.CSSProperties;
    tabIndex: number;
  };
}

export interface UseXlsxViewerControllerOptions {
  /**
   * Allows row and column resizing even while editing is disabled by `readOnly`.
   *
   * @default false
   */
  allowResizeInReadOnly?: boolean;
  /**
   * Defers loading until `continueDeferredLoad()` is called when the file is larger than this byte threshold.
   * Set to `0` to parse immediately.
   *
   * @default 0
   * @example
   * ```tsx
   * useXlsxViewerController({ file, deferLoadingAboveBytes: 10 * 1024 * 1024 })
   * ```
   */
  deferLoadingAboveBytes?: number;
  /**
   * Local workbook bytes to load. Use either `file` or `src`.
   *
   * @example
   * ```tsx
   * <XlsxViewer file={await uploadedFile.arrayBuffer()} fileName={uploadedFile.name} />
   * ```
   */
  file?: ArrayBuffer;
  /**
   * Optional display and download name for the workbook.
   *
   * @example
   * ```tsx
   * <XlsxViewer file={buffer} fileName="forecast.xlsx" />
   * ```
   */
  fileName?: string;
  /**
   * Rejects files larger than this byte limit and renders `fileTooLargeState`.
   *
   * @default 25 * 1024 * 1024
   */
  maxFileSizeBytes?: number;
  /**
   * Disables workbook edits, paste, fill, undo/redo, and other mutation actions.
   *
   * @default false
   */
  readOnly?: boolean;
  /**
   * Automatically enables read-only mode for files larger than this byte threshold.
   * Set to `0` to disable the automatic switch.
   *
   * @default 0
   */
  readOnlyAboveBytes?: number;
  /**
   * Includes hidden and very hidden workbook sheets in the sheet tabs and controller state.
   *
   * @default false
   */
  showHiddenSheets?: boolean;
  /**
   * Skips OOXML ZIP/XML parsing and relies on `@dukelib/sheets-wasm` metadata only.
   * This is automatically used for legacy `.xls` files.
   *
   * @default false
   */
  skipXmlParsing?: boolean;
  /**
   * Remote workbook URL to fetch and load. Use either `src` or `file`.
   *
   * @example
   * ```tsx
   * <XlsxViewer src="/workbooks/report.xlsx" />
   * ```
   */
  src?: string;
  /**
   * Parses supported workbook data in a Web Worker so large files do not block React rendering.
   *
   * @default true
   */
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
  getRowsBatchAsync?: (workbookSheetIndex: number, startRow: number, rowCount: number) => Promise<unknown[] | null>;
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

export type XlsxSheetThumbnailResolution =
  | number
  | {
      maxHeight?: number;
      maxWidth?: number;
    };

export interface UseXlsxViewerThumbnailsOptions {
  includeHeaders?: boolean;
  resolution?: XlsxSheetThumbnailResolution;
}

export interface XlsxSheetThumbnail {
  aspectRatio: number;
  contentHeight: number;
  contentWidth: number;
  height: number;
  paint: (canvas: HTMLCanvasElement | null) => boolean;
  sheet: XlsxSheetData;
  sheetIndex: number;
  sourceRange: XlsxCellRange;
  width: number;
  workbookSheetIndex: number;
}

export interface XlsxViewerThumbnails {
  paintThumbnail: (sheetIndex: number, canvas: HTMLCanvasElement | null) => boolean;
  thumbnails: XlsxSheetThumbnail[];
}

export interface XlsxTableHeaderMenuRenderProps {
  /** Address of the table header cell that owns this menu. */
  cell: XlsxCellAddress;
  /** Table column metadata for the active header. */
  column: XlsxTableColumn;
  /** Current sort direction for this column, or `null` when unsorted. */
  direction: XlsxTableSortDirection | null;
  /** Sorts the table by this column in ascending order. */
  sortAscending: () => void;
  /** Sorts the table by this column in descending order. */
  sortDescending: () => void;
  /** Table metadata for the active header. */
  table: XlsxTable;
  /** Built-in trigger text/icon. Useful when composing your own trigger button. */
  triggerIcon: string;
  /** Props that must be applied to your menu trigger button so grid selection does not receive the click. */
  triggerProps: React.ButtonHTMLAttributes<HTMLButtonElement>;
}

export interface XlsxCellStyleContext {
  /** Address of the cell being styled. */
  cell: XlsxCellAddress;
  /** True when the cell is part of a selected chart's highlighted source range. */
  hasChartHighlight: boolean;
  /** True when a conditional format (color scale, data bar, or icon set) applies to the cell. */
  hasConditionalFormat: boolean;
  /** True when the cell has a hyperlink. */
  hasHyperlink: boolean;
  /** True when the cell has a data validation rule. */
  hasValidation: boolean;
  /** True when the cell is the anchor of a merged range. */
  isMerged: boolean;
  /** True when the cell is a table header cell. */
  isTableHeader: boolean;
  /**
   * The fully resolved style the viewer computed for this cell, combining the workbook's
   * own formatting with the viewer's built-in styling. Treat this as read-only; returning
   * a partial style from `getCellStyle` merges on top of it.
   */
  resolvedStyle: React.CSSProperties;
  /** Display name of the sheet the cell belongs to. */
  sheetName: string;
  /** The cell's resolved display value. */
  value: string;
  /** Workbook sheet index of the cell's sheet. */
  workbookSheetIndex: number;
}

export interface XlsxViewerProviderProps extends UseXlsxViewerControllerOptions {
  /** Viewer UI and hooks that should share this provider's workbook controller. */
  children: React.ReactNode;
  /**
   * Existing controller to provide instead of creating one from `file`, `src`, or other controller options.
   *
   * @example
   * ```tsx
   * const controller = useXlsxViewerController({ src: "/report.xlsx" });
   * return <XlsxViewerProvider controller={controller}><XlsxViewer /></XlsxViewerProvider>;
   * ```
   */
  controller?: XlsxViewerController;
  /**
   * Uses the built-in dark worksheet/viewer palette.
   *
   * @default false
   */
  isDark?: boolean;
}

export interface XlsxViewerProps extends UseXlsxViewerControllerOptions {
  /**
   * Allows row and column resizing even while editing is disabled by `readOnly`.
   *
   * @default false
   */
  allowResizeInReadOnly?: boolean;
  /** Class name applied to the root viewer shell. */
  className?: string;
  /**
   * Monotonic revision counter that forces the viewer to re-run `getCellStyle` and repaint cells
   * when it changes. Use this to animate or refresh custom cell styling while keeping a stable
   * `getCellStyle` callback: store transient styling state in a ref, update it (for example from a
   * `requestAnimationFrame` loop or a timer), and bump this number to flush the new styles.
   *
   * Changing the `getCellStyle` callback identity already triggers a re-resolve, so this prop is
   * only needed when the callback itself stays stable but its output should change over time.
   *
   * @example
   * ```tsx
   * const [styleRevision, setStyleRevision] = React.useState(0);
   * React.useEffect(() => {
   *   const id = window.setInterval(() => setStyleRevision((value) => value + 1), 16);
   *   return () => window.clearInterval(id);
   * }, []);
   * return <XlsxViewer file={buffer} getCellStyle={getCellStyle} cellStyleRevision={styleRevision} />;
   * ```
   */
  cellStyleRevision?: number;
  /**
   * Existing controller to render. Takes precedence over provider context and source props on this viewer.
   *
   * @example
   * ```tsx
   * const controller = useXlsxViewerController({ file, fileName: "model.xlsx" });
   * return <XlsxViewer controller={controller} />;
   * ```
   */
  controller?: XlsxViewerController;
  /** Content shown when no workbook has been loaded. */
  emptyState?: React.ReactNode;
  /**
   * Animates the selection rectangle in canvas mode.
   *
   * @default true
   */
  enableCanvasSelectionAnimation?: boolean;
  /**
   * Enables pinch zoom and modifier-key wheel zoom inside the viewer.
   *
   * @default true
   */
  enableGestureZoom?: boolean;
  /**
   * Renders worksheets with the canvas renderer. Set to `false` to use the DOM table renderer.
   *
   * @default true
   * @example
   * ```tsx
   * <XlsxViewer experimentalCanvas={false} />
   * ```
   */
  experimentalCanvas?: boolean;
  /** Content shown for non-size load errors, or a function that receives the thrown error. */
  errorState?: React.ReactNode | ((error: Error) => React.ReactNode);
  /** Content shown when `maxFileSizeBytes` rejects a file. */
  fileTooLargeState?: React.ReactNode | ((props: XlsxFileTooLargeRenderProps) => React.ReactNode);
  /**
   * Returns extra CSS style overrides for an individual cell. Called for every rendered cell
   * with the cell's address, sheet, resolved style, and contextual flags. Return `undefined`
   * (or `null`) to leave a cell untouched, or a partial style object that merges on top of the
   * viewer's resolved style. Use this as an escape hatch for custom per-cell styling such as
   * highlights, outlines, or status tints, without forking the workbook data.
   *
   * The DOM renderer applies every returned CSS property. The canvas renderer
   * (`experimentalCanvas`, the default) honors the subset it can paint: `backgroundColor`,
   * `backgroundImage` gradients, `color`, the four `border*` sides, `padding`, `textAlign`,
   * `textDecoration`, `textOverflow`, and font properties. CSS-only effects such as `boxShadow`,
   * `outline`, or `animation` apply in the DOM renderer (`experimentalCanvas={false}`).
   *
   * Keep this callback stable (e.g. wrap it in `useCallback`) so cell styling is not recomputed
   * on every render. When the callback identity changes, the viewer re-resolves and repaints
   * cell styles.
   *
   * @example
   * ```tsx
   * const getCellStyle = React.useCallback<NonNullable<XlsxViewerProps["getCellStyle"]>>(
   *   ({ cell }) => (cell.row === 3 && cell.col === 1 ? { backgroundColor: "#dbeafe" } : undefined),
   *   []
   * );
   * return <XlsxViewer file={buffer} getCellStyle={getCellStyle} />;
   * ```
   */
  getCellStyle?: (context: XlsxCellStyleContext) => React.CSSProperties | null | undefined;
  /**
   * CSS height for the viewer container.
   *
   * @example
   * ```tsx
   * <XlsxViewer height="70vh" />
   * ```
   */
  height?: React.CSSProperties["height"];
  /**
   * Uses the built-in dark worksheet/viewer palette.
   *
   * @default false
   */
  isDark?: boolean;
  /** Legacy loading element rendered while the workbook is being parsed. Prefer `loadingState` for new code. */
  loadingComponent?: React.ReactElement;
  /** Content shown while the workbook is being parsed. */
  loadingState?: React.ReactNode;
  /** Replaces the chart loading placeholder. */
  renderChartLoading?: (props: XlsxChartLoadingRenderProps) => React.ReactNode;
  /**
   * Replaces worksheet images while preserving their calculated rect and style.
   *
   * @example
   * ```tsx
   * <XlsxViewer renderImage={({ image, style }) => <img src={image.src} style={style} alt={image.description ?? ""} />} />
   * ```
   */
  renderImage?: (props: XlsxImageRenderProps) => React.ReactNode;
  /** Replaces the selected-image outline and resize handles. */
  renderImageSelection?: (props: XlsxImageSelectionRenderProps) => React.ReactNode;
  /**
   * Replaces the scroll viewport shell. Spread `viewportProps` onto the actual scrollable viewport element.
   *
   * @example
   * ```tsx
   * <XlsxViewer
   *   renderScroller={({ children, viewportProps }) => (
   *     <ScrollAreaPrimitive.Root className="h-full min-h-0 w-full min-w-0 flex-1">
   *       <ScrollAreaPrimitive.Viewport {...viewportProps}>{children}</ScrollAreaPrimitive.Viewport>
   *       <ScrollBar />
   *       <ScrollAreaPrimitive.Corner />
   *     </ScrollAreaPrimitive.Root>
   *   )}
   * />
   * ```
   */
  renderScroller?: (props: XlsxScrollerRenderProps) => React.ReactNode;
  /**
   * Toggles the default rounded outer shell.
   *
   * @default true
   */
  rounded?: boolean;
  /**
   * Disables workbook edits, paste, fill, undo/redo, and other mutation actions.
   *
   * @default false
   */
  readOnly?: boolean;
  /**
   * Border and accent color for the current selection.
   *
   * @example
   * ```tsx
   * <XlsxViewer selectionColor="#2563eb" />
   * ```
   */
  selectionColor?: string;
  /** Fill color used for the translucent selection overlay. */
  selectionFillColor?: string;
  /** Accent color used for selected row and column headers. */
  selectionHeaderColor?: string;
  /**
   * Replaces the table header sort/filter trigger menu.
   * Apply `triggerProps` to your actual trigger button so clicks do not leak into grid selection.
   */
  renderTableHeaderMenu?: (props: XlsxTableHeaderMenuRenderProps) => React.ReactNode;
  /**
   * Shows worksheet images, charts, shapes, and form controls.
   *
   * @default true
   */
  showImages?: boolean;
  /**
   * Shows the built-in toolbar above the workbook grid.
   *
   * @default true
   */
  showDefaultToolbar?: boolean;
  /**
   * Replaces the toolbar area with a node or render function.
   *
   * @example
   * ```tsx
   * <XlsxViewer toolbar={(controller) => <button onClick={controller.download}>Download</button>} />
   * ```
   */
  toolbar?: React.ReactNode | ((controller: XlsxViewerController) => React.ReactNode);
}
