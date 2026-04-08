import * as React from "react";
import type { Workbook } from "@dukelib/sheets-wasm";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import {
  loadWorkbookChartAssets,
  updateWorkbookChartAnchor,
  updateWorkbookChartDefinition,
  type WorkbookChartAssets
} from "./charts";
import { resolveWorkbookColor, resolveWorkbookFillStyle } from "./colors";
import {
  mergeWorkbookImageAssets,
  parseWorkbookImageAssets,
  pxToSheetColumnWidth,
  rectToImageAnchor,
  resolveContentSheetAxisPixels,
  resolveSheetColumnWidthPixels,
  resolveRenderedSheetAxisPixels,
  resolveSheetRowHeightPixels,
  resizeImageRect,
  revokeWorkbookImageAssets,
  updateWorkbookImageAnchor,
  type WorkbookImageAssets,
  type WorkbookImageSheetOrigin,
  type WorkbookTableMetadata
} from "./images";
import { getSheetsWasmModule } from "./wasm";
import { XlsxWorkerClient } from "./worker-client";
import type {
  UseXlsxViewerControllerOptions,
  XlsxChart,
  XlsxChartsheet,
  XlsxCellAddress,
  XlsxCellRange,
  XlsxClipboardData,
  XlsxConditionalFormatRule,
  XlsxDataValidation,
  XlsxFreezePanes,
  XlsxImage,
  XlsxImageRect,
  XlsxImageResizeHandlePosition,
  XlsxResolvedCellStyle,
  XlsxShape,
  XlsxSheetData,
  XlsxSparkline,
  XlsxThemePalette,
  XlsxTable,
  XlsxTableStyleDefinition,
  XlsxTableSortDirection,
  XlsxTableSortState,
  XlsxViewerController,
  XlsxWorkbookTab
} from "./types";

const FORMULA_COUNT_THRESHOLD = 1000;
const DEFAULT_ROW_HEIGHT = 24;
const DEFAULT_COL_WIDTH = 80;
const XLSX_MIME_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
const CSV_MIME_TYPE = "text/csv;charset=utf-8";
const MIN_COL_WIDTH_PX = 30;
const MIN_ROW_HEIGHT_PX = 16;
const GRID_HEADER_HEIGHT = 24;
const GRID_ROW_HEADER_WIDTH = 40;
const HISTORY_LIMIT = 100;
const INTERNAL_CLIPBOARD_MIME = "application/x-react-xlsx-range+json";
const DEFAULT_DEFER_LOADING_ABOVE_BYTES = 0;
const DEFAULT_MAX_FILE_SIZE_BYTES = 25 * 1024 * 1024;
const MAX_INTERACTIVE_WORKSHEET_XML_BYTES = 200 * 1024 * 1024;
const MAX_INTERACTIVE_SHARED_STRINGS_BYTES = 50 * 1024 * 1024;
const MAX_INTERACTIVE_TOTAL_XML_BYTES = 256 * 1024 * 1024;
const EMU_PER_PIXEL = 9525;
const IMAGE_BATCH_ROW_COUNT = 256;
const DEFAULT_ZOOM_SCALE = 100;
const MIN_ZOOM_SCALE = 10;
const MAX_ZOOM_SCALE = 400;
const ZOOM_STEP = 10;
const DEFAULT_ZOOM_TAB_KEY = "__default__";

type IdleRequestHandle = number;

type IdleDeadline = {
  didTimeout: boolean;
  timeRemaining: () => number;
};

type IdleWindow = Window & {
  cancelIdleCallback?: (handle: IdleRequestHandle) => void;
  requestIdleCallback?: (
    callback: (deadline: IdleDeadline) => void,
    options?: {
      timeout: number;
    }
  ) => IdleRequestHandle;
};

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

function clampZoomScale(zoomScale: number) {
  if (!Number.isFinite(zoomScale)) {
    return DEFAULT_ZOOM_SCALE;
  }

  return Math.min(MAX_ZOOM_SCALE, Math.max(MIN_ZOOM_SCALE, Math.round(zoomScale)));
}

function resolveDefaultZoomScale(activeTab: XlsxWorkbookTab | null, activeSheet: XlsxSheetData | null) {
  if (activeTab?.kind !== "sheet") {
    return DEFAULT_ZOOM_SCALE;
  }

  return clampZoomScale(activeSheet?.zoomScale ?? DEFAULT_ZOOM_SCALE);
}

function resolveWorksheetZoomScale(
  worksheet: ReturnType<Workbook["getSheet"]>,
  sheetState?: Record<string, unknown> | null
) {
  const candidates = [
    typeof sheetState?.zoomScale === "number" ? sheetState.zoomScale : undefined,
    typeof worksheet.zoomScale === "number" ? worksheet.zoomScale : undefined
  ];
  const value = candidates.find((entry): entry is number => typeof entry === "number" && Number.isFinite(entry) && entry > 0);
  return clampZoomScale(value ?? DEFAULT_ZOOM_SCALE);
}

function resolveNextZoomScale(currentZoomScale: number, direction: 1 | -1) {
  if (direction > 0) {
    return Math.min(
      MAX_ZOOM_SCALE,
      currentZoomScale % ZOOM_STEP === 0
        ? currentZoomScale + ZOOM_STEP
        : Math.ceil(currentZoomScale / ZOOM_STEP) * ZOOM_STEP
    );
  }

  return Math.max(
    MIN_ZOOM_SCALE,
    currentZoomScale % ZOOM_STEP === 0
      ? currentZoomScale - ZOOM_STEP
      : Math.floor(currentZoomScale / ZOOM_STEP) * ZOOM_STEP
  );
}

type WorksheetApiImageInfo = {
  altText?: unknown;
  height?: unknown;
  source?: unknown;
  width?: unknown;
};

type WorksheetDirectImageAnchorInfo = {
  fromCol?: unknown;
  fromColOffset?: unknown;
  fromRow?: unknown;
  fromRowOffset?: unknown;
  toCol?: unknown;
  toColOffset?: unknown;
  toRow?: unknown;
  toRowOffset?: unknown;
};

type WorksheetDirectImageInfo = {
  anchor?: unknown;
  data?: unknown;
  format?: unknown;
  id?: unknown;
  mediaPath?: unknown;
  name?: unknown;
  widthEmu?: unknown;
  heightEmu?: unknown;
};

type WorksheetDirectShapeParagraphRunInfo = {
  bold?: unknown;
  color?: unknown;
  fontFamily?: unknown;
  fontSizePt?: unknown;
  italic?: unknown;
  text?: unknown;
  underline?: unknown;
};

type WorksheetDirectShapeParagraphInfo = {
  align?: unknown;
  runs?: unknown;
};

type WorksheetDirectShapeTextBoxInfo = {
  horizontalAlign?: unknown;
  insetPx?: {
    bottom?: unknown;
    left?: unknown;
    right?: unknown;
    top?: unknown;
  } | null;
  verticalAlign?: unknown;
};

type WorksheetDirectShapeInfo = {
  anchor?: unknown;
  description?: unknown;
  fill?: {
    color?: unknown;
    none?: unknown;
    opacity?: unknown;
  } | null;
  flipH?: unknown;
  flipV?: unknown;
  geometry?: unknown;
  geometryAdjustments?: unknown;
  hyperlink?: unknown;
  id?: unknown;
  name?: unknown;
  paragraphs?: unknown;
  rotationDeg?: unknown;
  scaleX?: unknown;
  scaleY?: unknown;
  stroke?: {
    color?: unknown;
    dash?: unknown;
    headEndType?: unknown;
    none?: unknown;
    opacity?: unknown;
    tailEndType?: unknown;
    widthPx?: unknown;
  } | null;
  svgPath?: unknown;
  svgViewBox?: {
    height?: unknown;
    width?: unknown;
  } | null;
  text?: unknown;
  textBox?: WorksheetDirectShapeTextBoxInfo | null;
};

type WorksheetApiRowCell = {
  col?: unknown;
  image?: WorksheetApiImageInfo | null;
};

type WorksheetApiRow = {
  cells?: unknown;
  index?: unknown;
};

type WorksheetWithRowsBatch = ReturnType<Workbook["getSheet"]> & {
  getRowsBatch?: (startRow: number, maxRows: number, options?: unknown) => unknown;
};

type ZipEntryMetadata = {
  compressedSize: number;
  name: string;
  uncompressedSize: number;
};

type WorkbookPreflightResult = {
  largestWorksheetXmlBytes: number;
  sharedStringsBytes: number;
  totalWorksheetXmlBytes: number;
  tooLarge: boolean;
};

function formatBinaryBytes(bytes: number) {
  if (!Number.isFinite(bytes) || bytes <= 0) {
    return "0 B";
  }

  const units = ["B", "KB", "MB", "GB"];
  let value = bytes;
  let unitIndex = 0;
  while (value >= 1024 && unitIndex < units.length - 1) {
    value /= 1024;
    unitIndex += 1;
  }

  return `${value >= 10 || unitIndex === 0 ? value.toFixed(0) : value.toFixed(1)} ${units[unitIndex]}`;
}

function findZipEndOfCentralDirectoryOffset(bytes: Uint8Array) {
  const minLength = 22;
  if (bytes.byteLength < minLength) {
    return -1;
  }

  const searchStart = Math.max(0, bytes.byteLength - (0xffff + minLength));
  for (let offset = bytes.byteLength - minLength; offset >= searchStart; offset -= 1) {
    if (
      bytes[offset] === 0x50 &&
      bytes[offset + 1] === 0x4b &&
      bytes[offset + 2] === 0x05 &&
      bytes[offset + 3] === 0x06
    ) {
      return offset;
    }
  }

  return -1;
}

function readZipCentralDirectoryEntries(buffer: ArrayBuffer): ZipEntryMetadata[] | null {
  const bytes = new Uint8Array(buffer);
  const eocdOffset = findZipEndOfCentralDirectoryOffset(bytes);
  if (eocdOffset < 0) {
    return null;
  }

  const view = new DataView(buffer, bytes.byteOffset, bytes.byteLength);
  const centralDirectorySize = view.getUint32(eocdOffset + 12, true);
  const centralDirectoryOffset = view.getUint32(eocdOffset + 16, true);
  const decoder = new TextDecoder();
  const entries: ZipEntryMetadata[] = [];

  let offset = centralDirectoryOffset;
  const endOffset = centralDirectoryOffset + centralDirectorySize;
  while (offset + 46 <= endOffset && offset + 46 <= bytes.byteLength) {
    if (view.getUint32(offset, true) !== 0x02014b50) {
      return null;
    }

    const compressedSize = view.getUint32(offset + 20, true);
    const uncompressedSize = view.getUint32(offset + 24, true);
    const fileNameLength = view.getUint16(offset + 28, true);
    const extraLength = view.getUint16(offset + 30, true);
    const commentLength = view.getUint16(offset + 32, true);
    const fileNameStart = offset + 46;
    const fileNameEnd = fileNameStart + fileNameLength;
    if (fileNameEnd > bytes.byteLength) {
      return null;
    }

    entries.push({
      compressedSize,
      name: decoder.decode(bytes.subarray(fileNameStart, fileNameEnd)),
      uncompressedSize
    });

    offset = fileNameEnd + extraLength + commentLength;
  }

  return entries;
}

function preflightWorkbookBuffer(buffer: ArrayBuffer): WorkbookPreflightResult | null {
  const entries = readZipCentralDirectoryEntries(buffer);
  if (!entries) {
    return null;
  }

  let largestWorksheetXmlBytes = 0;
  let totalWorksheetXmlBytes = 0;
  let sharedStringsBytes = 0;

  for (const entry of entries) {
    if (/^xl\/worksheets\/[^/]+\.xml$/i.test(entry.name)) {
      largestWorksheetXmlBytes = Math.max(largestWorksheetXmlBytes, entry.uncompressedSize);
      totalWorksheetXmlBytes += entry.uncompressedSize;
      continue;
    }

    if (entry.name === "xl/sharedStrings.xml") {
      sharedStringsBytes = entry.uncompressedSize;
    }
  }

  const tooLarge =
    largestWorksheetXmlBytes > MAX_INTERACTIVE_WORKSHEET_XML_BYTES ||
    sharedStringsBytes > MAX_INTERACTIVE_SHARED_STRINGS_BYTES ||
    totalWorksheetXmlBytes + sharedStringsBytes > MAX_INTERACTIVE_TOTAL_XML_BYTES;

  return {
    largestWorksheetXmlBytes,
    sharedStringsBytes,
    tooLarge,
    totalWorksheetXmlBytes
  };
}

function createWorkbookTooLargeError(preflight: WorkbookPreflightResult) {
  return new Error(
    `XLSX is too large to preview interactively. `
    + `Largest worksheet XML: ${formatBinaryBytes(preflight.largestWorksheetXmlBytes)}; `
    + `shared strings: ${formatBinaryBytes(preflight.sharedStringsBytes)}.`
  );
}

export class XlsxFileSizeLimitExceededError extends Error {
  fileSizeBytes: number;
  maxFileSizeBytes: number;

  constructor(fileSizeBytes: number, maxFileSizeBytes: number) {
    super(
      `XLSX file size ${formatBinaryBytes(fileSizeBytes)} exceeds the configured limit of ${formatBinaryBytes(maxFileSizeBytes)}.`
    );
    this.name = "XlsxFileSizeLimitExceededError";
    this.fileSizeBytes = fileSizeBytes;
    this.maxFileSizeBytes = maxFileSizeBytes;
  }
}

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

function buildSheetList(
  workbook: Workbook,
  sheetStatesByWorkbookSheetIndex?: Array<{
    cachedFormulaValues?: Record<string, string>;
    columnWidthCharacterWidthPx?: number;
    colWidthOverridesPx?: Record<number, number>;
    colStyleIds?: Record<number, number>;
    conditionalFormatRules?: XlsxConditionalFormatRule[];
    defaultColWidthPx?: number;
    defaultRowHeightPx?: number;
    hasHorizontalMerges?: boolean;
    hasVerticalMerges?: boolean;
    maxHorizontalMergeEndCol?: number;
    maxVerticalMergeEndRow?: number;
    hiddenCols?: number[];
    hiddenRows?: number[];
    rowHeightOverridesPx?: Record<number, number>;
    rowStyleIds?: Record<number, number>;
    showGridLines: boolean;
    sparklines?: XlsxSparkline[];
  } | null>,
  themePalette?: XlsxThemePalette | null,
  styleById?: Record<number, XlsxResolvedCellStyle> | null,
  namedCellStyleByName?: Record<string, XlsxResolvedCellStyle> | null,
  tableStyleByName?: Record<string, XlsxTableStyleDefinition> | null
): XlsxSheetData[] {
  const sheets: XlsxSheetData[] = [];

  for (let index = 0; index < workbook.sheetCount; index += 1) {
    const worksheet = workbook.getSheet(index);
    const sheetState = sheetStatesByWorkbookSheetIndex?.[index] ?? null;
    if (worksheet.visibility !== "visible") {
      continue;
    }

    const resolveColumnWidthPx = (col: number) => {
      const width = worksheet.getColumnWidth(col);
      if (width !== undefined && width !== null) {
        return resolveSheetColumnWidthPixels(width, sheetState?.columnWidthCharacterWidthPx);
      }

      return sheetState?.colWidthOverridesPx?.[col] ?? sheetState?.defaultColWidthPx ?? DEFAULT_COL_WIDTH;
    };

    const resolveRowHeightPx = (row: number) => {
      const height = worksheet.getRowHeight(row);
      if (height !== undefined && height !== null) {
        return Math.max(Math.round(height * 1.33), MIN_ROW_HEIGHT_PX);
      }

      return sheetState?.rowHeightOverridesPx?.[row] ?? sheetState?.defaultRowHeightPx ?? DEFAULT_ROW_HEIGHT;
    };

    const usedRange = worksheet.usedRange() as [number, number, number, number] | null;
    if (!usedRange) {
      sheets.push({
        cachedFormulaValues: sheetState?.cachedFormulaValues ?? {},
        columnWidthCharacterWidthPx: sheetState?.columnWidthCharacterWidthPx,
        colWidthOverridesPx: sheetState?.colWidthOverridesPx ?? {},
        colStyleIds: sheetState?.colStyleIds ?? {},
        conditionalFormatRules: sheetState?.conditionalFormatRules ?? [],
        dataValidations: parseWorksheetDataValidations(worksheet),
        defaultColWidthPx: sheetState?.defaultColWidthPx ?? DEFAULT_COL_WIDTH,
        defaultRowHeightPx: sheetState?.defaultRowHeightPx ?? DEFAULT_ROW_HEIGHT,
        freezePanes: parseWorksheetFreezePanes(worksheet),
        hasHorizontalMerges: sheetState?.hasHorizontalMerges ?? false,
        hasVerticalMerges: sheetState?.hasVerticalMerges ?? false,
        maxHorizontalMergeEndCol: sheetState?.maxHorizontalMergeEndCol ?? -1,
        maxVerticalMergeEndRow: sheetState?.maxVerticalMergeEndRow ?? -1,
        hiddenCols: sheetState?.hiddenCols ?? [],
        hiddenRows: sheetState?.hiddenRows ?? [],
        maxUsedCol: -1,
        maxUsedRow: -1,
        name: worksheet.name,
        namedCellStyleByName: namedCellStyleByName ?? {},
        rowCount: 0,
        colCount: 0,
        rowHeightOverridesPx: sheetState?.rowHeightOverridesPx ?? {},
        rowStyleIds: sheetState?.rowStyleIds ?? {},
        styleById: styleById ?? {},
        sparklines: sheetState?.sparklines ?? [],
        tableStyleByName: tableStyleByName ?? {},
        visibleRows: [],
        visibleCols: [],
        colWidths: [],
        rowHeights: [],
        showGridLines: sheetState?.showGridLines ?? true,
        themePalette: themePalette ?? { colorsByIndex: {} },
        workbookSheetIndex: index,
        zoomScale: resolveWorksheetZoomScale(worksheet, sheetState)
      });
      continue;
    }

    const [, , maxRow, maxCol] = usedRange;
    let visibleRowsCache: number[] | null = null;
    let visibleColsCache: number[] | null = null;
    let rowHeightsCache: number[] | null = null;
    let colWidthsCache: number[] | null = null;

    const getVisibleRows = () => {
      if (visibleRowsCache) {
        return visibleRowsCache;
      }

      const nextVisibleRows: number[] = [];
      for (let row = 0; row <= maxRow; row += 1) {
        if (!worksheet.isRowHidden(row)) {
          nextVisibleRows.push(row);
        }
      }

      visibleRowsCache = nextVisibleRows;
      return nextVisibleRows;
    };

    const getVisibleCols = () => {
      if (visibleColsCache) {
        return visibleColsCache;
      }

      const nextVisibleCols: number[] = [];
      for (let col = 0; col <= maxCol; col += 1) {
        if (!worksheet.isColumnHidden(col)) {
          nextVisibleCols.push(col);
        }
      }

      visibleColsCache = nextVisibleCols;
      return nextVisibleCols;
    };

    const getRowHeights = () => {
      if (rowHeightsCache) {
        return rowHeightsCache;
      }

      rowHeightsCache = getVisibleRows().map(resolveRowHeightPx);
      return rowHeightsCache;
    };

    const getColWidths = () => {
      if (colWidthsCache) {
        return colWidthsCache;
      }

      colWidthsCache = getVisibleCols().map(resolveColumnWidthPx);
      return colWidthsCache;
    };

    const sheet: XlsxSheetData = {
      cachedFormulaValues: sheetState?.cachedFormulaValues ?? {},
      columnWidthCharacterWidthPx: sheetState?.columnWidthCharacterWidthPx,
      colWidthOverridesPx: sheetState?.colWidthOverridesPx ?? {},
      colStyleIds: sheetState?.colStyleIds ?? {},
      conditionalFormatRules: sheetState?.conditionalFormatRules ?? [],
      dataValidations: parseWorksheetDataValidations(worksheet),
      defaultColWidthPx: sheetState?.defaultColWidthPx ?? DEFAULT_COL_WIDTH,
      defaultRowHeightPx: sheetState?.defaultRowHeightPx ?? DEFAULT_ROW_HEIGHT,
      freezePanes: parseWorksheetFreezePanes(worksheet),
      hasHorizontalMerges: sheetState?.hasHorizontalMerges ?? false,
      hasVerticalMerges: sheetState?.hasVerticalMerges ?? false,
      maxHorizontalMergeEndCol: sheetState?.maxHorizontalMergeEndCol ?? -1,
      maxVerticalMergeEndRow: sheetState?.maxVerticalMergeEndRow ?? -1,
      hiddenCols: sheetState?.hiddenCols ?? [],
      hiddenRows: sheetState?.hiddenRows ?? [],
      maxUsedCol: maxCol,
      maxUsedRow: maxRow,
      name: worksheet.name,
      namedCellStyleByName: namedCellStyleByName ?? {},
      rowHeightOverridesPx: sheetState?.rowHeightOverridesPx ?? {},
      rowStyleIds: sheetState?.rowStyleIds ?? {},
      showGridLines: sheetState?.showGridLines ?? true,
      styleById: styleById ?? {},
      sparklines: sheetState?.sparklines ?? [],
      tableStyleByName: tableStyleByName ?? {},
      themePalette: themePalette ?? { colorsByIndex: {} },
      workbookSheetIndex: index,
      zoomScale: resolveWorksheetZoomScale(worksheet, sheetState),
      get rowCount() {
        return getVisibleRows().length;
      },
      get colCount() {
        return getVisibleCols().length;
      },
      get visibleRows() {
        return getVisibleRows();
      },
      get visibleCols() {
        return getVisibleCols();
      },
      get colWidths() {
        return getColWidths();
      },
      get rowHeights() {
        return getRowHeights();
      }
    };

    sheets.push(sheet);
  }

  return sheets;
}

function buildVisibleSheetIndexMap(sheets: XlsxSheetData[]) {
  return new Map(sheets.map((sheet, index) => [sheet.workbookSheetIndex, index]));
}

function resolveInheritedCellStyle(sheet: XlsxSheetData | null | undefined, row: number, col: number) {
  if (!sheet) {
    return null;
  }

  const rowStyleId = sheet.rowStyleIds[row];
  if (rowStyleId !== undefined) {
    return sheet.styleById[rowStyleId] ?? null;
  }

  const colStyleId = sheet.colStyleIds[col];
  if (colStyleId !== undefined) {
    return sheet.styleById[colStyleId] ?? null;
  }

  return null;
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

function parseA1CellReference(reference: string): XlsxCellAddress | null {
  const match = /^([A-Z]+)(\d+)$/i.exec(reference.trim());
  if (!match) {
    return null;
  }

  const [, columnPart, rowPart] = match;
  let col = 0;
  for (const char of columnPart.toUpperCase()) {
    col = col * 26 + (char.charCodeAt(0) - 64);
  }

  return {
    col: col - 1,
    row: Number(rowPart) - 1
  };
}

function parseA1RangeReference(reference: string): XlsxCellRange | null {
  const [startRef, endRef = startRef] = reference.split(":");
  const start = parseA1CellReference(startRef ?? "");
  const end = parseA1CellReference(endRef ?? "");
  if (!start || !end) {
    return null;
  }

  return normalizeRange({ start, end });
}

function parseWorksheetFreezePanes(worksheet: ReturnType<Workbook["getSheet"]>): XlsxFreezePanes | null {
  const rawFreezePanes = worksheet.freezePanes as Record<string, unknown> | null | undefined;
  const row = typeof rawFreezePanes?.row === "number" && rawFreezePanes.row >= 0 ? rawFreezePanes.row : null;
  const col = typeof rawFreezePanes?.col === "number" && rawFreezePanes.col >= 0 ? rawFreezePanes.col : null;
  if (row === null && col === null) {
    return null;
  }

  return {
    col: col ?? 0,
    row: row ?? 0
  };
}

function parseWorksheetDataValidations(worksheet: ReturnType<Workbook["getSheet"]>): XlsxDataValidation[] {
  const rawDataValidations = Array.isArray(worksheet.dataValidations) ? worksheet.dataValidations : [];

  return rawDataValidations.flatMap((entry) => {
    if (!entry || typeof entry !== "object") {
      return [];
    }

    const validation = entry as Record<string, unknown>;
    const ranges = Array.isArray(validation.ranges)
      ? validation.ranges.flatMap((range) => {
          if (typeof range !== "string") {
            return [];
          }

          const parsedRange = parseA1RangeReference(range);
          return parsedRange ? [parsedRange] : [];
        })
      : [];
    const validationType = typeof validation.validationType === "string" ? validation.validationType : null;
    if (!validationType || ranges.length === 0) {
      return [];
    }

    return [{
      allowBlank: typeof validation.allowBlank === "boolean" ? validation.allowBlank : undefined,
      errorMessage: typeof validation.errorMessage === "string" ? validation.errorMessage : undefined,
      errorStyle: typeof validation.errorStyle === "string" ? validation.errorStyle : undefined,
      inputMessage: typeof validation.inputMessage === "string" ? validation.inputMessage : undefined,
      listSource: typeof validation.listSource === "string" ? validation.listSource : undefined,
      ranges,
      showDropdown: typeof validation.showDropdown === "boolean" ? validation.showDropdown : undefined,
      showErrorAlert: typeof validation.showErrorAlert === "boolean" ? validation.showErrorAlert : undefined,
      showInputMessage: typeof validation.showInputMessage === "boolean" ? validation.showInputMessage : undefined,
      validationType
    } satisfies XlsxDataValidation];
  });
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

function mapWorksheetTables(
  worksheet: ReturnType<Workbook["getSheet"]> | null,
  metadataForSheet?: WorkbookTableMetadata[] | null
): XlsxTable[] {
  const rawTables = (worksheet?.tables ?? []) as Array<Record<string, unknown>>;
  return rawTables.flatMap((table, index) => {
    const reference = typeof table.reference === "string" ? table.reference : "";
    const parsedRange = parseA1RangeReference(reference);
    if (!parsedRange) {
      return [];
    }

    const rawColumns = Array.isArray(table.columns) ? table.columns : [];
    const rawName = typeof table.name === "string" ? table.name : `Table${index + 1}`;
    const rawDisplayName =
      typeof table.displayName === "string"
        ? table.displayName
        : typeof table.name === "string"
          ? table.name
          : `Table ${index + 1}`;
    const metadata = metadataForSheet?.find((entry) =>
      (entry.name && entry.name === rawName)
      || (entry.displayName && entry.displayName === rawDisplayName)
      || (entry.reference && entry.reference === reference)
    );

    return [{
      columns: rawColumns.map((column, columnIndex) => ({
        id: typeof (column as { id?: unknown }).id === "number" ? ((column as { id?: number }).id ?? columnIndex + 1) : columnIndex + 1,
        index: columnIndex,
        name: typeof (column as { name?: unknown }).name === "string" ? ((column as { name?: string }).name ?? `Column ${columnIndex + 1}`) : `Column ${columnIndex + 1}`
      })),
      displayName: rawDisplayName,
      end: parsedRange.end,
      headerRowCount: typeof table.headerRowCount === "number" ? table.headerRowCount : 1,
      headerRowCellStyle: metadata?.headerRowCellStyle,
      name: rawName,
      reference,
      start: parsedRange.start,
      styleInfo: table.styleInfo as XlsxTable["styleInfo"] | undefined,
      totalsRowCount: typeof table.totalsRowCount === "number" ? table.totalsRowCount : 0,
      totalsRowShown: Boolean(table.totalsRowShown)
    }];
  });
}

function fileStem(fileName: string): string {
  const normalized = fileName.trim();
  const lastDot = normalized.lastIndexOf(".");
  return lastDot > 0 ? normalized.slice(0, lastDot) : normalized;
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

function sanitizeSavedWorkbookBytes(bytes: Uint8Array): Uint8Array {
  try {
    const archive = unzipSync(bytes);
    const stylesEntry = archive["xl/styles.xml"];
    if (stylesEntry) {
      const stylesXml = strFromU8(stylesEntry)
        .replace(/&amp;quot;/g, "&quot;")
        .replace(/&amp;apos;/g, "&apos;");
      archive["xl/styles.xml"] = strToU8(stylesXml);
    }

    return zipSync(archive, { level: 6 });
  } catch {
    return cloneBytes(bytes);
  }
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

function coerceUserEnteredValue(value: string): unknown {
  const trimmed = value.trim();
  if (!trimmed) {
    return "";
  }

  if (trimmed.startsWith("'")) {
    return trimmed.slice(1);
  }

  if (/^(true|false)$/i.test(trimmed)) {
    return trimmed.toLowerCase() === "true";
  }

  if (/^[+-]?(?:\d+\.?\d*|\.\d+)$/.test(trimmed)) {
    const parsed = Number(trimmed);
    if (Number.isFinite(parsed)) {
      return parsed;
    }
  }

  return value;
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

async function resolveWorkbookBuffer({ file, src }: UseXlsxViewerControllerOptions): Promise<ArrayBuffer> {
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

  return buffer;
}

async function parseWorkbookBuffer(buffer: ArrayBuffer): Promise<{
  shouldAutoCalculate: boolean;
  workbook: Workbook;
}> {
  const wasmModule = await getSheetsWasmModule();
  const workbook = wasmModule.Workbook.fromBytes(new Uint8Array(buffer));
  let totalFormulas = 0;

  for (let index = 0; index < workbook.sheetCount; index += 1) {
    totalFormulas += workbook.getSheet(index).formulaCount;
  }

  const shouldAutoCalculate = totalFormulas <= FORMULA_COUNT_THRESHOLD;
  if (shouldAutoCalculate) {
    workbook.calculate();
  }

  return {
    shouldAutoCalculate,
    workbook
  };
}

function scheduleLowPriorityTask(task: () => void) {
  if (typeof window === "undefined") {
    const timeoutHandle = setTimeout(task, 0);
    return () => clearTimeout(timeoutHandle);
  }

  const idleWindow = window as IdleWindow;
  if (typeof idleWindow.requestIdleCallback === "function") {
    const idleHandle = idleWindow.requestIdleCallback(() => {
      task();
    }, { timeout: 120 });
    return () => {
      idleWindow.cancelIdleCallback?.(idleHandle);
    };
  }

  const timeoutHandle = window.setTimeout(task, 0);
  return () => window.clearTimeout(timeoutHandle);
}

function asFiniteNumber(value: unknown) {
  return typeof value === "number" && Number.isFinite(value) ? value : null;
}

function inferImageMimeType(source: string) {
  if (source.startsWith("data:")) {
    const separatorIndex = source.indexOf(";");
    if (separatorIndex > 5) {
      return source.slice(5, separatorIndex);
    }
  }

  const normalized = source.split("?")[0]?.toLowerCase() ?? "";
  if (normalized.endsWith(".gif")) {
    return "image/gif";
  }
  if (normalized.endsWith(".jpg") || normalized.endsWith(".jpeg")) {
    return "image/jpeg";
  }
  if (normalized.endsWith(".svg")) {
    return "image/svg+xml";
  }
  if (normalized.endsWith(".webp")) {
    return "image/webp";
  }

  return "image/png";
}

function inferWorksheetDirectImageMimeType(info: WorksheetDirectImageInfo) {
  const format = typeof info.format === "string" ? info.format.trim().toLowerCase() : "";
  if (format === "gif") {
    return "image/gif";
  }
  if (format === "jpg" || format === "jpeg") {
    return "image/jpeg";
  }
  if (format === "svg") {
    return "image/svg+xml";
  }
  if (format === "webp") {
    return "image/webp";
  }
  if (format === "png") {
    return "image/png";
  }

  const mediaPath = typeof info.mediaPath === "string" ? info.mediaPath : "";
  if (mediaPath) {
    return inferImageMimeType(mediaPath);
  }

  return "image/png";
}

function createWorksheetDirectImageSource(
  data: unknown,
  mimeType: string,
  objectUrls: string[]
) {
  const bytes = data instanceof Uint8Array
    ? data
    : Array.isArray(data)
      ? Uint8Array.from(data.filter((value): value is number => typeof value === "number"))
      : null;
  if (!bytes || bytes.byteLength === 0) {
    return null;
  }

  const blobBuffer = new Uint8Array(bytes.byteLength);
  blobBuffer.set(bytes);
  const objectUrl = URL.createObjectURL(new Blob([blobBuffer.buffer], { type: mimeType }));
  objectUrls.push(objectUrl);
  return objectUrl;
}

function buildWorksheetDirectImageAnchor(
  rawAnchor: unknown,
  widthEmu: number,
  heightEmu: number
): XlsxImage["anchor"] {
  const anchor = rawAnchor && typeof rawAnchor === "object" ? rawAnchor as WorksheetDirectImageAnchorInfo : {};
  const fromCol = asFiniteNumber(anchor.fromCol) ?? 0;
  const fromRow = asFiniteNumber(anchor.fromRow) ?? 0;
  const fromColOffset = asFiniteNumber(anchor.fromColOffset) ?? 0;
  const fromRowOffset = asFiniteNumber(anchor.fromRowOffset) ?? 0;
  const toCol = asFiniteNumber(anchor.toCol);
  const toRow = asFiniteNumber(anchor.toRow);
  const toColOffset = asFiniteNumber(anchor.toColOffset) ?? 0;
  const toRowOffset = asFiniteNumber(anchor.toRowOffset) ?? 0;

  if (toCol !== null && toRow !== null) {
    return {
      from: {
        col: Math.max(0, Math.round(fromCol)),
        colOffsetEmu: Math.max(0, Math.round(fromColOffset)),
        row: Math.max(0, Math.round(fromRow)),
        rowOffsetEmu: Math.max(0, Math.round(fromRowOffset))
      },
      kind: "two-cell",
      to: {
        col: Math.max(0, Math.round(toCol)),
        colOffsetEmu: Math.max(0, Math.round(toColOffset)),
        row: Math.max(0, Math.round(toRow)),
        rowOffsetEmu: Math.max(0, Math.round(toRowOffset))
      }
    };
  }

  return {
    from: {
      col: Math.max(0, Math.round(fromCol)),
      colOffsetEmu: Math.max(0, Math.round(fromColOffset)),
      row: Math.max(0, Math.round(fromRow)),
      rowOffsetEmu: Math.max(0, Math.round(fromRowOffset))
    },
    kind: "one-cell",
    sizeEmu: {
      cx: Math.max(EMU_PER_PIXEL, Math.round(widthEmu)),
      cy: Math.max(EMU_PER_PIXEL, Math.round(heightEmu))
    }
  };
}

function normalizeWorksheetDirectShapeParagraphs(rawParagraphs: unknown, fallbackText: unknown): XlsxShape["paragraphs"] {
  const normalizedParagraphs: XlsxShape["paragraphs"] = [];

  if (Array.isArray(rawParagraphs)) {
    for (const entry of rawParagraphs) {
        const paragraph = entry && typeof entry === "object" ? entry as WorksheetDirectShapeParagraphInfo : {};
        const runs: XlsxShape["paragraphs"][number]["runs"] = [];
        if (Array.isArray(paragraph.runs)) {
          for (const runEntry of paragraph.runs) {
            const run = runEntry && typeof runEntry === "object" ? runEntry as WorksheetDirectShapeParagraphRunInfo : {};
            const text = typeof run.text === "string" ? run.text : "";
            if (!text) {
              continue;
            }
            runs.push({
              bold: typeof run.bold === "boolean" ? run.bold : undefined,
              color: typeof run.color === "string" && run.color.trim() ? run.color : undefined,
              fontFamily: typeof run.fontFamily === "string" && run.fontFamily.trim() ? run.fontFamily : undefined,
              fontSizePt: asFiniteNumber(run.fontSizePt) ?? undefined,
              italic: typeof run.italic === "boolean" ? run.italic : undefined,
              text,
              underline: typeof run.underline === "boolean" ? run.underline : undefined
            });
          }
        }
        if (runs.length === 0) {
          continue;
        }
        const align = paragraph.align;
        normalizedParagraphs.push({
          align: align === "center" || align === "justify" || align === "left" || align === "right" ? align : undefined,
          runs
        });
    }
  }

  if (normalizedParagraphs.length > 0) {
    return normalizedParagraphs;
  }

  const text = typeof fallbackText === "string" ? fallbackText : "";
  return text
    ? [{ runs: [{ text }] }]
    : [];
}

function buildWorksheetDirectApiShape(
  workbookSheetIndex: number,
  info: WorksheetDirectShapeInfo,
  zIndex: number
): XlsxShape {
  const fill = info.fill && typeof info.fill === "object"
    ? {
        color: typeof info.fill.color === "string" && info.fill.color.trim() ? info.fill.color : undefined,
        none: typeof info.fill.none === "boolean" ? info.fill.none : undefined,
        opacity: asFiniteNumber(info.fill.opacity) ?? undefined
      }
    : undefined;
  const stroke = info.stroke && typeof info.stroke === "object"
    ? {
        color: typeof info.stroke.color === "string" && info.stroke.color.trim() ? info.stroke.color : undefined,
        dash: typeof info.stroke.dash === "string" && info.stroke.dash.trim() ? info.stroke.dash : undefined,
        headEndType: typeof info.stroke.headEndType === "string" && info.stroke.headEndType.trim() ? info.stroke.headEndType : undefined,
        none: typeof info.stroke.none === "boolean" ? info.stroke.none : undefined,
        opacity: asFiniteNumber(info.stroke.opacity) ?? undefined,
        tailEndType: typeof info.stroke.tailEndType === "string" && info.stroke.tailEndType.trim() ? info.stroke.tailEndType : undefined,
        widthPx: asFiniteNumber(info.stroke.widthPx) ?? undefined
      }
    : undefined;
  const rawSvgViewBox = info.svgViewBox && typeof info.svgViewBox === "object" ? info.svgViewBox : null;
  const rawTextBox = info.textBox && typeof info.textBox === "object" ? info.textBox : null;
  const rawInset = rawTextBox?.insetPx && typeof rawTextBox.insetPx === "object" ? rawTextBox.insetPx : null;

  return {
    anchor: buildWorksheetDirectImageAnchor(
      info.anchor,
      DEFAULT_COL_WIDTH * EMU_PER_PIXEL,
      DEFAULT_ROW_HEIGHT * EMU_PER_PIXEL
    ),
    description: typeof info.description === "string" && info.description.trim() ? info.description : undefined,
    fill,
    flipH: typeof info.flipH === "boolean" ? info.flipH : undefined,
    flipV: typeof info.flipV === "boolean" ? info.flipV : undefined,
    geometry: typeof info.geometry === "string" && info.geometry.trim() ? info.geometry : "rect",
    geometryAdjustments: info.geometryAdjustments && typeof info.geometryAdjustments === "object"
      ? Object.fromEntries(
          Object.entries(info.geometryAdjustments as Record<string, unknown>)
            .map(([key, value]) => [key, asFiniteNumber(value)])
            .filter((entry): entry is [string, number] => typeof entry[1] === "number")
        )
      : undefined,
    hyperlink: typeof info.hyperlink === "string" && info.hyperlink.trim() ? info.hyperlink : undefined,
    id: `shape-${workbookSheetIndex}-${String(info.id ?? zIndex)}`,
    name: typeof info.name === "string" && info.name.trim() ? info.name : undefined,
    paragraphs: normalizeWorksheetDirectShapeParagraphs(info.paragraphs, info.text),
    rotationDeg: asFiniteNumber(info.rotationDeg) ?? undefined,
    scaleX: asFiniteNumber(info.scaleX) ?? undefined,
    scaleY: asFiniteNumber(info.scaleY) ?? undefined,
    sheetIndex: workbookSheetIndex,
    svgPath: typeof info.svgPath === "string" && info.svgPath.trim() ? info.svgPath : undefined,
    svgViewBox: rawSvgViewBox
      && asFiniteNumber(rawSvgViewBox.width) !== null
      && asFiniteNumber(rawSvgViewBox.height) !== null
      ? {
          height: asFiniteNumber(rawSvgViewBox.height) ?? 0,
          width: asFiniteNumber(rawSvgViewBox.width) ?? 0
        }
      : undefined,
    stroke,
    textBox: rawTextBox
      ? {
          horizontalAlign: rawTextBox.horizontalAlign === "center" || rawTextBox.horizontalAlign === "left"
            ? rawTextBox.horizontalAlign
            : undefined,
          insetPx: rawInset
            ? {
                bottom: asFiniteNumber(rawInset.bottom) ?? 0,
                left: asFiniteNumber(rawInset.left) ?? 0,
                right: asFiniteNumber(rawInset.right) ?? 0,
                top: asFiniteNumber(rawInset.top) ?? 0
              }
            : undefined,
          verticalAlign: rawTextBox.verticalAlign === "bottom" || rawTextBox.verticalAlign === "middle" || rawTextBox.verticalAlign === "top"
            ? rawTextBox.verticalAlign
            : undefined
        }
      : undefined,
    workbookSheetIndex,
    zIndex
  };
}

function buildWorksheetApiImage(
  workbookSheetIndex: number,
  row: number,
  col: number,
  info: WorksheetApiImageInfo,
  zIndex: number
): XlsxImage | null {
  if (typeof info.source !== "string" || !info.source) {
    return null;
  }

  const width = Math.max(1, Math.round(asFiniteNumber(info.width) ?? DEFAULT_COL_WIDTH));
  const height = Math.max(1, Math.round(asFiniteNumber(info.height) ?? DEFAULT_ROW_HEIGHT));
  const description = typeof info.altText === "string" && info.altText.trim() ? info.altText : undefined;

  return {
    anchor: {
      from: {
        col,
        colOffsetEmu: 0,
        row,
        rowOffsetEmu: 0
      },
      kind: "one-cell",
      sizeEmu: {
        cx: width * EMU_PER_PIXEL,
        cy: height * EMU_PER_PIXEL
      }
    },
    description,
    editable: false,
    id: `worksheet-image-${workbookSheetIndex}-${row}-${col}-${zIndex}`,
    mimeType: inferImageMimeType(info.source),
    sheetIndex: workbookSheetIndex,
    src: info.source,
    workbookSheetIndex,
    zIndex
  };
}

function buildWorksheetDirectApiImage(
  workbookSheetIndex: number,
  info: WorksheetDirectImageInfo,
  zIndex: number,
  objectUrls: string[]
): XlsxImage | null {
  const mimeType = inferWorksheetDirectImageMimeType(info);
  const src = createWorksheetDirectImageSource(info.data, mimeType, objectUrls);
  if (!src) {
    return null;
  }

  const widthEmu = Math.max(EMU_PER_PIXEL, Math.round(asFiniteNumber(info.widthEmu) ?? DEFAULT_COL_WIDTH * EMU_PER_PIXEL));
  const heightEmu = Math.max(EMU_PER_PIXEL, Math.round(asFiniteNumber(info.heightEmu) ?? DEFAULT_ROW_HEIGHT * EMU_PER_PIXEL));
  return {
    anchor: buildWorksheetDirectImageAnchor(info.anchor, widthEmu, heightEmu),
    editable: false,
    id: `worksheet-image-${workbookSheetIndex}-${String(info.id ?? zIndex)}`,
    mediaPath: typeof info.mediaPath === "string" && info.mediaPath.trim() ? info.mediaPath : undefined,
    mimeType,
    name: typeof info.name === "string" && info.name.trim() ? info.name : undefined,
    sheetIndex: workbookSheetIndex,
    src,
    workbookSheetIndex,
    zIndex
  };
}

function collectWorksheetBatchImages(workbook: Workbook) {
  const imagesByWorkbookSheetIndex = Array.from({ length: workbook.sheetCount }, () => [] as XlsxImage[]);

  for (let workbookSheetIndex = 0; workbookSheetIndex < workbook.sheetCount; workbookSheetIndex += 1) {
    const worksheet = workbook.getSheet(workbookSheetIndex) as WorksheetWithRowsBatch;
    if (typeof worksheet.getRowsBatch !== "function") {
      continue;
    }

    const usedRange = worksheet.usedRange() as [number, number, number, number] | null;
    const maxRow = usedRange?.[2] ?? -1;
    if (maxRow < 0) {
      continue;
    }

    let zIndex = 1;
    let sheetFailed = false;
    for (let startRow = 0; startRow <= maxRow; startRow += IMAGE_BATCH_ROW_COUNT) {
      let rows: unknown;
      try {
        rows = worksheet.getRowsBatch(startRow, IMAGE_BATCH_ROW_COUNT, { includeImages: true });
      } catch {
        sheetFailed = true;
        break;
      }

      if (!Array.isArray(rows)) {
        continue;
      }

      for (const rowEntry of rows as WorksheetApiRow[]) {
        const row = typeof rowEntry.index === "number" ? rowEntry.index : null;
        if (row === null || !Array.isArray(rowEntry.cells)) {
          continue;
        }

        for (const cellEntry of rowEntry.cells as WorksheetApiRowCell[]) {
          const col = typeof cellEntry.col === "number" ? cellEntry.col : null;
          if (col === null || !cellEntry.image || typeof cellEntry.image !== "object") {
            continue;
          }

          const image = buildWorksheetApiImage(workbookSheetIndex, row, col, cellEntry.image, zIndex);
          if (!image) {
            continue;
          }

          imagesByWorkbookSheetIndex[workbookSheetIndex].push(image);
          zIndex += 1;
        }
      }
    }

    if (sheetFailed) {
      imagesByWorkbookSheetIndex[workbookSheetIndex] = [];
    }
  }

  return imagesByWorkbookSheetIndex;
}

function collectWorksheetApiImages(workbook: Workbook, objectUrls: string[]) {
  const directImagesByWorkbookSheetIndex = Array.from({ length: workbook.sheetCount }, () => [] as XlsxImage[]);
  let didUseDirectImages = false;

  for (let workbookSheetIndex = 0; workbookSheetIndex < workbook.sheetCount; workbookSheetIndex += 1) {
    const worksheet = workbook.getSheet(workbookSheetIndex) as ReturnType<Workbook["getSheet"]> & {
      images?: unknown;
    };
    const rawImages = Array.isArray(worksheet.images) ? worksheet.images as WorksheetDirectImageInfo[] : [];
    if (rawImages.length === 0) {
      continue;
    }

    const nextImages = rawImages
      .map((info, index) => buildWorksheetDirectApiImage(workbookSheetIndex, info, index + 1, objectUrls))
      .filter((image): image is XlsxImage => Boolean(image));
    if (nextImages.length > 0) {
      directImagesByWorkbookSheetIndex[workbookSheetIndex] = nextImages;
      didUseDirectImages = true;
    }
  }

  if (didUseDirectImages) {
    return directImagesByWorkbookSheetIndex;
  }

  return collectWorksheetBatchImages(workbook);
}

function collectWorksheetApiShapes(workbook: Workbook) {
  return Array.from({ length: workbook.sheetCount }, (_, workbookSheetIndex) => {
    const worksheet = workbook.getSheet(workbookSheetIndex) as ReturnType<Workbook["getSheet"]> & {
      shapes?: unknown;
    };
    const rawShapes = Array.isArray(worksheet.shapes) ? worksheet.shapes as WorksheetDirectShapeInfo[] : [];
    return rawShapes
      .map((shape, index) => buildWorksheetDirectApiShape(workbookSheetIndex, shape, index + 1));
  });
}

function mergeParsedAndApiImages(parsedImages: XlsxImage[], apiImages: XlsxImage[]) {
  if (parsedImages.length === 0) {
    return apiImages;
  }
  if (apiImages.length === 0) {
    return parsedImages;
  }

  const normalizeTextKey = (value: string | undefined) => value?.trim().toLowerCase() ?? "";
  const anchorKey = (anchor: XlsxImage["anchor"]) => {
    if (anchor.kind === "absolute") {
      return [
        "absolute",
        Math.round(anchor.positionEmu.x),
        Math.round(anchor.positionEmu.y),
        Math.round(anchor.sizeEmu.cx),
        Math.round(anchor.sizeEmu.cy)
      ].join(":");
    }
    if (anchor.kind === "one-cell") {
      return [
        "one",
        anchor.from.col,
        anchor.from.row,
        Math.round(anchor.from.colOffsetEmu),
        Math.round(anchor.from.rowOffsetEmu),
        Math.round(anchor.sizeEmu.cx),
        Math.round(anchor.sizeEmu.cy)
      ].join(":");
    }
    return [
      "two",
      anchor.from.col,
      anchor.from.row,
      Math.round(anchor.from.colOffsetEmu),
      Math.round(anchor.from.rowOffsetEmu),
      anchor.to.col,
      anchor.to.row,
      Math.round(anchor.to.colOffsetEmu),
      Math.round(anchor.to.rowOffsetEmu)
    ].join(":");
  };
  const imageKeys = (image: XlsxImage) => {
    const keys = [
      `${normalizeTextKey(image.mediaPath)}|${normalizeTextKey(image.name)}|${anchorKey(image.anchor)}`,
      `${normalizeTextKey(image.mediaPath)}|${anchorKey(image.anchor)}`,
      `${normalizeTextKey(image.name)}|${anchorKey(image.anchor)}`,
      `${anchorKey(image.anchor)}`
    ];
    return keys.filter((key, index) => key && keys.indexOf(key) === index);
  };

  const apiBuckets = new Map<string, XlsxImage[]>();
  for (const apiImage of apiImages) {
    for (const key of imageKeys(apiImage)) {
      const bucket = apiBuckets.get(key);
      if (bucket) {
        bucket.push(apiImage);
      } else {
        apiBuckets.set(key, [apiImage]);
      }
    }
  }

  const usedApiImages = new Set<XlsxImage>();
  const takeApiMatch = (image: XlsxImage) => {
    for (const key of imageKeys(image)) {
      const bucket = apiBuckets.get(key);
      if (!bucket) {
        continue;
      }
      const match = bucket.find((candidate) => !usedApiImages.has(candidate));
      if (match) {
        usedApiImages.add(match);
        return match;
      }
    }
    return null;
  };

  const merged = parsedImages.map((image) => {
    const apiImage = takeApiMatch(image);
    if (!apiImage) {
      return image;
    }

    return {
      ...image,
      anchor: apiImage.anchor,
      mediaPath: apiImage.mediaPath ?? image.mediaPath,
      mimeType: apiImage.mimeType,
      name: apiImage.name ?? image.name,
      src: apiImage.src
    };
  });

  return merged;
}

function isZipWorkbook(bytes: Uint8Array) {
  return bytes.byteLength >= 4 && bytes[0] === 0x50 && bytes[1] === 0x4b;
}

function createBasicWorkbookAssets(workbook: Workbook): WorkbookImageAssets {
  const objectUrls: string[] = [];
  return {
    archive: {},
    imageOriginsById: new Map(),
    imagesByWorkbookSheetIndex: collectWorksheetApiImages(workbook, objectUrls),
    namedCellStyleByName: {},
    objectUrls,
    shapesByWorkbookSheetIndex: collectWorksheetApiShapes(workbook),
    sheetOrigins: Array.from({ length: workbook.sheetCount }, () => null as WorkbookImageSheetOrigin | null),
    sheetStatesByWorkbookSheetIndex: Array.from({ length: workbook.sheetCount }, () => null),
    styleById: {},
    tableMetadataByWorkbookSheetIndex: Array.from({ length: workbook.sheetCount }, () => [] as WorkbookTableMetadata[]),
    tableStyleByName: {},
    themePalette: { colorsByIndex: {} }
  };
}

function loadWorkbookImageAssets(bytes: Uint8Array, workbook: Workbook) {
  if (!isZipWorkbook(bytes)) {
    return createBasicWorkbookAssets(workbook);
  }

  const parsedAssets = parseWorkbookImageAssets(bytes);
  const apiImagesByWorkbookSheetIndex = collectWorksheetApiImages(workbook, parsedAssets.objectUrls);

  const imagesByWorkbookSheetIndex = Array.from(
    { length: Math.max(workbook.sheetCount, parsedAssets.imagesByWorkbookSheetIndex.length, apiImagesByWorkbookSheetIndex.length) },
    (_, index) => {
      const parsedImages = parsedAssets.imagesByWorkbookSheetIndex[index] ?? [];
      const apiImages = apiImagesByWorkbookSheetIndex[index] ?? [];
      return mergeParsedAndApiImages(parsedImages, apiImages);
    }
  );

  return {
    ...parsedAssets,
    imagesByWorkbookSheetIndex
  };
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
  const {
    deferLoadingAboveBytes = DEFAULT_DEFER_LOADING_ABOVE_BYTES,
    file,
    fileName,
    maxFileSizeBytes = DEFAULT_MAX_FILE_SIZE_BYTES,
    readOnly: requestedReadOnly = false,
    readOnlyAboveBytes = 0,
    src,
    useWorker = true
  } = options;
  const [isLoading, setIsLoading] = React.useState(Boolean(file ?? src));
  const [error, setError] = React.useState<Error | null>(null);
  const [workbook, setWorkbook] = React.useState<Workbook | null>(null);
  const [sheets, setSheets] = React.useState<XlsxSheetData[]>([]);
  const [chartsByWorkbookSheetIndex, setChartsByWorkbookSheetIndex] = React.useState<XlsxChart[][]>([]);
  const [chartsheets, setChartsheets] = React.useState<XlsxChartsheet[]>([]);
  const [tabs, setTabs] = React.useState<XlsxWorkbookTab[]>([]);
  const [isChartsLoading, setIsChartsLoading] = React.useState(false);
  const [workerTablesByWorkbookSheetIndex, setWorkerTablesByWorkbookSheetIndex] = React.useState<XlsxTable[][]>([]);
  const [imagesByWorkbookSheetIndex, setImagesByWorkbookSheetIndex] = React.useState<XlsxImage[][]>([]);
  const [shapesByWorkbookSheetIndex, setShapesByWorkbookSheetIndex] = React.useState<XlsxShape[][]>([]);
  const [activeSheetIndex, setActiveSheetIndexState] = React.useState(0);
  const [activeTabIndex, setActiveTabIndexState] = React.useState(0);
  const [zoomScaleOverridesByTabId, setZoomScaleOverridesByTabId] = React.useState<Record<string, number>>({});
  const [activeCell, setActiveCell] = React.useState<XlsxCellAddress | null>(null);
  const [selection, setSelection] = React.useState<XlsxCellRange | null>(null);
  const [selectedChartId, setSelectedChartId] = React.useState<string | null>(null);
  const [selectedImageId, setSelectedImageId] = React.useState<string | null>(null);
  const [revision, setRevision] = React.useState(0);
  const selectionAnchorRef = React.useRef<XlsxCellAddress | null>(null);
  const undoStackRef = React.useRef<HistoryEntry[]>([]);
  const redoStackRef = React.useRef<HistoryEntry[]>([]);
  const isApplyingHistoryRef = React.useRef(false);
  const [historyRevision, setHistoryRevision] = React.useState(0);
  const [shouldAutoCalculate, setShouldAutoCalculate] = React.useState(false);
  const [workerCellSnapshotRevision, setWorkerCellSnapshotRevision] = React.useState(0);
  const [isWorkerBacked, setIsWorkerBacked] = React.useState(false);
  const [sortState, setSortState] = React.useState<XlsxTableSortState | null>(null);
  const [forcedReadOnly, setForcedReadOnly] = React.useState(false);
  const deferredBufferRef = React.useRef<ArrayBuffer | null>(null);
  const [deferredLoadFileSize, setDeferredLoadFileSize] = React.useState<number | null>(null);
  const imageAssetsRef = React.useRef<WorkbookImageAssets | null>(null);
  const chartAssetsRef = React.useRef<WorkbookChartAssets | null>(null);
  const chartLoadRequestTokenRef = React.useRef(0);
  const chartDisplayFallbackCleanupRef = React.useRef<(() => void) | null>(null);
  const sheetOriginsRef = React.useRef<Array<WorkbookImageSheetOrigin | null>>([]);
  const workerClientRef = React.useRef<XlsxWorkerClient | null>(null);
  const workerCellSnapshotCacheRef = React.useRef(new Map<string, { displayValue: string; formula: string }>());
  const displayFileName = React.useMemo(() => resolveDisplayFileName(src, fileName), [fileName, src]);
  const shouldDeferLoading = deferLoadingAboveBytes > 0;
  const readOnly = requestedReadOnly || forcedReadOnly;
  const workerSupported = useWorker && typeof Worker !== "undefined";
  const shouldUseWorker = workerSupported && readOnly;
  const shouldForceReadOnlyForBuffer = React.useCallback((bufferByteLength: number) => (
    !requestedReadOnly && readOnlyAboveBytes > 0 && bufferByteLength > readOnlyAboveBytes
  ), [readOnlyAboveBytes, requestedReadOnly]);

  const disposeWorkerClient = React.useCallback(() => {
    workerClientRef.current?.dispose();
    workerClientRef.current = null;
  }, []);

  const getWorkerClient = React.useCallback(() => {
    if (!workerClientRef.current) {
      workerClientRef.current = new XlsxWorkerClient();
    }

    return workerClientRef.current;
  }, []);

  const clearImageAssets = React.useCallback(() => {
    revokeWorkbookImageAssets(imageAssetsRef.current);
    imageAssetsRef.current = null;
    sheetOriginsRef.current = [];
    setImagesByWorkbookSheetIndex([]);
    setShapesByWorkbookSheetIndex([]);
  }, []);

  const clearChartAssets = React.useCallback(() => {
    chartLoadRequestTokenRef.current += 1;
    chartDisplayFallbackCleanupRef.current?.();
    chartDisplayFallbackCleanupRef.current = null;
    chartAssetsRef.current = null;
    setChartsByWorkbookSheetIndex([]);
    setChartsheets([]);
    setTabs([]);
    setIsChartsLoading(false);
  }, []);

  const setImageAssets = React.useCallback((assets: WorkbookImageAssets | null) => {
    revokeWorkbookImageAssets(imageAssetsRef.current);
    imageAssetsRef.current = assets;
    sheetOriginsRef.current = assets?.sheetOrigins.slice() ?? [];
    setImagesByWorkbookSheetIndex(assets?.imagesByWorkbookSheetIndex ?? []);
    setShapesByWorkbookSheetIndex(assets?.shapesByWorkbookSheetIndex ?? []);
  }, []);

  const setChartAssets = React.useCallback((assets: WorkbookChartAssets | null) => {
    chartAssetsRef.current = assets;
    setChartsByWorkbookSheetIndex(assets?.chartsByWorkbookSheetIndex ?? []);
    setChartsheets(assets?.chartsheets ?? []);
    setTabs(assets?.tabs ?? []);
    setIsChartsLoading(false);
  }, []);

  const shouldFallbackFromWorkerError = React.useCallback((error: unknown) => {
    const message = error instanceof Error ? error.message : String(error ?? "");
    return (
      message.includes("DOMParser is not defined")
      || message.includes("XMLSerializer is not defined")
      || message.includes("Worker chart payload incomplete")
    );
  }, []);

  const hasIncompleteWorkerChartSnapshot = React.useCallback((snapshot: {
    chartsByWorkbookSheetIndex: XlsxChart[][];
  }) => {
    for (const sheetCharts of snapshot.chartsByWorkbookSheetIndex) {
      for (const chart of sheetCharts) {
        if (!chart.chartPath) {
          return true;
        }
        if (chart.chartType !== "Bubble") {
          continue;
        }
        for (const series of chart.series) {
          const pointCount = Math.max(series.values.length, series.categories.length);
          if (pointCount <= 1) {
            continue;
          }
          const numericBubbleSizes = (series.bubbleSizes ?? []).filter(
            (value): value is number => typeof value === "number" && Number.isFinite(value)
          );
          if (numericBubbleSizes.length < pointCount) {
            return true;
          }
        }
      }
    }

    return false;
  }, []);

  const ensureChartAssetsHydrated = React.useCallback((
    targetWorkbook: Workbook | null,
    targetSheets: XlsxSheetData[]
  ) => {
    if (chartAssetsRef.current || !targetWorkbook || !imageAssetsRef.current) {
      return chartAssetsRef.current;
    }

    const assets = loadWorkbookChartAssets(
      targetWorkbook,
      imageAssetsRef.current,
      buildVisibleSheetIndexMap(targetSheets)
    );
    chartAssetsRef.current = assets;
    return assets;
  }, []);

  const startChartDisplayHydration = React.useCallback((
    buffer: ArrayBuffer,
    targetWorkbook: Workbook,
    targetSheets: XlsxSheetData[]
  ) => {
    const visibleSheetIndexByWorkbookSheetIndex = buildVisibleSheetIndexMap(targetSheets);
    const quickAssets = loadWorkbookChartAssets(targetWorkbook, null, visibleSheetIndexByWorkbookSheetIndex);
    setChartAssets(quickAssets);

    const hasCharts = quickAssets.chartsByWorkbookSheetIndex.some((sheetCharts) => sheetCharts.length > 0);
    if (!hasCharts) {
      setIsChartsLoading(false);
      return;
    }

    setIsChartsLoading(true);
    const requestToken = chartLoadRequestTokenRef.current + 1;
    chartLoadRequestTokenRef.current = requestToken;
    chartDisplayFallbackCleanupRef.current?.();
    chartDisplayFallbackCleanupRef.current = null;
    let fallbackTriggered = false;
    const triggerFallback = () => {
      if (fallbackTriggered || requestToken !== chartLoadRequestTokenRef.current) {
        return;
      }
      fallbackTriggered = true;
      runMainThreadFallback();
    };
    const workerTimeoutHandle = typeof window !== "undefined"
      ? window.setTimeout(() => {
          triggerFallback();
        }, 1500)
      : null;

    const applyWorkerResult = (result: {
      chartsByWorkbookSheetIndex: XlsxChart[][];
      chartsheets: XlsxChartsheet[];
      tabs: XlsxWorkbookTab[];
    }) => {
      if (requestToken !== chartLoadRequestTokenRef.current) {
        return;
      }
      setChartsByWorkbookSheetIndex(result.chartsByWorkbookSheetIndex);
      setChartsheets(result.chartsheets);
      setTabs(result.tabs);
      setIsChartsLoading(false);
    };

    const runMainThreadFallback = () => {
      chartDisplayFallbackCleanupRef.current = scheduleLowPriorityTask(() => {
        if (requestToken !== chartLoadRequestTokenRef.current) {
          return;
        }
        try {
          const hydratedAssets = loadWorkbookChartAssets(
            targetWorkbook,
            imageAssetsRef.current,
            visibleSheetIndexByWorkbookSheetIndex
          );
          if (requestToken !== chartLoadRequestTokenRef.current) {
            return;
          }
          setChartAssets(hydratedAssets);
        } catch {
          if (requestToken !== chartLoadRequestTokenRef.current) {
            return;
          }
          setChartAssets(quickAssets);
        } finally {
          if (requestToken === chartLoadRequestTokenRef.current) {
            setIsChartsLoading(false);
          }
        }
      });
    };

    if (!workerSupported) {
      runMainThreadFallback();
      return;
    }

    void getWorkerClient().parseCharts(buffer)
      .then((result) => {
        if (workerTimeoutHandle !== null) {
          window.clearTimeout(workerTimeoutHandle);
        }
        if (fallbackTriggered) {
          return;
        }
        try {
          if (hasIncompleteWorkerChartSnapshot(result)) {
            triggerFallback();
            return;
          }
          applyWorkerResult(result);
        } catch {
          triggerFallback();
        }
      })
      .catch(() => {
        if (workerTimeoutHandle !== null) {
          window.clearTimeout(workerTimeoutHandle);
        }
        triggerFallback();
      });
  }, [getWorkerClient, hasIncompleteWorkerChartSnapshot, setChartAssets, workerSupported]);

  const loadWorkbookOnMainThread = React.useCallback(async (buffer: ArrayBuffer) => {
    const nextParsedWorkbook = await parseWorkbookBuffer(buffer);
    const nextImageAssets = loadWorkbookImageAssets(new Uint8Array(buffer), nextParsedWorkbook.workbook);
    return {
      imageAssets: nextImageAssets,
      parsedWorkbook: nextParsedWorkbook
    };
  }, []);

  const refreshWorkbookState = React.useCallback((targetWorkbook: Workbook) => {
    const nextSheets = buildSheetList(
      targetWorkbook,
      imageAssetsRef.current?.sheetStatesByWorkbookSheetIndex,
      imageAssetsRef.current?.themePalette,
      imageAssetsRef.current?.styleById,
      imageAssetsRef.current?.namedCellStyleByName,
      imageAssetsRef.current?.tableStyleByName
    );
    setSheets(nextSheets);
    setChartAssets(
      loadWorkbookChartAssets(
        targetWorkbook,
        imageAssetsRef.current,
        buildVisibleSheetIndexMap(nextSheets)
      )
    );
    setRevision((current) => current + 1);
  }, [setChartAssets]);

  React.useEffect(() => () => {
    chartDisplayFallbackCleanupRef.current?.();
    chartDisplayFallbackCleanupRef.current = null;
    revokeWorkbookImageAssets(imageAssetsRef.current);
    disposeWorkerClient();
  }, [disposeWorkerClient]);

  React.useEffect(() => {
    if (!file && !src) {
      disposeWorkerClient();
      setForcedReadOnly(false);
      setWorkbook(null);
      setSheets([]);
      clearChartAssets();
      setWorkerTablesByWorkbookSheetIndex([]);
      clearImageAssets();
      setError(null);
      setIsLoading(false);
      setIsWorkerBacked(false);
      deferredBufferRef.current = null;
      setDeferredLoadFileSize(null);
      setActiveSheetIndexState(0);
      setActiveTabIndexState(0);
      setActiveCell(null);
      setSelection(null);
      setSelectedChartId(null);
      setSelectedImageId(null);
      selectionAnchorRef.current = null;
      undoStackRef.current = [];
      redoStackRef.current = [];
      setHistoryRevision(0);
      setShouldAutoCalculate(false);
      workerCellSnapshotCacheRef.current.clear();
      setWorkerCellSnapshotRevision(0);
      setSortState(null);
      setZoomScaleOverridesByTabId({});
      setRevision(0);
      return;
    }

    let isCurrent = true;
    setIsLoading(true);
    setError(null);
    clearImageAssets();
    clearChartAssets();
    setWorkerTablesByWorkbookSheetIndex([]);
    setIsWorkerBacked(false);
    deferredBufferRef.current = null;
    setDeferredLoadFileSize(null);
    setActiveSheetIndexState(0);
    setActiveTabIndexState(0);
    setActiveCell(null);
    setSelection(null);
    setSelectedChartId(null);
    setSelectedImageId(null);
    selectionAnchorRef.current = null;
    undoStackRef.current = [];
    redoStackRef.current = [];
    setHistoryRevision(0);
    setShouldAutoCalculate(false);
    workerCellSnapshotCacheRef.current.clear();
    setWorkerCellSnapshotRevision(0);
    setSortState(null);
    setZoomScaleOverridesByTabId({});
    setRevision(0);
    if (!workerSupported) {
      disposeWorkerClient();
    }

    void resolveWorkbookBuffer({ file, src })
      .then(async (buffer) => {
        if (!isCurrent) {
          return;
        }

        if (maxFileSizeBytes > 0 && buffer.byteLength > maxFileSizeBytes) {
          throw new XlsxFileSizeLimitExceededError(buffer.byteLength, maxFileSizeBytes);
        }

        const preflight = preflightWorkbookBuffer(buffer);
        if (preflight?.tooLarge) {
          throw createWorkbookTooLargeError(preflight);
        }

        const shouldForceReadOnly = shouldForceReadOnlyForBuffer(buffer.byteLength);
        setForcedReadOnly(shouldForceReadOnly);
        const shouldUseWorkerForLoad = workerSupported && (requestedReadOnly || shouldForceReadOnly);

        if (shouldDeferLoading && buffer.byteLength > deferLoadingAboveBytes) {
          deferredBufferRef.current = buffer;
          setDeferredLoadFileSize(buffer.byteLength);
          setWorkbook(null);
          setSheets([]);
          clearChartAssets();
          setWorkerTablesByWorkbookSheetIndex([]);
          setIsLoading(false);
          return;
        }

        if (shouldUseWorkerForLoad) {
          try {
            const snapshot = await getWorkerClient().loadWorkbook(buffer);
            if (!isCurrent) {
              return;
            }
            if (hasIncompleteWorkerChartSnapshot(snapshot)) {
              throw new Error("Worker chart payload incomplete");
            }

            setWorkbook(null);
            setSheets(snapshot.sheets);
            setChartsByWorkbookSheetIndex(snapshot.chartsByWorkbookSheetIndex);
            setChartsheets(snapshot.chartsheets);
            setTabs(snapshot.tabs);
            chartAssetsRef.current = null;
            setWorkerTablesByWorkbookSheetIndex(snapshot.tablesByWorkbookSheetIndex);
            setShouldAutoCalculate(false);
            setIsWorkerBacked(true);
            setSortState(null);
            setIsChartsLoading(false);
            setIsLoading(false);
            return;
          } catch (workerError) {
            if (!shouldFallbackFromWorkerError(workerError)) {
              throw workerError;
            }

            disposeWorkerClient();
          }
        }

        const { imageAssets: nextImageAssets, parsedWorkbook: nextParsedWorkbook } = await loadWorkbookOnMainThread(buffer);
        if (!isCurrent) {
          revokeWorkbookImageAssets(nextImageAssets);
          return;
        }

        setImageAssets(nextImageAssets);
        setWorkbook(nextParsedWorkbook.workbook);
        const nextSheets = buildSheetList(
          nextParsedWorkbook.workbook,
          nextImageAssets.sheetStatesByWorkbookSheetIndex,
          nextImageAssets.themePalette,
          nextImageAssets.styleById,
          nextImageAssets.namedCellStyleByName,
          nextImageAssets.tableStyleByName
        );
        setSheets(nextSheets);
        startChartDisplayHydration(buffer, nextParsedWorkbook.workbook, nextSheets);
        setShouldAutoCalculate(nextParsedWorkbook.shouldAutoCalculate);
        setWorkerTablesByWorkbookSheetIndex([]);
        setIsWorkerBacked(false);
        setSortState(null);
        setIsLoading(false);
      })
      .catch((nextError: unknown) => {
        if (!isCurrent) {
          return;
        }

        setWorkbook(null);
        setSheets([]);
        clearChartAssets();
        setWorkerTablesByWorkbookSheetIndex([]);
        clearImageAssets();
        setShouldAutoCalculate(false);
        setIsWorkerBacked(false);
        setSortState(null);
        setError(nextError instanceof Error ? nextError : new Error("Could not load workbook."));
        setIsLoading(false);
      });

    return () => {
      isCurrent = false;
      if (!workerSupported) {
        disposeWorkerClient();
      }
    };
  }, [
    clearChartAssets,
    clearImageAssets,
    deferLoadingAboveBytes,
    disposeWorkerClient,
    file,
    getWorkerClient,
    hasIncompleteWorkerChartSnapshot,
    loadWorkbookOnMainThread,
    maxFileSizeBytes,
    requestedReadOnly,
    setImageAssets,
    startChartDisplayHydration,
    shouldFallbackFromWorkerError,
    shouldDeferLoading,
    shouldForceReadOnlyForBuffer,
    workerSupported,
    src
  ]);

  const activeTab = tabs[activeTabIndex] ?? null;
  const activeSheet = activeTab?.kind === "sheet"
    ? sheets[activeTab.sheetIndex ?? -1] ?? null
    : null;
  const deferredMetadataCell = React.useDeferredValue(activeCell);
  const deferredMetadataSheet = React.useDeferredValue(activeSheet);
  const activeZoomTabKey = activeTab?.id ?? DEFAULT_ZOOM_TAB_KEY;
  const defaultZoomScale = React.useMemo(
    () => resolveDefaultZoomScale(activeTab, activeSheet),
    [activeSheet, activeTab]
  );
  const zoomScale = React.useMemo(
    () => clampZoomScale(zoomScaleOverridesByTabId[activeZoomTabKey] ?? defaultZoomScale),
    [activeZoomTabKey, defaultZoomScale, zoomScaleOverridesByTabId]
  );
  const canZoomIn = zoomScale < MAX_ZOOM_SCALE;
  const canZoomOut = zoomScale > MIN_ZOOM_SCALE;

  React.useEffect(() => {
    setActiveCell(null);
    setSelection(null);
    setSelectedChartId(null);
    setSelectedImageId(null);
    selectionAnchorRef.current = null;
    setSortState(null);
  }, [activeTabIndex]);

  const setActiveSheetIndex = React.useCallback((index: number) => {
    setActiveSheetIndexState((currentIndex) => {
      if (index < 0 || index >= sheets.length) {
        return currentIndex;
      }
      const targetSheet = sheets[index];
      const tabIndex = tabs.findIndex((tab) => tab.kind === "sheet" && tab.workbookSheetIndex === targetSheet?.workbookSheetIndex);
      if (tabIndex >= 0) {
        setActiveTabIndexState(tabIndex);
      }
      return index;
    });
  }, [sheets, tabs]);

  const setActiveTabIndex = React.useCallback((index: number) => {
    setActiveTabIndexState((currentIndex) => {
      if (index < 0 || index >= tabs.length) {
        return currentIndex;
      }

      const targetTab = tabs[index];
      if (targetTab?.kind === "sheet" && typeof targetTab.sheetIndex === "number") {
        setActiveSheetIndexState(targetTab.sheetIndex);
      }
      return index;
    });
  }, [tabs]);

  const setZoomScale = React.useCallback((nextZoomScale: number) => {
    const normalizedZoomScale = clampZoomScale(nextZoomScale);
    setZoomScaleOverridesByTabId((current) => {
      if (current[activeZoomTabKey] === normalizedZoomScale) {
        return current;
      }

      return {
        ...current,
        [activeZoomTabKey]: normalizedZoomScale
      };
    });
  }, [activeZoomTabKey]);

  const resetZoom = React.useCallback(() => {
    setZoomScaleOverridesByTabId((current) => {
      if (current[activeZoomTabKey] === undefined) {
        return current;
      }

      const next = { ...current };
      delete next[activeZoomTabKey];
      return next;
    });
  }, [activeZoomTabKey]);

  const zoomIn = React.useCallback(() => {
    setZoomScale(resolveNextZoomScale(zoomScale, 1));
  }, [setZoomScale, zoomScale]);

  const zoomOut = React.useCallback(() => {
    setZoomScale(resolveNextZoomScale(zoomScale, -1));
  }, [setZoomScale, zoomScale]);

  React.useEffect(() => {
    setActiveTabIndexState((current) => {
      if (tabs.length === 0) {
        return 0;
      }
      return Math.min(current, tabs.length - 1);
    });
  }, [tabs.length]);

  const continueDeferredLoad = React.useCallback(() => {
    const deferredBuffer = deferredBufferRef.current;
    if (!deferredBuffer) {
      return;
    }

    setIsLoading(true);
    setError(null);

    if (maxFileSizeBytes > 0 && deferredBuffer.byteLength > maxFileSizeBytes) {
      deferredBufferRef.current = null;
      setDeferredLoadFileSize(null);
      setWorkbook(null);
      setSheets([]);
      clearChartAssets();
      setWorkerTablesByWorkbookSheetIndex([]);
      clearImageAssets();
      setShouldAutoCalculate(false);
      setIsWorkerBacked(false);
      setSortState(null);
      setError(new XlsxFileSizeLimitExceededError(deferredBuffer.byteLength, maxFileSizeBytes));
      setIsLoading(false);
      return;
    }

    const preflight = preflightWorkbookBuffer(deferredBuffer);
    if (preflight?.tooLarge) {
      deferredBufferRef.current = null;
      setDeferredLoadFileSize(null);
      setWorkbook(null);
      setSheets([]);
      clearChartAssets();
      setWorkerTablesByWorkbookSheetIndex([]);
      clearImageAssets();
      setShouldAutoCalculate(false);
      setIsWorkerBacked(false);
      setSortState(null);
      setError(createWorkbookTooLargeError(preflight));
      setIsLoading(false);
      return;
    }

    const shouldForceReadOnly = shouldForceReadOnlyForBuffer(deferredBuffer.byteLength);
    setForcedReadOnly(shouldForceReadOnly);
    const shouldUseWorkerForLoad = workerSupported && (requestedReadOnly || shouldForceReadOnly);

    if (shouldUseWorkerForLoad) {
      void getWorkerClient().loadWorkbook(deferredBuffer)
        .then((snapshot) => {
          if (hasIncompleteWorkerChartSnapshot(snapshot)) {
            throw new Error("Worker chart payload incomplete");
          }
          deferredBufferRef.current = null;
          setDeferredLoadFileSize(null);
          setWorkbook(null);
          setSheets(snapshot.sheets);
          setChartsByWorkbookSheetIndex(snapshot.chartsByWorkbookSheetIndex);
          setChartsheets(snapshot.chartsheets);
          setTabs(snapshot.tabs);
          chartAssetsRef.current = null;
          setWorkerTablesByWorkbookSheetIndex(snapshot.tablesByWorkbookSheetIndex);
          setShouldAutoCalculate(false);
          setIsWorkerBacked(true);
          setSortState(null);
          setIsChartsLoading(false);
          setIsLoading(false);
        })
        .catch(async (workerError: unknown) => {
          if (!shouldFallbackFromWorkerError(workerError)) {
            throw workerError;
          }

          disposeWorkerClient();
          const { imageAssets: nextImageAssets, parsedWorkbook: nextParsedWorkbook } = await loadWorkbookOnMainThread(deferredBuffer);
          deferredBufferRef.current = null;
          setDeferredLoadFileSize(null);
          setImageAssets(nextImageAssets);
          setWorkbook(nextParsedWorkbook.workbook);
          const nextSheets = buildSheetList(
            nextParsedWorkbook.workbook,
            nextImageAssets.sheetStatesByWorkbookSheetIndex,
            nextImageAssets.themePalette,
            nextImageAssets.styleById,
            nextImageAssets.namedCellStyleByName,
            nextImageAssets.tableStyleByName
          );
          setSheets(nextSheets);
          startChartDisplayHydration(deferredBuffer, nextParsedWorkbook.workbook, nextSheets);
          setShouldAutoCalculate(nextParsedWorkbook.shouldAutoCalculate);
          setWorkerTablesByWorkbookSheetIndex([]);
          setIsWorkerBacked(false);
          setSortState(null);
          setIsLoading(false);
        })
        .catch((nextError: unknown) => {
          deferredBufferRef.current = null;
          setDeferredLoadFileSize(null);
          setWorkbook(null);
          setSheets([]);
          clearChartAssets();
          setWorkerTablesByWorkbookSheetIndex([]);
          clearImageAssets();
          setShouldAutoCalculate(false);
          setIsWorkerBacked(false);
          setSortState(null);
          setError(nextError instanceof Error ? nextError : new Error("Could not load workbook."));
          setIsLoading(false);
        });
      return;
    }

    void parseWorkbookBuffer(deferredBuffer)
      .then((nextParsedWorkbook) => {
        const nextImageAssets = loadWorkbookImageAssets(new Uint8Array(deferredBuffer), nextParsedWorkbook.workbook);
        deferredBufferRef.current = null;
        setDeferredLoadFileSize(null);
        setImageAssets(nextImageAssets);
        setWorkbook(nextParsedWorkbook.workbook);
        const nextSheets = buildSheetList(
          nextParsedWorkbook.workbook,
          nextImageAssets.sheetStatesByWorkbookSheetIndex,
          nextImageAssets.themePalette,
          nextImageAssets.styleById,
          nextImageAssets.namedCellStyleByName,
          nextImageAssets.tableStyleByName
        );
        setSheets(nextSheets);
        startChartDisplayHydration(deferredBuffer, nextParsedWorkbook.workbook, nextSheets);
        setShouldAutoCalculate(nextParsedWorkbook.shouldAutoCalculate);
        setWorkerTablesByWorkbookSheetIndex([]);
        setIsWorkerBacked(false);
        setSortState(null);
        setIsLoading(false);
      })
      .catch((nextError: unknown) => {
        deferredBufferRef.current = null;
        setDeferredLoadFileSize(null);
        setWorkbook(null);
        setSheets([]);
        clearChartAssets();
        setWorkerTablesByWorkbookSheetIndex([]);
        clearImageAssets();
        setShouldAutoCalculate(false);
        setIsWorkerBacked(false);
        setSortState(null);
        setError(nextError instanceof Error ? nextError : new Error("Could not load workbook."));
        setIsLoading(false);
      });
  }, [
    clearChartAssets,
    clearImageAssets,
    disposeWorkerClient,
    getWorkerClient,
    loadWorkbookOnMainThread,
    requestedReadOnly,
    setImageAssets,
    startChartDisplayHydration,
    hasIncompleteWorkerChartSnapshot,
    maxFileSizeBytes,
    shouldFallbackFromWorkerError,
    shouldForceReadOnlyForBuffer,
    workerSupported
  ]);

  const maybeRecalculateWorkbook = React.useCallback((targetWorkbook: Workbook) => {
    if (!shouldAutoCalculate) {
      return;
    }

    targetWorkbook.calculate();
  }, [shouldAutoCalculate]);

  const getActiveWorksheet = React.useCallback(() => {
    if (!workbook || !activeSheet) {
      return null;
    }

    return workbook.getSheet(activeSheet.workbookSheetIndex);
  }, [activeSheet, workbook]);

  const activeTableMetadata = imageAssetsRef.current?.tableMetadataByWorkbookSheetIndex[activeSheet?.workbookSheetIndex ?? -1] ?? null;
  const tables = React.useMemo(
    () => (
      isWorkerBacked
        ? workerTablesByWorkbookSheetIndex[activeSheet?.workbookSheetIndex ?? -1] ?? []
        : mapWorksheetTables(getActiveWorksheet(), activeTableMetadata)
    ),
    [activeSheet?.workbookSheetIndex, activeTableMetadata, getActiveWorksheet, isWorkerBacked, revision, workerTablesByWorkbookSheetIndex]
  );

  const getCellSnapshotAsync = React.useCallback((workbookSheetIndex: number, row: number, col: number) => {
    if (!isWorkerBacked) {
      return Promise.resolve({
        displayValue: "",
        formula: ""
      });
    }

    return getWorkerClient().getCellSnapshot(workbookSheetIndex, row, col);
  }, [getWorkerClient, isWorkerBacked]);

  const getRowsBatchAsync = React.useCallback((workbookSheetIndex: number, startRow: number, rowCount: number) => {
    if (!isWorkerBacked) {
      return Promise.resolve(null);
    }

    return getWorkerClient().getRowsBatch(workbookSheetIndex, startRow, rowCount);
  }, [getWorkerClient, isWorkerBacked]);

  const visibleSheetIndexByWorkbookSheetIndex = React.useMemo(
    () => new Map(sheets.map((sheet, index) => [sheet.workbookSheetIndex, index])),
    [sheets]
  );

  const mapPublicChart = React.useCallback((chart: XlsxChart) => {
    const visibleSheetIndex = visibleSheetIndexByWorkbookSheetIndex.get(chart.workbookSheetIndex);
    return {
      ...chart,
      sheetIndex: visibleSheetIndex ?? chart.workbookSheetIndex
    };
  }, [visibleSheetIndexByWorkbookSheetIndex]);

  const mapPublicImage = React.useCallback((image: XlsxImage) => {
    const visibleSheetIndex = visibleSheetIndexByWorkbookSheetIndex.get(image.workbookSheetIndex);
    return {
      ...image,
      sheetIndex: visibleSheetIndex ?? image.workbookSheetIndex
    };
  }, [visibleSheetIndexByWorkbookSheetIndex]);

  const publicChartsByWorkbookSheetIndex = React.useMemo(
    () => chartsByWorkbookSheetIndex.map((sheetCharts) => sheetCharts.map(mapPublicChart)),
    [chartsByWorkbookSheetIndex, mapPublicChart]
  );
  const publicChartById = React.useMemo(() => {
    const lookup = new Map<string, XlsxChart>();
    for (const sheetCharts of publicChartsByWorkbookSheetIndex) {
      for (const chart of sheetCharts) {
        lookup.set(chart.id, chart);
      }
    }
    return lookup;
  }, [publicChartsByWorkbookSheetIndex]);

  const getSheetCharts = React.useCallback((sheetIndex = activeSheetIndex) => {
    const targetSheet = sheets[sheetIndex];
    if (!targetSheet) {
      return [];
    }

    return publicChartsByWorkbookSheetIndex[targetSheet.workbookSheetIndex] ?? [];
  }, [activeSheetIndex, publicChartsByWorkbookSheetIndex, sheets]);

  const getChartById = React.useCallback((id: string) => {
    return publicChartById.get(id) ?? null;
  }, [publicChartById]);

  const getChartsheetById = React.useCallback((id: string) => (
    chartsheets.find((chartsheet) => chartsheet.id === id) ?? null
  ), [chartsheets]);

  const charts = React.useMemo(() => {
    if (activeTab?.kind === "chartsheet" && typeof activeTab.chartsheetIndex === "number") {
      const chartsheet = chartsheets[activeTab.chartsheetIndex];
      return (chartsheet?.chartIds ?? []).map((id) => getChartById(id)).filter((value): value is XlsxChart => Boolean(value));
    }

    return getSheetCharts(activeSheetIndex);
  }, [activeSheetIndex, activeTab, chartsheets, getChartById, getSheetCharts]);

  const selectedChart = React.useMemo(
    () => (selectedChartId ? getChartById(selectedChartId) : null),
    [getChartById, selectedChartId]
  );

  const selectChart = React.useCallback((id: string | null) => {
    setSelectedImageId(null);
    setSelectedChartId(id);
  }, []);

  const clearSelectedChart = React.useCallback(() => {
    setSelectedChartId(null);
  }, []);

  const getSheetImages = React.useCallback((sheetIndex = activeSheetIndex) => {
    const targetSheet = sheets[sheetIndex];
    if (!targetSheet) {
      return [];
    }

    return (imagesByWorkbookSheetIndex[targetSheet.workbookSheetIndex] ?? []).map(mapPublicImage);
  }, [activeSheetIndex, imagesByWorkbookSheetIndex, mapPublicImage, sheets]);

  const images = React.useMemo(() => getSheetImages(activeSheetIndex), [activeSheetIndex, getSheetImages]);

  const getSheetShapes = React.useCallback((sheetIndex = activeSheetIndex) => {
    const targetSheet = sheets[sheetIndex];
    if (!targetSheet) {
      return [];
    }

    return (shapesByWorkbookSheetIndex[targetSheet.workbookSheetIndex] ?? []).map((shape) => {
      const visibleSheetIndex = visibleSheetIndexByWorkbookSheetIndex.get(shape.workbookSheetIndex);
      return {
        ...shape,
        sheetIndex: visibleSheetIndex ?? shape.workbookSheetIndex
      };
    });
  }, [activeSheetIndex, shapesByWorkbookSheetIndex, sheets, visibleSheetIndexByWorkbookSheetIndex]);

  const shapes = React.useMemo(() => getSheetShapes(activeSheetIndex), [activeSheetIndex, getSheetShapes]);

  const getImageById = React.useCallback((id: string) => {
    for (const sheetImages of imagesByWorkbookSheetIndex) {
      const match = sheetImages?.find((image) => image.id === id);
      if (match) {
        return mapPublicImage(match);
      }
    }

    return null;
  }, [imagesByWorkbookSheetIndex, mapPublicImage]);

  const selectedImage = React.useMemo(
    () => (selectedImageId ? getImageById(selectedImageId) : null),
    [getImageById, selectedImageId]
  );

  const selectImage = React.useCallback((id: string | null) => {
    setSelectedChartId(null);
    setSelectedImageId(id);
  }, []);

  const clearSelectedImage = React.useCallback(() => {
    setSelectedImageId(null);
  }, []);

  const getColumnWidthPx = React.useCallback((worksheet: ReturnType<Workbook["getSheet"]>, col: number) => {
    const sheetState = imageAssetsRef.current?.sheetStatesByWorkbookSheetIndex[activeSheet?.workbookSheetIndex ?? -1] ?? null;
    const width = worksheet.getColumnWidth(col);
    const showGridLines = activeSheet?.showGridLines ?? true;
    if (width !== undefined && width !== null) {
      return resolveRenderedSheetAxisPixels(
        resolveSheetColumnWidthPixels(width, sheetState?.columnWidthCharacterWidthPx),
        showGridLines
      );
    }

    return resolveRenderedSheetAxisPixels(
      sheetState?.colWidthOverridesPx?.[col] ?? sheetState?.defaultColWidthPx ?? DEFAULT_COL_WIDTH,
      showGridLines
    );
  }, [activeSheet?.showGridLines, activeSheet?.workbookSheetIndex]);

  const getRowHeightPx = React.useCallback((worksheet: ReturnType<Workbook["getSheet"]>, row: number) => {
    const sheetState = imageAssetsRef.current?.sheetStatesByWorkbookSheetIndex[activeSheet?.workbookSheetIndex ?? -1] ?? null;
    const height = worksheet.getRowHeight(row);
    const showGridLines = activeSheet?.showGridLines ?? true;
    if (height !== undefined && height !== null) {
      return resolveRenderedSheetAxisPixels(resolveSheetRowHeightPixels(height), showGridLines);
    }

    return resolveRenderedSheetAxisPixels(
      sheetState?.rowHeightOverridesPx?.[row] ?? sheetState?.defaultRowHeightPx ?? DEFAULT_ROW_HEIGHT,
      showGridLines
    );
  }, [activeSheet?.showGridLines, activeSheet?.workbookSheetIndex]);

  const getCellDisplayValue = React.useCallback((cell?: XlsxCellAddress | null) => {
    if (cell && activeSheet) {
      const workerSnapshot = workerCellSnapshotCacheRef.current.get(`${activeSheet.workbookSheetIndex}:${cell.row}:${cell.col}`);
      if (workerSnapshot) {
        return workerSnapshot.displayValue;
      }
    }

    const worksheet = getActiveWorksheet();
    if (!worksheet || !cell) {
      return "";
    }

    const formula = worksheet.getFormulaAt(cell.row, cell.col);
    const cachedFormulaValue = formula ? activeSheet?.cachedFormulaValues?.[cellAddressToA1(cell)] : undefined;
    const formatted = worksheet.getFormattedValueAt(cell.row, cell.col);
    if (formatted && !(formula && cachedFormulaValue !== undefined && formatted.startsWith("#"))) {
      return decodeHtmlEntities(formatted);
    }

    const calculated = worksheet.getCalculatedValueAt(cell.row, cell.col);
    if (formula && cachedFormulaValue !== undefined && calculated.is_error) {
      return cachedFormulaValue;
    }
    if (calculated.is_error) {
      return calculated.asError() ?? "";
    }
    if (calculated.is_empty) {
      return "";
    }

    return calculated.toString();
  }, [activeSheet, getActiveWorksheet]);

  const getCellFormula = React.useCallback((cell?: XlsxCellAddress | null) => {
    if (cell && activeSheet) {
      const workerSnapshot = workerCellSnapshotCacheRef.current.get(`${activeSheet.workbookSheetIndex}:${cell.row}:${cell.col}`);
      if (workerSnapshot) {
        return workerSnapshot.formula;
      }
    }

    const worksheet = getActiveWorksheet();
    if (!worksheet || !cell) {
      return "";
    }

    return worksheet.getFormulaAt(cell.row, cell.col) ?? "";
  }, [activeSheet, getActiveWorksheet]);

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
        const rawStyle = (
          worksheet.getCellStyleAt(row, col) as Record<string, unknown> | null | undefined
        ) ?? resolveInheritedCellStyle(activeSheet, row, col);
        const cellStyles: string[] = [
          "padding:2px 4px",
          "white-space:pre-wrap",
          "vertical-align:top"
        ];

        const fill = rawStyle?.fill as Record<string, unknown> | undefined;
        if (fill) {
          const fillStyle = resolveWorkbookFillStyle(fill, activeSheet?.themePalette);
          if (fillStyle.backgroundColor && fillStyle.backgroundColor.toLowerCase() !== "#ffffff") {
            cellStyles.push(`background-color:${fillStyle.backgroundColor}`);
          }
          if (fillStyle.backgroundImage) {
            cellStyles.push(`background-image:${fillStyle.backgroundImage}`);
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
          const fontColor = resolveWorkbookColor(font.color as Record<string, unknown> | undefined, activeSheet?.themePalette);
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

  React.useEffect(() => {
    if (!isWorkerBacked || !deferredMetadataSheet || !deferredMetadataCell) {
      return;
    }

    const cacheKey = `${deferredMetadataSheet.workbookSheetIndex}:${deferredMetadataCell.row}:${deferredMetadataCell.col}`;
    if (workerCellSnapshotCacheRef.current.has(cacheKey)) {
      return;
    }

    let isCurrent = true;
    void getCellSnapshotAsync(deferredMetadataSheet.workbookSheetIndex, deferredMetadataCell.row, deferredMetadataCell.col)
      .then((snapshot) => {
        if (!isCurrent) {
          return;
        }

        workerCellSnapshotCacheRef.current.set(cacheKey, snapshot);
        setWorkerCellSnapshotRevision((current) => current + 1);
      })
      .catch(() => {
        if (!isCurrent) {
          return;
        }

        workerCellSnapshotCacheRef.current.set(cacheKey, {
          displayValue: "",
          formula: ""
        });
        setWorkerCellSnapshotRevision((current) => current + 1);
      });

    return () => {
      isCurrent = false;
    };
  }, [deferredMetadataCell, deferredMetadataSheet, getCellSnapshotAsync, isWorkerBacked]);

  const activeCellAddress = React.useMemo(() => (activeCell ? cellAddressToA1(activeCell) : null), [activeCell]);
  const selectedRangeAddress = React.useMemo(() => (selection ? rangeToA1(selection) : null), [selection]);
  const selectedValue = React.useMemo(
    () => getCellDisplayValue(deferredMetadataCell),
    [deferredMetadataCell, getCellDisplayValue, revision, workerCellSnapshotRevision]
  );
  const selectedFormula = React.useMemo(
    () => getCellFormula(deferredMetadataCell),
    [deferredMetadataCell, getCellFormula, revision, workerCellSnapshotRevision]
  );
  const isLoadDeferred = deferredLoadFileSize !== null;
  const canLoadDeferred = !isLoading && isLoadDeferred;
  const canUndo = !readOnly && undoStackRef.current.length > 0;
  const canRedo = !readOnly && redoStackRef.current.length > 0;

  const createSavedWorkbookBytes = React.useCallback((targetWorkbook: Workbook) => {
    const sanitizedBytes = sanitizeSavedWorkbookBytes(targetWorkbook.saveXlsxBytes());
    return mergeWorkbookImageAssets(sanitizedBytes, imageAssetsRef.current, sheetOriginsRef.current);
  }, []);

  const createHistoryEntry = React.useCallback((): SnapshotHistoryEntry | null => {
    if (!workbook) {
      return null;
    }

    return {
      kind: "snapshot",
      activeCell,
      activeSheetIndex,
      bytes: createSavedWorkbookBytes(workbook),
      selection
    };
  }, [activeCell, activeSheetIndex, createSavedWorkbookBytes, selection, workbook]);

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
    const nextImageAssets = loadWorkbookImageAssets(entry.bytes, nextWorkbook);
    const nextSheets = buildSheetList(
      nextWorkbook,
      nextImageAssets.sheetStatesByWorkbookSheetIndex,
      nextImageAssets.themePalette,
      nextImageAssets.styleById,
      nextImageAssets.namedCellStyleByName,
      nextImageAssets.tableStyleByName
    );
    const nextSheetIndex = Math.max(0, Math.min(entry.activeSheetIndex, Math.max(0, nextSheets.length - 1)));

    setError(null);
    setIsLoading(false);
    setImageAssets(nextImageAssets);
    setWorkbook(nextWorkbook);
    setSheets(nextSheets);
    const nextChartAssets = loadWorkbookChartAssets(nextWorkbook, nextImageAssets, buildVisibleSheetIndexMap(nextSheets));
    setChartAssets(nextChartAssets);
    setActiveSheetIndexState(nextSheetIndex);
    const nextTabIndex = nextChartAssets.tabs.findIndex((tab) => tab.kind === "sheet" && tab.sheetIndex === nextSheetIndex);
    if (nextTabIndex >= 0) {
      setActiveTabIndexState(nextTabIndex);
    }
    setActiveCell(entry.activeCell);
    setSelection(entry.selection);
    selectionAnchorRef.current = entry.selection ? normalizeRange(entry.selection).start : entry.activeCell;
    setRevision((current) => current + 1);
  }, [setChartAssets, setImageAssets]);

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
    maybeRecalculateWorkbook(workbook);
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
  }, [maybeRecalculateWorkbook, refreshWorkbookState, sheets, workbook]);

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
    maybeRecalculateWorkbook(workbook);
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
  }, [maybeRecalculateWorkbook, refreshWorkbookState, sheets, workbook]);

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

  const sortTable = React.useCallback((tableName: string, columnIndex: number, direction: XlsxTableSortDirection) => {
    const worksheet = getActiveWorksheet();
    const targetTable = tables.find((table) => table.name === tableName || table.displayName === tableName);
    if (readOnly || !worksheet || !workbook || !activeSheet || !targetTable) {
      return;
    }

    const dataStartRow = targetTable.start.row + Math.max(targetTable.headerRowCount, 1);
    const totalsRowOffset = targetTable.totalsRowShown ? Math.max(targetTable.totalsRowCount, 1) : 0;
    const dataEndRow = targetTable.end.row - totalsRowOffset;
    const startCol = targetTable.start.col;
    const endCol = targetTable.end.col;
    const sortCol = startCol + columnIndex;

    if (columnIndex < 0 || sortCol > endCol || dataStartRow > dataEndRow) {
      return;
    }

    const rows: Array<{
      cells: CellMutationState[];
      index: number;
      sortBoolean: boolean | undefined;
      sortEmpty: boolean;
      sortNumber: number | undefined;
      sortText: string;
    }> = [];

    for (let row = dataStartRow; row <= dataEndRow; row += 1) {
      const cells: CellMutationState[] = [];
      for (let col = startCol; col <= endCol; col += 1) {
        cells.push({
          formula: worksheet.getFormulaAt(row, col) ?? null,
          value: worksheet.getCellAt(row, col).toJs()
        });
      }

      const calculated = worksheet.getCalculatedValueAt(row, sortCol);
      const formatted = decodeHtmlEntities(worksheet.getFormattedValueAt(row, sortCol) ?? "");
      rows.push({
        cells,
        index: row,
        sortBoolean: calculated.is_boolean ? calculated.asBoolean() : undefined,
        sortEmpty: calculated.is_empty || formatted.length === 0,
        sortNumber: calculated.is_number ? calculated.asNumber() : undefined,
        sortText: calculated.is_text ? (calculated.asText() ?? formatted) : formatted
      });
    }

    const sortedRows = [...rows].sort((left, right) => {
      if (left.sortEmpty !== right.sortEmpty) {
        return left.sortEmpty ? 1 : -1;
      }

      if (left.sortNumber !== undefined && right.sortNumber !== undefined) {
        return direction === "ascending" ? left.sortNumber - right.sortNumber : right.sortNumber - left.sortNumber;
      }

      if (left.sortBoolean !== undefined && right.sortBoolean !== undefined) {
        const leftValue = left.sortBoolean ? 1 : 0;
        const rightValue = right.sortBoolean ? 1 : 0;
        return direction === "ascending" ? leftValue - rightValue : rightValue - leftValue;
      }

      const comparedText = left.sortText.localeCompare(right.sortText, undefined, { numeric: true, sensitivity: "base" });
      return direction === "ascending" ? comparedText : -comparedText;
    });

    if (sortedRows.every((row, index) => row.index === rows[index]?.index)) {
      setSortState({ columnIndex, direction, tableName: targetTable.name });
      return;
    }

    const mutations: RangeCellMutation[] = [];
    for (let rowOffset = 0; rowOffset < rows.length; rowOffset += 1) {
      const targetRow = dataStartRow + rowOffset;
      const sourceRow = sortedRows[rowOffset];
      const beforeRow = rows[rowOffset];
      if (!sourceRow || !beforeRow) {
        continue;
      }

      for (let colOffset = 0; colOffset <= endCol - startCol; colOffset += 1) {
        const before = beforeRow.cells[colOffset];
        const after = sourceRow.cells[colOffset];
        if (!before || !after) {
          continue;
        }

        const cell = { row: targetRow, col: startCol + colOffset };
        applyCellMutationState(worksheet, cell, after);
        mutations.push({
          after,
          before,
          cell
        });
      }
    }

    maybeRecalculateWorkbook(workbook);
    refreshWorkbookState(workbook);
    setSortState({ columnIndex, direction, tableName: targetTable.name });
    recordRangeEditHistory(mutations, selection, activeCell);
  }, [
    activeCell,
    activeSheet,
    getActiveWorksheet,
    maybeRecalculateWorkbook,
    readOnly,
    recordRangeEditHistory,
    refreshWorkbookState,
    selection,
    tables,
    workbook
  ]);

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

    downloadBytes(createSavedWorkbookBytes(workbook), `${fileStem(displayFileName)}.xlsx`, XLSX_MIME_TYPE);
  }, [createSavedWorkbookBytes, displayFileName, workbook]);

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
    worksheet.setColumnWidth(
      col,
      pxToSheetColumnWidth(resolveContentSheetAxisPixels(widthPx, activeSheet.showGridLines))
    );
    refreshWorkbookState(workbook);
  }, [activeSheet, readOnly, recordHistoryBeforeMutation, refreshWorkbookState, workbook]);

  const resizeRow = React.useCallback((row: number, heightPx: number) => {
    if (readOnly || !workbook || !activeSheet) {
      return;
    }

    recordHistoryBeforeMutation();
    const worksheet = workbook.getSheet(activeSheet.workbookSheetIndex);
    worksheet.setRowHeight(
      row,
      pxToSheetRowHeight(resolveContentSheetAxisPixels(heightPx, activeSheet.showGridLines))
    );
    refreshWorkbookState(workbook);
  }, [activeSheet, readOnly, recordHistoryBeforeMutation, refreshWorkbookState, workbook]);

  const resolveAnchoredObjectRect = React.useCallback((
    anchor: XlsxImage["anchor"],
    worksheet: ReturnType<Workbook["getSheet"]>
  ): XlsxImageRect => {
    const resolveAxisSum = (
      index: number,
      getSize: (target: number) => number
    ) => {
      let total = 0;
      for (let cursor = 0; cursor < index; cursor += 1) {
        total += getSize(cursor);
      }
      return total;
    };

    if (anchor.kind === "absolute") {
      return {
        height: anchor.sizeEmu.cy / 9525,
        left: GRID_ROW_HEADER_WIDTH + anchor.positionEmu.x / 9525,
        top: GRID_HEADER_HEIGHT + anchor.positionEmu.y / 9525,
        width: anchor.sizeEmu.cx / 9525
      };
    }

    const left = GRID_ROW_HEADER_WIDTH + resolveAxisSum(anchor.from.col, (col) => getColumnWidthPx(worksheet, col)) + anchor.from.colOffsetEmu / 9525;
    const top = GRID_HEADER_HEIGHT + resolveAxisSum(anchor.from.row, (row) => getRowHeightPx(worksheet, row)) + anchor.from.rowOffsetEmu / 9525;

    if (anchor.kind === "one-cell") {
      return {
        height: anchor.sizeEmu.cy / 9525,
        left,
        top,
        width: anchor.sizeEmu.cx / 9525
      };
    }

    const right = GRID_ROW_HEADER_WIDTH + resolveAxisSum(anchor.to.col, (col) => getColumnWidthPx(worksheet, col)) + anchor.to.colOffsetEmu / 9525;
    const bottom = GRID_HEADER_HEIGHT + resolveAxisSum(anchor.to.row, (row) => getRowHeightPx(worksheet, row)) + anchor.to.rowOffsetEmu / 9525;

    return {
      height: Math.max(1, bottom - top),
      left,
      top,
      width: Math.max(1, right - left)
    };
  }, [getColumnWidthPx, getRowHeightPx]);

  const setChartRect = React.useCallback((id: string, rect: XlsxImageRect) => {
    const hydratedChartAssets = ensureChartAssetsHydrated(workbook, sheets);
    if (readOnly || !workbook || !activeSheet || !imageAssetsRef.current || !hydratedChartAssets) {
      return;
    }

    const worksheet = workbook.getSheet(activeSheet.workbookSheetIndex);
    const currentChart = getChartById(id);
    if (!currentChart || currentChart.editable === false || currentChart.workbookSheetIndex !== activeSheet.workbookSheetIndex) {
      return;
    }

    const nextAnchor = rectToImageAnchor(rect, currentChart.anchor, {
      contentOffsetLeft: GRID_ROW_HEADER_WIDTH,
      contentOffsetTop: GRID_HEADER_HEIGHT,
      getColumnWidthPx: (col) => getColumnWidthPx(worksheet, col),
      getRowHeightPx: (row) => getRowHeightPx(worksheet, row)
    });

    recordHistoryBeforeMutation();
    updateWorkbookChartAnchor(imageAssetsRef.current, hydratedChartAssets, id, nextAnchor);

    hydratedChartAssets.chartsByWorkbookSheetIndex = hydratedChartAssets.chartsByWorkbookSheetIndex.map((sheetCharts) => (
      sheetCharts.map((chart) => chart.id === id ? { ...chart, anchor: nextAnchor } : chart)
    ));

    setChartsByWorkbookSheetIndex((current) => current.map((sheetCharts) => (
      sheetCharts.map((chart) => chart.id === id ? { ...chart, anchor: nextAnchor } : chart)
    )));
    setRevision((current) => current + 1);
  }, [
    activeSheet,
    getChartById,
    getColumnWidthPx,
    getRowHeightPx,
    ensureChartAssetsHydrated,
    readOnly,
    recordHistoryBeforeMutation,
    sheets,
    workbook
  ]);

  const setImageRect = React.useCallback((id: string, rect: XlsxImageRect) => {
    if (readOnly || !workbook || !activeSheet || !imageAssetsRef.current) {
      return;
    }

    const worksheet = workbook.getSheet(activeSheet.workbookSheetIndex);
    const currentImage = getImageById(id);
    if (!currentImage || currentImage.editable === false || currentImage.workbookSheetIndex !== activeSheet.workbookSheetIndex) {
      return;
    }

    const nextAnchor = rectToImageAnchor(rect, currentImage.anchor, {
      contentOffsetLeft: GRID_ROW_HEADER_WIDTH,
      contentOffsetTop: GRID_HEADER_HEIGHT,
      getColumnWidthPx: (col) => getColumnWidthPx(worksheet, col),
      getRowHeightPx: (row) => getRowHeightPx(worksheet, row)
    });

    recordHistoryBeforeMutation();
    if (!updateWorkbookImageAnchor(imageAssetsRef.current, id, nextAnchor)) {
      return;
    }

    setImagesByWorkbookSheetIndex([...imageAssetsRef.current.imagesByWorkbookSheetIndex]);
    setRevision((current) => current + 1);
  }, [
    activeSheet,
    getColumnWidthPx,
    getImageById,
    getRowHeightPx,
    readOnly,
    recordHistoryBeforeMutation,
    workbook
  ]);

  const moveChartBy = React.useCallback((id: string, deltaX: number, deltaY: number) => {
    const currentChart = getChartById(id);
    if (!currentChart || currentChart.editable === false) {
      return;
    }

    const worksheet = getActiveWorksheet();
    if (!worksheet) {
      return;
    }

    const currentRect = resolveAnchoredObjectRect(currentChart.anchor, worksheet);
    setChartRect(id, {
      ...currentRect,
      left: currentRect.left + deltaX,
      top: currentRect.top + deltaY
    });
  }, [getActiveWorksheet, getChartById, resolveAnchoredObjectRect, setChartRect]);

  const moveImageBy = React.useCallback((id: string, deltaX: number, deltaY: number) => {
    const currentImage = getImageById(id);
    if (!currentImage || currentImage.editable === false) {
      return;
    }

    const currentRect = (() => {
      const worksheet = getActiveWorksheet();
      if (!worksheet) {
        return null;
      }

      const resolveAxisSum = (
        index: number,
        getSize: (target: number) => number
      ) => {
        let total = 0;
        for (let cursor = 0; cursor < index; cursor += 1) {
          total += getSize(cursor);
        }
        return total;
      };

      if (currentImage.anchor.kind === "absolute") {
        return {
          height: currentImage.anchor.sizeEmu.cy / 9525,
          left: GRID_ROW_HEADER_WIDTH + currentImage.anchor.positionEmu.x / 9525,
          top: GRID_HEADER_HEIGHT + currentImage.anchor.positionEmu.y / 9525,
          width: currentImage.anchor.sizeEmu.cx / 9525
        };
      }

      const left = GRID_ROW_HEADER_WIDTH + resolveAxisSum(currentImage.anchor.from.col, (col) => getColumnWidthPx(worksheet, col)) + currentImage.anchor.from.colOffsetEmu / 9525;
      const top = GRID_HEADER_HEIGHT + resolveAxisSum(currentImage.anchor.from.row, (row) => getRowHeightPx(worksheet, row)) + currentImage.anchor.from.rowOffsetEmu / 9525;

      if (currentImage.anchor.kind === "one-cell") {
        return {
          height: currentImage.anchor.sizeEmu.cy / 9525,
          left,
          top,
          width: currentImage.anchor.sizeEmu.cx / 9525
        };
      }

      const right = GRID_ROW_HEADER_WIDTH + resolveAxisSum(currentImage.anchor.to.col, (col) => getColumnWidthPx(worksheet, col)) + currentImage.anchor.to.colOffsetEmu / 9525;
      const bottom = GRID_HEADER_HEIGHT + resolveAxisSum(currentImage.anchor.to.row, (row) => getRowHeightPx(worksheet, row)) + currentImage.anchor.to.rowOffsetEmu / 9525;

      return {
        height: Math.max(1, bottom - top),
        left,
        top,
        width: Math.max(1, right - left)
      };
    })();

    if (!currentRect) {
      return;
    }

    setImageRect(id, {
      ...currentRect,
      left: currentRect.left + deltaX,
      top: currentRect.top + deltaY
    });
  }, [getActiveWorksheet, getColumnWidthPx, getImageById, getRowHeightPx, setImageRect]);

  const resizeChartBy = React.useCallback((
    id: string,
    handle: XlsxImageResizeHandlePosition,
    deltaX: number,
    deltaY: number
  ) => {
    const currentChart = getChartById(id);
    if (!currentChart || currentChart.editable === false) {
      return;
    }

    const worksheet = getActiveWorksheet();
    if (!worksheet) {
      return;
    }

    const currentRect = resolveAnchoredObjectRect(currentChart.anchor, worksheet);
    setChartRect(id, resizeImageRect(currentRect, handle, deltaX, deltaY, 48));
  }, [getActiveWorksheet, getChartById, resolveAnchoredObjectRect, setChartRect]);

  const resizeImageBy = React.useCallback((
    id: string,
    handle: XlsxImageResizeHandlePosition,
    deltaX: number,
    deltaY: number
  ) => {
    const currentImage = getImageById(id);
    if (!currentImage || currentImage.editable === false) {
      return;
    }

    const worksheet = getActiveWorksheet();
    if (!worksheet) {
      return;
    }

    const resolveAxisSum = (
      index: number,
      getSize: (target: number) => number
    ) => {
      let total = 0;
      for (let cursor = 0; cursor < index; cursor += 1) {
        total += getSize(cursor);
      }
      return total;
    };

    const left = currentImage.anchor.kind === "absolute"
      ? GRID_ROW_HEADER_WIDTH + currentImage.anchor.positionEmu.x / 9525
      : GRID_ROW_HEADER_WIDTH + resolveAxisSum(currentImage.anchor.from.col, (col) => getColumnWidthPx(worksheet, col)) + currentImage.anchor.from.colOffsetEmu / 9525;
    const top = currentImage.anchor.kind === "absolute"
      ? GRID_HEADER_HEIGHT + currentImage.anchor.positionEmu.y / 9525
      : GRID_HEADER_HEIGHT + resolveAxisSum(currentImage.anchor.from.row, (row) => getRowHeightPx(worksheet, row)) + currentImage.anchor.from.rowOffsetEmu / 9525;
    const width = currentImage.anchor.kind === "two-cell"
      ? Math.max(
          1,
          GRID_ROW_HEADER_WIDTH + resolveAxisSum(currentImage.anchor.to.col, (col) => getColumnWidthPx(worksheet, col)) + currentImage.anchor.to.colOffsetEmu / 9525 - left
        )
      : currentImage.anchor.sizeEmu.cx / 9525;
    const height = currentImage.anchor.kind === "two-cell"
      ? Math.max(
          1,
          GRID_HEADER_HEIGHT + resolveAxisSum(currentImage.anchor.to.row, (row) => getRowHeightPx(worksheet, row)) + currentImage.anchor.to.rowOffsetEmu / 9525 - top
        )
      : currentImage.anchor.sizeEmu.cy / 9525;

    const nextRect = resizeImageRect({ height, left, top, width }, handle, deltaX, deltaY);
    setImageRect(id, nextRect);
  }, [getActiveWorksheet, getColumnWidthPx, getImageById, getRowHeightPx, setImageRect]);

  const updateChart = React.useCallback((id: string, patch: Partial<XlsxChart>) => {
    const currentChart = getChartById(id);
    const hydratedChartAssets = ensureChartAssetsHydrated(workbook, sheets);
    if (readOnly || !currentChart) {
      return;
    }

    recordHistoryBeforeMutation();
    if (patch.anchor && imageAssetsRef.current && hydratedChartAssets) {
      updateWorkbookChartAnchor(imageAssetsRef.current, hydratedChartAssets, id, patch.anchor);
    }
    if (imageAssetsRef.current && hydratedChartAssets) {
      updateWorkbookChartDefinition(imageAssetsRef.current, hydratedChartAssets, id, patch);
    }

    setChartsByWorkbookSheetIndex((current) => current.map((sheetCharts) => (
      sheetCharts.map((chart) => chart.id === id ? { ...chart, ...patch } : chart)
    )));
    setRevision((current) => current + 1);
  }, [ensureChartAssetsHydrated, getChartById, readOnly, recordHistoryBeforeMutation, sheets, workbook]);

  const selectCell = React.useCallback((cell: XlsxCellAddress, options?: { extend?: boolean }) => {
    setSelectedChartId(null);
    setSelectedImageId(null);
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
    setSelectedChartId(null);
    setSelectedImageId(null);
    selectionAnchorRef.current = normalized.start;
    setActiveCell(normalized.end);
    setSelection(normalized);
  }, []);

  const clearSelection = React.useCallback(() => {
    selectionAnchorRef.current = null;
    setActiveCell(null);
    setSelection(null);
    setSelectedChartId(null);
    setSelectedImageId(null);
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

    maybeRecalculateWorkbook(workbook);
    refreshWorkbookState(workbook);
    recordRangeEditHistory(mutations, normalized, activeCell ?? normalized.start);
  }, [
    activeCell,
    captureCellMutationState,
    getActiveWorksheet,
    maybeRecalculateWorkbook,
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

    const nextValue = coerceUserEnteredValue(value);
    worksheet.setCell(cellAddressToA1(cell), nextValue);
    maybeRecalculateWorkbook(workbook);
    refreshWorkbookState(workbook);
    recordCellEditHistory(cell, before, { formula: null, value: nextValue });
  }, [captureCellMutationState, getActiveWorksheet, maybeRecalculateWorkbook, readOnly, recordCellEditHistory, refreshWorkbookState, workbook]);

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
    maybeRecalculateWorkbook(workbook);
    refreshWorkbookState(workbook);
    recordCellEditHistory(cell, before, {
      formula: trimmedFormula || null,
      value: trimmedFormula ? null : ""
    });
  }, [captureCellMutationState, getActiveWorksheet, maybeRecalculateWorkbook, readOnly, recordCellEditHistory, refreshWorkbookState, workbook]);

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

    maybeRecalculateWorkbook(workbook);
    refreshWorkbookState(workbook);
    setSelection(nextRange);
    setActiveCell(nextRange.end);
    selectionAnchorRef.current = nextRange.start;
    recordRangeEditHistory(mutations, nextRange, nextRange.end);
  }, [captureCellMutationState, getActiveWorksheet, maybeRecalculateWorkbook, readOnly, recordRangeEditHistory, refreshWorkbookState, selection, workbook]);

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
    sheetOriginsRef.current = [...sheetOriginsRef.current, null];
    setImagesByWorkbookSheetIndex((current) => [...current, []]);
    setShapesByWorkbookSheetIndex((current) => [...current, []]);
    const nextSheets = buildSheetList(
      workbook,
      imageAssetsRef.current?.sheetStatesByWorkbookSheetIndex,
      imageAssetsRef.current?.themePalette,
      imageAssetsRef.current?.styleById,
      imageAssetsRef.current?.namedCellStyleByName,
      imageAssetsRef.current?.tableStyleByName
    );
    setSheets(nextSheets);
    const nextChartAssets = imageAssetsRef.current
      ? loadWorkbookChartAssets(workbook, imageAssetsRef.current, buildVisibleSheetIndexMap(nextSheets))
      : null;
    if (imageAssetsRef.current) {
      setChartAssets(nextChartAssets);
    }
    const nextIndex = nextSheets.findIndex((sheet) => sheet.name === candidate);
    setActiveSheetIndexState(nextIndex >= 0 ? nextIndex : 0);
    const nextTabIndex = nextChartAssets?.tabs.findIndex((tab) => tab.kind === "sheet" && tab.name === candidate) ?? -1;
    if (nextTabIndex >= 0) {
      setActiveTabIndexState(nextTabIndex);
    }
    setRevision((current) => current + 1);
  }, [readOnly, recordHistoryBeforeMutation, setChartAssets, workbook]);

  const removeActiveSheet = React.useCallback(() => {
    if (readOnly || !workbook || !activeSheet) {
      return;
    }

    recordHistoryBeforeMutation();
    workbook.removeSheet(activeSheet.workbookSheetIndex);
    sheetOriginsRef.current = sheetOriginsRef.current.filter((_, index) => index !== activeSheet.workbookSheetIndex);
    setImagesByWorkbookSheetIndex((current) => current.filter((_, index) => index !== activeSheet.workbookSheetIndex));
    setShapesByWorkbookSheetIndex((current) => current.filter((_, index) => index !== activeSheet.workbookSheetIndex));
    if (imageAssetsRef.current) {
      imageAssetsRef.current.sheetStatesByWorkbookSheetIndex = imageAssetsRef.current.sheetStatesByWorkbookSheetIndex.filter(
        (_, index) => index !== activeSheet.workbookSheetIndex
      );
    }
    const nextSheets = buildSheetList(
      workbook,
      imageAssetsRef.current?.sheetStatesByWorkbookSheetIndex,
      imageAssetsRef.current?.themePalette,
      imageAssetsRef.current?.styleById,
      imageAssetsRef.current?.namedCellStyleByName,
      imageAssetsRef.current?.tableStyleByName
    );
    setSheets(nextSheets);
    if (imageAssetsRef.current) {
      setChartAssets(loadWorkbookChartAssets(workbook, imageAssetsRef.current, buildVisibleSheetIndexMap(nextSheets)));
    }
    setActiveSheetIndexState((current) => Math.max(0, Math.min(current, nextSheets.length - 1)));
    setRevision((current) => current + 1);
  }, [activeSheet, readOnly, recordHistoryBeforeMutation, setChartAssets, workbook]);

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
          const nextValue = coerceUserEnteredValue(rawValue);
          worksheet.setCell(cellAddressToA1(nextCell), nextValue);
          mutations.push({
            after: { formula: null, value: nextValue },
            before,
            cell: nextCell
          });
        }
      }
    }

    maybeRecalculateWorkbook(workbook);
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
  }, [activeCell, captureCellMutationState, getActiveWorksheet, maybeRecalculateWorkbook, readOnly, recordRangeEditHistory, refreshWorkbookState, selection, workbook]);

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

    maybeRecalculateWorkbook(workbook);
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
    maybeRecalculateWorkbook,
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
      activeTab,
      activeTabIndex,
      addSheet,
      canRedo,
      canDownload: Boolean(file ?? src),
      canExport: Boolean(workbook),
      canLoadDeferred,
      canUndo,
      canZoomIn,
      canZoomOut,
      charts,
      chartsheets,
      clearSelectedChart,
      clearSelectedCells,
      clearSelectedImage,
      clearSelection,
      continueDeferredLoad,
      copySelectionToClipboard,
      defaultZoomScale,
      deferredLoadFileSize,
      defineNamedRange,
      displayFileName,
      download,
      exportCsv,
      exportXlsx,
      error,
      fillSelection,
      getChartById,
      getChartsheetById,
      getImageById,
      getSheetCharts,
      getSheetImages,
      getSheetShapes,
      file,
      getClipboardData,
      getCellDisplayValue,
      getCellFormula,
      getCellSnapshotAsync: isWorkerBacked ? getCellSnapshotAsync : undefined,
      getActiveWorksheet,
      getRowsBatchAsync: isWorkerBacked ? getRowsBatchAsync : undefined,
      images,
      isLoadDeferred,
      isLoading,
      isChartsLoading,
      isWorkerBacked,
      mergeSelection,
      maxZoomScale: MAX_ZOOM_SCALE,
      minZoomScale: MIN_ZOOM_SCALE,
      moveChartBy,
      moveImageBy,
      pasteFromClipboard,
      pasteStructuredClipboardData,
      pasteText,
      removeActiveSheet,
      readOnly,
      recalculate,
      redo,
      resetZoom,
      revision,
      resizeChartBy,
      resizeImageBy,
      resizeColumn,
      resizeRow,
      setCellFormula,
      setCellValue,
      setZoomScale,
      setChartRect,
      setImageRect,
      selectedChart,
      selectedChartId,
      selectedFormula,
      selectedImage,
      selectedImageId,
      selectedRangeAddress,
      selectedValue,
      selectCell,
      selectChart,
      selectImage,
      selectRange,
      selection,
      setActiveSheetIndex,
      setActiveTabIndex,
      setSelectedCellFormula,
      setSelectedCellValue,
      sheets,
      shapes,
      src,
      sortState,
      sortTable,
      tabs,
      tables,
      undo,
      unmergeSelection,
      updateChart,
      workbook,
      zoomIn,
      zoomOut,
      zoomScale
    }),
    [
      activeCell,
      activeCellAddress,
      activeSheet,
      activeSheetIndex,
      activeTab,
      activeTabIndex,
      addSheet,
      canLoadDeferred,
      canRedo,
      canUndo,
      canZoomIn,
      canZoomOut,
      charts,
      chartsheets,
      clearSelectedChart,
      clearSelectedCells,
      clearSelectedImage,
      continueDeferredLoad,
      copySelectionToClipboard,
      defaultZoomScale,
      deferredLoadFileSize,
      defineNamedRange,
      displayFileName,
      download,
      error,
      exportCsv,
      exportXlsx,
      file,
      fillSelection,
      getChartById,
      getChartsheetById,
      getImageById,
      getSheetCharts,
      getSheetImages,
      getSheetShapes,
      getClipboardData,
      getCellDisplayValue,
      getCellFormula,
      getCellSnapshotAsync,
      getActiveWorksheet,
      historyRevision,
      images,
      isLoadDeferred,
      isLoading,
      isChartsLoading,
      isWorkerBacked,
      mergeSelection,
      resetZoom,
      moveChartBy,
      moveImageBy,
      pasteFromClipboard,
      pasteStructuredClipboardData,
      pasteText,
      removeActiveSheet,
      readOnly,
      recalculate,
      redo,
      revision,
      resizeChartBy,
      resizeImageBy,
      resizeColumn,
      resizeRow,
      setCellFormula,
      setCellValue,
      setZoomScale,
      setChartRect,
      setImageRect,
      selectedChart,
      selectedChartId,
      selectedFormula,
      selectedImage,
      selectedImageId,
      selectedRangeAddress,
      selectedValue,
      selectCell,
      selectChart,
      selectImage,
      selectRange,
      selection,
      setActiveSheetIndex,
      setActiveTabIndex,
      setSelectedCellFormula,
      setSelectedCellValue,
      sheets,
      shapes,
      sortState,
      sortTable,
      src,
      getRowsBatchAsync,
      tabs,
      tables,
      undo,
      unmergeSelection,
      updateChart,
      workbook,
      zoomIn,
      zoomOut,
      zoomScale
    ]
  );
}
