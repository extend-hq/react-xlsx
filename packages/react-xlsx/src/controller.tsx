import { createMergedRangeLookup } from "./merged-range-lookup";
import { restoreWorkbookXmlRecords } from "./workbook-xml";
import { localizeCellNumber } from "./number-locale";
import { resolveMessages } from "./messages";
import { applyInputNumberFormat, coerceUserEnteredValue, getUserEnteredFormula, isTextCell, normalizeFormulaReferences, parseClipboardText, quoteClipboardField } from "./cell-input";
import { loadSpreadsheetFonts } from "./spreadsheet-fonts";
import * as React from "react";
import { XlsxFileSizeLimitExceededError } from "./errors";
export { XlsxFileSizeLimitExceededError } from "./errors";
import type { DrawingAnchor as DukeDrawingAnchor, DrawingColor as DukeDrawingColor, DrawingInput as DukeDrawingInput, DrawingRunFont as DukeDrawingRunFont, DrawingText as DukeDrawingText, FormControlDrawing as DukeFormControlDrawing, FormControlKind as DukeFormControlKind, FormControlKindInput as DukeFormControlKindInput, Workbook } from "./duke-sheets/generated/duke_sheets_wasm.js";
import { applyChartSeriesFormula, buildChartSeriesFormula, refreshChartRenderSource, loadWorkbookChartAssets, updateWorkbookChartAnchor, updateWorkbookChartDefinition, type WorkbookChartAssets } from "./charts";
import { resolveWorkbookColor, resolveWorkbookFillStyle } from "./colors";
import { pxToOfficeColumnWidth, pxToSheetColumnWidth, rectToImageAnchor, refreshWorkbookFormControls, resolveContentSheetAxisPixels, resolveWorksheetDefaultColumnWidthPixels, resolveWorksheetDefaultRowHeightPixels, resolveWorksheetHiddenCols, resolveWorksheetHiddenRows, resolveWorksheetMergeMetadata, resolveSheetColumnWidthPixels, resolveRenderedSheetAxisPixels, resolveSheetRowHeightPixels, resizeImageRect, revokeWorkbookImageAssets, updateWorkbookImageAnchor, xlsxAnchorToDukeDrawingAnchor, type WorkbookImageAssets, type WorkbookImageSheetOrigin } from "./images";
import { tryRecalculate } from "./safe-calculate";
import { parseXlsxRange } from "./range-reference";
import { expandHighlightedRange } from "./range-highlight";
import { getSheetsWasmModule } from "./wasm";
import { XlsxWorkerClient } from "./worker-client";
import type {
  UseXlsxViewerControllerOptions,
  XlsxChart,
  XlsxChartElementSelection,
  XlsxChartsheet,
  XlsxCellAddress,
  XlsxCellRange,
  XlsxMergeMode,
  XlsxCellStyleColorInput,
  XlsxCellStyleInput,
  XlsxClipboardData,
  XlsxConditionalFormatRule,
  XlsxDataValidation,
  XlsxDrawingLayout,
  XlsxFormControl,
  XlsxFormControlCaptionInput,
  XlsxFormControlInput,
  XlsxFormControlPatch,
  XlsxFreezePanes,
  XlsxImage,
  XlsxImageRect,
  XlsxImageResizeHandlePosition,
  XlsxHighlightRangeOptions,
  XlsxRangeHighlight,
  XlsxResolvedCellStyle,
  XlsxShape,
  XlsxSheetData,
  XlsxSheetVisibility,
  XlsxSparkline,
  XlsxThemePalette,
  XlsxTable,
  XlsxTableStyleDefinition,
  XlsxTableSortDirection,
  XlsxTableSortState,
  XlsxViewerController,
  XlsxWorkbookTab
} from "./types";

const DEFAULT_ROW_HEIGHT = 24;
const DEFAULT_COL_WIDTH = 80;
const XLSX_MIME_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
const CSV_MIME_TYPE = "text/csv;charset=utf-8";

const MIN_ROW_HEIGHT_PX = 16;
const GRID_HEADER_HEIGHT = 24;
const GRID_ROW_HEADER_WIDTH = 40;
const HISTORY_LIMIT = 100;
const INTERNAL_CLIPBOARD_MIME = "application/x-react-xlsx-range+json";
const DEFAULT_DEFER_LOADING_ABOVE_BYTES = 0;
const DEFAULT_MAX_FILE_SIZE_BYTES = 25 * 1024 * 1024;

const DEFAULT_ZOOM_SCALE = 100;
const MIN_ZOOM_SCALE = 10;
const MAX_ZOOM_SCALE = 400;
const ZOOM_STEP = 10;
const DEFAULT_ZOOM_TAB_KEY = "__default__";

function normalizeWorksheetVisibility(value: unknown): XlsxSheetVisibility {
  return value === "hidden" || value === "veryHidden" ? value : "visible";
}

type IdleRequestHandle = number;

type IdleDeadline = {
  didTimeout: boolean;
  timeRemaining: () => number;
};

type SnapshotHistoryEntry = {
  kind: "snapshot";
  activeCell: XlsxCellAddress | null;
  activeSheetIndex: number;
  bytes: Uint8Array;
  assets: WorkbookImageAssets | null;
  origins: Array<WorkbookImageSheetOrigin | null>;
  selection: XlsxCellRange | null;
  sortState: XlsxTableSortState | null;
};

type CellMutationState = {
  valueSnapshot?: Uint8Array;
  formula: string | null;
  style: unknown;
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
  sortStateAfter: XlsxTableSortState | null;
  sortStateBefore: XlsxTableSortState | null;
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

function resolveSheetDisplayUsedRange(
  usedRange: [number, number, number, number],
  sheetState?: {
    maxContentCol?: number;
    maxContentRow?: number;
    maxHorizontalMergeEndCol?: number;
    maxVerticalMergeEndRow?: number;
    minContentCol?: number;
    minContentRow?: number;
  } | null
): [number, number, number, number] {
  const [minRow, minCol, maxRow, maxCol] = usedRange;
  const maxContentRow = sheetState?.maxContentRow ?? -1;
  const maxContentCol = sheetState?.maxContentCol ?? -1;
  const maxVerticalMergeEndRow = sheetState?.maxVerticalMergeEndRow ?? -1;
  const maxHorizontalMergeEndCol = sheetState?.maxHorizontalMergeEndCol ?? -1;
  const maxMeaningfulRow = Math.max(maxContentRow, maxVerticalMergeEndRow);
  const maxMeaningfulCol = Math.max(maxContentCol, maxHorizontalMergeEndCol);

  if (maxMeaningfulRow < 0 && maxMeaningfulCol < 0) {
    return usedRange;
  }

  return [
    sheetState?.minContentRow !== undefined && sheetState.minContentRow >= 0 ? Math.min(minRow, sheetState.minContentRow) : minRow,
    sheetState?.minContentCol !== undefined && sheetState.minContentCol >= 0 ? Math.min(minCol, sheetState.minContentCol) : minCol,
    maxMeaningfulRow >= 0
      ? (maxContentRow >= 0 ? Math.min(maxRow, maxMeaningfulRow) : Math.max(maxRow, maxMeaningfulRow))
      : maxRow,
    maxMeaningfulCol >= 0
      ? (maxContentCol >= 0 ? Math.min(maxCol, maxMeaningfulCol) : Math.max(maxCol, maxMeaningfulCol))
      : maxCol
  ];
}

function buildSheetList(
  workbook: Workbook,
  sheetStatesByWorkbookSheetIndex?: Array<{
    cellStyleIds?: Record<string, number>;
    autoFilterRanges?: XlsxCellRange[];
    cachedFormulaValues?: Record<string, string>;
    columnWidthCharacterWidthPx?: number;
    colWidthOverridesPx?: Record<number, number>;
    colStyleIds?: Record<number, number>;
    conditionalFormatRules?: XlsxConditionalFormatRule[];
    officeLayout?: XlsxSheetData["officeLayout"];
    defaultColWidthPx?: number;
    defaultRowHeightPx?: number;
    hasHorizontalMerges?: boolean;
    hasVerticalMerges?: boolean;
    maxHorizontalMergeEndCol?: number;
    maxVerticalMergeEndRow?: number;
    maxContentCol?: number;
    maxContentRow?: number;
    minContentCol?: number;
    minContentRow?: number;
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
  tableStyleByName?: Record<string, XlsxTableStyleDefinition> | null,
  showHiddenSheets = false
): XlsxSheetData[] {
  const sheets: XlsxSheetData[] = [];

  for (let index = 0; index < workbook.sheetCount; index += 1) {
    const worksheet = workbook.getSheet(index);
    const sheetState = sheetStatesByWorkbookSheetIndex?.[index] ?? null;
    const mergeMetadata = resolveWorksheetMergeMetadata(worksheet);
    const effectiveSheetState = {
      ...sheetState,
      ...mergeMetadata
    };
    const defaultColWidthPx = sheetState?.defaultColWidthPx ?? resolveWorksheetDefaultColumnWidthPixels(
      worksheet,
      sheetState?.columnWidthCharacterWidthPx,
      DEFAULT_COL_WIDTH
    );
    const defaultRowHeightPx = sheetState?.defaultRowHeightPx ?? resolveWorksheetDefaultRowHeightPixels(
      worksheet,
      DEFAULT_ROW_HEIGHT
    );
    const visibility = normalizeWorksheetVisibility(worksheet.visibility);
    if (!showHiddenSheets && visibility !== "visible") {
      continue;
    }

    const resolveColumnWidthPx = (col: number) => {
      const width = worksheet.getColumnWidth(col);
      if (width !== undefined && width !== null) {
        return resolveSheetColumnWidthPixels(width, sheetState?.columnWidthCharacterWidthPx);
      }

      return sheetState?.colWidthOverridesPx?.[col] ?? defaultColWidthPx;
    };

    const resolveRowHeightPx = (row: number) => {
      const height = worksheet.getRowHeight(row);
      if (height !== undefined && height !== null) {
        return Math.max(Math.round(height * 1.33), MIN_ROW_HEIGHT_PX);
      }

      return sheetState?.rowHeightOverridesPx?.[row] ?? defaultRowHeightPx;
    };

    const usedRange = worksheet.usedRange() as [number, number, number, number] | null;
    if (!usedRange) {
      sheets.push({
        autoFilterRanges: sheetState?.autoFilterRanges ?? [],
        cachedFormulaValues: sheetState?.cachedFormulaValues ?? {},
        columnWidthCharacterWidthPx: sheetState?.columnWidthCharacterWidthPx,
        colWidthOverridesPx: sheetState?.colWidthOverridesPx ?? {},
        colStyleIds: sheetState?.colStyleIds ?? {},
        conditionalFormatRules: sheetState?.conditionalFormatRules ?? [],
        dataValidations: parseWorksheetDataValidations(worksheet),
        officeLayout: sheetState?.officeLayout,
        defaultColWidthPx,
        defaultRowHeightPx,
        freezePanes: parseWorksheetFreezePanes(worksheet),
        hasHorizontalMerges: mergeMetadata.hasHorizontalMerges,
        hasVerticalMerges: mergeMetadata.hasVerticalMerges,
        maxHorizontalMergeEndCol: mergeMetadata.maxHorizontalMergeEndCol,
        maxVerticalMergeEndRow: mergeMetadata.maxVerticalMergeEndRow,
        mergedRanges: mergeMetadata.mergedRanges,
        hiddenCols: [],
        hiddenRows: [],
        minUsedCol: -1,
        minUsedRow: -1,
        maxUsedCol: -1,
        maxUsedRow: -1,
        name: worksheet.name,
        visibility,
        namedCellStyleByName: namedCellStyleByName ?? {},
        rowCount: 0,
        colCount: 0,
        rowHeightOverridesPx: sheetState?.rowHeightOverridesPx ?? {},
        rowStyleIds: sheetState?.rowStyleIds ?? {},
        cellStyleIds: sheetState?.cellStyleIds ?? {},
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

    const [minRow, minCol, maxRow, maxCol] = resolveSheetDisplayUsedRange(usedRange, effectiveSheetState);
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
      autoFilterRanges: sheetState?.autoFilterRanges ?? [],
      cachedFormulaValues: sheetState?.cachedFormulaValues ?? {},
      columnWidthCharacterWidthPx: sheetState?.columnWidthCharacterWidthPx,
      colWidthOverridesPx: sheetState?.colWidthOverridesPx ?? {},
      colStyleIds: sheetState?.colStyleIds ?? {},
      conditionalFormatRules: sheetState?.conditionalFormatRules ?? [],
      dataValidations: parseWorksheetDataValidations(worksheet),
      officeLayout: sheetState?.officeLayout,
      defaultColWidthPx,
      defaultRowHeightPx,
      freezePanes: parseWorksheetFreezePanes(worksheet),
      hasHorizontalMerges: mergeMetadata.hasHorizontalMerges,
      hasVerticalMerges: mergeMetadata.hasVerticalMerges,
      maxHorizontalMergeEndCol: mergeMetadata.maxHorizontalMergeEndCol,
      maxVerticalMergeEndRow: mergeMetadata.maxVerticalMergeEndRow,
      mergedRanges: mergeMetadata.mergedRanges,
      hiddenCols: resolveWorksheetHiddenCols(worksheet, maxCol),
      hiddenRows: resolveWorksheetHiddenRows(worksheet, maxRow),
      minUsedCol: minCol,
      minUsedRow: minRow,
      maxUsedCol: maxCol,
      maxUsedRow: maxRow,
      name: worksheet.name,
      visibility,
      namedCellStyleByName: namedCellStyleByName ?? {},
      rowHeightOverridesPx: sheetState?.rowHeightOverridesPx ?? {},
      rowStyleIds: sheetState?.rowStyleIds ?? {},
      cellStyleIds: sheetState?.cellStyleIds ?? {},
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

type ResolvedWorkbookReference = {
  range: XlsxCellRange;
  workbookSheetIndex: number;
  worksheet: ReturnType<Workbook["getSheet"]>;
};

function resolveWorkbookReference(
  workbook: Workbook,
  defaultWorkbookSheetIndex: number,
  rawReference: string,
  resolvingNamedRange = false
): ResolvedWorkbookReference | null {
  const reference = rawReference.trim().replace(/^=/, "");
  if (!reference) {
    return null;
  }

  let workbookSheetIndex = defaultWorkbookSheetIndex;
  let rangeReference = reference;
  const bangIndex = reference.lastIndexOf("!");
  if (bangIndex >= 0) {
    let sheetName = reference.slice(0, bangIndex).trim();
    rangeReference = reference.slice(bangIndex + 1).trim();
    if (sheetName.startsWith("'") && sheetName.endsWith("'")) {
      sheetName = sheetName.slice(1, -1).replace(/''/g, "'");
    }
    const resolvedSheetIndex = workbook.sheetIndex(sheetName);
    if (resolvedSheetIndex === undefined) {
      return null;
    }
    workbookSheetIndex = resolvedSheetIndex;
  } else if (!resolvingNamedRange) {
    const namedRange = workbook.getNamedRange(reference);
    if (namedRange) {
      return resolveWorkbookReference(workbook, defaultWorkbookSheetIndex, namedRange, true);
    }
  }

  const range = parseA1RangeReference(rangeReference.replace(/\$/g, ""));
  if (!range || workbookSheetIndex < 0 || workbookSheetIndex >= workbook.sheetCount) {
    return null;
  }

  try {
    return {
      range,
      workbookSheetIndex,
      worksheet: workbook.getSheet(workbookSheetIndex)
    };
  } catch {
    return null;
  }
}

function parseColorHex(hex: string | undefined) {
  const normalized = hex?.replace(/^#/, "");
  return normalized && /^[0-9a-f]{6}$/i.test(normalized)
    ? [0, 2, 4].map((offset) => Number.parseInt(normalized.slice(offset, offset + 2), 16))
    : null;
}

function xlsxColorToDukeColor(color: XlsxCellStyleColorInput | undefined): DukeDrawingColor | undefined {
  if (!color) {
    return undefined;
  }
  const normalizedHex = color.hex?.replace(/^#/, "");
  const colorType = color.colorType
    ?? (normalizedHex?.length === 8 ? "argb" : normalizedHex?.length === 6 ? "rgb" : undefined)
    ?? (color.r !== undefined && color.g !== undefined && color.b !== undefined
      ? color.a !== undefined ? "argb" : "rgb"
      : undefined)
    ?? (color.themeIndex !== undefined ? "theme" : color.paletteIndex !== undefined ? "indexed" : undefined);
  switch (colorType) {
    case "auto":
      return { colorType: "auto" };
    case "rgb": {
      const channels = color.r !== undefined && color.g !== undefined && color.b !== undefined
        ? [color.r, color.g, color.b]
        : parseColorHex(color.hex);
      return channels ? { b: channels[2]!, colorType: "rgb", g: channels[1]!, r: channels[0]! } : undefined;
    }
    case "argb": {
      const normalized = color.hex?.replace(/^#/, "");
      const channels = normalized && /^[0-9a-f]{8}$/i.test(normalized)
        ? [0, 2, 4, 6].map((offset) => Number.parseInt(normalized.slice(offset, offset + 2), 16))
        : color.a !== undefined && color.r !== undefined && color.g !== undefined && color.b !== undefined
          ? [color.a, color.r, color.g, color.b]
          : null;
      return channels
        ? { a: channels[0]!, b: channels[3]!, colorType: "argb", g: channels[2]!, r: channels[1]! }
        : undefined;
    }
    case "theme":
      return color.themeIndex === undefined
        ? undefined
        : { colorType: "theme", index: color.themeIndex, tint: color.tint ?? 0 };
    case "indexed":
      return color.paletteIndex === undefined
        ? undefined
        : { colorType: "indexed", index: color.paletteIndex };
  }
}

function xlsxFontToDukeFont(font: NonNullable<Exclude<XlsxFormControlCaptionInput, string>["runs"][number]["font"]>): DukeDrawingRunFont {
  return {
    bold: font.bold,
    charset: font.charset,
    color: xlsxColorToDukeColor(font.color),
    family: font.family,
    italic: font.italic,
    name: font.name,
    scheme: font.scheme,
    size: font.size,
    strikethrough: font.strikethrough,
    underline: font.underline,
    verticalAlign: font.verticalAlign
  };
}

function xlsxCaptionToDukeText(caption: XlsxFormControlCaptionInput): DukeDrawingText {
  if (typeof caption === "string") {
    return { runs: [{ text: caption }] };
  }
  return {
    horizontalAlignment: caption.horizontalAlignment,
    runs: caption.runs.map((run) => ({
      font: run.font ? xlsxFontToDukeFont(run.font) : undefined,
      text: run.text
    })),
    verticalAlignment: caption.verticalAlignment
  };
}

function xlsxFormControlKindToDukeKind(kind: XlsxFormControlInput["kind"]): DukeFormControlKindInput {
  switch (kind.kind) {
    case "button":
    case "label":
      return { caption: xlsxCaptionToDukeText(kind.caption), kind: kind.kind };
    case "checkbox":
      return { ...kind, caption: xlsxCaptionToDukeText(kind.caption) };
    case "optionButton":
      return { ...kind, caption: xlsxCaptionToDukeText(kind.caption) };
    case "groupBox":
      return { ...kind, caption: xlsxCaptionToDukeText(kind.caption) };
    case "editbox":
      return {
        caption: xlsxCaptionToDukeText(kind.caption),
        kind: "unknown",
        legacyObjectType: kind.legacyObjectType,
        objectType: "EditBox"
      };
    case "unknown":
      return {
        caption: kind.caption ? xlsxCaptionToDukeText(kind.caption) : undefined,
        kind: "unknown",
        legacyObjectType: kind.legacyObjectType,
        objectType: kind.objectType
      };
    default:
      return { ...kind };
  }
}

function dukeFormControlKindToInput(kind: DukeFormControlKind): DukeFormControlKindInput {
  if (kind.kind === "optionButton") {
    return { ...kind, state: kind.state === "checked" ? "checked" : "unchecked" };
  }
  if (kind.kind === "unknown") {
    return {
      caption: kind.caption,
      kind: "unknown",
      legacyObjectType: kind.legacyObjectType,
      objectType: kind.objectType
    };
  }
  return { ...kind };
}

function formControlInputFromDukeControl(control: DukeFormControlDrawing): DukeDrawingInput {
  return {
    altText: control.altText,
    ...(control.anchor ? { anchor: control.anchor } : { transform: control.transform! }),
    formControl: {
      kind: dukeFormControlKindToInput(control.formControl.kind),
      macroName: control.formControl.macroName,
      rawClientData: control.formControl.rawClientData,
      rawObj: control.formControl.rawObj,
      rawProperties: control.formControl.rawProperties
    },
    hidden: control.hidden,
    kind: "formControl",
    locked: control.locked,
    name: control.name,
    printable: control.printable,
    title: control.title
  };
}

type DukeFormControlDrawingInput = Extract<DukeDrawingInput, { kind: "formControl" }> & {
  anchor: DukeDrawingAnchor;
};

function xlsxFormControlInputToDukeDrawing(input: XlsxFormControlInput): DukeFormControlDrawingInput {
  return {
    altText: input.altText,
    anchor: xlsxAnchorToDukeDrawingAnchor(input.anchor, input.editAs),
    formControl: {
      kind: xlsxFormControlKindToDukeKind(input.kind),
      macroName: input.macroName,
      rawClientData: input.rawClientData,
      rawObj: input.kind.kind === "editbox" || input.kind.kind === "unknown" ? input.kind.rawObj : undefined,
      rawProperties: input.kind.kind === "editbox" || input.kind.kind === "unknown" ? input.kind.rawProperties : undefined
    },
    hidden: input.hidden,
    kind: "formControl",
    locked: input.locked,
    name: input.name,
    printable: input.printable,
    title: input.title
  };
}

function reconcileCheckedOptionButtons(worksheet: ReturnType<Workbook["getSheet"]>) {
  const checkedPaths = worksheet.formControls.flatMap((control) => (
    control.formControl.kind.kind === "optionButton" && control.formControl.kind.state === "checked"
      ? [control.drawingPath]
      : []
  ));
  for (const path of checkedPaths) {
    const current = worksheet.formControls.find((control) => (
      control.drawingPath.length === path.length
      && control.drawingPath.every((segment, index) => segment === path[index])
    ));
    if (current?.formControl.kind.kind === "optionButton" && current.formControl.kind.state === "checked") {
      worksheet.setFormControlCheckState(current.drawingPath, "checked");
    }
  }
}

function formControlInteractionChanged(
  current: DukeFormControlKind,
  next: DukeFormControlKind
) {
  if (current.kind !== next.kind) {
    return true;
  }
  if ("cellLink" in current && "cellLink" in next && current.cellLink !== next.cellLink) {
    return true;
  }
  switch (current.kind) {
    case "checkbox":
    case "optionButton":
      return next.kind === current.kind && current.state !== next.state;
    case "dropdown":
      return next.kind === "dropdown" && current.selected !== next.selected;
    case "listBox":
      return next.kind === "listBox" && (
        current.selection !== next.selection
        || current.selected.length !== next.selected.length
        || current.selected.some((value, index) => value !== next.selected[index])
      );
    case "scrollbar":
    case "spinner":
      return next.kind === current.kind && current.value !== next.value;
    default:
      return false;
  }
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

function rangesIntersect(a: XlsxCellRange, b: XlsxCellRange): boolean {
  return a.start.row <= b.end.row && a.end.row >= b.start.row
    && a.start.col <= b.end.col && a.end.col >= b.start.col;
}

function mapWorksheetTables(
  worksheet: ReturnType<Workbook["getSheet"]> | null,
  autoFilterRanges: XlsxCellRange[] = []
): XlsxTable[] {
  const rawTables = (worksheet?.tables ?? []) as Array<Record<string, unknown>>;
  const mappedTables = rawTables.flatMap((table, index) => {
    const rawColumns = Array.isArray(table.columns) ? table.columns : [];
    const rawName = typeof table.name === "string" ? table.name : `Table${index + 1}`;
    const rawDisplayName =
      typeof table.displayName === "string"
        ? table.displayName
        : typeof table.name === "string"
          ? table.name
          : `Table ${index + 1}`;
    const rawReference = typeof table.reference === "string" ? table.reference : "";
    const reference = rawReference;
    const parsedRange = parseA1RangeReference(reference);
    if (!parsedRange) {
      return [];
    }

    return [{
      columns: rawColumns.map((column, columnIndex) => ({
        id: typeof (column as { id?: unknown }).id === "number" ? ((column as { id?: number }).id ?? columnIndex + 1) : columnIndex + 1,
        index: columnIndex,
        name: typeof (column as { name?: unknown }).name === "string" ? ((column as { name?: string }).name ?? `Column ${columnIndex + 1}`) : `Column ${columnIndex + 1}`
      })),
      displayName: rawDisplayName,
      end: parsedRange.end,
      headerRowCount: resolveWorkbookTableCount(table.headerRowCount, 1),
      headerRowCellStyle: typeof table.headerRowCellStyle === "string" ? table.headerRowCellStyle : undefined,
      name: rawName,
      reference,
      start: parsedRange.start,
      styleInfo: table.styleInfo as XlsxTable["styleInfo"] | undefined,
      totalsRowCount: resolveWorkbookTableCount(table.totalsRowCount, 0),
      totalsRowShown: resolveWorkbookTableBoolean(table.totalsRowShown)
    }];
  });

  const existingReferences = new Set(mappedTables.map((table) => table.reference));
  const mappedAutoFilterTables = autoFilterRanges.flatMap((range, index) => {
    const reference = rangeToA1(range);
    if (existingReferences.has(reference)) {
      return [];
    }

    const columnCount = Math.max(0, range.end.col - range.start.col + 1);
    const columns = Array.from({ length: columnCount }, (_, columnIndex) => {
      const headerValue = worksheet
        ? decodeHtmlEntities(worksheet.getFormattedValueAt(range.start.row, range.start.col + columnIndex) ?? "")
        : "";
      return {
        id: columnIndex + 1,
        index: columnIndex,
        name: headerValue.trim().length > 0 ? headerValue : `Column ${columnIndex + 1}`
      };
    });

    return [{
      columns,
      displayName: `AutoFilter ${index + 1}`,
      end: range.end,
      headerRowCount: 1,
      name: `__autofilter_${index + 1}_${reference}`,
      reference,
      start: range.start,
      totalsRowCount: 0,
      totalsRowShown: false
    } satisfies XlsxTable];
  });

  return [...mappedTables, ...mappedAutoFilterTables];
}

function resolveWorkbookTableCount(value: unknown, fallback: number) {
  if (typeof value === "number" && Number.isFinite(value) && value >= 0) {
    return value;
  }

  if (typeof value === "string") {
    const parsed = Number.parseInt(value, 10);
    if (Number.isFinite(parsed) && parsed >= 0) {
      return parsed;
    }
  }

  return fallback;
}

function resolveWorkbookTableBoolean(value: unknown) {
  if (typeof value === "boolean") {
    return value;
  }

  if (typeof value === "number") {
    return value !== 0;
  }

  if (typeof value === "string") {
    const normalized = value.trim().toLowerCase();
    if (normalized === "0" || normalized === "false" || normalized === "") {
      return false;
    }
    if (normalized === "1" || normalized === "true") {
      return true;
    }
  }

  return false;
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

function pushHistoryEntry(stack: HistoryEntry[], entry: HistoryEntry) {
  stack.push(entry);
  if (stack.length > HISTORY_LIMIT) {
    stack.shift();
  }
}

function normalizeCellValue(value: unknown) {
  return value ?? "";
}

function cloneCellStyle(style: unknown): unknown {
  if (!style || typeof style !== "object") {
    return style;
  }

  if (typeof structuredClone === "function") {
    try {
      return structuredClone(style);
    } catch {
      // Fall through to the JSON clone below.
    }
  }

  try {
    return JSON.parse(JSON.stringify(style));
  } catch {
    return style;
  }
}

function applyCellMutationState(
  worksheet: ReturnType<Workbook["getSheet"]>,
  cell: XlsxCellAddress,
  state: CellMutationState
) {
  if (state.formula) {
    worksheet.setFormula(cellAddressToA1(cell), state.formula);
  } else if (state.valueSnapshot) {
    worksheet.restoreValueSnapshotAt(cell.row, cell.col, state.valueSnapshot);
  } else {
    worksheet.setCell(cellAddressToA1(cell), normalizeCellValue(state.value));
  }

  if (state.style && typeof state.style === "object") {
    worksheet.setCellStyleAt(cell.row, cell.col, state.style);
  }
}

function escapeHtml(value: string) {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function createAbortError() {
  if (typeof DOMException !== "undefined") {
    return new DOMException("Aborted", "AbortError");
  }

  const error = new Error("Aborted");
  error.name = "AbortError";
  return error;
}

function isAbortError(error: unknown) {
  return error instanceof Error && error.name === "AbortError";
}

async function resolveWorkbookBuffer(
  { file, src }: UseXlsxViewerControllerOptions,
  signal?: AbortSignal
): Promise<ArrayBuffer> {
  let buffer: ArrayBuffer;

  if (signal?.aborted) {
    throw createAbortError();
  }

  if (file) {
    buffer = file;
  } else if (src) {
    let response: Response;
    try {
      response = await fetch(src, { signal });
    } catch (error) {
      if (isAbortError(error)) {
        throw error;
      }

      throw new Error(
        "Failed to fetch workbook. The remote URL may be blocked by CORS, unavailable, or not directly downloadable from the browser."
      );
    }
    if (!response.ok) {
      throw new Error(`Failed to fetch workbook (status ${response.status})`);
    }
    buffer = await response.arrayBuffer();
  } else {
    throw new Error("Either `file` or `src` must be provided.");
  }

  return buffer;
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
    allowResizeInReadOnly = false,
    deferLoadingAboveBytes = DEFAULT_DEFER_LOADING_ABOVE_BYTES,
    file,
    fileName,
    locale = "en-US",
    messages: messageOverrides,
    maxFileSizeBytes = DEFAULT_MAX_FILE_SIZE_BYTES,
    readOnly: requestedReadOnly = false,
    readOnlyAboveBytes = 0,
    showHiddenSheets = false,
    skipXmlParsing = false,
    src,
    useWorker = true
  } = options;
  const messages = React.useMemo(() => resolveMessages(locale, messageOverrides), [locale, messageOverrides]);
  const [isLoading, setIsLoading] = React.useState(Boolean(file ?? src));
  const [error, setError] = React.useState<Error | null>(null);
  const [workbook, setWorkbook] = React.useState<Workbook | null>(null);
  const [sheets, setSheets] = React.useState<XlsxSheetData[]>([]);
  const [chartsByWorkbookSheetIndex, setChartsByWorkbookSheetIndex] = React.useState<XlsxChart[][]>([]);
  const [chartsheets, setChartsheets] = React.useState<XlsxChartsheet[]>([]);
  const [tabs, setTabs] = React.useState<XlsxWorkbookTab[]>([]);
  const [isChartsLoading, setIsChartsLoading] = React.useState(false);
  const [workerTablesByWorkbookSheetIndex, setWorkerTablesByWorkbookSheetIndex] = React.useState<XlsxTable[][]>([]);
  const [formControlsByWorkbookSheetIndex, setFormControlsByWorkbookSheetIndex] = React.useState<XlsxFormControl[][]>([]);
  const [imagesByWorkbookSheetIndex, setImagesByWorkbookSheetIndex] = React.useState<XlsxImage[][]>([]);
  const [shapesByWorkbookSheetIndex, setShapesByWorkbookSheetIndex] = React.useState<XlsxShape[][]>([]);
  const [activeSheetIndex, setActiveSheetIndexState] = React.useState(0);
  const [activeTabIndex, setActiveTabIndexState] = React.useState(0);
  const [zoomScaleOverridesByTabId, setZoomScaleOverridesByTabId] = React.useState<Record<string, number>>({});
  const [activeCell, setActiveCell] = React.useState<XlsxCellAddress | null>(null);
  const [selection, setSelection] = React.useState<XlsxCellRange | null>(null);
  const [highlightedRanges, setHighlightedRanges] = React.useState<XlsxRangeHighlight[]>([]);
  const [selectedChartId, setSelectedChartId] = React.useState<string | null>(null);
  const [selectedChartElement, setSelectedChartElement] = React.useState<XlsxChartElementSelection | null>(null);
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
  const releaseXmlRecordsRef = React.useRef<(() => void) | null>(null);
  const loadGenerationRef = React.useRef(0);
  const deferredBufferRef = React.useRef<ArrayBuffer | null>(null);
  const [deferredLoadFileSize, setDeferredLoadFileSize] = React.useState<number | null>(null);
  const retainedObjectUrlsRef = React.useRef(new Set<string>());
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
  const canResizeReadOnly = requestedReadOnly && allowResizeInReadOnly && !forcedReadOnly;
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
    releaseXmlRecordsRef.current?.();
    releaseXmlRecordsRef.current = null;
    for (const url of retainedObjectUrlsRef.current) URL.revokeObjectURL(url);
    retainedObjectUrlsRef.current.clear();
    revokeWorkbookImageAssets(imageAssetsRef.current);
    imageAssetsRef.current = null;
    sheetOriginsRef.current = [];
    setFormControlsByWorkbookSheetIndex([]);
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
    for (const url of imageAssetsRef.current?.objectUrls ?? []) retainedObjectUrlsRef.current.add(url);
    imageAssetsRef.current = assets;
    sheetOriginsRef.current = assets?.sheetOrigins.slice() ?? [];
    setFormControlsByWorkbookSheetIndex(assets?.formControlsByWorkbookSheetIndex ?? []);
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

  const ensureChartAssetsHydrated = React.useCallback((
    targetWorkbook: Workbook | null,
    targetSheets: XlsxSheetData[]
  ) => {
    const currentAssets = chartAssetsRef.current;
    if (
      currentAssets
      && (
        currentAssets.chartOriginsById.size > 0
        || !targetWorkbook
        || !imageAssetsRef.current
      )
    ) {
      return currentAssets;
    }

    if (!targetWorkbook || !imageAssetsRef.current) {
      return chartAssetsRef.current;
    }

    const assets = loadWorkbookChartAssets(
      targetWorkbook,
      imageAssetsRef.current,
      buildVisibleSheetIndexMap(targetSheets),
      showHiddenSheets
    );
    chartAssetsRef.current = assets;
    return assets;
  }, [showHiddenSheets]);

  const loadPreparedWorkbook = React.useCallback(async (buffer: ArrayBuffer, isCurrent: () => boolean = () => true) => {
    const defaultFonts = loadSpreadsheetFonts(["Calibri", "Arial"]);
    const forced = shouldForceReadOnlyForBuffer(buffer.byteLength);
    const editable = !requestedReadOnly && !forced;
    const snapshot = await getWorkerClient().loadWorkbook(buffer, skipXmlParsing, showHiddenSheets, editable || !useWorker);
    if (!isCurrent()) return;
    await Promise.all([defaultFonts, loadSpreadsheetFonts(snapshot.fontFamilies)]);
    if (!isCurrent()) return;
    const nextWorkbook = snapshot.workbookSnapshot
      ? (await getSheetsWasmModule()).Workbook.fromSnapshot(snapshot.workbookSnapshot)
      : null;
    if (!isCurrent()) { nextWorkbook?.free(); return; }
    releaseXmlRecordsRef.current?.();
    releaseXmlRecordsRef.current = restoreWorkbookXmlRecords(snapshot.xmlRecords);
    setForcedReadOnly(forced);
    setImageAssets(snapshot.imageAssets);
    setWorkbook(nextWorkbook);
    setSheets(snapshot.sheets);
    setChartAssets({ chartsByWorkbookSheetIndex: snapshot.chartsByWorkbookSheetIndex, chartsheets: snapshot.chartsheets, tabs: snapshot.tabs, chartOriginsById: snapshot.chartOriginsById });
    setFormControlsByWorkbookSheetIndex(snapshot.formControlsByWorkbookSheetIndex);
    setWorkerTablesByWorkbookSheetIndex(snapshot.tablesByWorkbookSheetIndex);
    setShouldAutoCalculate(editable && snapshot.shouldAutoCalculate);
    setIsWorkerBacked(!nextWorkbook);
    setSortState(null);
    setIsLoading(false);
  }, [getWorkerClient, requestedReadOnly, setChartAssets, setImageAssets, shouldForceReadOnlyForBuffer, showHiddenSheets, skipXmlParsing, useWorker]);

  const refreshWorkbookState = React.useCallback((targetWorkbook: Workbook) => {
    const currentFormControls = imageAssetsRef.current?.formControlsByWorkbookSheetIndex ?? [];
    const nextFormControls = refreshWorkbookFormControls(
      targetWorkbook,
      currentFormControls,
      imageAssetsRef.current?.themePalette
    );
    if (imageAssetsRef.current) {
      imageAssetsRef.current.formControlsByWorkbookSheetIndex = nextFormControls;
    }
    setFormControlsByWorkbookSheetIndex(nextFormControls);
    const nextSheets = buildSheetList(
      targetWorkbook,
      imageAssetsRef.current?.sheetStatesByWorkbookSheetIndex,
      imageAssetsRef.current?.themePalette,
      imageAssetsRef.current?.styleById,
      imageAssetsRef.current?.namedCellStyleByName,
      imageAssetsRef.current?.tableStyleByName,
      showHiddenSheets
    );
    setSheets(nextSheets);
    setChartAssets(
      loadWorkbookChartAssets(
        targetWorkbook,
        imageAssetsRef.current,
        buildVisibleSheetIndexMap(nextSheets),
        showHiddenSheets
      )
    );
    setRevision((current) => current + 1);
  }, [setChartAssets, showHiddenSheets]);

  React.useEffect(() => () => {
    releaseXmlRecordsRef.current?.();
    releaseXmlRecordsRef.current = null;
    chartDisplayFallbackCleanupRef.current?.();
    chartDisplayFallbackCleanupRef.current = null;
    for (const url of retainedObjectUrlsRef.current) URL.revokeObjectURL(url);
    retainedObjectUrlsRef.current.clear();
    revokeWorkbookImageAssets(imageAssetsRef.current);
    disposeWorkerClient();
  }, [disposeWorkerClient]);

  React.useEffect(() => {
    loadGenerationRef.current += 1;
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
      setHighlightedRanges([]);
      setSelectedChartId(null);
      setSelectedChartElement(null);
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
    const abortController = new AbortController();
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
    setHighlightedRanges([]);
    setSelectedChartId(null);
    setSelectedChartElement(null);
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
    disposeWorkerClient();

    void resolveWorkbookBuffer({ file, src }, abortController.signal)
      .then(async (buffer) => {
        if (!isCurrent || abortController.signal.aborted) {
          return;
        }

        if (maxFileSizeBytes > 0 && buffer.byteLength > maxFileSizeBytes) {
          throw new XlsxFileSizeLimitExceededError(buffer.byteLength, maxFileSizeBytes);
        }

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
        await loadPreparedWorkbook(buffer, () => isCurrent && !abortController.signal.aborted);
      })
      .catch((nextError: unknown) => {
        if (!isCurrent || isAbortError(nextError)) {
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
      loadGenerationRef.current += 1;
      abortController.abort();
      disposeWorkerClient();
    };
  }, [
    clearChartAssets,
    clearImageAssets,
    deferLoadingAboveBytes,
    disposeWorkerClient,
    file,
    getWorkerClient,
    loadPreparedWorkbook,
    maxFileSizeBytes,
    setImageAssets,
    shouldDeferLoading,
    shouldForceReadOnlyForBuffer,
    src,
    showHiddenSheets
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
    setSelectedChartElement(null);
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

  const highlightRanges = React.useCallback((references: string | readonly string[], options: XlsxHighlightRangeOptions = {}) => {
    const inputs = typeof references === "string" ? [references] : references;
    if (!Array.isArray(inputs) || isLoading || error) {
      return false;
    }
    const nextHighlights: XlsxRangeHighlight[] = [];
    let firstSheetIndex = activeSheetIndex;
    for (const reference of inputs) {
      const parsed = parseXlsxRange(reference);
      if (!parsed) {
        return false;
      }
      const requestedSheetName = parsed.sheetName?.toLowerCase();
      const sheetIndex = requestedSheetName === undefined
        ? options.sheetIndex ?? activeSheetIndex
        : sheets.findIndex((sheet) => sheet.name.toLowerCase() === requestedSheetName);
      const sheet = Number.isInteger(sheetIndex) ? sheets[sheetIndex] : undefined;
      if (!sheet) {
        return false;
      }
      if (nextHighlights.length === 0) {
        firstSheetIndex = sheetIndex;
      }
      nextHighlights.push({
        ...parsed,
        autoScroll: nextHighlights.length === 0 && (options.autoScroll ?? true),
        sheetName: sheet.name,
        workbookSheetIndex: sheet.workbookSheetIndex
      });
    }
    setHighlightedRanges(nextHighlights);
    if (nextHighlights[0]?.autoScroll) {
      setActiveSheetIndex(firstSheetIndex);
    }
    return true;
  }, [activeSheetIndex, error, isLoading, setActiveSheetIndex, sheets]);

  const clearHighlightedRanges = React.useCallback(() => setHighlightedRanges([]), []);

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
    const buffer = deferredBufferRef.current;
    if (!buffer) return;
    if (maxFileSizeBytes > 0 && buffer.byteLength > maxFileSizeBytes) {
      setError(new XlsxFileSizeLimitExceededError(buffer.byteLength, maxFileSizeBytes));
      return;
    }
    const generation = loadGenerationRef.current;
    const isCurrent = () => generation === loadGenerationRef.current;
    deferredBufferRef.current = null;
    setDeferredLoadFileSize(null);
    setIsLoading(true);
    setError(null);
    void loadPreparedWorkbook(buffer, isCurrent).catch((nextError: unknown) => {
      if (!isCurrent()) return;
      setWorkbook(null);
      setSheets([]);
      clearChartAssets();
      clearImageAssets();
      setWorkerTablesByWorkbookSheetIndex([]);
      setShouldAutoCalculate(false);
      setIsWorkerBacked(false);
      setError(nextError instanceof Error ? nextError : new Error("Could not load workbook."));
      setIsLoading(false);
    });
  }, [clearChartAssets, clearImageAssets, loadPreparedWorkbook, maxFileSizeBytes]);

  const maybeRecalculateWorkbook = React.useCallback((targetWorkbook: Workbook) => {
    if (!shouldAutoCalculate) {
      return;
    }

    const result = tryRecalculate(targetWorkbook);
    if (!result.calculated) {
      setShouldAutoCalculate(false);
    }
  }, [shouldAutoCalculate]);

  const getActiveWorksheet = React.useCallback(() => {
    if (!workbook || !activeSheet) {
      return null;
    }

    return workbook.getSheet(activeSheet.workbookSheetIndex);
  }, [activeSheet, workbook]);

  const getFormControlWorksheet = React.useCallback((sheetIndex = activeSheetIndex) => {
    const targetSheet = sheets[sheetIndex];
    if (!workbook || !targetSheet) {
      return null;
    }
    try {
      return {
        workbookSheetIndex: targetSheet.workbookSheetIndex,
        worksheet: workbook.getSheet(targetSheet.workbookSheetIndex)
      };
    } catch {
      return null;
    }
  }, [activeSheetIndex, sheets, workbook]);

  const tables = React.useMemo(
    () => (
      isWorkerBacked
        ? workerTablesByWorkbookSheetIndex[activeSheet?.workbookSheetIndex ?? -1] ?? []
        : mapWorksheetTables(getActiveWorksheet(), activeSheet?.autoFilterRanges ?? [])
    ),
    [activeSheet?.autoFilterRanges, activeSheet?.workbookSheetIndex, getActiveWorksheet, isWorkerBacked, revision, workerTablesByWorkbookSheetIndex]
  );

  const getCellSnapshotAsync = React.useCallback((workbookSheetIndex: number, row: number, col: number) => {
    if (!isWorkerBacked) {
      return Promise.resolve({
        displayValue: "",
        formula: ""
      });
    }

    return getWorkerClient().getCellSnapshot(workbookSheetIndex, row, col, locale);
  }, [getWorkerClient, isWorkerBacked, locale]);

  const getRowsBatchAsync = React.useCallback((workbookSheetIndex: number, startRow: number, rowCount: number) => {
    if (!isWorkerBacked) {
      return Promise.resolve(null);
    }

    return getWorkerClient().getRowsBatch(workbookSheetIndex, startRow, rowCount, locale);
  }, [getWorkerClient, isWorkerBacked, locale]);

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

  const mapPublicFormControl = React.useCallback((control: XlsxFormControl) => {
    const visibleSheetIndex = visibleSheetIndexByWorkbookSheetIndex.get(control.workbookSheetIndex);
    return {
      ...control,
      sheetIndex: visibleSheetIndex ?? control.workbookSheetIndex
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

  React.useEffect(() => {
    if (!selectedChartId) {
      if (selectedChartElement) {
        setSelectedChartElement(null);
      }
      return;
    }

    if (!selectedChart) {
      setSelectedChartId(null);
      setSelectedChartElement(null);
      return;
    }

    if (!selectedChartElement) {
      setSelectedChartElement({ chartId: selectedChartId, kind: "chart" });
      return;
    }

    if (selectedChartElement.chartId !== selectedChartId) {
      setSelectedChartElement({ chartId: selectedChartId, kind: "chart" });
      return;
    }

    if (selectedChartElement.kind !== "chart" && selectedChartElement.kind !== "legend") {
      const selectedSeries = selectedChart.series[selectedChartElement.seriesIndex];
      if (!selectedSeries || selectedSeries.id !== selectedChartElement.seriesId) {
        setSelectedChartElement({ chartId: selectedChartId, kind: "chart" });
      }
    }
  }, [selectedChart, selectedChartElement, selectedChartId]);

  const selectChart = React.useCallback((id: string | null) => {
    setSelectedImageId(null);
    setSelectedChartId(id);
    setSelectedChartElement(id ? { chartId: id, kind: "chart" } : null);
  }, []);

  const clearSelectedChart = React.useCallback(() => {
    setSelectedChartId(null);
    setSelectedChartElement(null);
  }, []);

  const clearSelectedChartElement = React.useCallback(() => {
    setSelectedChartElement(selectedChartId ? { chartId: selectedChartId, kind: "chart" } : null);
  }, [selectedChartId]);

  const selectChartElement = React.useCallback((selection: XlsxChartElementSelection | null) => {
    setSelectedImageId(null);
    setSelectedChartId(selection?.chartId ?? null);
    setSelectedChartElement(selection);
  }, []);

  const getSheetImages = React.useCallback((sheetIndex = activeSheetIndex) => {
    const targetSheet = sheets[sheetIndex];
    if (!targetSheet) {
      return [];
    }

    return (imagesByWorkbookSheetIndex[targetSheet.workbookSheetIndex] ?? []).map(mapPublicImage);
  }, [activeSheetIndex, imagesByWorkbookSheetIndex, mapPublicImage, sheets]);

  const images = React.useMemo(() => getSheetImages(activeSheetIndex), [activeSheetIndex, getSheetImages]);

  const getSheetFormControls = React.useCallback((sheetIndex = activeSheetIndex) => {
    const targetSheet = sheets[sheetIndex];
    if (!targetSheet) {
      return [];
    }

    return (formControlsByWorkbookSheetIndex[targetSheet.workbookSheetIndex] ?? []).map(mapPublicFormControl);
  }, [activeSheetIndex, formControlsByWorkbookSheetIndex, mapPublicFormControl, sheets]);

  const formControls = React.useMemo(() => getSheetFormControls(activeSheetIndex), [activeSheetIndex, getSheetFormControls]);

  const getFormControlItems = React.useCallback((controlIndex: number, sheetIndex = activeSheetIndex) => {
    const target = getFormControlWorksheet(sheetIndex);
    if (!workbook || !target) {
      const targetSheet = sheets[sheetIndex];
      if (!targetSheet) {
        return [];
      }
      const control = (formControlsByWorkbookSheetIndex[targetSheet.workbookSheetIndex] ?? [])
        .find((entry) => entry.controlIndex === controlIndex);
      return control?.items?.slice() ?? [];
    }
    const control = target.worksheet.formControls[controlIndex];
    const kind = control?.formControl.kind;
    const inputRange = kind?.kind === "listBox" || kind?.kind === "dropdown"
      ? kind.inputRange
      : undefined;
    if (!inputRange) {
      return [];
    }
    const source = resolveWorkbookReference(workbook, target.workbookSheetIndex, inputRange);
    if (!source) {
      return [];
    }

    const items: string[] = [];
    for (let row = source.range.start.row; row <= source.range.end.row; row += 1) {
      for (let col = source.range.start.col; col <= source.range.end.col; col += 1) {
        items.push(source.worksheet.getFormattedValueAt(row, col));
      }
    }
    return items;
  }, [activeSheetIndex, formControlsByWorkbookSheetIndex, getFormControlWorksheet, sheets, workbook]);

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
    setSelectedChartElement(null);
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
      const workerSnapshot = workerCellSnapshotCacheRef.current.get(`${locale}:${activeSheet.workbookSheetIndex}:${cell.row}:${cell.col}`);
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
      return localizeCellNumber(worksheet, cell.row, cell.col, decodeHtmlEntities(formatted), locale);
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
  }, [activeSheet, getActiveWorksheet, locale]);

  const getCellFormula = React.useCallback((cell?: XlsxCellAddress | null) => {
    if (cell && activeSheet) {
      const workerSnapshot = workerCellSnapshotCacheRef.current.get(`${locale}:${activeSheet.workbookSheetIndex}:${cell.row}:${cell.col}`);
      if (workerSnapshot) {
        return workerSnapshot.formula;
      }
    }

    const worksheet = getActiveWorksheet();
    if (!worksheet || !cell) {
      return "";
    }

    return worksheet.getFormulaAt(cell.row, cell.col) ?? "";
  }, [activeSheet, getActiveWorksheet, locale]);

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

        textCells.push(quoteClipboardField(value));
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

    const cacheKey = `${locale}:${deferredMetadataSheet.workbookSheetIndex}:${deferredMetadataCell.row}:${deferredMetadataCell.col}`;
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
  }, [deferredMetadataCell, deferredMetadataSheet, getCellSnapshotAsync, isWorkerBacked, locale]);

  const activeCellAddress = React.useMemo(() => (activeCell ? cellAddressToA1(activeCell) : null), [activeCell]);
  const selectedRangeAddress = React.useMemo(() => (selection ? rangeToA1(selection) : null), [selection]);
  const selectedValue = React.useMemo(
    () => getCellDisplayValue(deferredMetadataCell),
    [deferredMetadataCell, getCellDisplayValue, revision, workerCellSnapshotRevision]
  );
  const selectedCellFormula = React.useMemo(
    () => getCellFormula(deferredMetadataCell),
    [deferredMetadataCell, getCellFormula, revision, workerCellSnapshotRevision]
  );
  const getChartSeriesFormula = React.useCallback((chartId: string, seriesIndex: number) => (
    buildChartSeriesFormula(getChartById(chartId), seriesIndex)
  ), [getChartById]);
  const selectedChartFormula = React.useMemo(() => {
    if (
      !selectedChartElement
      || selectedChartElement.kind === "chart"
      || selectedChartElement.kind === "legend"
      || selectedChartElement.seriesIndex < 0
    ) {
      return null;
    }

    return getChartSeriesFormula(selectedChartElement.chartId, selectedChartElement.seriesIndex);
  }, [getChartSeriesFormula, selectedChartElement]);
  const selectedFormulaTarget = React.useMemo(() => {
    if (selectedChartFormula && selectedChartElement && selectedChartElement.kind !== "chart" && selectedChartElement.kind !== "legend") {
      return {
        chartId: selectedChartElement.chartId,
        kind: "chartSeries" as const,
        seriesId: selectedChartElement.seriesId,
        seriesIndex: selectedChartElement.seriesIndex
      };
    }

    return {
      cell: deferredMetadataCell,
      kind: "cell" as const
    };
  }, [deferredMetadataCell, selectedChartElement, selectedChartFormula]);
  const selectedFormula = selectedChartFormula ?? selectedCellFormula;
  const isLoadDeferred = deferredLoadFileSize !== null;
  const canLoadDeferred = !isLoading && isLoadDeferred;
  const canUndo = !readOnly && undoStackRef.current.length > 0;
  const canRedo = !readOnly && redoStackRef.current.length > 0;

  const renderChartAsync = React.useCallback<NonNullable<XlsxViewerController["renderChartAsync"]>>((...args) => getWorkerClient().renderChart(...args), [getWorkerClient]);

  const createSavedWorkbookBytes = React.useCallback((targetWorkbook: Workbook) => {
    return getWorkerClient().exportWorkbook(targetWorkbook.exportSnapshot(), imageAssetsRef.current, sheetOriginsRef.current);
  }, [getWorkerClient]);

  const createHistoryEntry = React.useCallback((): SnapshotHistoryEntry | null => {
    if (!workbook) {
      return null;
    }

    return {
      kind: "snapshot",
      activeCell,
      activeSheetIndex,
      bytes: workbook.exportSnapshot(),
      assets: imageAssetsRef.current ? structuredClone(imageAssetsRef.current) : null,
      origins: structuredClone(sheetOriginsRef.current),
      selection,
      sortState
    };
  }, [activeCell, activeSheetIndex, selection, sortState, workbook]);

  const captureCellMutationState = React.useCallback((cell: XlsxCellAddress): CellMutationState | null => {
    const worksheet = getActiveWorksheet();
    if (!worksheet) {
      return null;
    }

    return {
      formula: worksheet.getFormulaAt(cell.row, cell.col) ?? null,
      style: worksheet.getCellStyleAt(cell.row, cell.col),
      valueSnapshot: worksheet.getValueSnapshotAt(cell.row, cell.col),
      value: worksheet.getCellAt(cell.row, cell.col).toJs()
    };
  }, [getActiveWorksheet]);

  const restoreHistoryEntry = React.useCallback(async (entry: SnapshotHistoryEntry) => {
    const wasmModule = await getSheetsWasmModule();
    const nextWorkbook = wasmModule.Workbook.fromSnapshot(entry.bytes);
    const nextImageAssets = entry.assets ? structuredClone(entry.assets) : null;
    const nextSheets = buildSheetList(
      nextWorkbook,
      nextImageAssets?.sheetStatesByWorkbookSheetIndex,
      nextImageAssets?.themePalette,
      nextImageAssets?.styleById,
      nextImageAssets?.namedCellStyleByName,
      nextImageAssets?.tableStyleByName,
      showHiddenSheets
    );
    const nextSheetIndex = Math.max(0, Math.min(entry.activeSheetIndex, Math.max(0, nextSheets.length - 1)));

    setError(null);
    setIsLoading(false);
    setImageAssets(nextImageAssets);
    sheetOriginsRef.current = structuredClone(entry.origins);
    setWorkbook(nextWorkbook);
    setSheets(nextSheets);
    const nextChartAssets = loadWorkbookChartAssets(nextWorkbook, nextImageAssets, buildVisibleSheetIndexMap(nextSheets), showHiddenSheets);
    setChartAssets(nextChartAssets);
    setActiveSheetIndexState(nextSheetIndex);
    const nextTabIndex = nextChartAssets.tabs.findIndex((tab) => tab.kind === "sheet" && tab.sheetIndex === nextSheetIndex);
    if (nextTabIndex >= 0) {
      setActiveTabIndexState(nextTabIndex);
    }
    setActiveCell(entry.activeCell);
    setSelection(entry.selection);
    setSortState(entry.sortState);
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
    setSortState(direction === "undo" ? entry.sortStateBefore : entry.sortStateAfter);
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

  const addFormControl = React.useCallback((input: XlsxFormControlInput, sheetIndex = activeSheetIndex) => {
    const target = getFormControlWorksheet(sheetIndex);
    if (readOnly || !workbook || !target) {
      return null;
    }

    recordHistoryBeforeMutation();
    target.worksheet.addDrawing(xlsxFormControlInputToDukeDrawing(input));
    const controlIndex = target.worksheet.formControlCount - 1;
    const addedControl = target.worksheet.formControls[controlIndex];
    const addedKind = addedControl?.formControl.kind;
    if (addedControl && (addedKind?.kind === "checkbox" || addedKind?.kind === "optionButton")) {
      target.worksheet.setFormControlCheckState(addedControl.drawingPath, addedKind.state);
    } else if (addedKind?.kind === "groupBox") {
      reconcileCheckedOptionButtons(target.worksheet);
    } else if (addedKind && "cellLink" in addedKind && addedKind.cellLink) {
      workbook.syncFormControls();
    }
    maybeRecalculateWorkbook(workbook);
    refreshWorkbookState(workbook);
    return controlIndex;
  }, [activeSheetIndex, getFormControlWorksheet, maybeRecalculateWorkbook, readOnly, recordHistoryBeforeMutation, refreshWorkbookState, workbook]);

  const updateFormControl = React.useCallback((
    controlIndex: number,
    patch: XlsxFormControlPatch,
    sheetIndex = activeSheetIndex
  ) => {
    const target = getFormControlWorksheet(sheetIndex);
    const currentControl = target?.worksheet.formControls[controlIndex];
    if (readOnly || !workbook || !target || !currentControl) {
      return false;
    }

    recordHistoryBeforeMutation();
    const currentInput = formControlInputFromDukeControl(currentControl);
    if (currentInput.kind !== "formControl" || !currentControl.anchor) {
      return false;
    }
    const nextAnchor = patch.anchor
      ? xlsxAnchorToDukeDrawingAnchor(
          patch.anchor,
          patch.editAs ?? (currentControl.anchor.type === "twoCell" ? currentControl.anchor.editAs : undefined)
        )
      : patch.editAs !== undefined && currentControl.anchor.type === "twoCell"
        ? { ...currentControl.anchor, editAs: patch.editAs }
        : currentControl.anchor;
    let nextKind = patch.kind
      ? xlsxFormControlKindToDukeKind(patch.kind)
      : dukeFormControlKindToInput(currentControl.formControl.kind);
    if (currentControl.formControl.kind.kind === "unknown" && nextKind.kind === "unknown") {
      nextKind = {
        ...nextKind,
        legacyObjectType: nextKind.legacyObjectType ?? currentControl.formControl.kind.legacyObjectType
      };
    }
    const patchedRawObj = patch.kind?.kind === "editbox" || patch.kind?.kind === "unknown"
      ? patch.kind.rawObj
      : undefined;
    const patchedRawProperties = patch.kind?.kind === "editbox" || patch.kind?.kind === "unknown"
      ? patch.kind.rawProperties
      : undefined;
    const nextInput: DukeFormControlDrawingInput = {
      altText: "altText" in patch ? patch.altText : currentControl.altText,
      anchor: nextAnchor,
      formControl: {
        ...currentInput.formControl,
        kind: nextKind,
        macroName: "macroName" in patch ? patch.macroName : currentControl.formControl.macroName,
        rawClientData: "rawClientData" in patch ? patch.rawClientData : currentControl.formControl.rawClientData,
        rawObj: patchedRawObj ?? currentControl.formControl.rawObj,
        rawProperties: patchedRawProperties ?? currentControl.formControl.rawProperties
      },
      hidden: "hidden" in patch ? patch.hidden : currentControl.hidden,
      kind: "formControl",
      locked: "locked" in patch ? patch.locked : currentControl.locked,
      name: "name" in patch ? patch.name : currentControl.name,
      printable: "printable" in patch ? patch.printable : currentControl.printable,
      title: "title" in patch ? patch.title : currentControl.title
    };

    target.worksheet.setDrawing(currentControl.drawingPath, nextInput);
    const updatedControl = target.worksheet.formControls[controlIndex];
    const updatedKind = updatedControl?.formControl.kind;
    const interactionChanged = updatedKind
      ? formControlInteractionChanged(currentControl.formControl.kind, updatedKind)
      : false;
    const groupTopologyChanged = patch.anchor !== undefined || (
      updatedKind !== undefined
      && updatedKind.kind !== currentControl.formControl.kind.kind
      && (updatedKind.kind === "groupBox" || currentControl.formControl.kind.kind === "groupBox")
    );
    if (interactionChanged && updatedControl && (updatedKind?.kind === "checkbox" || updatedKind?.kind === "optionButton")) {
      target.worksheet.setFormControlCheckState(updatedControl.drawingPath, updatedKind.state);
    } else if (groupTopologyChanged) {
      reconcileCheckedOptionButtons(target.worksheet);
    } else if (interactionChanged && updatedKind && "cellLink" in updatedKind && updatedKind.cellLink) {
      workbook.syncFormControls();
    }
    maybeRecalculateWorkbook(workbook);
    refreshWorkbookState(workbook);
    return true;
  }, [activeSheetIndex, getFormControlWorksheet, maybeRecalculateWorkbook, readOnly, recordHistoryBeforeMutation, refreshWorkbookState, workbook]);

  const removeFormControl = React.useCallback((controlIndex: number, sheetIndex = activeSheetIndex) => {
    const target = getFormControlWorksheet(sheetIndex);
    if (readOnly || !workbook || !target || !target.worksheet.formControls[controlIndex]) {
      return false;
    }

    recordHistoryBeforeMutation();
    const removedControl = target.worksheet.formControls[controlIndex]!;
    const removedKind = removedControl.formControl.kind.kind;
    target.worksheet.removeDrawing(removedControl.drawingPath);
    if (removedKind === "groupBox" || removedKind === "optionButton") {
      reconcileCheckedOptionButtons(target.worksheet);
      maybeRecalculateWorkbook(workbook);
    }
    refreshWorkbookState(workbook);
    return true;
  }, [activeSheetIndex, getFormControlWorksheet, maybeRecalculateWorkbook, readOnly, recordHistoryBeforeMutation, refreshWorkbookState, workbook]);

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
    activeCellAfter: XlsxCellAddress | null,
    sortStateAfter: XlsxTableSortState | null = sortState
  ) => {
    if (!activeSheet || isApplyingHistoryRef.current || (mutations.length === 0 && sortStateAfter === sortState)) {
      return;
    }

    pushHistoryEntry(undoStackRef.current, {
      kind: "range-edit",
      activeCellAfter,
      activeCellBefore: activeCell,
      mutations,
      selectionAfter,
      selectionBefore: selection,
      sheetIndex: activeSheet.workbookSheetIndex,
      sortStateAfter,
      sortStateBefore: sortState
    });
    redoStackRef.current = [];
    setHistoryRevision((current) => current + 1);
  }, [activeCell, activeSheet, selection, sortState]);

  const sortTable = React.useCallback((tableName: string, columnIndex: number, direction: XlsxTableSortDirection) => {
    const worksheet = getActiveWorksheet();
    const targetTable = tables.find((table) => table.name === tableName || table.displayName === tableName);
    if (!worksheet || !workbook || !activeSheet || !targetTable) {
      return;
    }

    const dataStartRow = targetTable.start.row + Math.max(targetTable.headerRowCount, 1);
    const totalsRowOffset = targetTable.totalsRowCount;
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
          style: worksheet.getCellStyleAt(row, col),
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

    const nextSortState = { columnIndex, direction, tableName: targetTable.name };
    if (sortedRows.every((row, index) => row.index === rows[index]?.index)) {
      if (sortState?.tableName === targetTable.name && sortState.columnIndex === columnIndex && sortState.direction === direction) {
        return;
      }
      setSortState(nextSortState);
      recordRangeEditHistory([], selection, activeCell, nextSortState);
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
    setSortState(nextSortState);
    recordRangeEditHistory(mutations, selection, activeCell, nextSortState);
  }, [
    activeCell,
    activeSheet,
    getActiveWorksheet,
    maybeRecalculateWorkbook,
    readOnly,
    recordRangeEditHistory,
    refreshWorkbookState,
    selection,
    sortState,
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

  const getXlsxBytes = React.useCallback(async (): Promise<Uint8Array | null> => {
    if (!workbook) {
      return null;
    }

    return createSavedWorkbookBytes(workbook);
  }, [createSavedWorkbookBytes, workbook]);

  const exportXlsx = React.useCallback(async () => {
    const bytes = await getXlsxBytes();
    if (!bytes) {
      return;
    }

    downloadBytes(bytes, `${fileStem(displayFileName)}.xlsx`, XLSX_MIME_TYPE);
  }, [displayFileName, getXlsxBytes]);

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

    const result = tryRecalculate(workbook);
    if (result.calculated) {
      refreshWorkbookState(workbook);
      return;
    }

    // Trap poisons the Workbook pointer; skip refreshWorkbookState so we
    // don't crash reading cells from it.
    setShouldAutoCalculate(false);
  }, [refreshWorkbookState, workbook]);

  const applyReadOnlyResizeOverride = React.useCallback((
    axis: "column" | "row",
    actualIndex: number,
    sizePx: number
  ) => {
    if (!activeSheet) {
      return;
    }

    const contentSizePx = activeSheet.officeLayout ? sizePx : resolveContentSheetAxisPixels(sizePx, activeSheet.showGridLines);
    const targetWorkbookSheetIndex = activeSheet.workbookSheetIndex;
    setSheets((currentSheets) => currentSheets.map((sheet) => {
      if (sheet.workbookSheetIndex !== targetWorkbookSheetIndex) {
        return sheet;
      }

      if (axis === "column") {
        const nextColWidthOverridesPx = {
          ...sheet.colWidthOverridesPx,
          [actualIndex]: contentSizePx
        };
        const nextColWidths = [...sheet.colWidths];
        const visibleColIndex = sheet.visibleCols.indexOf(actualIndex);
        if (visibleColIndex >= 0) {
          nextColWidths[visibleColIndex] = contentSizePx;
        }

        return {
          ...sheet,
          colWidthOverridesPx: nextColWidthOverridesPx,
          colWidths: nextColWidths
        };
      }

      const nextRowHeightOverridesPx = {
        ...sheet.rowHeightOverridesPx,
        [actualIndex]: contentSizePx
      };
      const nextRowHeights = [...sheet.rowHeights];
      const visibleRowIndex = sheet.visibleRows.indexOf(actualIndex);
      if (visibleRowIndex >= 0) {
        nextRowHeights[visibleRowIndex] = contentSizePx;
      }

      return {
        ...sheet,
        rowHeightOverridesPx: nextRowHeightOverridesPx,
        rowHeights: nextRowHeights
      };
    }));
    setRevision((current) => current + 1);
  }, [activeSheet]);

  const resizeColumn = React.useCallback((col: number, widthPx: number) => {
    if ((readOnly && !canResizeReadOnly) || !activeSheet) {
      return;
    }

    if (isWorkerBacked) {
      applyReadOnlyResizeOverride("column", col, widthPx);
      return;
    }

    if (!workbook) {
      return;
    }

    recordHistoryBeforeMutation();
    const worksheet = workbook.getSheet(activeSheet.workbookSheetIndex);
    worksheet.setColumnWidth(
      col,
      activeSheet.officeLayout
        ? pxToOfficeColumnWidth(widthPx, activeSheet.officeLayout.columnDigitWidthPt)
        : pxToSheetColumnWidth(resolveContentSheetAxisPixels(widthPx, activeSheet.showGridLines))
    );
    refreshWorkbookState(workbook);
  }, [
    activeSheet,
    applyReadOnlyResizeOverride,
    canResizeReadOnly,
    isWorkerBacked,
    readOnly,
    recordHistoryBeforeMutation,
    refreshWorkbookState,
    workbook
  ]);

  const resizeRow = React.useCallback((row: number, heightPx: number) => {
    if ((readOnly && !canResizeReadOnly) || !activeSheet) {
      return;
    }

    if (isWorkerBacked) {
      applyReadOnlyResizeOverride("row", row, heightPx);
      return;
    }

    if (!workbook) {
      return;
    }

    recordHistoryBeforeMutation();
    const worksheet = workbook.getSheet(activeSheet.workbookSheetIndex);
    worksheet.setRowHeight(
      row,
      activeSheet.officeLayout
        ? Math.max(heightPx, MIN_ROW_HEIGHT_PX) * (3 / 4)
        : pxToSheetRowHeight(resolveContentSheetAxisPixels(heightPx, activeSheet.showGridLines))
    );
    refreshWorkbookState(workbook);
  }, [
    activeSheet,
    applyReadOnlyResizeOverride,
    canResizeReadOnly,
    isWorkerBacked,
    readOnly,
    recordHistoryBeforeMutation,
    refreshWorkbookState,
    workbook
  ]);

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

  const setChartRect = React.useCallback((id: string, rect: XlsxImageRect, layout?: XlsxDrawingLayout) => {
    const hydratedChartAssets = ensureChartAssetsHydrated(workbook, sheets);
    console.info("[react-xlsx debug] setChartRect", {
      hasActiveSheet: Boolean(activeSheet),
      hasHydratedChartAssets: Boolean(hydratedChartAssets),
      hasImageAssets: Boolean(imageAssetsRef.current),
      hasWorkbook: Boolean(workbook),
      id,
      readOnly,
      rect
    });
    if (readOnly || !workbook || !activeSheet || !imageAssetsRef.current || !hydratedChartAssets) {
      return;
    }

    const worksheet = workbook.getSheet(activeSheet.workbookSheetIndex);
    const currentChart = getChartById(id);
    console.info("[react-xlsx debug] currentChart", {
      activeWorkbookSheetIndex: activeSheet.workbookSheetIndex,
      editable: currentChart?.editable,
      found: Boolean(currentChart),
      originCount: hydratedChartAssets.chartOriginsById.size,
      workbookSheetIndex: currentChart?.workbookSheetIndex
    });
    if (!currentChart || currentChart.editable === false || currentChart.workbookSheetIndex !== activeSheet.workbookSheetIndex) {
      return;
    }

    const nextAnchor = rectToImageAnchor(rect, currentChart.anchor, {
      contentOffsetLeft: GRID_ROW_HEADER_WIDTH,
      contentOffsetTop: GRID_HEADER_HEIGHT,
      getColumnWidthPx: (col) => layout?.columnWidths[col] ?? getColumnWidthPx(worksheet, col),
      getRowHeightPx: (row) => layout?.rowHeights[row] ?? getRowHeightPx(worksheet, row)
    });

    recordHistoryBeforeMutation();
    const didUpdateAnchor = updateWorkbookChartAnchor(imageAssetsRef.current, hydratedChartAssets, id, nextAnchor);
    console.info("[react-xlsx debug] updateWorkbookChartAnchor", { didUpdateAnchor, nextAnchor });

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

  const setImageRect = React.useCallback((id: string, rect: XlsxImageRect, layout?: XlsxDrawingLayout) => {
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
      getColumnWidthPx: (col) => layout?.columnWidths[col] ?? getColumnWidthPx(worksheet, col),
      getRowHeightPx: (row) => layout?.rowHeights[row] ?? getRowHeightPx(worksheet, row)
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
      sheetCharts.map((chart) => {
        if (chart.id !== id) return chart;
        const nextChart = { ...chart, ...patch, renderSource: patch.renderSource ?? chart.renderSource };
        return imageAssetsRef.current ? refreshChartRenderSource(nextChart, imageAssetsRef.current, workbook) : nextChart;
      })
    )));
    setRevision((current) => current + 1);
  }, [ensureChartAssetsHydrated, getChartById, readOnly, recordHistoryBeforeMutation, sheets, workbook]);

  const setChartSeriesFormula = React.useCallback((chartId: string, seriesIndex: number, formula: string) => {
    if (readOnly) {
      return false;
    }

    const chart = getChartById(chartId);
    if (!chart || chart.editable === false) {
      return false;
    }

    const nextChart = applyChartSeriesFormula(chart, seriesIndex, formula, workbook);
    if (!nextChart) {
      return false;
    }

    updateChart(chartId, {
      series: nextChart.series,
      typeGroups: nextChart.typeGroups,
      raw: nextChart.raw
    });
    const selectedSeries = nextChart.series[seriesIndex];
    if (selectedSeries) {
      setSelectedChartElement((current) => (
        current && current.chartId === chartId && current.kind !== "chart"
          ? {
              ...current,
              seriesId: selectedSeries.id,
              seriesIndex
            }
          : current
      ));
    }
    return true;
  }, [getChartById, readOnly, updateChart, workbook]);

  const findSelectionMerge = React.useMemo(() => createMergedRangeLookup(activeSheet?.mergedRanges ?? []), [activeSheet?.mergedRanges]);

  const selectCell = React.useCallback((cell: XlsxCellAddress, options?: { extend?: boolean }) => {
    const merged = findSelectionMerge(cell);
    const anchorCell = merged?.start ?? cell;
    setSelectedChartId(null);
    setSelectedChartElement(null);
    setSelectedImageId(null);
    setActiveCell(anchorCell);
    if (options?.extend && selectionAnchorRef.current) {
      setSelection(expandHighlightedRange(normalizeRange({ start: selectionAnchorRef.current, end: cell }), activeSheet?.mergedRanges ?? []));
      return;
    }
    selectionAnchorRef.current = anchorCell;
    setSelection(merged ?? { start: cell, end: cell });
  }, [activeSheet?.mergedRanges, findSelectionMerge]);

  const selectRange = React.useCallback((range: XlsxCellRange) => {
    const normalized = expandHighlightedRange(normalizeRange(range), activeSheet?.mergedRanges ?? []);
    setSelectedChartId(null);
    setSelectedChartElement(null);
    setSelectedImageId(null);
    selectionAnchorRef.current = normalized.start;
    setActiveCell(findSelectionMerge(normalized.end)?.start ?? normalized.end);
    setSelection(normalized);
  }, [activeSheet?.mergedRanges, findSelectionMerge]);

  const clearSelection = React.useCallback(() => {
    selectionAnchorRef.current = null;
    setActiveCell(null);
    setSelection(null);
    setSelectedChartId(null);
    setSelectedChartElement(null);
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
        const after = captureCellMutationState(cell);
        if (!after) {
          continue;
        }
        mutations.push({
          after,
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

    if (locale !== "en-US" && value === getCellDisplayValue(cell)) return;
    const nextValue = isTextCell(worksheet, cell) ? (value.startsWith("'") ? value.slice(1) : value) : coerceUserEnteredValue(value, locale);
    worksheet.setCell(cellAddressToA1(cell), nextValue);
    if (typeof nextValue === "number") applyInputNumberFormat(worksheet, cell, value);
    const after = captureCellMutationState(cell);
    if (!after) {
      return;
    }
    maybeRecalculateWorkbook(workbook);
    refreshWorkbookState(workbook);
    recordCellEditHistory(cell, before, after);
  }, [captureCellMutationState, getActiveWorksheet, getCellDisplayValue, locale, maybeRecalculateWorkbook, readOnly, recordCellEditHistory, refreshWorkbookState, workbook]);

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
      worksheet.setFormula(cellAddressToA1(cell), normalizeFormulaReferences(formula));
      applyInputNumberFormat(worksheet, cell, formula);
    }
    const after = captureCellMutationState(cell);
    if (!after) {
      return;
    }
    maybeRecalculateWorkbook(workbook);
    refreshWorkbookState(workbook);
    recordCellEditHistory(cell, before, after);
  }, [captureCellMutationState, getActiveWorksheet, maybeRecalculateWorkbook, readOnly, recordCellEditHistory, refreshWorkbookState, workbook]);

  const setCellStyle = React.useCallback((cell: XlsxCellAddress, style: XlsxCellStyleInput) => {
    const worksheet = getActiveWorksheet();
    if (readOnly || !worksheet || !workbook) {
      return;
    }

    const before = captureCellMutationState(cell);
    if (!before) {
      return;
    }

    worksheet.setCellStyleAt(cell.row, cell.col, style);
    const after = captureCellMutationState(cell);
    if (!after) {
      return;
    }

    refreshWorkbookState(workbook);
    recordCellEditHistory(cell, before, after);
  }, [captureCellMutationState, getActiveWorksheet, readOnly, recordCellEditHistory, refreshWorkbookState, workbook]);

  const setCellStyles = React.useCallback((updates: ReadonlyArray<{ cell: XlsxCellAddress; style: XlsxCellStyleInput }>) => {
    const worksheet = getActiveWorksheet();
    if (readOnly || !worksheet || !workbook) return;
    const beforeStates = new Map<string, { cell: XlsxCellAddress; before: CellMutationState }>();
    for (const { cell, style } of updates) {
      const key = `${cell.row}:${cell.col}`;
      if (!beforeStates.has(key)) {
        const before = captureCellMutationState(cell);
        if (!before) continue;
        beforeStates.set(key, { cell: { ...cell }, before });
      }
      worksheet.setCellStyleAt(cell.row, cell.col, style);
    }
    if (!beforeStates.size) return;
    const mutations: RangeCellMutation[] = [];
    for (const { cell, before } of beforeStates.values()) {
      const after = captureCellMutationState(cell);
      if (after) mutations.push({ cell, before, after });
    }
    refreshWorkbookState(workbook);
    recordRangeEditHistory(mutations, selection, activeCell);
  }, [activeCell, captureCellMutationState, getActiveWorksheet, readOnly, recordRangeEditHistory, refreshWorkbookState, selection, workbook]);

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

  const setSelectedFormula = React.useCallback((formula: string) => {
    if (selectedFormulaTarget?.kind === "chartSeries") {
      return setChartSeriesFormula(selectedFormulaTarget.chartId, selectedFormulaTarget.seriesIndex, formula);
    }

    if (!activeCell) {
      return false;
    }

    setCellFormula(activeCell, formula);
    return true;
  }, [activeCell, selectedFormulaTarget, setCellFormula, setChartSeriesFormula]);

  const setSelectedCellStyle = React.useCallback((style: XlsxCellStyleInput) => {
    if (!activeCell) {
      return;
    }

    setCellStyle(activeCell, style);
  }, [activeCell, setCellStyle]);

  const setRangeStyle = React.useCallback((range: XlsxCellRange, style: XlsxCellStyleInput) => {
    const worksheet = getActiveWorksheet();
    if (readOnly || !worksheet || !workbook) {
      return;
    }

    const normalized = normalizeRange(range);
    const beforeStates: Array<{ before: CellMutationState; cell: XlsxCellAddress }> = [];
    for (let row = normalized.start.row; row <= normalized.end.row; row += 1) {
      for (let col = normalized.start.col; col <= normalized.end.col; col += 1) {
        const cell = { row, col };
        const before = captureCellMutationState(cell);
        if (!before) {
          continue;
        }
        beforeStates.push({
          before,
          cell
        });
      }
    }

    if (beforeStates.length === 0) {
      return;
    }

    worksheet.setRangeStyle(rangeToA1(normalized), style);
    const mutations: RangeCellMutation[] = [];
    for (const mutation of beforeStates) {
      const after = captureCellMutationState(mutation.cell);
      if (!after) {
        continue;
      }
      mutations.push({
        after,
        before: mutation.before,
        cell: mutation.cell
      });
    }

    refreshWorkbookState(workbook);
    recordRangeEditHistory(mutations, selection, activeCell);
  }, [activeCell, captureCellMutationState, getActiveWorksheet, readOnly, recordRangeEditHistory, refreshWorkbookState, selection, workbook]);

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
        const sourceStyle = cloneCellStyle(worksheet.getCellStyleAt(sourceRow, sourceCol));

        if (sourceFormula) {
          worksheet.setFormula(cellAddressToA1(targetCell), sourceFormula);
        } else {
          const sourceValue = normalizeCellValue(worksheet.getCellAt(sourceRow, sourceCol).toJs());
          worksheet.setCell(cellAddressToA1(targetCell), sourceValue);
        }

        if (sourceStyle && typeof sourceStyle === "object") {
          worksheet.setCellStyleAt(targetCell.row, targetCell.col, sourceStyle);
        }

        const after = captureCellMutationState(targetCell);
        if (!after) {
          continue;
        }
        mutations.push({
          after,
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

  const mergeSelection = React.useCallback((mode: XlsxMergeMode = "cells") => {
    const worksheet = getActiveWorksheet();
    if (readOnly || !worksheet || !selection || !workbook) {
      return;
    }

    const mergedRanges = resolveWorksheetMergeMetadata(worksheet).mergedRanges;
    const range = expandHighlightedRange(normalizeRange(selection), mergedRanges);
    const isSingleCell = range.start.row === range.end.row && range.start.col === range.end.col;
    if ((isSingleCell && mode !== "center") || (mode === "across" && range.start.col === range.end.col)) {
      return;
    }

    recordHistoryBeforeMutation();
    for (const merged of mergedRanges) {
      if (rangesIntersect(range, merged)) {
        worksheet.unmergeCells(rangeToA1(merged));
      }
    }
    if (mode === "across") {
      for (let row = range.start.row; row <= range.end.row; row += 1) {
        worksheet.mergeCells(rangeToA1({
          start: { row, col: range.start.col },
          end: { row, col: range.end.col }
        }));
      }
    } else {
      if (!isSingleCell) worksheet.mergeCells(rangeToA1(range));
      if (mode === "center") {
        worksheet.setCellStyleAt(range.start.row, range.start.col, { alignment: { horizontal: "center" } });
      }
    }
    refreshWorkbookState(workbook);
    selectionAnchorRef.current = range.start;
    setSelection(range);
    setActiveCell(range.start);
  }, [getActiveWorksheet, readOnly, recordHistoryBeforeMutation, refreshWorkbookState, selection, workbook]);

  const unmergeSelection = React.useCallback(() => {
    const worksheet = getActiveWorksheet();
    if (readOnly || !worksheet || !selection || !workbook) {
      return;
    }

    const range = normalizeRange(selection);
    const mergedRanges = resolveWorksheetMergeMetadata(worksheet).mergedRanges
      .filter((merged) => rangesIntersect(range, merged));
    if (mergedRanges.length === 0) return;

    recordHistoryBeforeMutation();
    for (const merged of mergedRanges) {
      worksheet.unmergeCells(rangeToA1(merged));
    }
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
    setFormControlsByWorkbookSheetIndex((current) => [...current, []]);
    setImagesByWorkbookSheetIndex((current) => [...current, []]);
    setShapesByWorkbookSheetIndex((current) => [...current, []]);
    if (imageAssetsRef.current) {
      imageAssetsRef.current.formControlsByWorkbookSheetIndex = [
        ...imageAssetsRef.current.formControlsByWorkbookSheetIndex,
        []
      ];
    }
    const nextSheets = buildSheetList(
      workbook,
      imageAssetsRef.current?.sheetStatesByWorkbookSheetIndex,
      imageAssetsRef.current?.themePalette,
      imageAssetsRef.current?.styleById,
      imageAssetsRef.current?.namedCellStyleByName,
      imageAssetsRef.current?.tableStyleByName,
      showHiddenSheets
    );
    setSheets(nextSheets);
    const nextChartAssets = imageAssetsRef.current
      ? loadWorkbookChartAssets(workbook, imageAssetsRef.current, buildVisibleSheetIndexMap(nextSheets), showHiddenSheets)
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
    setFormControlsByWorkbookSheetIndex((current) => current.filter((_, index) => index !== activeSheet.workbookSheetIndex));
    setImagesByWorkbookSheetIndex((current) => current.filter((_, index) => index !== activeSheet.workbookSheetIndex));
    setShapesByWorkbookSheetIndex((current) => current.filter((_, index) => index !== activeSheet.workbookSheetIndex));
    if (imageAssetsRef.current) {
      imageAssetsRef.current.formControlsByWorkbookSheetIndex = imageAssetsRef.current.formControlsByWorkbookSheetIndex.filter(
        (_, index) => index !== activeSheet.workbookSheetIndex
      );
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
      imageAssetsRef.current?.tableStyleByName,
      showHiddenSheets
    );
    setSheets(nextSheets);
    if (imageAssetsRef.current) {
      setChartAssets(loadWorkbookChartAssets(workbook, imageAssetsRef.current, buildVisibleSheetIndexMap(nextSheets), showHiddenSheets));
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
        const textCell = isTextCell(worksheet, nextCell);
        const formula = textCell ? null : getUserEnteredFormula(rawValue, locale);
        if (formula) {
          worksheet.setFormula(cellAddressToA1(nextCell), normalizeFormulaReferences(formula));
          applyInputNumberFormat(worksheet, nextCell, formula);
          const after = captureCellMutationState(nextCell);
          if (!after) {
            continue;
          }
          mutations.push({
            after,
            before,
            cell: nextCell
          });
        } else {
          const nextValue = textCell ? rawValue : coerceUserEnteredValue(rawValue, locale, "paste");
          worksheet.setCell(cellAddressToA1(nextCell), nextValue);
          if (typeof nextValue === "number") applyInputNumberFormat(worksheet, nextCell, rawValue);
          const after = captureCellMutationState(nextCell);
          if (!after) {
            continue;
          }
          mutations.push({
            after,
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
  }, [activeCell, captureCellMutationState, getActiveWorksheet, locale, maybeRecalculateWorkbook, readOnly, recordRangeEditHistory, refreshWorkbookState, selection, workbook]);

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
          const after = captureCellMutationState(nextCell);
          if (!after) {
            continue;
          }
          mutations.push({
            after,
            before,
            cell: nextCell
          });
        }
      } else {
        worksheet.setCell(cellAddressToA1(nextCell), cell.value);
        if (before) {
          const after = captureCellMutationState(nextCell);
          if (!after) {
            continue;
          }
          mutations.push({
            after,
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
    if (typeof navigator === "undefined" || !navigator.clipboard) {
      return false;
    }

    if (isWorkerBacked && activeSheet && selection) {
      const range = normalizeRange(selection);
      const textPromise = (async () => {
        const lines: string[] = [];
        for (let startRow = range.start.row; startRow <= range.end.row; startRow += 128) {
          const count = Math.min(128, range.end.row - startRow + 1);
          const batch = await getRowsBatchAsync(activeSheet.workbookSheetIndex, startRow, count);
          if (!batch) throw new Error("Unable to read the selected cells.");
          const values = new Map<string, string>();
          for (const entry of batch as Array<{ index: number; cells?: Array<{ col: number; value?: unknown; formula?: string; isMergedSecondary?: boolean }> }>) {
            for (const cell of entry.cells ?? []) {
              if (cell.col < range.start.col || cell.col > range.end.col || cell.isMergedSecondary) continue;
              const value = String(cell.value ?? "");
              const cached = cell.formula ? activeSheet.cachedFormulaValues[cellAddressToA1({ row: entry.index, col: cell.col })] : undefined;
              values.set(`${entry.index}:${cell.col}`, decodeHtmlEntities(cached !== undefined && value.startsWith("#") ? cached : value));
            }
          }
          for (let row = startRow; row < startRow + count; row++) {
            const cells: string[] = [];
            for (let col = range.start.col; col <= range.end.col; col++) cells.push(values.get(`${row}:${col}`) ?? "");
            lines.push(cells.join("\t"));
          }
        }
        return lines.join("\n");
      })();
      if (typeof ClipboardItem === "function" && navigator.clipboard.write) {
        await navigator.clipboard.write([new ClipboardItem({
          "text/plain": textPromise.then((text) => new Blob([text], { type: "text/plain" }))
        })]);
      } else {
        await navigator.clipboard.writeText(await textPromise);
      }
      return true;
    }

    const clipboardData = getClipboardData();
    if (!clipboardData) return false;

    if (typeof ClipboardItem === "function" && navigator.clipboard.write) {
      const customType = `web ${INTERNAL_CLIPBOARD_MIME}`;
      try {
        await navigator.clipboard.write([new ClipboardItem({
          [customType]: new Blob([clipboardData.structured], { type: INTERNAL_CLIPBOARD_MIME }),
          "text/html": new Blob([clipboardData.html], { type: "text/html" }),
          "text/plain": new Blob([clipboardData.text], { type: "text/plain" })
        })]);
        return true;
      } catch (error) {
        if (!(error instanceof TypeError) &&
          (!(error instanceof DOMException) || (error.name !== "NotAllowedError" && error.name !== "NotSupportedError"))) {
          throw error;
        }
        await navigator.clipboard.write([new ClipboardItem({
          "text/html": new Blob([clipboardData.html], { type: "text/html" }),
          "text/plain": new Blob([clipboardData.text], { type: "text/plain" })
        })]);
        return true;
      }
    }

    await navigator.clipboard.writeText(clipboardData.text);
    return true;
  }, [activeSheet, getClipboardData, getRowsBatchAsync, isWorkerBacked, selection]);

  const pasteFromClipboard = React.useCallback(async () => {
    if (readOnly || typeof navigator === "undefined" || !navigator.clipboard) {
      return false;
    }

    if (navigator.clipboard.read) {
      const items = await navigator.clipboard.read();
      for (const item of items) {
        const type = [`web ${INTERNAL_CLIPBOARD_MIME}`, INTERNAL_CLIPBOARD_MIME].find((type) => item.types.includes(type));
        if (type) {
          const blob = await item.getType(type);
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
      addFormControl,
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
      clearSelectedChartElement,
      clearSelectedCells,
      clearSelectedImage,
      clearSelection,
      clearHighlightedRanges,
      highlightRanges,
      highlightedRanges,
      continueDeferredLoad,
      copySelectionToClipboard,
      defaultZoomScale,
      deferredLoadFileSize,
      defineNamedRange,
      displayFileName,
      locale,
      messages,
      download,
      exportCsv,
      exportXlsx,
      getXlsxBytes,
      renderChartAsync,
      error,
      fillSelection,
      formControls,
      getChartById,
      getChartSeriesFormula,
      getChartsheetById,
      getImageById,
      getSheetCharts,
      getSheetFormControls,
      getSheetImages,
      getSheetShapes,
      file,
      getClipboardData,
      getCellDisplayValue,
      getCellFormula,
      getFormControlItems,
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
      removeFormControl,
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
      setCellStyle,
      setCellStyles,
      setCellValue,
      setRangeStyle,
      setSelectedFormula,
      setZoomScale,
      setChartRect,
      setChartSeriesFormula,
      setImageRect,
      selectedChart,
      selectedChartElement,
      selectedChartFormula,
      selectedChartId,
      selectedCellFormula,
      selectedFormula,
      selectedFormulaTarget,
      selectedImage,
      selectedImageId,
      selectedRangeAddress,
      selectedValue,
      selectCell,
      selectChart,
      selectChartElement,
      selectImage,
      selectRange,
      selection,
      setActiveSheetIndex,
      setActiveTabIndex,
      setSelectedCellFormula,
      setSelectedCellStyle,
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
      updateFormControl,
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
      addFormControl,
      addSheet,
      canLoadDeferred,
      canRedo,
      canUndo,
      canZoomIn,
      canZoomOut,
      charts,
      chartsheets,
      clearSelectedChart,
      clearSelectedChartElement,
      clearSelectedCells,
      clearSelectedImage,
      continueDeferredLoad,
      clearHighlightedRanges,
      highlightRanges,
      highlightedRanges,
      copySelectionToClipboard,
      defaultZoomScale,
      deferredLoadFileSize,
      defineNamedRange,
      displayFileName,
      locale,
      messages,
      download,
      error,
      exportCsv,
      exportXlsx,
      getXlsxBytes,
      renderChartAsync,
      file,
      fillSelection,
      formControls,
      getChartById,
      getChartSeriesFormula,
      getChartsheetById,
      getImageById,
      getSheetCharts,
      getSheetFormControls,
      getSheetImages,
      getSheetShapes,
      getClipboardData,
      getCellDisplayValue,
      getCellFormula,
      getFormControlItems,
      getCellSnapshotAsync,
      getActiveWorksheet,
      historyRevision,
      formControls,
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
      removeFormControl,
      readOnly,
      recalculate,
      redo,
      revision,
      resizeChartBy,
      resizeImageBy,
      resizeColumn,
      resizeRow,
      setCellFormula,
      setCellStyle,
      setCellStyles,
      setCellValue,
      setRangeStyle,
      setSelectedFormula,
      setZoomScale,
      setChartRect,
      setChartSeriesFormula,
      setImageRect,
      selectedChart,
      selectedChartElement,
      selectedChartFormula,
      selectedChartId,
      selectedCellFormula,
      selectedFormula,
      selectedFormulaTarget,
      selectedImage,
      selectedImageId,
      selectedRangeAddress,
      selectedValue,
      selectCell,
      selectChart,
      selectChartElement,
      selectImage,
      selectRange,
      selection,
      setActiveSheetIndex,
      setActiveTabIndex,
      setSelectedCellFormula,
      setSelectedCellStyle,
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
      updateFormControl,
      workbook,
      zoomIn,
      zoomOut,
      zoomScale
    ]
  );
}
