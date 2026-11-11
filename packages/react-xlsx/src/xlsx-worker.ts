import type { ChartWorkerResult, RenderChartAsync } from "./chart-renderer-types";
import { loadWorkbookImageAssets } from "./workbook-assets";
import { configureWorkbookArchiveDecoder, configureWorkbookXmlDecoder, installWorkerXmlAdapters, takeWorkbookXmlRecords } from "./workbook-xml";
import { loadSpreadsheetFonts } from "./spreadsheet-fonts";
import type { XlsxWorkerLoadResult } from "./worker-client";
import { localizeCellNumber, localizeRows } from "./number-locale";
import type { Workbook } from "./duke-sheets/generated/duke_sheets_wasm.js";
import { strFromU8, strToU8 } from "fflate";
import { hydrateWorkbookChartStyles, loadWorkbookChartAssets } from "./charts";
import { buildThemePaletteFromHex } from "./colors";
import { getCellGeneralAlignment } from "./cell-alignment";
import { getColumnFontMeasurements, getUnmeasuredWorkbookFont, restoreColumnFontMeasurements, collectWorkbookFormControls, mergeWorkbookImageAssets, type WorkbookImageAssets, type WorkbookImageSheetOrigin, resolveSheetColumnWidthPixels, resolveWorksheetDefaultColumnWidthPixels, resolveWorksheetDefaultRowHeightPixels, resolveWorksheetMergeMetadata } from "./images";
import type { WorkbookStructureAssets } from "./images";
import { safeCalculate } from "./safe-calculate";
import { getSheetsWasmModule, setWasmSource, type WorkerWasmSource } from "./wasm";
import type { XlsxChart, XlsxChartsheet, XlsxCellAddress, XlsxCellRange, XlsxDataValidation, XlsxFormControl, XlsxFreezePanes, XlsxSheetData, XlsxSheetVisibility, XlsxTable, XlsxThemePalette, XlsxWorkbookTab } from "./types";

const DEFAULT_ROW_HEIGHT = 24;
const DEFAULT_COL_WIDTH = 80;
const DEFAULT_ZOOM_SCALE = 100;
const FORMULA_COUNT_THRESHOLD = 1000;

type WorkerSheetState = Partial<NonNullable<WorkbookStructureAssets["sheetStatesByWorkbookSheetIndex"][number]>>;

function isLegacyXlsWorkbook(bytes: Uint8Array) {
  return bytes.byteLength >= 8
    && bytes[0] === 0xd0
    && bytes[1] === 0xcf
    && bytes[2] === 0x11
    && bytes[3] === 0xe0
    && bytes[4] === 0xa1
    && bytes[5] === 0xb1
    && bytes[6] === 0x1a
    && bytes[7] === 0xe1;
}

function shouldSkipXmlParsingForWorkbook(bytes: Uint8Array, skipXmlParsing = false) {
  return skipXmlParsing || isLegacyXlsWorkbook(bytes);
}

function normalizeWorksheetVisibility(value: unknown): XlsxSheetVisibility {
  return value === "hidden" || value === "veryHidden" ? value : "visible";
}

type WorkerRequest =
  | { id: number; type: "renderChart"; payload: Parameters<RenderChartAsync> }
  | { id: number; type: "export"; payload: { snapshot: Uint8Array; assets: WorkbookImageAssets | null; origins: Array<WorkbookImageSheetOrigin | null> } }

  | {
      id: number;
      type: "load";
      payload: {
        buffer: ArrayBuffer;
        includeSnapshot?: boolean;
        columnFontMeasurements?: Array<[string, number]>;
        showHiddenSheets?: boolean;
        skipXmlParsing?: boolean;
        wasmSource?: WorkerWasmSource;
      };
    }
  | {
      id: number;
      type: "getCellSnapshot";
      payload: {
        workbookSheetIndex: number;
        row: number;
        col: number;
        locale?: string;
      };
    }
  | {
      id: number;
      type: "getRowsBatch";
      payload: {
        workbookSheetIndex: number;
        startRow: number;
        rowCount: number;
        locale?: string;
      };
    };

type WorkerSuccessResponse = {
  id: number;
  success: true;
  result:
    | ChartWorkerResult
    | Uint8Array
    | {
        chartsByWorkbookSheetIndex: XlsxChart[][];
        chartsheets: XlsxChartsheet[];
        tabs: XlsxWorkbookTab[];
      }
    | {
        chartsByWorkbookSheetIndex: XlsxChart[][];
        chartsheets: XlsxChartsheet[];
        formControlsByWorkbookSheetIndex: XlsxFormControl[][];
        sheets: XlsxSheetData[];
        tablesByWorkbookSheetIndex: XlsxTable[][];
        tabs: XlsxWorkbookTab[];
      }
    | {
        displayValue: string;
        formula: string;
      }
    | unknown[]
    | null;
};

type WorkerErrorResponse = {
  id: number;
  success: false;
  error: string;
};

type WorkerResponse = WorkerSuccessResponse | WorkerErrorResponse;

let workbook: Workbook | null = null;
let workbookFontFamilies: string[] = [];
let chartsByWorkbookSheetIndex: XlsxChart[][] = [];
let chartsheets: XlsxChartsheet[] = [];
let formControlsByWorkbookSheetIndex: XlsxFormControl[][] = [];
let sheets: XlsxSheetData[] = [];
let tablesByWorkbookSheetIndex: XlsxTable[][] = [];
let tabs: XlsxWorkbookTab[] = [];

function normalizeRange(range: XlsxCellRange): XlsxCellRange {
  return {
    start: {
      col: Math.min(range.start.col, range.end.col),
      row: Math.min(range.start.row, range.end.row)
    },
    end: {
      col: Math.max(range.start.col, range.end.col),
      row: Math.max(range.start.row, range.end.row)
    }
  };
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

  return normalizeRange({ end, start });
}

type ResolvedWorkbookReference = {
  range: XlsxCellRange;
  worksheet: ReturnType<Workbook["getSheet"]>;
};

function resolveWorkbookReference(
  targetWorkbook: Workbook,
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
    const resolvedSheetIndex = targetWorkbook.sheetIndex(sheetName);
    if (resolvedSheetIndex === undefined) {
      return null;
    }
    workbookSheetIndex = resolvedSheetIndex;
  } else if (!resolvingNamedRange) {
    const namedRange = targetWorkbook.getNamedRange(reference);
    if (namedRange) {
      return resolveWorkbookReference(targetWorkbook, defaultWorkbookSheetIndex, namedRange, true);
    }
  }

  const range = parseA1RangeReference(rangeReference.replace(/\$/g, ""));
  if (!range || workbookSheetIndex < 0 || workbookSheetIndex >= targetWorkbook.sheetCount) {
    return null;
  }

  try {
    return {
      range,
      worksheet: targetWorkbook.getSheet(workbookSheetIndex)
    };
  } catch {
    return null;
  }
}

function resolveFormControlItems(
  targetWorkbook: Workbook,
  workbookSheetIndex: number,
  control: XlsxFormControl
) {
  if ((control.kind !== "dropdown" && control.kind !== "listbox") || !control.inputRange) {
    return [];
  }

  const source = resolveWorkbookReference(targetWorkbook, workbookSheetIndex, control.inputRange);
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

function resolveWorksheetZoomScale(
  worksheet: ReturnType<Workbook["getSheet"]>,
  sheetState?: { zoomScale?: number } | null
) {
  const candidates = [
    sheetState?.zoomScale,
    typeof worksheet.zoomScale === "number" ? worksheet.zoomScale : undefined
  ];
  const value = candidates.find((entry): entry is number => typeof entry === "number" && Number.isFinite(entry) && entry > 0);
  return value ?? DEFAULT_ZOOM_SCALE;
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
  nextWorkbook: Workbook,
  themePalette: XlsxThemePalette,
  structureAssets?: WorkbookStructureAssets | null,
  sheetLayoutStates?: Array<WorkerSheetState | null>,
  showHiddenSheets = false
) {
  const sheetsByWorkbookSheetIndex: XlsxSheetData[] = [];

  for (let index = 0; index < nextWorkbook.sheetCount; index += 1) {
    const worksheet = nextWorkbook.getSheet(index);
    const sheetState = structureAssets?.sheetStatesByWorkbookSheetIndex[index] ?? sheetLayoutStates?.[index] ?? null;
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
        return Math.max(Math.round(height * 1.33), 16);
      }

      return sheetState?.rowHeightOverridesPx?.[row] ?? defaultRowHeightPx;
    };

    const usedRange = worksheet.usedRange() as [number, number, number, number] | null;
    if (!usedRange) {
      sheetsByWorkbookSheetIndex.push({
        autoFilterRanges: sheetState?.autoFilterRanges ?? [],
        cachedFormulaValues: sheetState?.cachedFormulaValues ?? {},
        columnWidthCharacterWidthPx: sheetState?.columnWidthCharacterWidthPx,
        colCount: 0,
        colStyleIds: sheetState?.colStyleIds ?? {},
        colWidthOverridesPx: sheetState?.colWidthOverridesPx ?? {},
        colWidths: [],
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
        hiddenCols: sheetState?.hiddenCols ?? [],
        hiddenRows: sheetState?.hiddenRows ?? [],
        minUsedCol: -1,
        minUsedRow: -1,
        maxUsedCol: -1,
        maxUsedRow: -1,
        name: worksheet.name,
        visibility,
        namedCellStyleByName: structureAssets?.namedCellStyleByName ?? {},
        rowCount: 0,
        rowHeightOverridesPx: sheetState?.rowHeightOverridesPx ?? {},
        rowHeights: [],
        rowStyleIds: sheetState?.rowStyleIds ?? {},
        cellStyleIds: sheetState?.cellStyleIds ?? {},
        showGridLines: sheetState?.showGridLines ?? true,
        sparklines: sheetState?.sparklines ?? [],
        styleById: structureAssets?.styleById ?? {},
        tableStyleByName: structureAssets?.tableStyleByName ?? {},
        themePalette,
        visibleCols: [],
        visibleRows: [],
        workbookSheetIndex: index,
        zoomScale: resolveWorksheetZoomScale(worksheet, sheetState)
      });
      continue;
    }

    const [minRow, minCol, maxRow, maxCol] = resolveSheetDisplayUsedRange(usedRange, effectiveSheetState);
    const visibleRows: number[] = [];
    const hiddenRows: number[] = [];
    for (let row = 0; row <= maxRow; row += 1) {
      if (worksheet.isRowHidden(row)) {
        hiddenRows.push(row);
      } else {
        visibleRows.push(row);
      }
    }

    const visibleCols: number[] = [];
    const hiddenCols: number[] = [];
    for (let col = 0; col <= maxCol; col += 1) {
      if (worksheet.isColumnHidden(col)) {
        hiddenCols.push(col);
      } else {
        visibleCols.push(col);
      }
    }

    sheetsByWorkbookSheetIndex.push({
      autoFilterRanges: sheetState?.autoFilterRanges ?? [],
      cachedFormulaValues: sheetState?.cachedFormulaValues ?? {},
      columnWidthCharacterWidthPx: sheetState?.columnWidthCharacterWidthPx,
      colCount: visibleCols.length,
      colStyleIds: sheetState?.colStyleIds ?? {},
      colWidthOverridesPx: sheetState?.colWidthOverridesPx ?? {},
      colWidths: visibleCols.map(resolveColumnWidthPx),
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
      hiddenCols,
      hiddenRows,
      minUsedCol: minCol,
      minUsedRow: minRow,
      maxUsedCol: maxCol,
      maxUsedRow: maxRow,
      name: worksheet.name,
      visibility,
      namedCellStyleByName: structureAssets?.namedCellStyleByName ?? {},
      rowCount: visibleRows.length,
      rowHeightOverridesPx: sheetState?.rowHeightOverridesPx ?? {},
      rowHeights: visibleRows.map(resolveRowHeightPx),
      rowStyleIds: sheetState?.rowStyleIds ?? {},
      cellStyleIds: sheetState?.cellStyleIds ?? {},
      showGridLines: sheetState?.showGridLines ?? true,
      sparklines: sheetState?.sparklines ?? [],
      styleById: structureAssets?.styleById ?? {},
      tableStyleByName: structureAssets?.tableStyleByName ?? {},
      themePalette,
      visibleCols,
      visibleRows,
      workbookSheetIndex: index,
      zoomScale: resolveWorksheetZoomScale(worksheet, sheetState)
    });
  }

  return sheetsByWorkbookSheetIndex;
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
    const reference = `${cellAddressToA1(range.start)}:${cellAddressToA1(range.end)}`;
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

function getCellDisplayValue(worksheet: ReturnType<Workbook["getSheet"]>, row: number, col: number, activeSheet?: XlsxSheetData | null) {
  const formula = worksheet.getFormulaAt(row, col);
  const cachedFormulaValue = formula ? activeSheet?.cachedFormulaValues?.[cellAddressToA1({ row, col })] : undefined;
  const formatted = worksheet.getFormattedValueAt(row, col);
  if (formatted && !(formula && cachedFormulaValue !== undefined && formatted.startsWith("#"))) {
    return decodeHtmlEntities(formatted);
  }

  const cellValue = worksheet.getCalculatedValueAt(row, col);
  if (formula && cachedFormulaValue !== undefined && cellValue.is_error) {
    return cachedFormulaValue;
  }
  if (cellValue.is_error) {
    return cellValue.asError() ?? "";
  }
  if (cellValue.is_empty) {
    return "";
  }

  return cellValue.toString();
}

function cellAddressToA1(cell: XlsxCellAddress) {
  let col = cell.col + 1;
  let label = "";
  while (col > 0) {
    const remainder = (col - 1) % 26;
    label = String.fromCharCode(65 + remainder) + label;
    col = Math.floor((col - 1) / 26);
  }
  return `${label}${cell.row + 1}`;
}

async function loadWorkbook(buffer: ArrayBuffer, skipXmlParsing = false, showHiddenSheets = false, includeSnapshot = false): Promise<XlsxWorkerLoadResult> {
  const wasmModule = await getSheetsWasmModule();
  configureWorkbookXmlDecoder(wasmModule.parseXmlTree);
  configureWorkbookArchiveDecoder(wasmModule.decodeArchive);
  installWorkerXmlAdapters();
  const bytes = new Uint8Array(buffer);
  const preflight = wasmModule.inspectArchive(bytes);
  if (preflight.tooLarge) throw new Error("XLSX is too large to preview interactively.");
  const effectiveSkipXmlParsing = shouldSkipXmlParsingForWorkbook(bytes, skipXmlParsing);
  let activeWorkbook = wasmModule.Workbook.fromBytesForViewer(bytes);
  const fontFamilies = activeWorkbook.fontFamilies();
  const archive = activeWorkbook.takeViewerArchive() as Record<string, Uint8Array>;
  let totalFormulas = 0;
  for (let index = 0; index < activeWorkbook.sheetCount; index += 1) totalFormulas += activeWorkbook.getSheet(index).formulaCount;
  let shouldAutoCalculate = totalFormulas === 0;
  if (totalFormulas > 0 && totalFormulas <= FORMULA_COUNT_THRESHOLD) {
    const snapshot = activeWorkbook.exportSnapshot();
    const result = safeCalculate(activeWorkbook, { reparse: () => wasmModule.Workbook.fromSnapshot(snapshot) });
    activeWorkbook = result.workbook;
    shouldAutoCalculate = result.calculated;
  }
  const nextWorkbook = activeWorkbook;
  const unmeasuredFont = effectiveSkipXmlParsing ? null : getUnmeasuredWorkbookFont(archive);
  if (unmeasuredFont) await loadSpreadsheetFonts([unmeasuredFont], globalThis as unknown as Document);
  workbookFontFamilies = fontFamilies;
  const imageAssets = await loadWorkbookImageAssets(bytes, nextWorkbook, effectiveSkipXmlParsing, archive);
  const structureAssets = imageAssets;
  const themePalette = structureAssets?.themePalette ?? buildThemePaletteFromHex(nextWorkbook.themePalette);
  formControlsByWorkbookSheetIndex = collectWorkbookFormControls(nextWorkbook, themePalette).map(
    (controls, sheetIndex) => controls.map(control => ({ ...control, items: resolveFormControlItems(nextWorkbook, sheetIndex, control) }))
  );
  workbook?.free();
  workbook = nextWorkbook;
  sheets = buildSheetList(nextWorkbook, themePalette, structureAssets, undefined, showHiddenSheets);
  tablesByWorkbookSheetIndex = Array.from({ length: nextWorkbook.sheetCount }, (_, index) =>
    mapWorksheetTables(nextWorkbook.getSheet(index), structureAssets?.sheetStatesByWorkbookSheetIndex[index]?.autoFilterRanges ?? [])
  );
  const visibleSheetIndexByWorkbookSheetIndex = new Map(sheets.map((sheet, index) => [sheet.workbookSheetIndex, index]));
  const chartAssets = loadWorkbookChartAssets(nextWorkbook, null, visibleSheetIndexByWorkbookSheetIndex, showHiddenSheets);
  if (imageAssets) chartAssets.chartOriginsById = hydrateWorkbookChartStyles(chartAssets.chartsByWorkbookSheetIndex, imageAssets, chartAssets.chartsheets, nextWorkbook);
  chartsByWorkbookSheetIndex = chartAssets.chartsByWorkbookSheetIndex;
  chartsheets = chartAssets.chartsheets;
  tabs = chartAssets.tabs;
  return {
    chartsByWorkbookSheetIndex, chartsheets, formControlsByWorkbookSheetIndex, sheets, tablesByWorkbookSheetIndex, tabs,
    chartOriginsById: chartAssets.chartOriginsById,
    columnFontMeasurements: getColumnFontMeasurements(),
    fontFamilies, imageAssets, xmlRecords: takeWorkbookXmlRecords(includeSnapshot ? undefined : new Set(chartsByWorkbookSheetIndex.flat().flatMap(chart => chart.renderSource ? [chart.renderSource.xml] : []))), shouldAutoCalculate,
    workbookSnapshot: includeSnapshot ? nextWorkbook.exportSnapshot() : undefined
  };
}

function respond(message: WorkerResponse) {
  const buffers = new Set<ArrayBuffer>();
  const seen = new Set<object>();
  const visit = (value: unknown) => {
    if (!value || typeof value !== "object" || seen.has(value)) return;
    seen.add(value);
    if (ArrayBuffer.isView(value) && value.buffer instanceof ArrayBuffer) buffers.add(value.buffer);
    else if (value instanceof ArrayBuffer) buffers.add(value);
    else if (value instanceof Map || value instanceof Set) for (const entry of value.values()) visit(entry);
    else for (const child of Object.values(value)) visit(child);
  };
  visit(message);
  (self as unknown as { postMessage: (message: WorkerResponse, transfer: Transferable[]) => void }).postMessage(message, [...buffers]);
}

async function handleMessage(message: WorkerRequest) {
  switch (message.type) {
    case "renderChart": {
      const [chart, source, width, height] = message.payload;
      const [{ loadOfficeCharts }, { renderChartInWorker }] = await Promise.all([
        source.experimental ? import("@extend-ai/rust-office-charts/full") : import("@extend-ai/rust-office-charts"),
        import("./chart-worker-render")
      ]);
      await loadSpreadsheetFonts(workbookFontFamilies, globalThis as unknown as Document);
      const renderer = await loadOfficeCharts(source.experimental && typeof DecompressionStream !== "function" ? new URL("./office_charts_full_bg.wasm", import.meta.url) : undefined);
      try { return renderChartInWorker(renderer, chart, source, width, height); }
      finally { takeWorkbookXmlRecords(); }
    }
    case "export": {
      const wasmModule = await getSheetsWasmModule();
      configureWorkbookXmlDecoder(wasmModule.parseXmlTree);
      configureWorkbookArchiveDecoder(wasmModule.decodeArchive);
      installWorkerXmlAdapters();
      const restored = wasmModule.Workbook.fromSnapshot(message.payload.snapshot);
      try {
        const bytes = restored.saveXlsxBytes();
        const archive = wasmModule.decodeArchive(bytes) as Record<string, Uint8Array>;
        const styles = archive["xl/styles.xml"];
        if (styles) archive["xl/styles.xml"] = strToU8(strFromU8(styles).replace(/&amp;quot;/g, "&quot;").replace(/&amp;apos;/g, "&apos;"));
        const result = mergeWorkbookImageAssets(bytes, message.payload.assets, message.payload.origins, archive);
        return result;
      } finally { restored.free(); takeWorkbookXmlRecords(); }
    }
    case "load": {
      if (message.payload.wasmSource !== undefined) {
        setWasmSource(message.payload.wasmSource);
      }
      restoreColumnFontMeasurements(message.payload.columnFontMeasurements ?? []);
      return loadWorkbook(message.payload.buffer, message.payload.skipXmlParsing, message.payload.showHiddenSheets, message.payload.includeSnapshot);
    }
    case "getCellSnapshot": {
      if (!workbook) {
        return {
          displayValue: "",
          formula: ""
        };
      }

      const targetSheet = sheets.find((sheet) => sheet.workbookSheetIndex === message.payload.workbookSheetIndex) ?? null;
      const worksheet = workbook.getSheet(message.payload.workbookSheetIndex);
      return {
        displayValue: localizeCellNumber(worksheet, message.payload.row, message.payload.col, getCellDisplayValue(worksheet, message.payload.row, message.payload.col, targetSheet), message.payload.locale),
        formula: worksheet.getFormulaAt(message.payload.row, message.payload.col) ?? ""
      };
    }
    case "getRowsBatch": {
      if (!workbook) {
        return null;
      }

      const worksheet = workbook.getSheet(message.payload.workbookSheetIndex) as ReturnType<Workbook["getSheet"]> & {
        getRowsBatch?: (startRow: number, maxRows: number, options?: unknown) => unknown;
      };
      if (typeof worksheet.getRowsBatch !== "function") {
        return null;
      }

      const rows = worksheet.getRowsBatch(message.payload.startRow, message.payload.rowCount, {
        includeFormulas: true,
        includeHyperlinks: true,
        includeMergeInfo: true,
        includeStyles: true,
        useFormattedValues: true
      }) as Array<{ index: number; cells: Array<{ col: number; isMergedSecondary?: boolean; generalAlignment?: string }> }> | null;
      for (const row of rows ?? []) {
        for (const cell of row.cells) {
          if (!cell.isMergedSecondary) {
            cell.generalAlignment = getCellGeneralAlignment(worksheet, row.index, cell.col);
          }
        }
      }
      localizeRows(worksheet, rows, message.payload.locale);
      return rows;
    }
    default:
      return null;
  }
}

self.addEventListener("message", (event: MessageEvent<WorkerRequest>) => {
  const message = event.data;
  void handleMessage(message)
    .then((result) => {
      respond({
        id: message.id,
        result,
        success: true
      });
    })
    .catch((error: unknown) => {
      respond({
        error: error instanceof Error ? error.message : "Worker request failed.",
        id: message.id,
        success: false
      } satisfies WorkerErrorResponse);
    });
});

self.postMessage({ type: "ready" });
