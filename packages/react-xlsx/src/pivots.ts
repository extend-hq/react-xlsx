import type { Workbook } from "./duke-sheets/generated/duke_sheets_wasm.js";
import type { DrawingAnchor } from "./duke-sheets/generated/duke_sheets_wasm.js";
import type { XlsxPivotDrillDown, XlsxPivotSource, XlsxPivotSourceField, XlsxPivotTable, XlsxPivotTableInput } from "./pivot-types";
import type { XlsxPivotChart, XlsxPivotChartInput } from "./pivot-types";
import type { XlsxCellRange } from "./types";
import { parseXlsxRange } from "./range-reference";

export interface PivotWorkbook {
  readonly pivotTables: XlsxPivotTable[];
  getPivotSourceFields(source: XlsxPivotSource): XlsxPivotSourceField[];
  upsertPivotTable(input: XlsxPivotTableInput): XlsxPivotTable;
  refreshPivotTable(id: string): XlsxPivotTable;
  removePivotTable(id: string): boolean;
  pivotDrillDown(id: string, row: number, col: number): XlsxPivotDrillDown;
  createPivotChart(id: string, input: XlsxPivotChartInput & { workbookName: string }): XlsxPivotChart;
  updatePivotChart(sheetName: string, drawingName: string, input: { title?: string; chartType?: string; anchor?: DrawingAnchor }): void;
  modifiedPivotCharts(): Array<{ sheet: string; name: string; sourceName?: string | null; formatId?: number | null }>;
}

export function pivotWorkbook(workbook: Workbook): PivotWorkbook {
  const candidate = workbook as unknown as PivotWorkbook;
  if (typeof candidate.upsertPivotTable !== "function") {
    throw new Error("Pivot tables require the current bundled spreadsheet runtime.");
  }
  return candidate;
}

export function readPivotTables(workbook: Workbook): XlsxPivotTable[] {
  const pivots = (workbook as unknown as Partial<PivotWorkbook>).pivotTables;
  return Array.isArray(pivots) ? pivots : [];
}

export function findIntersectingPivot(pivots: readonly XlsxPivotTable[], sheetName: string | undefined, range: XlsxCellRange) {
  if (!sheetName) return null;
  const top = Math.min(range.start.row, range.end.row);
  const bottom = Math.max(range.start.row, range.end.row);
  const left = Math.min(range.start.col, range.end.col);
  const right = Math.max(range.start.col, range.end.col);
  return pivots.find(pivot => {
    if (pivot.target.sheet.toLowerCase() !== sheetName.toLowerCase() || !pivot.outputRange) return false;
    const bounds = parseXlsxRange(pivot.outputRange)?.range;
    return bounds && top <= bounds.end.row && bottom >= bounds.start.row && left <= bounds.end.col && right >= bounds.start.col;
  }) ?? null;
}

export function pivotFieldHighlightRanges(pivot: XlsxPivotTable): XlsxCellRange[] {
  const bounds = pivot.outputRange && parseXlsxRange(pivot.outputRange)?.range;
  if (!bounds || !pivot.headers?.length) return [];
  const rowHeaders = pivot.headers.filter(header => header.axis === "rows");
  const columnHeaders = pivot.headers.filter(header => header.axis === "columns");
  const pageHeaders = pivot.headers.filter(header => header.axis === "reportFilters");
  const headerTop = bounds.start.row + (pageHeaders.length ? Math.max(...pageHeaders.map(header => header.row)) - bounds.start.row + 2 : 0);
  const lastHeaderRow = rowHeaders.length ? Math.max(...rowHeaders.map(header => header.row))
    : columnHeaders.length ? Math.max(...columnHeaders.map(header => header.row)) + Number(pivot.values.length > 1 && !pivot.valuesOnRows)
    : headerTop;
  const dataCol = rowHeaders.length ? Math.max(...rowHeaders.map(header => header.col)) + 1 + Number(Boolean(pivot.valuesOnRows) && pivot.layout !== "compact")
    : columnHeaders.length ? Math.min(...columnHeaders.map(header => header.col)) : bounds.start.col + 1;
  const ranges: XlsxCellRange[] = pageHeaders.map(header => ({ start: { row: header.row, col: bounds.start.col }, end: { row: header.row, col: header.col } }));
  for (const header of rowHeaders) ranges.push({ start: { row: header.row, col: header.col }, end: { row: bounds.end.row, col: header.col } });
  if (dataCol <= bounds.end.col && headerTop <= lastHeaderRow) ranges.push({ start: { row: headerTop, col: dataCol }, end: { row: lastHeaderRow, col: bounds.end.col } });
  if (dataCol <= bounds.end.col && lastHeaderRow < bounds.end.row) ranges.push({ start: { row: lastHeaderRow + 1, col: dataCol }, end: bounds.end });
  return ranges;
}
