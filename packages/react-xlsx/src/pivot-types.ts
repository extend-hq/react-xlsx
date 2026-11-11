import type { XlsxCellAddress } from "./types";

export type XlsxPivotAggregation = "sum" | "count" | "countNums" | "average" | "min" | "max" | "product" | "var" | "varp" | "stdDev" | "stdDevp";
export type XlsxPivotShowAs = "normal" | "percentOfGrandTotal" | "percentOfRowTotal" | "percentOfColumnTotal" | "percentOfParentRowTotal" | "percentOfParentColumnTotal" | "difference" | "percentDifference" | "percent" | "runningTotal" | "percentRunningTotal" | "rankAscending" | "rankDescending" | "index";
/** Worksheet source range, table, or supported static defined name. */
export interface XlsxPivotSource { sheet: string; range: string; table?: string | null }
/** Existing destination worksheet and its A1-style anchor address. */
export interface XlsxPivotTarget { sheet: string; cell: string }
/** Use the typed key for filtering and base items; the caption is display text. */
export interface XlsxPivotMember { key: string; caption: string }
/** Indexes identify zero-based source columns. */
export interface XlsxPivotSourceField { index: number; name: string; dataType: string; items: XlsxPivotMember[] }
export interface XlsxPivotGrouping {
  kind: "number" | "date" | "manual";
  unit: "years" | "quarters" | "months" | "days" | "hours" | "minutes" | "seconds";
  interval: number;
  start?: number | null;
  end?: number | null;
  groups: { name: string; members: string[] }[];
}
export interface XlsxPivotAxisField {
  /** Zero-based source field index; date hierarchy levels may share a field. */
  field: number;
  name?: string | null;
  sort: "ascending" | "descending" | "manual" | "valueAscending" | "valueDescending";
  /** Zero-based index in the report's values array for value-based sorting. */
  sortMeasure?: number | null;
  items?: XlsxPivotMember[];
  subtotalFunctions?: XlsxPivotAggregation[];
  showItemsWithNoData?: boolean;
  subtotals: boolean;
  collapsed: string[];
  grouping?: XlsxPivotGrouping | null;
}
export interface XlsxPivotFilter {
  field: number;
  /** Identifies the effective grouped level when a source field has multiple levels. */
  groupingUnit?: string | null;
  kind: "members" | "label" | "value" | "top10";
  /** Typed member keys for a member filter. */
  values: string[];
  operator: "equal" | "notEqual" | "contains" | "notContains" | "beginsWith" | "notBeginsWith" | "endsWith" | "notEndsWith" | "greaterThan" | "greaterThanOrEqual" | "lessThan" | "lessThanOrEqual" | "between" | "notBetween";
  value: string;
  value2: string;
  /** Zero-based index in the report's values array. */
  measure: number;
  top: boolean;
  count: number;
  percent: boolean;
}
export interface XlsxPivotValueField {
  field: number;
  /** Stable calculated-field name, independent of its displayed caption. */
  sourceName?: string | null;
  name: string;
  aggregation: XlsxPivotAggregation;
  showAs: XlsxPivotShowAs;
  baseField?: number | null;
  baseGroupingUnit?: string | null;
  /** Typed member key, or "previous"/"next", used by relative Show Values As calculations. */
  baseItem?: string | null;
  numberFormat?: string | null;
  formula?: string | null;
}
export interface XlsxPivotCalculatedItem { field: number; name: string; formula: string }
export type XlsxPivotChartType = "ColumnClustered" | "ColumnStacked" | "ColumnPercentStacked" | "BarClustered" | "BarStacked" | "BarPercentStacked" | "Line" | "LineStacked" | "Area" | "AreaStacked" | "AreaPercentStacked" | "Pie" | "Doughnut";
export interface XlsxPivotChartInput {
  sheet?: string;
  cell?: string;
  chartType?: XlsxPivotChartType;
  title?: string;
  width?: number;
  height?: number;
  measure?: number | null;
}
export interface XlsxPivotChart extends XlsxPivotChartInput {
  sheet: string;
  drawingName: string;
  chartType: XlsxPivotChartType;
  title: string;
  workbookName: string;
  formatId: number;
  measure?: number | null;
}
/** Expansion control with zero-based worksheet coordinates. */
export interface XlsxPivotToggle {
  row: number;
  col: number;
  axis: "rows" | "columns";
  field: number;
  depth?: number;
  key: string;
  collapsed: boolean;
  caption: string;
}
export interface XlsxPivotHeaderField {
  field: number;
  depth: number;
  name: string;
  groupingUnit?: string | null;
}
/** Sort/filter header with zero-based worksheet coordinates. */
export interface XlsxPivotHeader {
  row: number;
  col: number;
  axis: "rows" | "columns" | "reportFilters";
  caption: string;
  fields: XlsxPivotHeaderField[];
}
export interface XlsxPivotTable {
  id: string;
  name: string;
  source: XlsxPivotSource;
  target: XlsxPivotTarget;
  rows: XlsxPivotAxisField[];
  columns: XlsxPivotAxisField[];
  filters: XlsxPivotFilter[];
  reportFilters: number[];
  values: XlsxPivotValueField[];
  layout: "compact" | "outline" | "tabular";
  valuesOnRows?: boolean;
  repeatLabels: boolean;
  /** Preserve existing report formatting when rebuilding results. Defaults to true. */
  preserveFormatting?: boolean;
  rowGrandTotals: boolean;
  columnGrandTotals: boolean;
  subtotalsPosition: "top" | "bottom" | "none";
  styleName: string;
  fields: XlsxPivotSourceField[];
  outputRange?: string | null;
  /** Reasons the report can be viewed but cannot be edited or refreshed. */
  unsupportedFeatures: string[];
  /** Refresh from worksheet source data when the workbook opens. */
  refreshOnLoad: boolean;
  /** Automatic refresh failure; saved results remain available. */
  refreshError?: string;
  calculatedItems: XlsxPivotCalculatedItem[];
  charts?: XlsxPivotChart[];
  toggles?: XlsxPivotToggle[];
  headers?: XlsxPivotHeader[];
}
/** Supply an existing id to update a report; omit it to create one. */
export type XlsxPivotTableInput = Pick<XlsxPivotTable, "source" | "target"> & Partial<Omit<XlsxPivotTable, "source" | "target">>;
export interface XlsxPivotDrillDown { fields: string[]; rows: (string | number | boolean | null)[][] }

export interface XlsxViewerPivots {
  /** Workbook-wide report metadata, including reports on inactive worksheets. */
  pivotTables: XlsxPivotTable[];
  /** Shared visibility for default/custom field lists and active-report guides. */
  pivotFieldListVisible: boolean;
  /** Explicitly show or hide the field list; hiding suppresses automatic reopening. */
  setPivotFieldListVisible: (visible: boolean) => void;
  /** Inspect source columns and typed member keys without creating a report. */
  getPivotSourceFields: (source: XlsxPivotSource) => XlsxPivotSourceField[];
  /** Find a report at zero-based worksheet coordinates; defaults to the active sheet. */
  getPivotTableAtCell: (cell: XlsxCellAddress, sheetName?: string) => XlsxPivotTable | null;
  /** Create or update a report atomically. Invalid or disallowed mutations throw. */
  upsertPivotTable: (input: XlsxPivotTableInput) => XlsxPivotTable;
  /** Rebuild the saved cache and report from source data as an undoable mutation. */
  refreshPivotTable: (id: string) => XlsxPivotTable;
  /** Refresh reports in one transaction; dependency chains are not reordered. */
  refreshAllPivotTables: () => XlsxPivotTable[];
  /** Remove a report and detach linked PivotCharts; returns whether it was found. */
  removePivotTable: (id: string) => boolean;
  /** Read source records behind a result at zero-based worksheet coordinates. */
  pivotDrillDown: (id: string, row: number, col: number) => XlsxPivotDrillDown;
  /** Create an undoable details worksheet and return its name. Coordinates are zero-based. */
  showPivotDetails: (id: string, row: number, col: number) => string;
  /** Create a linked chart whose data stays synchronized with the report. */
  createPivotChart: (id: string, options?: XlsxPivotChartInput) => XlsxPivotChart;
  /** Expand or collapse a group using a toggle from the report's metadata. */
  togglePivotGroup: (id: string, toggle: XlsxPivotToggle) => XlsxPivotTable;
}
