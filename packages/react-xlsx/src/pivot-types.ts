import type { XlsxCellAddress } from "./types";

export type XlsxPivotAggregation = "sum" | "count" | "countNums" | "average" | "min" | "max" | "product" | "var" | "varp" | "stdDev" | "stdDevp";
export type XlsxPivotShowAs = "normal" | "percentOfGrandTotal" | "percentOfRowTotal" | "percentOfColumnTotal" | "percentOfParentRowTotal" | "percentOfParentColumnTotal" | "difference" | "percentDifference" | "percent" | "runningTotal" | "percentRunningTotal" | "rankAscending" | "rankDescending" | "index";
export interface XlsxPivotSource { sheet: string; range: string; table?: string | null }
export interface XlsxPivotTarget { sheet: string; cell: string }
export interface XlsxPivotMember { key: string; caption: string }
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
  field: number;
  name?: string | null;
  sort: "ascending" | "descending" | "manual" | "valueAscending" | "valueDescending";
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
  groupingUnit?: string | null;
  kind: "members" | "label" | "value" | "top10";
  values: string[];
  operator: "equal" | "notEqual" | "contains" | "notContains" | "beginsWith" | "notBeginsWith" | "endsWith" | "notEndsWith" | "greaterThan" | "greaterThanOrEqual" | "lessThan" | "lessThanOrEqual" | "between" | "notBetween";
  value: string;
  value2: string;
  measure: number;
  top: boolean;
  count: number;
  percent: boolean;
}
export interface XlsxPivotValueField {
  field: number;
  sourceName?: string | null;
  name: string;
  aggregation: XlsxPivotAggregation;
  showAs: XlsxPivotShowAs;
  baseField?: number | null;
  baseGroupingUnit?: string | null;
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
  preserveFormatting?: boolean;
  rowGrandTotals: boolean;
  columnGrandTotals: boolean;
  subtotalsPosition: "top" | "bottom" | "none";
  styleName: string;
  fields: XlsxPivotSourceField[];
  outputRange?: string | null;
  unsupportedFeatures: string[];
  refreshOnLoad: boolean;
  refreshError?: string;
  calculatedItems: XlsxPivotCalculatedItem[];
  charts?: XlsxPivotChart[];
  toggles?: XlsxPivotToggle[];
  headers?: XlsxPivotHeader[];
}
export type XlsxPivotTableInput = Pick<XlsxPivotTable, "source" | "target"> & Partial<Omit<XlsxPivotTable, "source" | "target">>;
export interface XlsxPivotDrillDown { fields: string[]; rows: (string | number | boolean | null)[][] }

export interface XlsxViewerPivots {
  pivotTables: XlsxPivotTable[];
  pivotFieldListVisible: boolean;
  setPivotFieldListVisible: (visible: boolean) => void;
  getPivotSourceFields: (source: XlsxPivotSource) => XlsxPivotSourceField[];
  getPivotTableAtCell: (cell: XlsxCellAddress, sheetName?: string) => XlsxPivotTable | null;
  upsertPivotTable: (input: XlsxPivotTableInput) => XlsxPivotTable;
  refreshPivotTable: (id: string) => XlsxPivotTable;
  refreshAllPivotTables: () => XlsxPivotTable[];
  removePivotTable: (id: string) => boolean;
  pivotDrillDown: (id: string, row: number, col: number) => XlsxPivotDrillDown;
  showPivotDetails: (id: string, row: number, col: number) => string;
  createPivotChart: (id: string, options?: XlsxPivotChartInput) => XlsxPivotChart;
  togglePivotGroup: (id: string, toggle: XlsxPivotToggle) => XlsxPivotTable;
}
