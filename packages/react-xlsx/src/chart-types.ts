import type { XlsxChart } from "./types";

export interface ChartTypeOption {
  id: string;
  label: string;
  family: string;
  chartType: string;
  node: string;
  is3d?: boolean;
  grouping?: string;
  barDir?: string;
  marker?: boolean;
  scatterStyle?: string;
  radarStyle?: string;
  ofPieType?: string;
  wireframe?: boolean;
  bubble3d?: boolean;
  stockSeries?: number;
  combo?: boolean;
  layout?: string;
}

const options: ChartTypeOption[] = [];
function add(family: string, id: string, label: string, chartType: string, node: string, rest: Partial<ChartTypeOption> = {}) {
  options.push({ family, id, label, chartType, node, ...rest });
}
for (const family of ["Column", "Bar"]) {
  for (const is3d of [false, true]) {
    for (const [suffix, label, grouping] of [["Clustered", "Clustered", "clustered"], ["Stacked", "Stacked", "stacked"], ["PercentStacked", "100% Stacked", "percentStacked"]]) {
      add(family, `${family}${suffix}${is3d ? "3D" : ""}`, `${is3d ? "3-D " : ""}${label} ${family}`, `${family}${suffix}`, is3d ? "bar3DChart" : "barChart", { grouping, is3d, barDir: family === "Bar" ? "bar" : "col" });
    }
  }
}
for (const family of ["Line", "Area"]) {
  for (const [suffix, label, grouping] of [["", family, "standard"], ["Stacked", `Stacked ${family}`, "stacked"], ["PercentStacked", `100% Stacked ${family}`, "percentStacked"]]) {
    add(family, `${family}${suffix}`, label, `${family}${suffix}`, `${family.toLowerCase()}Chart`, { grouping, marker: false });
    if (family === "Line") add(family, `${family}${suffix}Markers`, `${label} with Markers`, `${family}${suffix}`, "lineChart", { grouping, marker: true });
    if (family === "Area" || suffix === "") add(family, `${family}${suffix}3D`, `3-D ${label}`, `${family}${suffix}`, `${family.toLowerCase()}3DChart`, { grouping, is3d: true });
  }
}
add("Pie", "Pie", "Pie", "Pie", "pieChart");
add("Pie", "Pie3D", "3-D Pie", "Pie3D", "pie3DChart", { is3d: true });
add("Pie", "PieOfPie", "Pie of Pie", "BarOfPie", "ofPieChart", { ofPieType: "pie" });
add("Pie", "BarOfPie", "Bar of Pie", "BarOfPie", "ofPieChart", { ofPieType: "bar" });
add("Pie", "Doughnut", "Doughnut", "Doughnut", "doughnutChart");
for (const [id, label, scatterStyle, marker] of [
  ["Scatter", "Scatter", "marker", true],
  ["ScatterSmoothMarkers", "Scatter with Smooth Lines and Markers", "smoothMarker", true],
  ["ScatterSmooth", "Scatter with Smooth Lines", "smooth", false],
  ["ScatterLinesMarkers", "Scatter with Straight Lines and Markers", "lineMarker", true],
  ["ScatterLines", "Scatter with Straight Lines", "line", false]
] as const) add("XY (Scatter)", id, label, id.startsWith("ScatterSmooth") ? "ScatterSmooth" : id.startsWith("ScatterLines") ? "ScatterLines" : "Scatter", "scatterChart", { scatterStyle, marker });
add("XY (Scatter)", "Bubble", "Bubble", "Bubble", "bubbleChart", { bubble3d: false });
add("XY (Scatter)", "Bubble3D", "3-D Bubble", "Bubble", "bubbleChart", { bubble3d: true });
for (const [id, label, count] of [["StockHLC", "High-Low-Close", 3], ["StockOHLC", "Open-High-Low-Close", 4], ["StockVHLC", "Volume-High-Low-Close", 4], ["StockVOHLC", "Volume-Open-High-Low-Close", 5]] as const) {
  add("Stock", id, label, "Stock", "stockChart", { stockSeries: count });
}
for (const is3d of [true, false]) for (const wireframe of [false, true]) {
  add("Surface", `Surface${is3d ? "3D" : "Contour"}${wireframe ? "Wireframe" : ""}`, `${wireframe ? "Wireframe " : ""}${is3d ? "3-D Surface" : "Contour"}`, "Surface", is3d ? "surface3DChart" : "surfaceChart", { is3d, wireframe });
}
for (const [id, label, radarStyle] of [["Radar", "Radar", "standard"], ["RadarMarkers", "Radar with Markers", "marker"], ["RadarFilled", "Filled Radar", "filled"]]) {
  add("Radar", id, label, "Radar", "radarChart", { radarStyle, marker: radarStyle === "marker" });
}
add("Combo", "ColumnLine", "Clustered Column – Line", "ColumnClustered", "barChart", { grouping: "clustered", barDir: "col", combo: true });
for (const [family, id, layout] of [["Treemap", "Treemap", "treemap"], ["Sunburst", "Sunburst", "sunburst"], ["Histogram", "Histogram", "clusteredColumn"], ["Histogram", "Pareto", "paretoLine"], ["Box & Whisker", "BoxWhisker", "boxWhisker"], ["Waterfall", "Waterfall", "waterfall"], ["Funnel", "Funnel", "funnel"], ["Map", "RegionMap", "regionMap"]]) {
  add(family, id, id === "BoxWhisker" ? "Box & Whisker" : id === "RegionMap" ? "Filled Map" : id, id, "", { layout });
}

export const chartTypeOptions: readonly ChartTypeOption[] = options;

export function chartTypeUnavailable(chart: XlsxChart, option: ChartTypeOption): boolean {
  if (chart.editable === false || !chart.chartPath) return true;
  if (chart.raw?.pivotSource) return !["ColumnClustered", "ColumnStacked", "ColumnPercentStacked", "BarClustered", "BarStacked", "BarPercentStacked", "Line", "LineStacked", "Area", "AreaStacked", "AreaPercentStacked", "Pie", "Doughnut"].includes(option.id);
  if (option.stockSeries && chart.series.length !== option.stockSeries) return true;
  if (option.combo && chart.series.length < 2) return true;
  return false;
}

export function matchesChartType(chart: XlsxChart, option: ChartTypeOption) {
  if (option.layout) return chart.chartExLayout === option.layout;
  const groups = chart.typeGroups ?? [];
  const volumeStock = groups.some(group => group.chartType === "Stock") && groups.some(group => group.chartType.startsWith("Column"));
  if (option.stockSeries) return (chart.chartType === "Stock" || volumeStock) && chart.series.length === option.stockSeries
    && volumeStock === (option.id === "StockVHLC" || option.id === "StockVOHLC");
  if (volumeStock) return false;
  return chart.chartType === option.chartType && Boolean(chart.is3d) === Boolean(option.is3d)
    && (!option.scatterStyle || chart.scatterStyle === option.scatterStyle)
    && (!option.radarStyle || chart.radarStyle === option.radarStyle)
    && (option.wireframe === undefined || Boolean(chart.wireframe) === option.wireframe)
    && (option.bubble3d === undefined || Boolean(chart.bubble3d) === option.bubble3d)
    && (option.marker === undefined || option.node === "scatterChart" || chart.series.some(series => series.marker?.symbol && series.marker.symbol !== "none") === option.marker)
    && (option.ofPieType === undefined || chart.raw?.ofPieType === option.ofPieType)
    && Boolean(chart.typeGroups && chart.typeGroups.length > 1) === Boolean(option.combo);
}
