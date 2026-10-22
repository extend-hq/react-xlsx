import { hierarchy as d3Hierarchy, partition as d3Partition, treemap as d3Treemap, treemapBinary, treemapDice, treemapSquarify } from "d3-hierarchy";
import * as React from "react";
import { shadeSurfaceMatteColor, surfaceNormal, transformSurfaceNormal } from "./surface-lighting";
import { clipSurfaceBand, normalizeSurfaceValue, surfaceHeightRatio, surfaceBandEdges, triangulateSurfaceQuad } from "./surface-geometry";
import { scaleBand, scaleLinear, scalePoint } from "d3-scale";
import {
  arc as d3Arc,
  area as d3Area,
  curveCatmullRom,
  curveLinear,
  curveLinearClosed,
  line as d3Line,
  pie as d3Pie
} from "d3-shape";
import type { CurveFactory } from "d3-shape";
import { regionMapResources, surfaceRenderer } from "./chart-loader";
import { useChartModule } from "./use-chart-module";
import type { RegionMapResources, RegionMapFeature } from "./region-map-geography";
import type { HierarchyNode, HierarchyRectangularNode } from "d3-hierarchy";
import type { FeatureCollection, Geometry } from "geojson";
import type { XlsxChart, XlsxChartAxis, XlsxChartElementSelection, XlsxChartSeries, XlsxChartTypeGroup, XlsxImageRect } from "./types";

type ChartRendererPalette = {
  border: string;
  mutedText: string;
  surface: string;
  text: string;
};

export type ChartSvgProps = {
  chart: XlsxChart;
  onChartElementDoubleClick?: (selection: XlsxChartElementSelection, event: React.MouseEvent<SVGSVGElement>) => void;
  onChartElementPointerDown?: (selection: XlsxChartElementSelection, event: React.PointerEvent<SVGSVGElement>) => void;
  palette: ChartRendererPalette;
  rect: XlsxImageRect;
  selectedChartElement?: XlsxChartElementSelection | null;
};

type RegionMapPlotProps = Pick<ChartSvgProps, "chart" | "palette" | "selectedChartElement"> & { layout: ChartLayout };

function RegionMapPlot({ chart, palette, layout, selectedChartElement }: RegionMapPlotProps) {
  const resources = useChartModule(regionMapResources);
  return resources
    ? renderRegionMapChart(chart, palette, layout, selectedChartElement, resources)
    : null;
}

function SurfaceChartComposite({
  loadingFallback,
  ...props
}: React.ComponentProps<typeof import("./surface-regl").MemoSurfaceChartComposite> & { loadingFallback: React.ReactNode }) {
  const renderer = useChartModule(surfaceRenderer);
  return renderer ? <renderer.MemoSurfaceChartComposite {...props} /> : <>{loadingFallback}</>;
}

type LegendItem = {
  color: string;
  label: string;
};

type PlotRect = {
  height: number;
  left: number;
  top: number;
  width: number;
};

type ChartLayout = {
  height: number;
  legendBounds?: PlotRect;
  legendItems: LegendItem[];
  legendPosition: string | undefined;
  plot: PlotRect;
  titleHeight: number;
  width: number;
};

type BarRect = {
  capEnd?: boolean;
  capStart?: boolean;
  categoryIndex: number;
  bottomScale?: number;
  color: string;
  depthOrder?: number;
  depthOffsetX?: number;
  depthOffsetY?: number;
  depthX?: number;
  depthY?: number;
  gradientId?: string;
  height: number;
  invertedNegative?: boolean;
  isHorizontal: boolean;
  key: string;
  left: number;
  shape3d?: string;
  seriesIndex: number;
  stroke: string;
  strokeWidth: number;
  topScale?: number;
  value: number;
  width: number;
  top: number;
};

type ChartElementDataOptions = {
  selectionMode?: "seriesFirst";
};

function chartElementDataProps(seriesIndex: number, pointIndex?: number, options?: ChartElementDataOptions) {
  return {
    "data-xlsx-chart-point-index": typeof pointIndex === "number" ? String(pointIndex) : undefined,
    "data-xlsx-chart-selection-mode": options?.selectionMode,
    "data-xlsx-chart-series-index": String(seriesIndex),
    style: {
      cursor: "pointer",
      pointerEvents: "all"
    } as React.CSSProperties
  };
}

function barChartElementDataProps(seriesIndex: number, pointIndex: number) {
  return chartElementDataProps(seriesIndex, pointIndex, { selectionMode: "seriesFirst" });
}

function resolveChartElementTarget(target: EventTarget | null) {
  if (!(target instanceof Element)) {
    return null;
  }

  return target.closest("[data-xlsx-chart-series-index]") as HTMLElement | SVGElement | null;
}

function resolveChartSelectionFromTarget(
  chart: XlsxChart,
  target: EventTarget | null,
  selectedChartElement: XlsxChartElementSelection | null | undefined
): XlsxChartElementSelection | null {
  const element = resolveChartElementTarget(target);
  if (!element) {
    return selectedChartElement?.chartId === chart.id
      ? { chartId: chart.id, kind: "chart" }
      : { chartId: chart.id, kind: "chart" };
  }

  const seriesIndex = Number(element.dataset.xlsxChartSeriesIndex);
  if (!Number.isInteger(seriesIndex) || seriesIndex < 0 || seriesIndex >= chart.series.length) {
    return null;
  }

  const series = chart.series[seriesIndex];
  if (!series) {
    return null;
  }

  const rawPointIndex = element.dataset.xlsxChartPointIndex;
  const rawElementKind = element.dataset.xlsxChartElementKind;
  const selectionMode = element.dataset.xlsxChartSelectionMode;
  const pointIndex = rawPointIndex == null || rawPointIndex === "" ? null : Number(rawPointIndex);
  const hasPoint = pointIndex != null && Number.isInteger(pointIndex) && pointIndex >= 0;
  const sameSelectedSeries = selectedChartElement?.chartId === chart.id
    && selectedChartElement.kind !== "chart"
    && selectedChartElement.seriesIndex === seriesIndex;
  if (hasPoint && (selectionMode !== "seriesFirst" || sameSelectedSeries)) {
    return {
      chartId: chart.id,
      kind: "point",
      pointIndex,
      seriesId: series.id,
      seriesIndex
    };
  }

  return {
    chartId: chart.id,
    kind: rawElementKind === "legendEntry" ? "legendEntry" : "series",
    seriesId: series.id,
    seriesIndex
  };
}

function isSelectedChartSeries(
  selectedChartElement: XlsxChartElementSelection | null | undefined,
  chartId: string,
  seriesIndex: number
) {
  return selectedChartElement?.chartId === chartId
    && selectedChartElement.kind !== "chart"
    && selectedChartElement.seriesIndex === seriesIndex;
}

function isSelectedChartPoint(
  selectedChartElement: XlsxChartElementSelection | null | undefined,
  chartId: string,
  seriesIndex: number,
  pointIndex: number
) {
  return selectedChartElement?.chartId === chartId
    && selectedChartElement.kind === "point"
    && selectedChartElement.seriesIndex === seriesIndex
    && selectedChartElement.pointIndex === pointIndex;
}

function isSelectedChartPointOrSeries(
  selectedChartElement: XlsxChartElementSelection | null | undefined,
  chartId: string,
  seriesIndex: number,
  pointIndex: number
) {
  return isSelectedChartPoint(selectedChartElement, chartId, seriesIndex, pointIndex)
    || (
      isSelectedChartSeries(selectedChartElement, chartId, seriesIndex)
      && selectedChartElement?.kind !== "point"
    );
}

function renderSelectionRectHandles(
  key: string,
  left: number,
  top: number,
  width: number,
  height: number,
  color = "#64748b"
) {
  return renderSelectionPointHandles(
    key,
    [
      { x: left, y: top },
      { x: left + width, y: top },
      { x: left, y: top + height },
      { x: left + width, y: top + height }
    ],
    color
  );
}

function renderSelectionPointHandles(
  key: string,
  points: Array<{ x: number; y: number }>,
  color = "#64748b"
) {
  const radius = 3;

  return (
    <g key={key} pointerEvents="none">
      {points.map((point, index) => (
        <circle
          cx={point.x}
          cy={point.y}
          fill="#ffffff"
          key={`${key}-handle-${index}`}
          r={radius}
          stroke={color}
          strokeWidth={1.1}
          style={{ filter: "drop-shadow(0 1px 2px rgba(15, 23, 42, 0.18))" }}
          vectorEffect="non-scaling-stroke"
        />
      ))}
    </g>
  );
}

type ChartHierarchyDatum = {
  children?: ChartHierarchyDatum[];
  colorIndex?: number;
  name: string;
  value?: number;
};

type ComboRenderableGroup = {
  axisIds: number[];
  categoryAxis: XlsxChartAxis | null;
  chartType: string;
  gapWidth?: number;
  is3d?: boolean;
  raw?: Record<string, unknown>;
  series: XlsxChartSeries[];
  valueAxis: XlsxChartAxis | null;
};

type ChartStage = {
  color: string;
  isSubtotal: boolean;
  label: string;
  value: number;
};

type BoxWhiskerStats = {
  lowerFence: number;
  lowerWhisker: number;
  max: number;
  mean: number;
  median: number;
  min: number;
  outliers: number[];
  q1: number;
  q3: number;
  upperFence: number;
  upperWhisker: number;
  visiblePoints: number[];
};

type SurfaceDomain = {
  maxValue: number;
  minValue: number;
  safeMax: number;
  ticks: number[];
};

function parseRgbColor(color: string) {
  const match = /^#?([0-9a-f]{6})$/i.exec(color);
  if (!match) {
    return null;
  }
  return {
    blue: Number.parseInt(match[1].slice(4, 6), 16),
    green: Number.parseInt(match[1].slice(2, 4), 16),
    red: Number.parseInt(match[1].slice(0, 2), 16)
  };
}

function mixRgbColor(color: string, mixWith: string, ratio: number) {
  const base = parseRgbColor(color);
  const target = parseRgbColor(mixWith);
  if (!base || !target) {
    return color;
  }
  const clamped = Math.max(0, Math.min(1, ratio));
  const mixChannel = (left: number, right: number) => Math.round(left + (right - left) * clamped);
  return `#${[
    mixChannel(base.red, target.red),
    mixChannel(base.green, target.green),
    mixChannel(base.blue, target.blue)
  ].map((value) => value.toString(16).padStart(2, "0")).join("")}`;
}

function lightenColor(color: string, ratio: number) {
  return mixRgbColor(color, "#ffffff", ratio);
}

function darkenColor(color: string, ratio: number) {
  return mixRgbColor(color, "#000000", ratio);
}

function clamp(value: number, min: number, max: number) {
  return Math.max(min, Math.min(max, value));
}

const DEFAULT_CHART_FONT_STACK = [
  "\"Aptos\"",
  "Calibri",
  "Carlito",
  "\"Segoe UI\"",
  "Tahoma",
  "Arial",
  "sans-serif"
].join(", ");

const DEFAULT_CHART_TEXT_COLOR = "#000000";
const DEFAULT_CHART_MUTED_TEXT_COLOR = "#7f7f7f";

function escapeCssFontFamilyToken(value: string) {
  const trimmed = value.trim();
  if (!trimmed) {
    return "";
  }
  if (
    trimmed.startsWith("\"")
    || trimmed.startsWith("'")
    || /^(serif|sans-serif|monospace|cursive|fantasy|system-ui|math|emoji|fangsong)$/i.test(trimmed)
  ) {
    return trimmed;
  }
  return /\s/.test(trimmed) ? `"${trimmed.replace(/"/g, "\\\"")}"` : trimmed;
}

function buildChartFontFamily(fontFamily: string | undefined) {
  if (!fontFamily || fontFamily.trim().length === 0) {
    return DEFAULT_CHART_FONT_STACK;
  }
  const tokens = fontFamily
    .split(",")
    .map(escapeCssFontFamilyToken)
    .filter((token) => token.length > 0);
  if (tokens.length === 0) {
    return DEFAULT_CHART_FONT_STACK;
  }
  return [...tokens, "\"Segoe UI\"", "Tahoma", "Arial", "sans-serif"].join(", ");
}

function resolveChartTextColor(chart: XlsxChart) {
  return chart.textColor ?? chart.titleColor ?? DEFAULT_CHART_TEXT_COLOR;
}

function resolveChartAxisTextColor(chart: XlsxChart) {
  return chart.axisLabelColor ?? chart.textColor ?? chart.titleColor ?? DEFAULT_CHART_TEXT_COLOR;
}

function resolveChartMutedTextColor(chart: XlsxChart) {
  return chart.textColor ?? chart.axisLabelColor ?? DEFAULT_CHART_MUTED_TEXT_COLOR;
}

function safeNumber(value: unknown) {
  if (typeof value === "number" && Number.isFinite(value)) {
    return value;
  }
  if (typeof value === "string") {
    const normalized = value.trim().replace(/,/g, "");
    if (!normalized) {
      return null;
    }
    const parsed = Number(normalized);
    return Number.isFinite(parsed) ? parsed : null;
  }
  return null;
}

function chartSeriesColor(chart: XlsxChart, seriesIndex: number) {
  const series = chart.series[seriesIndex];
  const paletteColor = chart.chartColorPalette?.[seriesIndex % Math.max(1, chart.chartColorPalette.length)];
  return series?.color ?? series?.lineColor ?? paletteColor ?? chart.textColor ?? "#222222";
}

function chartSeriesStrokeColor(chart: XlsxChart, seriesIndex: number) {
  const series = chart.series[seriesIndex];
  const paletteColor = chart.chartColorPalette?.[seriesIndex % Math.max(1, chart.chartColorPalette.length)];
  return series?.lineColor ?? series?.color ?? paletteColor ?? chart.textColor ?? "#222222";
}

function variesBarColorsByCategory(chart: XlsxChart) {
  return chart.series.length === 1
    && chart.varyColors !== false
    && !chart.chartExLayout
    && ["ColumnClustered", "ColumnStacked", "ColumnPercentStacked", "BarClustered", "BarStacked", "BarPercentStacked"].includes(chart.chartType);
}

function chartPointColor(chart: XlsxChart, pointIndex: number, seriesIndex = 0) {
  const pointStyle = chart.series[seriesIndex]?.dataPointStyles?.find((entry) => entry.index === pointIndex);
  if (pointStyle?.color) {
    return pointStyle.color;
  }
  const rawPoint = chart.series[seriesIndex]?.dataPoints?.[pointIndex];
  if (rawPoint && typeof rawPoint === "object") {
    const pointRecord = rawPoint as Record<string, unknown>;
    if (typeof pointRecord.color === "string") {
      return pointRecord.color;
    }
    if (typeof pointRecord.fillColor === "string") {
      return pointRecord.fillColor;
    }
    if (typeof pointRecord.solidFillHex === "string") {
      const normalized = pointRecord.solidFillHex.replace(/^#/, "");
      if (/^[0-9a-fA-F]{6}$/.test(normalized)) {
        return `#${normalized.toLowerCase()}`;
      }
    }
    const shapeProperties = pointRecord.shapeProperties && typeof pointRecord.shapeProperties === "object"
      ? pointRecord.shapeProperties as Record<string, unknown>
      : null;
    if (shapeProperties && typeof shapeProperties.solidFillHex === "string") {
      const normalized = shapeProperties.solidFillHex.replace(/^#/, "");
      if (/^[0-9a-fA-F]{6}$/.test(normalized)) {
        return `#${normalized.toLowerCase()}`;
      }
    }
  }
  const palette = chart.chartColorPalette;
  if (palette && palette.length > 0) {
    const offset = chart.chartColorPaletteOffset ?? 0;
    return palette[(pointIndex + offset) % palette.length] ?? palette[pointIndex % palette.length];
  }
  return chartSeriesColor(chart, seriesIndex);
}

function isHistogramLikeSeries(series: XlsxChartSeries | null | undefined) {
  const raw = series?.raw && typeof series.raw === "object"
    ? series.raw as Record<string, unknown>
    : null;
  return Array.isArray(raw?.chartExHistogramBins) && raw.chartExHistogramBins.length > 0;
}

function isHistogramLikeChart(chart: XlsxChart) {
  return chart.series.some((series) => isHistogramLikeSeries(series));
}

function resolveRegionMapBaseColor(chart: XlsxChart, seriesIndex: number) {
  return chart.series[seriesIndex]?.color
    ?? chart.series[seriesIndex]?.lineColor
    ?? chart.chartColorPalette?.[0]
    ?? chartSeriesColor(chart, seriesIndex);
}

function resolveRegionMapDataColor(chart: XlsxChart, seriesIndex: number) {
  const pointColor = normalizeRendererHexColor(chartPointColor(chart, 0, seriesIndex));
  if (pointColor) {
    return pointColor;
  }

  const palette = Array.isArray(chart.chartColorPalette) ? chart.chartColorPalette : [];
  if (palette.length > 0) {
    const offset = chart.chartColorPaletteOffset ?? 0;
    const paletteColor = normalizeRendererHexColor(
      palette[((offset % palette.length) + palette.length) % palette.length]
    );
    if (paletteColor) {
      return paletteColor;
    }
  }

  return normalizeRendererHexColor(chart.series[seriesIndex]?.color ?? chart.series[seriesIndex]?.lineColor)
    ?? "#4f81bd";
}

function resolveRegionMapValueColors(series: XlsxChartSeries | null | undefined) {
  const raw = series?.raw && typeof series.raw === "object"
    ? series.raw as Record<string, unknown>
    : null;
  const colors = Array.isArray(raw?.valueColors)
    ? raw.valueColors
      .map((value) => normalizeRendererHexColor(value))
      .filter((value): value is string => Boolean(value))
    : [];
  return colors.length >= 2 ? colors : null;
}

function resolveRegionMapColorStrings(series: XlsxChartSeries | null | undefined) {
  const raw = series?.raw && typeof series.raw === "object"
    ? series.raw as Record<string, unknown>
    : null;
  const values = Array.isArray(raw?.chartExColorStrings)
    ? raw.chartExColorStrings
      .map((value) => (typeof value === "string" ? value.trim() : ""))
    : [];
  return values.some((value) => value.length > 0) ? values : null;
}

function resolveRegionMapLayoutProperties(series: XlsxChartSeries | null | undefined) {
  const raw = series?.raw && typeof series.raw === "object"
    ? series.raw as Record<string, unknown>
    : null;
  return raw?.layoutProperties && typeof raw.layoutProperties === "object"
    ? raw.layoutProperties as Record<string, unknown>
    : null;
}

function resolveRegionMapValueColorFromStops(stops: string[], ratio: number) {
  if (stops.length === 0) {
    return "#4f81bd";
  }
  if (stops.length === 1) {
    return stops[0];
  }
  const clamped = clamp(ratio, 0, 1);
  const scaled = clamped * (stops.length - 1);
  const lowerIndex = Math.floor(scaled);
  const upperIndex = Math.min(stops.length - 1, lowerIndex + 1);
  const mixRatio = scaled - lowerIndex;
  return mixRgbColor(stops[lowerIndex] ?? stops[0], stops[upperIndex] ?? stops[stops.length - 1], mixRatio);
}

function resolveRegionMapValueColor(chart: XlsxChart, seriesIndex: number, ratio: number) {
  const explicitStops = resolveRegionMapValueColors(chart.series[seriesIndex] ?? null);
  if (explicitStops) {
    return resolveRegionMapValueColorFromStops(explicitStops, ratio);
  }
  const baseColor = resolveRegionMapDataColor(chart, seriesIndex);
  return resolveRegionMapValueColorFromStops([
    lightenColor(baseColor, 0.82),
    lightenColor(baseColor, 0.28),
    darkenColor(baseColor, 0.08)
  ], ratio);
}

function resolveRegionMapNoDataColor(chart: XlsxChart, seriesIndex: number) {
  const baseColor = resolveRegionMapBaseColor(chart, seriesIndex);
  return normalizeRendererHexColor(baseColor) ?? chartSeriesColor(chart, seriesIndex);
}

function buildRegionMapLegendItems(chart: XlsxChart): LegendItem[] {
  const primarySeriesIndex = Math.max(0, chart.series.findIndex((series) => series.hidden !== true));
  const categoricalValues = resolveRegionMapColorStrings(chart.series[primarySeriesIndex] ?? null);
  if (categoricalValues) {
    const uniqueValues = Array.from(new Set(categoricalValues.filter(Boolean)));
    return uniqueValues.map((value, index) => ({
      color: chartPointColor(chart, index, primarySeriesIndex),
      label: value
    }));
  }
  const values = (chart.series[primarySeriesIndex]?.values ?? [])
    .map((value) => safeNumber(value))
    .filter((value): value is number => value != null);
  if (values.length === 0) {
    return [];
  }
  const minValue = Math.min(...values);
  const maxValue = Math.max(...values);
  return [minValue, maxValue].map((value, index) => ({
    color: resolveRegionMapValueColor(chart, primarySeriesIndex, index),
    label: formatTickValue(value)
  }));
}

function normalizeBuiltinPieStyleId(styleId: number | undefined) {
  if (typeof styleId !== "number" || !Number.isFinite(styleId)) {
    return null;
  }
  return styleId >= 100 ? styleId - 100 : styleId;
}

function getBuiltinPiePalette(chart: XlsxChart, seriesIndex: number) {
  const normalized = normalizeBuiltinPieStyleId(chart.chartStyleId);
  if (normalized !== 32) {
    return null;
  }
  const baseColor = chart.series[seriesIndex]?.color
    ?? chart.series[seriesIndex]?.lineColor
    ?? chart.chartColorPalette?.[0]
    ?? null;
  if (!baseColor) {
    return null;
  }
  return [
    lightenColor(baseColor, 0.16),
    darkenColor(baseColor, 0.42),
    baseColor,
    darkenColor(baseColor, 0.18),
    lightenColor(baseColor, 0.08),
    darkenColor(baseColor, 0.3)
  ];
}

function resolvePiePointColor(chart: XlsxChart, pointIndex: number, seriesIndex = 0) {
  const pointStyle = chart.series[seriesIndex]?.dataPointStyles?.find((entry) => entry.index === pointIndex);
  if (pointStyle?.color) {
    return pointStyle.color;
  }
  const builtinPalette = getBuiltinPiePalette(chart, seriesIndex);
  if (builtinPalette && builtinPalette.length > 0) {
    return builtinPalette[pointIndex % builtinPalette.length] ?? builtinPalette[0];
  }
  return chartPointColor(chart, pointIndex, seriesIndex);
}

function selectPrimaryPieSeriesIndex(chart: XlsxChart) {
  let bestIndex = 0;
  let bestScore = Number.NEGATIVE_INFINITY;

  chart.series.forEach((series, index) => {
    let positiveCount = 0;
    let total = 0;
    for (const rawValue of series.values) {
      const value = safeNumber(rawValue);
      if (value != null && value > 0) {
        positiveCount += 1;
        total += value;
      }
    }
    const score = positiveCount * 1_000_000 + total;
    if (score > bestScore) {
      bestScore = score;
      bestIndex = index;
    }
  });

  return bestIndex;
}

function chartSeriesBarColors(
  chart: XlsxChart,
  seriesIndex: number,
  value: number,
  negativeFillMode: "chartArea" | "none" | "series"
) {
  const series = chart.series[seriesIndex];
  const defaultFill = chartSeriesColor(chart, seriesIndex);
  const defaultStroke = chartSeriesStrokeColor(chart, seriesIndex);
  const defaultStrokeWidth = typeof series?.lineWidthPx === "number" && Number.isFinite(series.lineWidthPx)
    ? Math.max(0.8, Math.min(4, series.lineWidthPx))
    : 1;
  if (value < 0 && series?.invertIfNegative) {
    const resolvedNegativeFill = negativeFillMode === "none"
      ? "none"
      : series.negativeColor
        ?? (negativeFillMode === "chartArea" ? chart.chartAreaFillColor : undefined)
        ?? defaultFill;
    return {
      fill: resolvedNegativeFill,
      stroke: series.negativeLineColor ?? defaultStroke,
      strokeWidth: defaultStrokeWidth
    };
  }
  return {
    fill: defaultFill,
    stroke: defaultStroke,
    strokeWidth: defaultStrokeWidth
  };
}

function resolveCategoryBandPadding(gapWidth: number | undefined) {
  const normalizedGap = typeof gapWidth === "number" && Number.isFinite(gapWidth)
    ? clamp(gapWidth, 0, 500)
    : 150;
  const inner = clamp(normalizedGap / (100 + normalizedGap), 0.05, 0.88);
  const outer = clamp(inner * 0.5, 0, 0.45);
  return { inner, outer };
}

function normalizeLegendPosition(position: string | undefined) {
  switch (position) {
    case "bottom":
      return "bottom";
    case "left":
      return "left";
    case "right":
      return "right";
    case "top":
      return "top";
    case "b":
      return "bottom";
    case "l":
      return "left";
    case "r":
      return "right";
    case "t":
      return "top";
    default:
      return position;
  }
}

function normalizeChartMarkerSymbol(value: string | undefined) {
  if (!value || value === "none") {
    return "none";
  }
  if (value === "auto") {
    return "circle";
  }
  return value;
}

function normalizeRenderableChartType(chart: XlsxChart) {
  if (chart.chartType === "ScatterSmooth") {
    return "ScatterSmooth";
  }
  if (chart.chartType === "Pie" && chart.is3d) {
    return "Pie3D";
  }
  if (
    chart.chartType === "Pie"
    && chart.series.some((series) => Array.isArray(series.dataPoints) && series.dataPoints.some((point) => (
      point != null
      && typeof point === "object"
      && "explosion" in point
      && typeof (point as { explosion?: unknown }).explosion === "number"
      && ((point as { explosion?: number }).explosion ?? 0) > 0
    )))
  ) {
    return "PieExploded";
  }
  if (chart.chartType === "Unsupported(c:ofPieChart)") {
    return "BarOfPie";
  }
  return chart.chartType;
}

function normalizeCategoryLabel(value: unknown) {
  if (value == null) {
    return "";
  }
  return String(value)
    .replace(/[\u200B-\u200D\uFEFF]/g, "")
    .trim();
}

function estimateReferencePointCount(formula: string | undefined) {
  if (!formula) {
    return 0;
  }
  const bang = formula.lastIndexOf("!");
  const rawRange = (bang >= 0 ? formula.slice(bang + 1) : formula).replace(/\$/g, "");
  const match = /^([A-Za-z]+)(\d+):([A-Za-z]+)(\d+)$/.exec(rawRange.trim());
  if (!match) {
    return 0;
  }
  const startRow = Number(match[2]);
  const endRow = Number(match[4]);
  if (!Number.isFinite(startRow) || !Number.isFinite(endRow)) {
    return 0;
  }
  return Math.abs(endRow - startRow) + 1;
}

function getCategoryLabels(chart: XlsxChart) {
  const primaryCategories = chart.series[0]?.categories ?? [];
  const includeReferenceCategoryCount = chart.plotVisibleOnly !== true;
  const referenceCategoryCount = includeReferenceCategoryCount
    ? Math.max(
      0,
      ...chart.series.map((series) => Math.max(
        estimateReferencePointCount(series.categoriesRef?.formula),
        estimateReferencePointCount(series.valuesRef?.formula)
      ))
    )
    : 0;
  const categoryCount = Math.max(
    referenceCategoryCount,
    primaryCategories.length,
    ...chart.series.map((series) => Math.max(series.categories.length, series.values.length))
  );
  if (categoryCount <= 0) {
    return [];
  }
  const hasAnyExplicitCategory = chart.series.some((series) => (
    series.categories.some((value) => normalizeCategoryLabel(value).length > 0)
  ));
  const fallbackToImplicitOrdinal = chart.series.some((series) => {
    const categoriesLength = series.categories.length;
    if (categoriesLength === 0) {
      return false;
    }
    return series.categories.every((value) => normalizeCategoryLabel(value).length === 0);
  }) || !hasAnyExplicitCategory;
  return Array.from({ length: categoryCount }, (_, categoryIndex) => {
    const primary = primaryCategories[categoryIndex];
    if (primary != null) {
      const normalizedPrimary = normalizeCategoryLabel(primary);
      if (normalizedPrimary.length > 0) {
        return formatCategoryLabel(chart, primary, normalizedPrimary);
      }
    }
    const fallback = chart.series
      .map((series) => series.categories[categoryIndex])
      .find((value) => normalizeCategoryLabel(value).length > 0);
    if (fallback != null) {
      return formatCategoryLabel(chart, fallback, normalizeCategoryLabel(fallback));
    }
    return fallbackToImplicitOrdinal ? String(categoryIndex + 1) : "";
  });
}

function isComboChart(chart: XlsxChart) {
  const typeGroups = chart.typeGroups ?? [];
  if (typeGroups.length < 2) {
    return false;
  }
  const distinctChartTypes = new Set(typeGroups.map((group) => group.chartType));
  return distinctChartTypes.size > 1;
}

function getComboLegendSeries(chart: XlsxChart) {
  if (!isComboChart(chart)) {
    return chart.series.map((series, index) => ({
      color: chartSeriesColor(chart, index),
      label: series.name ?? `Series ${index + 1}`
    }));
  }
  return (chart.typeGroups ?? []).flatMap((group) => (
    group.series.map((series, seriesIndex) => ({
      color: series.lineColor ?? series.markerColor ?? series.color ?? chartSeriesColor(chart, seriesIndex),
      label: series.name ?? `Series ${seriesIndex + 1}`
    }))
  ));
}

function findAxisForGroup(
  chart: XlsxChart,
  axisIds: number[],
  positionMatcher: (position: string | undefined) => boolean,
  allowAnyMatch = false
) {
  const positionedMatch = chart.axes.find((axis) => (
    axis.id != null
    && axisIds.includes(axis.id)
    && positionMatcher(axis.position)
  ));
  if (positionedMatch) {
    return positionedMatch;
  }
  if (!allowAnyMatch) {
    return null;
  }
  return chart.axes.find((axis) => axis.id != null && axisIds.includes(axis.id)) ?? null;
}

function buildComboGroups(chart: XlsxChart): ComboRenderableGroup[] {
  return (chart.typeGroups ?? []).map((group) => {
    const axisIds = group.axisIds ?? [];
    const categoryAxis = findAxisForGroup(chart, axisIds, (position) => position === "b" || position === "t")
      ?? chart.categoryAxis
      ?? null;
    const valueAxis = findAxisForGroup(chart, axisIds, (position) => position === "l" || position === "r", true)
      ?? chart.valueAxis
      ?? null;
    return {
      axisIds,
      categoryAxis,
      chartType: group.chartType,
      gapWidth: group.gapWidth,
      is3d: group.is3d,
      raw: group.raw,
      series: group.series,
      valueAxis
    };
  }).filter((group) => group.series.length > 0);
}

function resolveRenderableSeriesValue(rawValue: unknown, displayBlanksAs: string | undefined) {
  const numeric = safeNumber(rawValue);
  if (numeric != null) {
    return numeric;
  }
  return displayBlanksAs === "zero" ? 0 : null;
}

function coerceLooseNumber(value: unknown): number | null {
  const strict = safeNumber(value);
  if (strict != null) {
    return strict;
  }
  if (Array.isArray(value)) {
    for (const entry of value) {
      const nested = coerceLooseNumber(entry);
      if (nested != null) {
        return nested;
      }
    }
    return null;
  }
  if (value && typeof value === "object") {
    const record = value as Record<string, unknown>;
    for (const key of ["x", "value", "v", "num", "number", "raw"]) {
      const nested = coerceLooseNumber(record[key]);
      if (nested != null) {
        return nested;
      }
    }
    return null;
  }
  if (typeof value === "string") {
    const match = /-?\d+(?:\.\d+)?/.exec(value);
    if (!match) {
      return null;
    }
    const parsed = Number(match[0]);
    return Number.isFinite(parsed) ? parsed : null;
  }
  return null;
}

function buildPieEntries(chart: XlsxChart, seriesIndex = selectPrimaryPieSeriesIndex(chart)) {
  const categories = getCategoryLabels(chart);
  const values = chart.series[seriesIndex]?.values ?? [];
  return values
    .map((rawValue, index) => ({
      color: resolvePiePointColor(chart, index, seriesIndex),
      index,
      label: normalizeCategoryLabel(categories[index]),
      value: Math.max(0, safeNumber(rawValue) ?? 0)
    }))
    .filter((entry) => entry.value > 0);
}

function resolveChartStageSubtotal(series: XlsxChart["series"][number]) {
  const raw = series.raw && typeof series.raw === "object"
    ? series.raw as Record<string, unknown>
    : null;
  const layoutProperties = raw?.layoutProperties && typeof raw.layoutProperties === "object"
    ? raw.layoutProperties as Record<string, unknown>
    : null;
  const subtotals = Array.isArray(layoutProperties?.subtotals) ? layoutProperties.subtotals : [];
  return subtotals.length > 0 || layoutProperties?.aggregation === true;
}

function buildChartStages(chart: XlsxChart) {
  if (chart.chartType === "Funnel" || chart.chartType === "Waterfall") {
    const primarySeriesIndex = Math.max(0, chart.series.findIndex((series) => series.hidden !== true));
    const primarySeries = chart.series[primarySeriesIndex] ?? null;
    if (!primarySeries) {
      return [];
    }

    const labels = getCategoryLabels(chart);
    return primarySeries.values
      .map((rawValue, index): ChartStage | null => {
        const value = safeNumber(rawValue);
        if (value == null || !Number.isFinite(value)) {
          return null;
        }
        return {
          color: chart.varyColors
            ? chartPointColor(chart, index, primarySeriesIndex)
            : chartSeriesColor(chart, primarySeriesIndex),
          isSubtotal: false,
          label: normalizeCategoryLabel(labels[index]) || String(index + 1),
          value
        };
      })
      .filter((stage): stage is ChartStage => stage != null);
  }

  return chart.series
    .map((series, index): ChartStage | null => {
      const value = series.values.reduce<number>((sum, entry) => sum + (safeNumber(entry) ?? 0), 0);
      if (!Number.isFinite(value)) {
        return null;
      }
      const formula = typeof series.valuesRef?.formula === "string" ? series.valuesRef.formula : "";
      const label = series.name
        ?? (formula.length > 0 ? formula.replace(/^.*!/, "").replace(/\$/g, "") : `Series ${index + 1}`);
      return {
        color: chartSeriesColor(chart, typeof series.formatIdx === "number" ? series.formatIdx : index),
        isSubtotal: resolveChartStageSubtotal(series),
        label,
        value
      };
    })
    .filter((stage): stage is ChartStage => stage != null);
}

function buildHierarchyData(chart: XlsxChart) {
  const root: ChartHierarchyDatum = {
    children: [],
    name: chart.title ?? chart.name ?? "Root"
  };
  const rootChildren = root.children ?? [];
  const topLevelIndexByName = new Map<string, number>();
  const primaryHierarchySeries = chart.series.find((series) => {
    const raw = series.raw && typeof series.raw === "object" ? series.raw as Record<string, unknown> : null;
    return Array.isArray(raw?.chartExHierarchyCategories);
  }) ?? null;
  const primaryHierarchyPaths = (() => {
    if (!primaryHierarchySeries) {
      return null;
    }
    const raw = primaryHierarchySeries.raw as Record<string, unknown>;
    return Array.isArray(raw.chartExHierarchyCategories)
      ? raw.chartExHierarchyCategories.map((entry) => Array.isArray(entry)
        ? entry.map((value) => normalizeCategoryLabel(value)).filter((value) => value.length > 0)
        : [])
      : null;
  })();
  const rowCount = primaryHierarchyPaths
    ? primaryHierarchyPaths.length
    : Math.max(0, ...chart.series.map((series) => series.values.length));

  for (let rowIndex = 0; rowIndex < rowCount; rowIndex += 1) {
    const path = primaryHierarchyPaths?.[rowIndex] ?? chart.series
      .map((series) => normalizeCategoryLabel(series.categories[rowIndex] ?? series.values[rowIndex]))
      .filter((value) => value.length > 0);
    if (path.length === 0) {
      continue;
    }

    let current = root;
    path.forEach((part, levelIndex) => {
      current.children = current.children ?? [];
      const preserveDuplicateLeaf = levelIndex === path.length - 1 && path.length === 1;
      let next = preserveDuplicateLeaf
        ? undefined
        : current.children.find((child) => child.name === part);
      if (!next) {
        next = { children: [], name: part };
        if (levelIndex === 0) {
          const topLevelIndex = topLevelIndexByName.size;
          topLevelIndexByName.set(part, topLevelIndex);
          next.colorIndex = topLevelIndex;
        } else {
          next.colorIndex = current.colorIndex;
        }
        current.children.push(next);
      }
      current = next;
    });

    const valueSource = primaryHierarchySeries ?? chart.series[chart.series.length - 1] ?? null;
    current.value = (current.value ?? 0) + Math.max(0.0001, safeNumber(valueSource?.values[rowIndex]) ?? 1);
  }

  return rootChildren.length > 0 ? root : null;
}

function resolveBoxWhiskerQuartileMethod(series: XlsxChartSeries) {
  const raw = series.raw && typeof series.raw === "object" ? series.raw as Record<string, unknown> : null;
  const layoutProperties = raw?.layoutProperties && typeof raw.layoutProperties === "object"
    ? raw.layoutProperties as Record<string, unknown>
    : null;
  const statistics = layoutProperties?.statistics && typeof layoutProperties.statistics === "object"
    ? layoutProperties.statistics as Record<string, unknown>
    : null;
  return statistics?.quartileMethod === "inclusive" ? "inclusive" : "exclusive";
}

function resolveBoxWhiskerVisibility(series: XlsxChartSeries) {
  const raw = series.raw && typeof series.raw === "object" ? series.raw as Record<string, unknown> : null;
  const layoutProperties = raw?.layoutProperties && typeof raw.layoutProperties === "object"
    ? raw.layoutProperties as Record<string, unknown>
    : null;
  const visibility = layoutProperties?.visibility && typeof layoutProperties.visibility === "object"
    ? layoutProperties.visibility as Record<string, unknown>
    : null;
  return {
    meanLine: visibility?.meanLine === true,
    meanMarker: visibility?.meanMarker !== false,
    nonoutliers: visibility?.nonoutliers === true,
    outliers: visibility?.outliers !== false
  };
}

function computePercentile(sortedValues: number[], percentile: number, method: "exclusive" | "inclusive") {
  const count = sortedValues.length;
  if (count === 0) {
    return 0;
  }
  if (count === 1) {
    return sortedValues[0] ?? 0;
  }

  const rank = method === "exclusive"
    ? percentile * (count + 1)
    : 1 + percentile * (count - 1);
  if (rank <= 1) {
    return sortedValues[0] ?? 0;
  }
  if (rank >= count) {
    return sortedValues[count - 1] ?? 0;
  }

  const lowerIndex = Math.floor(rank) - 1;
  const upperIndex = Math.ceil(rank) - 1;
  const fraction = rank - Math.floor(rank);
  const lower = sortedValues[Math.max(0, lowerIndex)] ?? sortedValues[0] ?? 0;
  const upper = sortedValues[Math.max(0, upperIndex)] ?? sortedValues[count - 1] ?? 0;
  return lower + (upper - lower) * fraction;
}

function computeBoxWhiskerStats(series: XlsxChartSeries): BoxWhiskerStats | null {
  const sortedValues = series.values
    .map((value) => safeNumber(value))
    .filter((value): value is number => value != null)
    .sort((left, right) => left - right);
  if (sortedValues.length === 0) {
    return null;
  }

  const quartileMethod = resolveBoxWhiskerQuartileMethod(series);
  const q1 = computePercentile(sortedValues, 0.25, quartileMethod);
  const median = computePercentile(sortedValues, 0.5, quartileMethod);
  const q3 = computePercentile(sortedValues, 0.75, quartileMethod);
  const iqr = q3 - q1;
  const lowerFence = q1 - iqr * 1.5;
  const upperFence = q3 + iqr * 1.5;
  const visiblePoints = sortedValues.filter((value) => value >= lowerFence && value <= upperFence);
  const outliers = sortedValues.filter((value) => value < lowerFence || value > upperFence);
  const lowerWhisker = visiblePoints[0] ?? sortedValues[0] ?? 0;
  const upperWhisker = visiblePoints[visiblePoints.length - 1] ?? sortedValues[sortedValues.length - 1] ?? 0;
  const mean = sortedValues.reduce((sum, value) => sum + value, 0) / sortedValues.length;

  return {
    lowerFence,
    lowerWhisker,
    max: sortedValues[sortedValues.length - 1] ?? 0,
    mean,
    median,
    min: sortedValues[0] ?? 0,
    outliers,
    q1,
    q3,
    upperFence,
    upperWhisker,
    visiblePoints
  };
}

function resolveHierarchyNodeColor(chart: XlsxChart, node: HierarchyNode<ChartHierarchyDatum>) {
  const lineage = node.ancestors().reverse();
  const topLevel = lineage[1];
  const baseIndex = topLevel?.data.colorIndex ?? 0;
  const baseColor = chartSeriesColor(chart, baseIndex);
  const depth = Math.max(0, node.depth - 1);
  return depth === 0 ? baseColor : lightenColor(baseColor, clamp(depth * 0.16, 0, 0.5));
}

function resolveTreemapNodeColor(chart: XlsxChart, node: HierarchyNode<ChartHierarchyDatum>) {
  const lineage = node.ancestors().reverse();
  const topLevel = lineage[1];
  return chartSeriesColor(chart, topLevel?.data.colorIndex ?? 0);
}

function excelTreemapTile(node: HierarchyRectangularNode<ChartHierarchyDatum>, x0: number, y0: number, x1: number, y1: number) {
  if (!node.children || node.children.length === 0) {
    return;
  }
  if (node.depth === 0) {
    treemapDice(node, x0, y0, x1, y1);
    node.children.forEach((child) => excelTreemapTile(child, child.x0, child.y0, child.x1, child.y1));
    return;
  }
  treemapBinary(node, x0, y0, x1, y1);
  node.children.forEach((child) => excelTreemapTile(child, child.x0, child.y0, child.x1, child.y1));
}

function resolveSurfaceBaseColor(chart: XlsxChart, palette: Pick<ChartRendererPalette, "text">) {
  return chart.chartColorPalette?.[0]
    ?? chart.series[0]?.color
    ?? chart.series[0]?.lineColor
    ?? chart.axisLineColor
    ?? chart.textColor
    ?? palette.text;
}

function resolveExplicitSurfaceSeriesColor(chart: XlsxChart, kind: "fill" | "line") {
  const shape = chart.series[0]?.shapeProperties;
  if (kind === "fill" ? shape?.xmlFillHidden === true : shape?.xmlLineHidden === true) return "transparent";
  const color = kind === "fill"
    ? shape?.xmlFillColor ?? shape?.solidFillHex
    : shape?.xmlLineColor ?? shape?.lineColorHex;
  return normalizeRendererHexColor(color);
}

function getSurfaceColorStops(chart: XlsxChart, palette: ChartRendererPalette) {
  const explicitFill = resolveExplicitSurfaceSeriesColor(chart, "fill");
  if (explicitFill) return [explicitFill];
  const colors = (chart.chartColorPalette ?? []).filter((value): value is string => typeof value === "string" && value.length > 0);
  return colors.length > 0 ? colors : [resolveSurfaceBaseColor(chart, palette)];
}

function resolveSurfaceBandPaletteColor(chart: XlsxChart, palette: ChartRendererPalette, domain: SurfaceDomain, value: number) {
  const bandIndex = resolveSurfaceBandIndex(domain, value);
  const explicitColor = resolveSurfaceBandFormatColor(chart, "bandFormatColors", bandIndex);
  if (explicitColor) return explicitColor;
  const stops = getSurfaceColorStops(chart, palette);
  if (stops.length === 0) {
    return resolveSurfaceBaseColor(chart, palette);
  }
  return stops[Math.min(stops.length - 1, Math.max(0, bandIndex))] ?? stops[stops.length - 1] ?? resolveSurfaceBaseColor(chart, palette);
}

function getSurfaceDomain(chart: XlsxChart): SurfaceDomain | null {
  const numericValues = chart.series.flatMap((series) => (
    series.values
      .map((value) => safeNumber(value))
      .filter((value): value is number => value != null)
  ));
  if (numericValues.length === 0) {
    return null;
  }
  const explicitMin = typeof chart.valueAxis?.min === "number" && Number.isFinite(chart.valueAxis.min)
    ? chart.valueAxis.min
    : null;
  const explicitMax = typeof chart.valueAxis?.max === "number" && Number.isFinite(chart.valueAxis.max)
    ? chart.valueAxis.max
    : null;
  const rawMin = Math.min(...numericValues);
  const rawMax = Math.max(...numericValues);
  const spanBase = Math.max(1e-6, rawMax - Math.min(0, rawMin));
  const roughStep = spanBase / 5;
  const magnitude = 10 ** Math.floor(Math.log10(Math.max(roughStep, 1e-6)));
  const normalized = roughStep / magnitude;
  const niceNormalized = normalized <= 1 ? 1 : normalized <= 2 ? 2 : normalized <= 5 ? 5 : 10;
  const step = typeof chart.valueAxis?.majorUnit === "number" && chart.valueAxis.majorUnit > 0
    ? chart.valueAxis.majorUnit
    : niceNormalized * magnitude;
  const minValue = explicitMin ?? (rawMin >= 0 ? 0 : Math.floor(rawMin / step) * step);
  const maxValue = explicitMax ?? Math.ceil(rawMax / step) * step;
  const safeMax = maxValue <= minValue ? minValue + step : maxValue;
  const ticks: number[] = [];
  for (let current = minValue; current <= safeMax + step * 0.001; current += step) {
    ticks.push(Number(current.toFixed(8)));
  }
  const boundedTicks = [...new Set([...ticks.filter((tick) => tick < safeMax), safeMax])];
  return {
    maxValue,
    minValue,
    safeMax,
    ticks: boundedTicks.length >= 2 ? boundedTicks : [minValue, safeMax]
  };
}

function resolveSurfaceBandColor(chart: XlsxChart, palette: ChartRendererPalette, domain: SurfaceDomain, value: number) {
  return resolveSurfaceBandPaletteColor(chart, palette, domain, value);
}

function resolveSurfaceBandIndex(domain: SurfaceDomain, value: number) {
  const ticks = domain.ticks;
  for (let index = 0; index < ticks.length - 1; index += 1) {
    const end = ticks[index + 1] ?? domain.safeMax;
    if (value <= end || index === ticks.length - 2) {
      return index;
    }
  }
  return Math.max(0, ticks.length - 2);
}

function resolveSurfaceBandFormatColor(chart: XlsxChart, key: "bandFormatColors" | "bandFormatLineColors", index: number) {
  const colors = chart.raw?.[key];
  if (!colors || typeof colors !== "object" || !Object.prototype.hasOwnProperty.call(colors, index)) return null;
  const color = (colors as Record<number, unknown>)[index];
  return typeof color === "string" && color.length > 0 ? color : null;
}

function resolveSurfaceWireframeColor(chart: XlsxChart, palette: ChartRendererPalette, domain: SurfaceDomain, value: number) {
  const bandIndex = resolveSurfaceBandIndex(domain, value);
  return resolveSurfaceBandFormatColor(chart, "bandFormatLineColors", bandIndex)
    ?? resolveExplicitSurfaceSeriesColor(chart, "line")
    ?? resolveSurfaceBandPaletteColor(chart, palette, domain, value);
}

function resolveSurfacePlotRect(chart: XlsxChart, layout: ChartLayout) {
  if (!isContourSurfaceChart(chart)) {
    return layout.plot;
  }
  const columnCount = Math.max(1, getCategoryLabels(chart).length);
  const depthScale = Math.max(0.01, (chart.view3d?.depthPercent ?? 100) / 100);
  const categoryExtent = chart.valueAxis?.crossBetween !== "midCat" && columnCount > 1 ? columnCount / (columnCount - 1) : 1;
  const targetAspect = categoryExtent / depthScale;
  let width = layout.plot.width;
  let height = layout.plot.height;
  if (width / Math.max(1e-6, height) > targetAspect) {
    width = height * targetAspect;
  } else {
    height = width / Math.max(1e-6, targetAspect);
  }
  return {
    height,
    left: layout.plot.left + (layout.plot.width - width) / 2,
    top: layout.plot.top + (layout.plot.height - height) / 2,
    width
  };
}

function buildSurfaceLegendItems(chart: XlsxChart, palette: ChartRendererPalette) {
  const domain = getSurfaceDomain(chart);
  if (!domain) {
    return [];
  }
  const items: LegendItem[] = [];
  for (let index = 0; index < domain.ticks.length - 1; index += 1) {
    const start = domain.ticks[index] ?? domain.minValue;
    const end = domain.ticks[index + 1] ?? domain.safeMax;
    const midpoint = start + (end - start) * 0.5;
    items.push({
      color: chart.wireframe
        ? resolveSurfaceWireframeColor(chart, palette, domain, midpoint)
        : resolveSurfaceBandColor(chart, palette, domain, midpoint),
      label: `${formatTickValue(start)}-${formatTickValue(end)}`
    });
  }
  return items.reverse();
}

function isContourSurfaceChart(chart: XlsxChart) {
  const rawChartType = chart.raw && typeof chart.raw === "object" && typeof (chart.raw as Record<string, unknown>).xmlChartType === "string"
    ? (chart.raw as Record<string, unknown>).xmlChartType
    : "";
  if (rawChartType === "surfaceChart") {
    return true;
  }
  if (rawChartType === "surface3DChart") {
    return false;
  }
  return chart.chartType === "Surface" && chart.is3d !== true;
}

function renderSurfaceAxes(chart: XlsxChart, layout: ChartLayout) {
  const plot = resolveSurfacePlotRect(chart, layout);
  const categories = getCategoryLabels(chart);
  const seriesLabels = chart.series.map((series, index) => normalizeCategoryLabel(series.name) || `Q${index + 1}`);
  const labelColor = resolveChartAxisTextColor(chart);
  const axisColor = chart.axisLineColor ?? chart.chartAreaBorderColor ?? "#888888";
  const rowCount = Math.max(1, seriesLabels.length);
  const columnCount = Math.max(1, categories.length);
  const columnPositions = categories.map((_, index) => (
    plot.left + (isContourSurfaceChart(chart) && chart.valueAxis?.crossBetween !== "midCat"
      ? ((index + 0.5) / columnCount) * plot.width
      : (columnCount <= 1 ? plot.width / 2 : (index / (columnCount - 1)) * plot.width))
  ));
  const rowPositions = seriesLabels.map((_, index) => (
    plot.top + plot.height - (rowCount <= 1 ? plot.height / 2 : (index / (rowCount - 1)) * plot.height)
  ));

  return (
    <g>
      <rect
        fill="none"
        height={plot.height}
        stroke={lightenColor(axisColor, 0.18)}
        strokeWidth={0.8}
        width={plot.width}
        x={plot.left}
        y={plot.top}
      />
      {categories.map((label, index) => (
        <text
          key={`surface-x-label-${index}`}
          fill={labelColor}
          fontSize={10}
          textAnchor="middle"
          x={columnPositions[index] ?? plot.left}
          y={plot.top + plot.height + 14}
        >
          {label}
        </text>
      ))}
      {seriesLabels.map((label, index) => (
        <text
          key={`surface-y-label-${index}`}
          fill={labelColor}
          fontSize={10}
          textAnchor="start"
          x={plot.left + plot.width + 8}
          y={(rowPositions[index] ?? plot.top) + 3}
        >
          {label}
        </text>
      ))}
    </g>
  );
}

function normalizeRendererHexColor(value: unknown) {
  if (typeof value !== "string") {
    return null;
  }
  const trimmed = value.trim();
  if (!trimmed) {
    return null;
  }
  const normalized = trimmed.startsWith("#") ? trimmed.slice(1) : trimmed;
  return /^[0-9a-f]{6}$/i.test(normalized) ? `#${normalized.toLowerCase()}` : null;
}

function buildStockLegendItems(chart: XlsxChart, _palette: ChartRendererPalette): LegendItem[] {
  return chart.series.map((series, index) => ({
    color: chartSeriesStrokeColor(chart, index),
    label: series.name ?? `Series ${index + 1}`
  }));
}

function getLegendItems(chart: XlsxChart, chartType: string, palette: ChartRendererPalette): LegendItem[] {
  if (!chart.legend) {
    return [];
  }
  if (isComboChart(chart)) {
    return getComboLegendSeries(chart);
  }
  if (chartType === "Stock") {
    return buildStockLegendItems(chart, palette);
  }
  if (chartType === "Surface") {
    return buildSurfaceLegendItems(chart, palette);
  }
  if (chartType === "RegionMap") {
    return buildRegionMapLegendItems(chart);
  }
  if (variesBarColorsByCategory(chart)) {
    return getCategoryLabels(chart).map((label, index) => ({
      color: chartPointColor(chart, index),
      label
    }));
  }
  if (chartType === "Sunburst" || chartType === "Treemap") {
    const hierarchyData = buildHierarchyData(chart);
    if (!hierarchyData?.children) {
      return [];
    }
    return hierarchyData.children.map((child, index) => ({
      color: chartSeriesColor(chart, child.colorIndex ?? index),
      label: child.name
    }));
  }
  if (chartType === "Pie" || chartType === "Pie3D" || chartType === "PieExploded" || chartType === "Doughnut" || chartType === "BarOfPie") {
    if (chartType === "BarOfPie") {
      const categories = getCategoryLabels(chart);
      const values = chart.series[0]?.values ?? [];
      return categories
        .map((label, index) => ({
          color: chartPointColor(chart, index),
          label: normalizeCategoryLabel(label),
          value: safeNumber(values[index]) ?? 0
        }))
        .filter((entry) => entry.value > 0 && entry.label.trim().length > 0)
        .map((entry) => ({
          color: entry.color,
          label: entry.label
        }));
    }
    return buildPieEntries(chart).map((entry) => ({
      color: entry.color,
      label: entry.label
    }));
  }
  const isXyLegend = (
    chartType === "Scatter"
    || chartType === "ScatterLines"
    || chartType === "ScatterSmooth"
    || chartType === "Bubble"
  );
  return chart.series.map((series, index) => ({
    color: isXyLegend
      ? (series.lineColor ?? series.markerColor ?? series.color ?? chartSeriesColor(chart, index))
      : chartSeriesColor(chart, index),
    label: series.name ?? `Series ${index + 1}`
  }));
}

function buildLayout(chart: XlsxChart, rect: XlsxImageRect, legendItems: LegendItem[]): ChartLayout {
  const width = Math.max(80, Math.round(rect.width));
  const height = Math.max(60, Math.round(rect.height));
  const isSurfaceChart = chart.chartType === "Surface"
    || (chart.raw && typeof chart.raw === "object" && typeof (chart.raw as Record<string, unknown>).xmlChartType === "string" && String((chart.raw as Record<string, unknown>).xmlChartType).includes("surface"));
  const titleHeight = chart.title ? (isSurfaceChart ? 30 : 24) : 8;
  const legendPosition = normalizeLegendPosition(chart.legend?.position);

  const legendVertical = legendItems.length > 0 && (legendPosition === "right" || legendPosition === "left");
  const compactWidth = width <= 280;
  const compactHeight = height <= 200;
  const finalLegendPosition = compactWidth || compactHeight
    ? (legendVertical ? "bottom" : legendPosition)
    : legendPosition;
  const layoutLegendPosition = chart.chartType === "RegionMap" ? finalLegendPosition : legendPosition;
  const primarySeries = chart.series.find((series) => series.hidden !== true) ?? chart.series[0];
  const horizontalMapGradient = chart.chartType === "RegionMap"
    && legendItems.length > 0 && !resolveRegionMapColorStrings(primarySeries)
    && (layoutLegendPosition === "top" || layoutLegendPosition === "bottom");
  const horizontalLegendHeight = horizontalMapGradient
    ? Math.min(48, Math.max(24, height - titleHeight - 8))
    : 24;
  const plotLeft = 42 + (layoutLegendPosition === "left" ? 98 : 0);
  const plotRight = 16 + (layoutLegendPosition === "right" ? 98 : 0);
  const plotTop = titleHeight + (layoutLegendPosition === "top" ? horizontalLegendHeight : 0);
  const plotBottom = (horizontalMapGradient ? 8 : 28) + (layoutLegendPosition === "bottom" ? horizontalLegendHeight : 0);
  const plotWidth = Math.max(horizontalMapGradient ? 1 : 40, width - plotLeft - plotRight);

  return {
    height,
    legendBounds: horizontalMapGradient ? {
      left: plotLeft,
      top: layoutLegendPosition === "top" ? titleHeight : height - horizontalLegendHeight,
      width: plotWidth,
      height: horizontalLegendHeight
    } : undefined,
    legendItems,
    legendPosition: finalLegendPosition,
    plot: {
      height: Math.max(horizontalMapGradient ? 1 : 40, height - plotTop - plotBottom),
      left: plotLeft,
      top: plotTop,
      width: plotWidth
    },
    titleHeight,
    width
  };
}

function buildNiceStep(minValue: number, maxValue: number, preferredTicks = 5) {
  const span = Math.max(1e-6, maxValue - minValue);
  const roughStep = span / Math.max(1, preferredTicks);
  const magnitude = 10 ** Math.floor(Math.log10(roughStep));
  const normalized = roughStep / magnitude;
  const niceNormalized = normalized <= 1 ? 1 : normalized <= 2 ? 2 : normalized <= 5 ? 5 : 10;
  return niceNormalized * magnitude;
}

function buildNumericTickValues(minValue: number, maxValue: number, majorUnit?: number) {
  if (!Number.isFinite(minValue) || !Number.isFinite(maxValue) || maxValue <= minValue) {
    return [];
  }
  const step = typeof majorUnit === "number" && majorUnit > 0
    ? majorUnit
    : buildNiceStep(minValue, maxValue, 5);
  if (!Number.isFinite(step) || step <= 0) {
    return [];
  }
  const start = Math.floor(minValue / step) * step;
  const values: number[] = [];
  for (let current = start; current <= maxValue + step * 0.001; current += step) {
    if (current >= minValue - step * 0.001 && current <= maxValue + step * 0.001) {
      values.push(Number(current.toFixed(8)));
    }
  }
  return values;
}

function resolveNumericAxisDomain(minValue: number, maxValue: number, majorUnit?: number, includeZero = false) {
  if (!Number.isFinite(minValue) || !Number.isFinite(maxValue)) {
    return {
      max: 1,
      min: 0,
      ticks: [0, 1]
    };
  }
  let domainMin = includeZero ? Math.min(0, minValue) : minValue;
  let domainMax = includeZero ? Math.max(0, maxValue) : maxValue;
  if (domainMax <= domainMin) {
    domainMax = domainMin + 1;
  }
  const step = typeof majorUnit === "number" && majorUnit > 0
    ? majorUnit
    : buildNiceStep(domainMin, domainMax, 5);
  const roundedMin = typeof majorUnit === "number" && majorUnit > 0
    ? domainMin
    : Math.floor(domainMin / step) * step;
  const roundedMax = typeof majorUnit === "number" && majorUnit > 0
    ? domainMax
    : Math.ceil(domainMax / step) * step;
  const finalMin = includeZero ? Math.min(0, roundedMin) : roundedMin;
  const finalMax = roundedMax <= finalMin ? finalMin + step : roundedMax;
  const ticks = buildNumericTickValues(finalMin, finalMax, step);
  return {
    max: finalMax,
    min: finalMin,
    ticks: ticks.length > 0 ? ticks : [finalMin, finalMax]
  };
}

function resolveAxisDomainWithChartOverrides(
  axis: XlsxChartAxis | null | undefined,
  minValue: number,
  maxValue: number,
  includeZero = false
) {
  const hasExplicitMin = typeof axis?.min === "number" && Number.isFinite(axis.min);
  const hasExplicitMax = typeof axis?.max === "number" && Number.isFinite(axis.max);
  const rawMin = hasExplicitMin ? Number(axis?.min) : minValue;
  const rawMax = hasExplicitMax ? Number(axis?.max) : maxValue;
  const domain = resolveNumericAxisDomain(rawMin, rawMax, axis?.majorUnit, includeZero);
  return {
    hasExplicitMax,
    hasExplicitMin,
    majorUnit: axis?.majorUnit,
    max: hasExplicitMax ? Number(axis?.max) : domain.max,
    min: hasExplicitMin ? Number(axis?.min) : domain.min,
    ticks: (hasExplicitMin || hasExplicitMax)
      ? buildNumericTickValues(
          hasExplicitMin ? Number(axis?.min) : domain.min,
          hasExplicitMax ? Number(axis?.max) : domain.max,
          axis?.majorUnit
        )
      : domain.ticks
  };
}

function resolve3dFrameOffsets(chart: XlsxChart, baseDepthX = 11, baseDepthY = 8) {
  const depthRatio = clamp((chart.view3d?.depthPercent ?? 100) / 100, 0.35, 3.2);
  const rotX = clamp(chart.view3d?.rotX ?? 20, -80, 80);
  const rotY = clamp(chart.view3d?.rotY ?? 20, -80, 80);
  const usePerspective = chart.view3d?.rAngAx === false;
  const perspectiveStrength = clamp((chart.view3d?.perspective ?? (usePerspective ? 26 : 0)) / 100, 0, 1);
  const projectedOrigin = projectCartesian3dPoint(
    0,
    0,
    0,
    rotX * (Math.PI / 180),
    rotY * (Math.PI / 180),
    usePerspective,
    perspectiveStrength
  );
  const projectedDepthPoint = projectCartesian3dPoint(
    0,
    0,
    1,
    rotX * (Math.PI / 180),
    rotY * (Math.PI / 180),
    usePerspective,
    perspectiveStrength
  );
  const depthDirectionX = projectedDepthPoint.x - projectedOrigin.x;
  const depthDirectionY = projectedDepthPoint.y - projectedOrigin.y;
  const horizontalBias = clamp(1.42 + Math.abs(rotY) / 48 + depthRatio * 0.06, 1.5, 1.95);
  const verticalBias = clamp(1.18 + Math.abs(rotX) / 78, 1.22, 1.5);
  const biasedDepthDirectionX = depthDirectionX * horizontalBias;
  const biasedDepthDirectionY = depthDirectionY * verticalBias;
  const depthDirectionMagnitude = Math.hypot(biasedDepthDirectionX, biasedDepthDirectionY) || 1;
  const horizontalFactor = clamp(Math.abs(rotY) / 22, 0.55, 2.2);
  const verticalFactor = clamp(Math.abs(rotX) / 18, 0.45, 1.9);
  const legacyDepthX = Math.max(6, baseDepthX * depthRatio * horizontalFactor) * (rotY < 0 ? -1 : 1);
  const legacyDepthY = -Math.max(4, baseDepthY * depthRatio * verticalFactor);
  const depthMagnitude = Math.max(10, Math.hypot(legacyDepthX, legacyDepthY) * 1.34);
  const depthX = biasedDepthDirectionX / depthDirectionMagnitude * depthMagnitude;
  const depthY = biasedDepthDirectionY / depthDirectionMagnitude * depthMagnitude;
  return {
    depthRatio,
    depthX,
    depthY,
    insetBottom: Math.max(6, Math.abs(depthY) + 5),
    insetLeft: depthX < 0 ? Math.abs(depthX) + 4 : 0,
    insetRight: depthX > 0 ? depthX + 4 : 0,
    insetTop: Math.max(4, Math.abs(depthY) + 2)
  };
}

function formatTickValue(value: number) {
  if (Math.abs(value) >= 1000) {
    return value.toLocaleString(undefined, { maximumFractionDigits: 0 });
  }
  return value.toLocaleString(undefined, { maximumFractionDigits: 2 });
}

function formatPercentTickValue(value: number) {
  return `${formatTickValue(value)}%`;
}

function isLikelyDateFormatCode(formatCode: string | undefined) {
  if (!formatCode) {
    return false;
  }
  const normalized = formatCode
    .toLowerCase()
    .replace(/\[[^\]]*]/g, " ")
    .replace(/"[^"]*"/g, " ")
    .replace(/\\./g, " ");
  return /(?:d|m|y)/.test(normalized);
}

function excelSerialToDate(serial: number, use1904: boolean) {
  const wholeDays = Math.trunc(serial);
  const milliseconds = Math.round((serial - wholeDays) * 86_400_000);
  const baseUtc = use1904
    ? Date.UTC(1904, 0, 1)
    : Date.UTC(1899, 11, 30);
  return new Date(baseUtc + wholeDays * 86_400_000 + milliseconds);
}

function formatExcelDateSerial(value: number, formatCode: string | undefined, use1904: boolean) {
  const date = excelSerialToDate(value, use1904);
  const normalized = (formatCode ?? "").toLowerCase();
  const options: Intl.DateTimeFormatOptions = {};

  if (/yyyy/.test(normalized)) {
    options.year = "numeric";
  } else if (/yy/.test(normalized)) {
    options.year = "2-digit";
  }

  if (/mmmm/.test(normalized)) {
    options.month = "long";
  } else if (/mmm/.test(normalized)) {
    options.month = "short";
  } else if (/(^|[^a-z])m([^a-z]|$)|(^|[^a-z])mm([^a-z]|$)/.test(normalized)) {
    options.month = "numeric";
  }

  if (/dddd/.test(normalized)) {
    options.weekday = "long";
  } else if (/ddd/.test(normalized)) {
    options.weekday = "short";
  }

  if (/d/.test(normalized)) {
    options.day = "numeric";
  }

  if (Object.keys(options).length === 0) {
    options.month = "short";
    options.day = "numeric";
  }

  return new Intl.DateTimeFormat(undefined, options).format(date);
}

function formatCategoryLabel(chart: XlsxChart, value: unknown, fallback = "") {
  const rawRecord = chart.raw && typeof chart.raw === "object"
    ? chart.raw as Record<string, unknown>
    : null;
  const numeric = typeof value === "number" && Number.isFinite(value)
    ? value
    : null;
  const formatCode = chart.categoryAxis?.numberFormat?.formatCode;
  if (numeric != null && isLikelyDateFormatCode(formatCode)) {
    return formatExcelDateSerial(numeric, formatCode, rawRecord?.date1904 === true);
  }

  const normalized = normalizeCategoryLabel(value);
  if (normalized.length > 0) {
    return normalized;
  }
  return fallback;
}

function truncateSvgText(value: string, maxWidth: number, fontSize = 10) {
  if (!value || maxWidth <= 0) {
    return "";
  }
  const charWidth = Math.max(4.2, fontSize * 0.56);
  const maxChars = Math.max(1, Math.floor(maxWidth / charWidth));
  if (value.length <= maxChars) {
    return value;
  }
  if (maxChars <= 1) {
    return "…";
  }
  return `${value.slice(0, Math.max(0, maxChars - 1)).trimEnd()}…`;
}

function markerRadiusPx(sizePt: number | undefined) {
  return (Number.isFinite(sizePt) && Number(sizePt) > 0 ? Number(sizePt) : 5) * 2 / 3;
}

function markerSymbolPath(symbol: string, radius: number) {
  const diagonal = radius / Math.SQRT2;
  switch (symbol) {
    case "none":
      return "";
    case "diamond":
      return `M0,${-radius}L${radius},0L0,${radius}L${-radius},0Z`;
    case "square":
      return `M${-radius},${-radius}H${radius}V${radius}H${-radius}Z`;
    case "triangle":
    case "triangle-up":
      return `M0,${-radius}L${radius},${radius}H${-radius}Z`;
    case "cross":
    case "x":
      return `M${-diagonal},${-diagonal}L${diagonal},${diagonal}M${-diagonal},${diagonal}L${diagonal},${-diagonal}`;
    case "plus":
      return `M${-radius},0H${radius}M0,${-radius}V${radius}`;
    case "star":
      return `M${-radius},0H${radius}M${-radius / 2},${-radius * Math.sqrt(3) / 2}L${radius / 2},${radius * Math.sqrt(3) / 2}M${-radius / 2},${radius * Math.sqrt(3) / 2}L${radius / 2},${-radius * Math.sqrt(3) / 2}`;
    case "dash":
      return `M${-radius},0H${radius}`;
    case "dot":
      return `M${-radius / 2},0H${radius / 2}`;
    default:
      return `M${radius},0A${radius},${radius} 0 1,0 ${-radius},0A${radius},${radius} 0 1,0 ${radius},0Z`;
  }
}

const TWO_PI = Math.PI * 2;

function normalizePieArc(startAngle: number, endAngle: number) {
  let start = startAngle;
  let end = endAngle;
  while (end < start) {
    end += TWO_PI;
  }
  return { end, start };
}

function resolvePieFrontSegments(startAngle: number, endAngle: number): Array<[number, number]> {
  const { start, end } = normalizePieArc(startAngle, endAngle);
  const segments: Array<[number, number]> = [];
  const minBand = Math.floor(start / TWO_PI) - 1;
  const maxBand = Math.ceil(end / TWO_PI) + 1;
  for (let band = minBand; band <= maxBand; band += 1) {
    const frontStart = Math.PI / 2 + band * TWO_PI;
    const frontEnd = Math.PI * 1.5 + band * TWO_PI;
    const segmentStart = Math.max(start, frontStart);
    const segmentEnd = Math.min(end, frontEnd);
    if (segmentEnd - segmentStart > 1e-4) {
      segments.push([segmentStart, segmentEnd]);
    }
  }
  return segments;
}

function pieEllipsePoint(
  centerX: number,
  centerY: number,
  radius: number,
  tilt: number,
  angle: number,
  depth = 0
) {
  return {
    x: centerX + Math.sin(angle) * radius,
    y: centerY - Math.cos(angle) * radius * tilt + depth
  };
}

function toSvgNumber(value: number) {
  return Number(value.toFixed(3));
}

function buildPieOuterWallPath(
  centerX: number,
  centerY: number,
  radius: number,
  tilt: number,
  depth: number,
  startAngle: number,
  endAngle: number,
  fullArc = false
) {
  const ry = radius * tilt;
  const segments = fullArc ? [[startAngle, endAngle] as [number, number]] : resolvePieFrontSegments(startAngle, endAngle);
  return segments.map(([segmentStart, segmentEnd]) => {
    const topStart = pieEllipsePoint(centerX, centerY, radius, tilt, segmentStart, 0);
    const topEnd = pieEllipsePoint(centerX, centerY, radius, tilt, segmentEnd, 0);
    const bottomStart = pieEllipsePoint(centerX, centerY, radius, tilt, segmentStart, depth);
    const bottomEnd = pieEllipsePoint(centerX, centerY, radius, tilt, segmentEnd, depth);
    const largeArc = segmentEnd - segmentStart > Math.PI ? 1 : 0;
    return [
      `M ${toSvgNumber(topStart.x)} ${toSvgNumber(topStart.y)}`,
      `A ${toSvgNumber(radius)} ${toSvgNumber(ry)} 0 ${largeArc} 1 ${toSvgNumber(topEnd.x)} ${toSvgNumber(topEnd.y)}`,
      `L ${toSvgNumber(bottomEnd.x)} ${toSvgNumber(bottomEnd.y)}`,
      `A ${toSvgNumber(radius)} ${toSvgNumber(ry)} 0 ${largeArc} 0 ${toSvgNumber(bottomStart.x)} ${toSvgNumber(bottomStart.y)}`,
      "Z"
    ].join(" ");
  });
}

function isPieFrontFacingAngle(angle: number) {
  const normalized = ((angle % TWO_PI) + TWO_PI) % TWO_PI;
  return normalized > Math.PI / 2 && normalized < Math.PI * 1.5;
}

function buildPieRadialWallPath(
  centerX: number,
  centerY: number,
  radius: number,
  tilt: number,
  depth: number,
  angle: number,
  forceVisible = false
) {
  if (!forceVisible && !isPieFrontFacingAngle(angle)) {
    return null;
  }
  const topCenter = { x: centerX, y: centerY };
  const bottomCenter = { x: centerX, y: centerY + depth };
  const topEdge = pieEllipsePoint(centerX, centerY, radius, tilt, angle, 0);
  const bottomEdge = pieEllipsePoint(centerX, centerY, radius, tilt, angle, depth);
  return [
    `M ${toSvgNumber(topCenter.x)} ${toSvgNumber(topCenter.y)}`,
    `L ${toSvgNumber(topEdge.x)} ${toSvgNumber(topEdge.y)}`,
    `L ${toSvgNumber(bottomEdge.x)} ${toSvgNumber(bottomEdge.y)}`,
    `L ${toSvgNumber(bottomCenter.x)} ${toSvgNumber(bottomCenter.y)}`,
    "Z"
  ].join(" ");
}

function renderTitle(chart: XlsxChart, layout: ChartLayout, palette: ChartRendererPalette) {
  if (!chart.title) {
    return null;
  }
  const fontSize = 12;
  const text = truncateSvgText(chart.title, Math.max(40, layout.width - 12), fontSize);
  const baselineY = layout.titleHeight >= 30 ? 19 : 16;
  return (
    <text
      fill={chart.titleColor ?? chart.textColor ?? DEFAULT_CHART_TEXT_COLOR}
      fontFamily={buildChartFontFamily(chart.titleFontFamily ?? chart.fontFamily)}
      fontSize={fontSize}
      fontWeight={600}
      textAnchor="middle"
      x={layout.width / 2}
      y={baselineY}
    >
      {text}
    </text>
  );
}

function renderLegend(chart: XlsxChart, layout: ChartLayout, palette: ChartRendererPalette) {
  if (!chart.legend || layout.legendItems.length === 0) {
    return null;
  }
  const textColor = resolveChartAxisTextColor(chart);
  const legendPos = layout.legendPosition;
  const items = layout.legendItems;
  const swatchSize = 8;
  const textOffset = swatchSize + 4;

  if (chart.chartType === "RegionMap") {
    const seriesIndex = Math.max(0, chart.series.findIndex((series) => series.hidden !== true));
    const series = chart.series[seriesIndex];
    if (!resolveRegionMapColorStrings(series)) {
      const vertical = legendPos === "left" || legendPos === "right";
      const x = vertical ? (legendPos === "right" ? layout.plot.left + layout.plot.width + 8 : 8) : layout.plot.left;
      const showSeriesName = vertical || (layout.legendBounds?.height ?? 48) >= 44;
      const y = vertical ? layout.plot.top + 26 : (layout.legendBounds?.top ?? layout.height - 48) + (showSeriesName ? 20 : 4);
      const length = vertical ? Math.min(100, Math.max(24, layout.plot.height - 42)) : Math.min(160, layout.plot.width);
      const gradientId = `map-legend-${chart.id}`;
      const stops = resolveRegionMapValueColors(series);
      const count = stops?.length ?? 3;
      return (
        <g>
          <defs>
            <linearGradient id={gradientId} x1="0%" x2={vertical ? "0%" : "100%"} y1={vertical ? "100%" : "0%"} y2="0%">
              {Array.from({ length: count }, (_, index) => (
                <stop key={index} offset={`${index / (count - 1) * 100}%`} stopColor={resolveRegionMapValueColor(chart, seriesIndex, index / (count - 1))} />
              ))}
            </linearGradient>
          </defs>
          {showSeriesName && <text fill={textColor} fontSize={10} x={x} y={y - 12}>{series?.name ?? ""}</text>}
          <rect fill={`url(#${gradientId})`} x={x} y={y} width={vertical ? 8 : length} height={vertical ? length : 8} />
          <text fill={textColor} fontSize={10} x={vertical ? x + 12 : x} y={vertical ? y + length : y + 20}>{items[0]?.label}</text>
          <text fill={textColor} fontSize={10} textAnchor={vertical ? "start" : "end"} x={vertical ? x + 12 : x + length} y={vertical ? y + 7 : y + 20}>{items[1]?.label}</text>
        </g>
      );
    }
  }

  if (legendPos === "left" || legendPos === "right") {
    const x = legendPos === "right"
      ? layout.plot.left + layout.plot.width + 8
      : 8;
    const startY = chart.chartType === "Surface"
      ? layout.plot.top + Math.max(6, (layout.plot.height - items.length * 18) / 2 + 9)
      : layout.plot.top + 6;
    return (
      <g>
        {items.map((item, index) => {
          const y = startY + index * 18;
          return (
              <g
                {...(index < chart.series.length ? chartElementDataProps(index) : {})}
                data-xlsx-chart-element-kind={index < chart.series.length ? "legendEntry" : undefined}
                key={`legend-${index}`}
                transform={`translate(${x}, ${y})`}
              >
              <rect fill={chart.chartType === "Surface" && chart.wireframe ? "none" : item.color} stroke={chart.chartType === "Surface" && chart.wireframe ? item.color : undefined} height={swatchSize} width={swatchSize} x={0} y={-7} />
              <text fill={textColor} fontSize={10} x={textOffset} y={0}>
                {item.label}
              </text>
            </g>
          );
        })}
      </g>
    );
  }

  const rowY = legendPos === "top" ? (layout.titleHeight + 12) : (layout.height - 8);
  const totalWidth = items.reduce((sum, item) => sum + 24 + Math.min(96, item.label.length * 5.4), 0);
  let cursorX = Math.max(8, (layout.width - totalWidth) / 2);
  return (
    <g>
      {items.map((item, index) => {
        const labelWidth = Math.min(96, item.label.length * 5.4);
        const node = (
            <g
              {...(index < chart.series.length ? chartElementDataProps(index) : {})}
              data-xlsx-chart-element-kind={index < chart.series.length ? "legendEntry" : undefined}
              key={`legend-${index}`}
              transform={`translate(${cursorX}, ${rowY})`}
            >
            <rect fill={chart.chartType === "Surface" && chart.wireframe ? "none" : item.color} stroke={chart.chartType === "Surface" && chart.wireframe ? item.color : undefined} height={swatchSize} width={swatchSize} x={0} y={-7} />
            <text fill={textColor} fontSize={10} x={textOffset} y={0}>
              {item.label}
            </text>
          </g>
        );
        cursorX += 24 + labelWidth;
        return node;
      })}
    </g>
  );
}

function renderCartesianAxes(
  chart: XlsxChart,
  palette: ChartRendererPalette,
  plot: PlotRect,
  isHorizontal: boolean,
  categoryLabels: string[],
  categoryPositions: number[],
  valueTicks: number[],
  mapValue: (value: number) => number,
  formatValueTick: (value: number) => string = formatTickValue
) {
  const axisColor = chart.axisLineColor ?? chart.chartAreaBorderColor ?? palette.border;
  const labelColor = resolveChartAxisTextColor(chart);
  const zeroPosition = mapValue(0);

  return (
    <g>
      {valueTicks.map((tick) => {
        const valuePosition = mapValue(tick);
        if (isHorizontal) {
          return (
            <g key={`grid-v-${tick}`}>
              <line
                stroke={lightenColor(axisColor, 0.7)}
                strokeWidth={1}
                x1={valuePosition}
                x2={valuePosition}
                y1={plot.top}
                y2={plot.top + plot.height}
              />
              <text
                fill={labelColor}
                fontSize={10}
                textAnchor="middle"
                x={valuePosition}
                y={plot.top + plot.height + 14}
              >
                {formatValueTick(tick)}
              </text>
            </g>
          );
        }
        return (
          <g key={`grid-h-${tick}`}>
            <line
              stroke={lightenColor(axisColor, 0.7)}
              strokeWidth={1}
              x1={plot.left}
              x2={plot.left + plot.width}
              y1={valuePosition}
              y2={valuePosition}
            />
            <text
              fill={labelColor}
              fontSize={10}
              textAnchor="end"
              x={plot.left - 6}
              y={valuePosition + 3}
            >
              {formatValueTick(tick)}
            </text>
          </g>
        );
      })}
      {categoryPositions.map((position, index) => {
        const label = categoryLabels[index] ?? "";
        if (isHorizontal) {
          return (
            <text
              key={`cat-y-${index}`}
              fill={labelColor}
              fontSize={10}
              textAnchor="end"
              x={plot.left - 6}
              y={position + 3}
            >
              {label}
            </text>
          );
        }
        return (
          <text
            key={`cat-x-${index}`}
            fill={labelColor}
            fontSize={10}
            textAnchor="middle"
            x={position}
            y={plot.top + plot.height + 14}
          >
            {label}
          </text>
        );
      })}
      {isHorizontal ? (
        <>
          <line
            stroke={axisColor}
            strokeWidth={1.2}
            x1={plot.left}
            x2={plot.left + plot.width}
            y1={plot.top + plot.height}
            y2={plot.top + plot.height}
          />
          <line
            stroke={axisColor}
            strokeWidth={1.2}
            x1={zeroPosition}
            x2={zeroPosition}
            y1={plot.top}
            y2={plot.top + plot.height}
          />
        </>
      ) : (
        <>
          <line
            stroke={axisColor}
            strokeWidth={1.2}
            x1={plot.left}
            x2={plot.left}
            y1={plot.top}
            y2={plot.top + plot.height}
          />
          <line
            stroke={axisColor}
            strokeWidth={1.2}
            x1={plot.left}
            x2={plot.left + plot.width}
            y1={zeroPosition}
            y2={zeroPosition}
          />
        </>
      )}
    </g>
  );
}

function scaleProjectedVector(vector: { x: number; y: number }, scale: number) {
  return {
    x: vector.x * scale,
    y: vector.y * scale
  };
}

function sampleProjectedEllipseArc(
  center: { x: number; y: number },
  axisA: { x: number; y: number },
  axisB: { x: number; y: number },
  startAngle: number,
  endAngle: number
) {
  const steps = Math.max(12, Math.ceil(Math.abs(endAngle - startAngle) / (Math.PI / 14)));
  const points: Array<{ x: number; y: number }> = [];
  for (let step = 0; step <= steps; step += 1) {
    const ratio = step / steps;
    const angle = startAngle + ((endAngle - startAngle) * ratio);
    points.push({
      x: center.x + (axisA.x * Math.cos(angle)) + (axisB.x * Math.sin(angle)),
      y: center.y + (axisA.y * Math.cos(angle)) + (axisB.y * Math.sin(angle))
    });
  }
  return points;
}

function buildRibbonSvgPath(startArc: Array<{ x: number; y: number }>, endArc: Array<{ x: number; y: number }>) {
  if (startArc.length === 0 || endArc.length === 0) {
    return "";
  }
  return buildLinearSvgPath([...startArc, ...endArc.slice().reverse()], true);
}

function renderRadialFrustum(
  bar: BarRect,
  normalizedShape: "cone" | "cylinder",
  frontFill: string,
  sideFill: string,
  topFill: string
) {
  const depthX = bar.depthX ?? (bar.isHorizontal ? 10 : 9);
  const depthY = bar.depthY ?? -7;
  const frontX = bar.left + (bar.depthOffsetX ?? 0);
  const frontY = bar.top + (bar.depthOffsetY ?? 0);
  const frontW = bar.width;
  const frontH = bar.height;
  const centerX = frontX + frontW / 2;
  const centerY = frontY + frontH / 2;
  const showStartCap = bar.capStart !== false;
  const showEndCap = bar.capEnd !== false;
  const startScale = normalizedShape === "cone" ? clamp(bar.bottomScale ?? 1, 0, 1) : 1;
  const endScale = normalizedShape === "cone" ? clamp(bar.topScale ?? 1, 0, 1) : 1;
  const radialDepthScale = normalizedShape === "cylinder"
    ? (bar.isHorizontal ? 0.52 : 0.48)
    : (bar.isHorizontal ? 0.7 : 0.58);
  const depthAxis = scaleProjectedVector({ x: depthX / 2, y: depthY / 2 }, radialDepthScale);
  const secondaryStrokeWidth = Math.max(0.6, bar.strokeWidth * 0.65);

  if (bar.isHorizontal) {
    const startCenter = { x: frontX + depthAxis.x, y: centerY + depthAxis.y };
    const endCenter = { x: frontX + frontW + depthAxis.x, y: centerY + depthAxis.y };
    const startAxisA = { x: 0, y: (frontH * startScale) / 2 };
    const endAxisA = { x: 0, y: (frontH * endScale) / 2 };
    const startAxisB = scaleProjectedVector(depthAxis, startScale);
    const endAxisB = scaleProjectedVector(depthAxis, endScale);
    const backStartArc = sampleProjectedEllipseArc(startCenter, startAxisA, startAxisB, 0, Math.PI);
    const backEndArc = sampleProjectedEllipseArc(endCenter, endAxisA, endAxisB, 0, Math.PI);
    const frontStartArc = sampleProjectedEllipseArc(startCenter, startAxisA, startAxisB, Math.PI, Math.PI * 2);
    const frontEndArc = sampleProjectedEllipseArc(endCenter, endAxisA, endAxisB, Math.PI, Math.PI * 2);
    const backWallPath = buildRibbonSvgPath(backEndArc, backStartArc);
    const frontWallPath = buildRibbonSvgPath(frontEndArc, frontStartArc);
    const endCapPath = endScale > 0.018
      ? buildLinearSvgPath(sampleProjectedEllipseArc(endCenter, endAxisA, endAxisB, 0, Math.PI * 2), true)
      : "";
    const startCapPath = startScale > 0.028
      ? buildLinearSvgPath(sampleProjectedEllipseArc(startCenter, startAxisA, startAxisB, 0, Math.PI * 2), true)
      : "";
    return (
      <g key={`${bar.key}-3d-horizontal-${normalizedShape}`}>
        {backWallPath ? <path d={backWallPath} fill={sideFill} stroke={bar.stroke} strokeWidth={secondaryStrokeWidth} /> : null}
        {showEndCap && endCapPath ? <path d={endCapPath} fill={topFill} stroke={bar.stroke} strokeWidth={secondaryStrokeWidth} /> : null}
        {frontWallPath ? <path d={frontWallPath} fill={frontFill} stroke={bar.stroke} strokeWidth={bar.strokeWidth} /> : null}
        {showStartCap && startCapPath ? <path d={startCapPath} fill={frontFill} stroke={bar.stroke} strokeWidth={secondaryStrokeWidth} /> : null}
      </g>
    );
  }

  const startCenter = { x: centerX + depthAxis.x, y: frontY + frontH + depthAxis.y };
  const endCenter = { x: centerX + depthAxis.x, y: frontY + depthAxis.y };
  const startAxisA = { x: (frontW * startScale) / 2, y: 0 };
  const endAxisA = { x: (frontW * endScale) / 2, y: 0 };
  const startAxisB = scaleProjectedVector(depthAxis, startScale);
  const endAxisB = scaleProjectedVector(depthAxis, endScale);
  const backStartArc = sampleProjectedEllipseArc(startCenter, startAxisA, startAxisB, 0, Math.PI);
  const backEndArc = sampleProjectedEllipseArc(endCenter, endAxisA, endAxisB, 0, Math.PI);
  const frontStartArc = sampleProjectedEllipseArc(startCenter, startAxisA, startAxisB, Math.PI, Math.PI * 2);
  const frontEndArc = sampleProjectedEllipseArc(endCenter, endAxisA, endAxisB, Math.PI, Math.PI * 2);
  const backWallPath = buildRibbonSvgPath(backEndArc, backStartArc);
  const frontWallPath = buildRibbonSvgPath(frontEndArc, frontStartArc);
  const endCapPath = endScale > 0.018
    ? buildLinearSvgPath(sampleProjectedEllipseArc(endCenter, endAxisA, endAxisB, 0, Math.PI * 2), true)
    : "";
  return (
    <g key={`${bar.key}-3d-${normalizedShape}`}>
      {backWallPath ? <path d={backWallPath} fill={sideFill} stroke={bar.stroke} strokeWidth={secondaryStrokeWidth} /> : null}
      {showEndCap && endCapPath ? <path d={endCapPath} fill={topFill} stroke={bar.stroke} strokeWidth={secondaryStrokeWidth} /> : null}
      {frontWallPath ? <path d={frontWallPath} fill={frontFill} stroke={bar.stroke} strokeWidth={bar.strokeWidth} /> : null}
    </g>
  );
}

function renderExtrudedRect(bar: BarRect) {
  const normalizedShape = (() => {
    switch (bar.shape3d) {
      case "cone":
      case "coneToMax":
        return "cone";
      case "cylinder":
        return "cylinder";
      case "pyramid":
      case "pyramidToMax":
        return "pyramid";
      default:
        return "box";
    }
  })();
  const depthX = bar.depthX ?? (bar.isHorizontal ? 10 : 9);
  const depthY = bar.depthY ?? -7;
  const frontX = bar.left + (bar.depthOffsetX ?? 0);
  const frontY = bar.top + (bar.depthOffsetY ?? 0);
  const frontW = bar.width;
  const frontH = bar.height;
  const frontX2 = frontX + frontW;
  const frontY2 = frontY + frontH;

  const sideAnchorX = frontX2;
  const sideDepthX = depthX;
  const centerX = frontX + frontW / 2;
  const bottomScale = clamp(bar.bottomScale ?? 1, 0.04, 1);
  const topScale = clamp(bar.topScale ?? 1, 0.04, 1);
  const bottomHalfWidth = (frontW * bottomScale) / 2;
  const topHalfWidth = (frontW * topScale) / 2;
  const bottomLeft = centerX - bottomHalfWidth;
  const bottomRight = centerX + bottomHalfWidth;
  const topLeft = centerX - topHalfWidth;
  const topRight = centerX + topHalfWidth;
  const showStartCap = bar.capStart !== false;
  const showEndCap = bar.capEnd !== false;

  const topFace = `${frontX},${frontY} ${frontX2},${frontY} ${frontX2 + sideDepthX},${frontY + depthY} ${frontX + sideDepthX},${frontY + depthY}`;
  const sideFace = `${sideAnchorX},${frontY} ${sideAnchorX},${frontY2} ${sideAnchorX + sideDepthX},${frontY2 + depthY} ${sideAnchorX + sideDepthX},${frontY + depthY}`;
  const frontFill = bar.gradientId ? `url(#${bar.gradientId})` : bar.color;
  const sideFill = bar.invertedNegative ? bar.color : darkenColor(bar.color, 0.22);
  const topFill = bar.invertedNegative ? lightenColor(bar.color, 0.04) : lightenColor(bar.color, 0.24);

  if (normalizedShape === "cylinder" || normalizedShape === "cone") {
    return renderRadialFrustum(bar, normalizedShape, frontFill, sideFill, topFill);
  }

  if (normalizedShape === "box") {
    return (
      <g key={`${bar.key}-3d`}>
        <polygon fill={sideFill} points={sideFace} stroke={bar.stroke} strokeWidth={Math.max(0.6, bar.strokeWidth * 0.65)} />
        <polygon fill={topFill} points={topFace} stroke={bar.stroke} strokeWidth={Math.max(0.6, bar.strokeWidth * 0.65)} />
        <rect fill={frontFill} height={frontH} stroke={bar.stroke} strokeWidth={bar.strokeWidth} width={frontW} x={frontX} y={frontY} />
      </g>
    );
  }

  if (bar.isHorizontal) {
    const centerY = frontY + frontH / 2;
    const startHalfHeight = (frontH * clamp(bar.bottomScale ?? 1, 0.04, 1)) / 2;
    const endHalfHeight = (frontH * clamp(bar.topScale ?? 1, 0.04, 1)) / 2;
    const startTop = centerY - startHalfHeight;
    const startBottom = centerY + startHalfHeight;
    const endTop = centerY - endHalfHeight;
    const endBottom = centerY + endHalfHeight;
    const topFacePoints = `${frontX},${startTop} ${frontX2},${endTop} ${frontX2 + sideDepthX},${endTop + depthY} ${frontX + sideDepthX},${startTop + depthY}`;
    const farSidePoints = `${frontX2},${endTop} ${frontX2},${endBottom} ${frontX2 + sideDepthX},${endBottom + depthY} ${frontX2 + sideDepthX},${endTop + depthY}`;

    const taperedFarFace = `${frontX2},${endTop} ${frontX2 + sideDepthX},${endTop + depthY} ${frontX2 + sideDepthX},${endBottom + depthY} ${frontX2},${endBottom}`;
    const frontPolygon = `${frontX},${startTop} ${frontX2},${endTop} ${frontX2},${endBottom} ${frontX},${startBottom}`;

    return (
      <g key={`${bar.key}-3d-horizontal-${normalizedShape}`}>
        <polygon fill={sideFill} points={taperedFarFace} stroke={bar.stroke} strokeWidth={Math.max(0.6, bar.strokeWidth * 0.65)} />
        <polygon fill={topFill} points={topFacePoints} stroke={bar.stroke} strokeWidth={Math.max(0.6, bar.strokeWidth * 0.65)} />
        <polygon fill={frontFill} points={frontPolygon} stroke={bar.stroke} strokeWidth={bar.strokeWidth} />
      </g>
    );
  }

  const taperedTopFace = `${topLeft},${frontY} ${topRight},${frontY} ${topRight + sideDepthX},${frontY + depthY} ${topLeft + sideDepthX},${frontY + depthY}`;
  const taperedSideFace = `${topRight},${frontY} ${bottomRight},${frontY2} ${bottomRight + sideDepthX},${frontY2 + depthY} ${topRight + sideDepthX},${frontY + depthY}`;
  const pyramidFrontFace = `${topLeft},${frontY} ${topRight},${frontY} ${bottomRight},${frontY2} ${bottomLeft},${frontY2}`;

  return (
    <g key={`${bar.key}-3d-${normalizedShape}`}>
      <polygon fill={sideFill} points={taperedSideFace} stroke={bar.stroke} strokeWidth={Math.max(0.6, bar.strokeWidth * 0.65)} />
      {showEndCap ? (
        <polygon fill={topFill} points={taperedTopFace} stroke={bar.stroke} strokeWidth={Math.max(0.6, bar.strokeWidth * 0.65)} />
      ) : null}
      <polygon fill={frontFill} points={pyramidFrontFace} stroke={bar.stroke} strokeWidth={bar.strokeWidth} />
    </g>
  );
}

function buildLinearSvgPath(points: Array<{ x: number; y: number }>, close = false) {
  if (points.length === 0) {
    return "";
  }
  const commands = points.map((point, index) => (
    `${index === 0 ? "M" : "L"} ${toSvgNumber(point.x)} ${toSvgNumber(point.y)}`
  ));
  if (close) {
    commands.push("Z");
  }
  return commands.join(" ");
}

function projectCartesian3dPoint(
  x: number,
  y: number,
  z: number,
  rotXRad: number,
  rotYRad: number,
  usePerspective: boolean,
  perspectiveStrength: number
) {
  const cosX = Math.cos(rotXRad);
  const sinX = Math.sin(rotXRad);
  const cosY = Math.cos(rotYRad);
  const sinY = Math.sin(rotYRad);

  const x1 = x * cosY + z * sinY;
  const z1 = -x * sinY + z * cosY;
  const y1 = y * cosX - z1 * sinX;
  const z2 = y * sinX + z1 * cosX;
  const perspective = usePerspective
    ? 1 / Math.max(0.22, 1 + z2 * (0.26 + perspectiveStrength * 0.54))
    : 1;

  return {
    depth: z2,
    x: x1 * perspective,
    y: y1 * perspective
  };
}

function renderLineOrAreaChart3d(
  chart: XlsxChart,
  palette: ChartRendererPalette,
  layout: ChartLayout,
  categories: string[],
  stackedPointsBySeries: Array<Array<{ defined: boolean; y: number | null; y0: number | null; y1: number | null }>>,
  minValue: number,
  maxValue: number,
  isAreaChart: boolean,
  isStackedSeries: boolean,
  selectedChartElement?: XlsxChartElementSelection | null
) {
  const plot = layout.plot;
  const valueDomain = resolveAxisDomainWithChartOverrides(
    chart.valueAxis,
    minValue,
    maxValue,
    isAreaChart || chart.valueAxis?.crosses === "autoZero"
  );
  minValue = valueDomain.min;
  maxValue = valueDomain.max;
  const valueSpan = Math.max(1e-6, maxValue - minValue);
  const seriesCount = Math.max(1, chart.series.length);
  const categoryCount = Math.max(1, categories.length);
  const depthScale = clamp((chart.view3d?.depthPercent ?? 100) / 100, 0.5, 4);
  const halfDepth = (seriesCount <= 1 ? 0.82 : 1.02) * depthScale;
  const frontZ = halfDepth;
  const backZ = -halfDepth;
  const baseValue = isAreaChart ? Math.max(minValue, Math.min(maxValue, 0)) : minValue;
  const rotXRad = clamp(chart.view3d?.rotX ?? 18, -80, 80) * (Math.PI / 180);
  const rotYRad = clamp(chart.view3d?.rotY ?? 24, -80, 80) * (Math.PI / 180);
  const usePerspective = chart.view3d?.rAngAx === false;
  const perspectiveStrength = clamp((chart.view3d?.perspective ?? (usePerspective ? 26 : 0)) / 100, 0, 1);

  const normalizeX = (categoryIndex: number) => (
    categoryCount <= 1 ? 0 : ((categoryIndex / (categoryCount - 1)) - 0.5) * 2
  );
  const normalizeY = (value: number) => (
    -((((value - minValue) / valueSpan) - 0.5) * 2)
  );
  const stackedAreaDepthSpan = isAreaChart && isStackedSeries
    ? Math.max(0.18, (frontZ - backZ) * 0.56)
    : 0;
  const stackedAreaBackZ = frontZ - stackedAreaDepthSpan;
  const normalizeZ = (seriesIndex: number) => (
    seriesCount <= 1 || (isAreaChart && isStackedSeries)
      ? frontZ * 0.94
      : (((seriesIndex / (seriesCount - 1)) - 0.5) * (frontZ - backZ))
  );
  const projectSeriesPoint = (x: number, value: number, z: number) => projectCartesian3dPoint(
    x,
    normalizeY(value),
    z,
    rotXRad,
    rotYRad,
    usePerspective,
    perspectiveStrength
  );

  const cubeCorners = [
    { x: -1, y: -1, z: backZ },
    { x: 1, y: -1, z: backZ },
    { x: 1, y: 1, z: backZ },
    { x: -1, y: 1, z: backZ },
    { x: -1, y: -1, z: frontZ },
    { x: 1, y: -1, z: frontZ },
    { x: 1, y: 1, z: frontZ },
    { x: -1, y: 1, z: frontZ }
  ].map((corner) => projectCartesian3dPoint(
    corner.x,
    corner.y,
    corner.z,
    rotXRad,
    rotYRad,
    usePerspective,
    perspectiveStrength
  ));

  const projectedSeries = stackedPointsBySeries.map((seriesPoints, seriesIndex) => {
    const z = normalizeZ(seriesIndex);
    return seriesPoints.map((point, categoryIndex) => {
      const x = normalizeX(categoryIndex);
      const topValue = isStackedSeries ? (point.y1 ?? point.y ?? baseValue) : (point.y ?? baseValue);
      const bottomValue = isAreaChart
        ? (isStackedSeries ? (point.y0 ?? baseValue) : baseValue)
        : baseValue;
      const top = projectSeriesPoint(x, topValue, z);
      const bottom = projectSeriesPoint(x, bottomValue, z);
      const topBack = isAreaChart && isStackedSeries
        ? projectSeriesPoint(x, topValue, stackedAreaBackZ)
        : null;
      const bottomBack = isAreaChart && isStackedSeries
        ? projectSeriesPoint(x, bottomValue, stackedAreaBackZ)
        : null;
      return {
        bottom,
        bottomBack,
        categoryIndex,
        defined: point.defined,
        depth: top.depth,
        depthBack: topBack?.depth ?? top.depth,
        top,
        topBack
      };
    });
  });

  const bounds = [
    ...cubeCorners,
    ...projectedSeries.flatMap((series) => series.flatMap((point) => point.defined
      ? [point.top, point.bottom, ...(point.topBack ? [point.topBack] : []), ...(point.bottomBack ? [point.bottomBack] : [])]
      : []))
  ];
  const minX = Math.min(...bounds.map((point) => point.x));
  const maxX = Math.max(...bounds.map((point) => point.x));
  const minY = Math.min(...bounds.map((point) => point.y));
  const maxY = Math.max(...bounds.map((point) => point.y));
  const scale = Math.min(
    plot.width / Math.max(0.4, maxX - minX),
    plot.height / Math.max(0.4, maxY - minY)
  ) * 0.82;
  const centerRawX = (minX + maxX) / 2;
  const centerRawY = (minY + maxY) / 2;
  const centerX = plot.left + plot.width / 2;
  const centerY = plot.top + plot.height / 2;
  const toScreenPoint = (point: { depth: number; x: number; y: number }) => ({
    depth: point.depth,
    x: centerX + (point.x - centerRawX) * scale,
    y: centerY + (point.y - centerRawY) * scale
  });

  const screenCorners = cubeCorners.map(toScreenPoint);
  const seriesGeometry = projectedSeries.map((series, seriesIndex) => ({
    averageDepth: series.reduce((sum, point) => sum + ((point.depth + point.depthBack) / 2), 0) / Math.max(1, series.length),
    points: series.map((point) => ({
      bottom: toScreenPoint(point.bottom),
      bottomBack: point.bottomBack ? toScreenPoint(point.bottomBack) : null,
      categoryIndex: point.categoryIndex,
      defined: point.defined,
      top: toScreenPoint(point.top),
      topBack: point.topBack ? toScreenPoint(point.topBack) : null
    })),
    seriesIndex
  })).sort((left, right) => left.averageDepth - right.averageDepth);

  const edgeColor = chart.axisLineColor ?? chart.chartAreaBorderColor ?? palette.border;
  const gridColor = lightenColor(edgeColor, 0.56);
  const labelColor = resolveChartAxisTextColor(chart);
  const boxEdges: Array<[number, number, boolean]> = [
    [0, 1, false], [1, 2, false], [2, 3, false], [3, 0, false],
    [4, 5, true], [5, 6, true], [6, 7, true], [7, 4, true],
    [0, 4, false], [1, 5, true], [2, 6, true], [3, 7, false]
  ];
  const yTicks = valueDomain.ticks;
  const xLabelPoints = categories.map((category, categoryIndex) => ({
    label: category,
    point: toScreenPoint(projectCartesian3dPoint(
      normalizeX(categoryIndex),
      -1.06,
      frontZ,
      rotXRad,
      rotYRad,
      usePerspective,
      perspectiveStrength
    ))
  }));
  const yLabelPoints = yTicks.map((tick) => ({
    point: toScreenPoint(projectCartesian3dPoint(
      -1.04,
      normalizeY(tick),
      frontZ,
      rotXRad,
      rotYRad,
      usePerspective,
      perspectiveStrength
    )),
    tick
  }));

  return (
    <g>
      {yTicks.map((tick) => {
        const yNorm = normalizeY(tick);
        const leftBack = toScreenPoint(projectCartesian3dPoint(-1, yNorm, backZ, rotXRad, rotYRad, usePerspective, perspectiveStrength));
        const rightBack = toScreenPoint(projectCartesian3dPoint(1, yNorm, backZ, rotXRad, rotYRad, usePerspective, perspectiveStrength));
        const leftFront = toScreenPoint(projectCartesian3dPoint(-1, yNorm, frontZ, rotXRad, rotYRad, usePerspective, perspectiveStrength));
        return (
          <g key={`line3d-grid-${tick}`}>
            <line stroke={gridColor} strokeWidth={1} x1={leftBack.x} x2={rightBack.x} y1={leftBack.y} y2={rightBack.y} />
            <line stroke={gridColor} strokeWidth={0.9} x1={leftBack.x} x2={leftFront.x} y1={leftBack.y} y2={leftFront.y} />
          </g>
        );
      })}
      {boxEdges.map(([startIndex, endIndex, emphasized], index) => {
        const start = screenCorners[startIndex];
        const end = screenCorners[endIndex];
        return (
          <line
            key={`line3d-box-${index}`}
            stroke={emphasized ? edgeColor : lightenColor(edgeColor, 0.34)}
            strokeWidth={emphasized ? 1.35 : 1}
            x1={start.x}
            x2={end.x}
            y1={start.y}
            y2={end.y}
          />
        );
      })}
      {seriesGeometry.map(({ points, seriesIndex }) => {
        const definedPoints = points.filter((point) => point.defined);
        if (definedPoints.length === 0) {
          return null;
        }
        const linePoints = definedPoints.map((point) => point.top);
        const areaPoints = isAreaChart
          ? [
              ...definedPoints.map((point) => point.top),
              ...definedPoints.slice().reverse().map((point) => point.bottom)
            ]
          : [];
        const areaBackPoints = isAreaChart && isStackedSeries
          ? [
              ...definedPoints.map((point) => point.topBack ?? point.top),
              ...definedPoints.slice().reverse().map((point) => point.bottomBack ?? point.bottom)
            ]
          : [];
        const strokeColor = chartSeriesStrokeColor(chart, seriesIndex);
        const fillColor = chartSeriesColor(chart, seriesIndex);
        const markerSymbol = normalizeChartMarkerSymbol(chart.series[seriesIndex]?.markerSymbol);
        const markerPath = markerSymbolPath(markerSymbol, markerRadiusPx(chart.series[seriesIndex]?.markerSize));
        const slabFaces = isAreaChart && isStackedSeries
          ? definedPoints.slice(1).map((point, pointIndex) => {
              const previous = definedPoints[pointIndex];
              if (!previous?.topBack || !point.topBack || !previous.bottomBack || !point.bottomBack) {
                return null;
              }
              const topFace = buildLinearSvgPath([previous.top, point.top, point.topBack, previous.topBack], true);
              const bottomFace = buildLinearSvgPath([previous.bottom, point.bottom, point.bottomBack, previous.bottomBack], true);
              return (
                <React.Fragment key={`line3d-area-face-${seriesIndex}-${pointIndex}`}>
                  <path
                    {...chartElementDataProps(seriesIndex)}
                    d={topFace}
                    fill={lightenColor(fillColor, 0.08)}
                    fillOpacity={0.8}
                    stroke={darkenColor(fillColor, 0.16)}
                    strokeWidth={0.8}
                  />
                  <path
                    {...chartElementDataProps(seriesIndex)}
                    d={bottomFace}
                    fill={darkenColor(fillColor, 0.2)}
                    fillOpacity={0.34}
                    stroke={darkenColor(fillColor, 0.24)}
                    strokeWidth={0.6}
                  />
                </React.Fragment>
              );
            })
          : [];
        const firstDefinedPoint = definedPoints[0] ?? null;
        const lastDefinedPoint = definedPoints[definedPoints.length - 1] ?? null;
        const startCap = isAreaChart && isStackedSeries && firstDefinedPoint?.topBack && firstDefinedPoint?.bottomBack
          ? buildLinearSvgPath([
              firstDefinedPoint.top,
              firstDefinedPoint.bottom,
              firstDefinedPoint.bottomBack,
              firstDefinedPoint.topBack
            ], true)
          : "";
        const endCap = isAreaChart && isStackedSeries && lastDefinedPoint?.topBack && lastDefinedPoint?.bottomBack
          ? buildLinearSvgPath([
              lastDefinedPoint.top,
              lastDefinedPoint.bottom,
              lastDefinedPoint.bottomBack,
              lastDefinedPoint.topBack
            ], true)
          : "";

        return (
          <g key={`line3d-series-${seriesIndex}`}>
            {isAreaChart && isStackedSeries && areaBackPoints.length >= 3 ? (
              <path
                {...chartElementDataProps(seriesIndex)}
                d={buildLinearSvgPath(areaBackPoints, true)}
                fill={darkenColor(fillColor, 0.18)}
                fillOpacity={0.44}
                stroke={darkenColor(fillColor, 0.26)}
                strokeWidth={0.8}
              />
            ) : null}
            {slabFaces}
            {startCap ? (
              <path
                {...chartElementDataProps(seriesIndex)}
                d={startCap}
                fill={darkenColor(fillColor, 0.24)}
                fillOpacity={0.54}
                stroke={darkenColor(fillColor, 0.3)}
                strokeWidth={0.7}
              />
            ) : null}
            {endCap ? (
              <path
                {...chartElementDataProps(seriesIndex)}
                d={endCap}
                fill={darkenColor(fillColor, 0.14)}
                fillOpacity={0.6}
                stroke={darkenColor(fillColor, 0.24)}
                strokeWidth={0.7}
              />
            ) : null}
            {isAreaChart && areaPoints.length >= 3 ? (
              <path
                {...chartElementDataProps(seriesIndex)}
                d={buildLinearSvgPath(areaPoints, true)}
                fill={fillColor}
                fillOpacity={0.74}
                stroke={darkenColor(fillColor, 0.12)}
                strokeWidth={0.9}
              />
            ) : null}
            {!isAreaChart
              ? definedPoints.map((point, pointIndex) => (
                  <line
                    key={`line3d-drop-${seriesIndex}-${pointIndex}`}
                    stroke={lightenColor(strokeColor, 0.44)}
                    strokeDasharray="2 2"
                    strokeWidth={0.9}
                    x1={point.top.x}
                    x2={point.bottom.x}
                    y1={point.top.y}
                    y2={point.bottom.y}
                  />
                ))
              : null}
            <path
              {...chartElementDataProps(seriesIndex)}
              d={buildLinearSvgPath(linePoints)}
              fill="none"
              stroke={strokeColor}
              strokeLinejoin="round"
              strokeWidth={Math.max(1.8, chart.series[seriesIndex]?.lineWidthPx ?? 2)}
            />
            {markerPath.length > 0
              ? definedPoints.map((point, pointIndex) => (
                  <g
                    key={`line3d-marker-${seriesIndex}-${pointIndex}`}
                    transform={`translate(${toSvgNumber(point.top.x)}, ${toSvgNumber(point.top.y)})`}
                  >
                    <path
                      {...chartElementDataProps(seriesIndex, point.categoryIndex)}
                      d={markerPath}
                      fill={chart.series[seriesIndex]?.markerColor ?? fillColor}
                      stroke={chart.series[seriesIndex]?.markerLineColor ?? chart.chartAreaFillColor ?? palette.surface}
                      strokeWidth={1}
                    />
                  </g>
                ))
              : null}
            {definedPoints.map((point) => (
              <circle
                {...chartElementDataProps(seriesIndex, point.categoryIndex)}
                cx={point.top.x}
                cy={point.top.y}
                fill="transparent"
                key={`line3d-hit-${seriesIndex}-${point.categoryIndex}`}
                r={7}
                stroke="none"
              />
            ))}
            {definedPoints.map((point) => (
              isSelectedChartPointOrSeries(selectedChartElement, chart.id, seriesIndex, point.categoryIndex)
                ? renderSelectionPointHandles(
                    `line3d-selection-${seriesIndex}-${point.categoryIndex}`,
                    [{ x: point.top.x, y: point.top.y }]
                  )
                : null
            ))}
          </g>
        );
      })}
      {yLabelPoints.map(({ point, tick }) => (
        <text
          key={`line3d-y-label-${tick}`}
          fill={labelColor}
          fontSize={10}
          textAnchor="end"
          x={point.x - 6}
          y={point.y + 3}
        >
          {formatTickValue(tick)}
        </text>
      ))}
      {xLabelPoints.map(({ label, point }, index) => (
        <text
          key={`line3d-x-label-${index}`}
          fill={labelColor}
          fontSize={10}
          textAnchor="middle"
          x={point.x}
          y={point.y + 14}
        >
          {label}
        </text>
      ))}
    </g>
  );
}

function renderBarChart(chart: XlsxChart, palette: ChartRendererPalette, layout: ChartLayout, chartType: string, selectedChartElement?: XlsxChartElementSelection | null) {
  if (chart.series.length === 0) {
    return null;
  }
  const sourceCategories = getCategoryLabels(chart);
  if (sourceCategories.length === 0) {
    return null;
  }
  const isPercentStacked = chartType === "ColumnPercentStacked" || chartType === "BarPercentStacked";
  const isHorizontal = chartType === "BarClustered" || chartType === "BarStacked" || chartType === "BarPercentStacked";
  const isHistogramLike = isHistogramLikeChart(chart) && chartType === "ColumnClustered";
  const shouldReverseCategories = isHorizontal && chart.categoryAxis?.orientation !== "maxMin";
  const categories = shouldReverseCategories ? sourceCategories.slice().reverse() : sourceCategories;
  const isStacked = chartType === "ColumnStacked" || chartType === "BarStacked" || isPercentStacked;
  const negativeFillMode: "chartArea" | "none" | "series" = isPercentStacked
    ? "chartArea"
    : chartType === "ColumnClustered"
      ? "none"
      : "series";
  const categoryCount = categories.length;
  const seriesCount = chart.series.length;
  const normalized3dShape = (() => {
    switch (chart.shape3d) {
      case "cone":
      case "coneToMax":
        return "cone";
      case "cylinder":
        return "cylinder";
      case "pyramid":
      case "pyramidToMax":
        return "pyramid";
      default:
        return "box";
    }
  })();
  const frameOffsets = chart.is3d ? resolve3dFrameOffsets(chart, isHorizontal ? 9 : 11, isHorizontal ? 7 : 8) : null;
  const plot = chart.is3d && frameOffsets
    ? {
        ...layout.plot,
        height: Math.max(20, layout.plot.height - frameOffsets.insetBottom - 2),
        left: layout.plot.left + frameOffsets.insetLeft,
        top: layout.plot.top + frameOffsets.insetTop,
        width: Math.max(20, layout.plot.width - frameOffsets.insetLeft - frameOffsets.insetRight - 2)
      }
    : layout.plot;

  const matrix = chart.series.map((series) => {
    const values = Array.from(
      { length: categoryCount },
      (_, categoryIndex) => safeNumber(series.values[categoryIndex]) ?? 0
    );
    return shouldReverseCategories ? values.reverse() : values;
  });
  const rawMatrix = matrix.map((row) => row.slice());

  if (isPercentStacked) {
    for (let categoryIndex = 0; categoryIndex < categoryCount; categoryIndex += 1) {
      const total = matrix.reduce((sum, row) => sum + Math.max(0, row[categoryIndex] ?? 0), 0);
      for (let seriesIndex = 0; seriesIndex < seriesCount; seriesIndex += 1) {
        matrix[seriesIndex][categoryIndex] = total > 0 ? ((matrix[seriesIndex][categoryIndex] ?? 0) / total) * 100 : 0;
      }
    }
  }

  const positiveTotals = Array.from({ length: categoryCount }, (_, categoryIndex) => (
    matrix.reduce((sum, row) => sum + Math.max(0, row[categoryIndex] ?? 0), 0)
  ));
  const negativeTotals = Array.from({ length: categoryCount }, (_, categoryIndex) => (
    Math.abs(matrix.reduce((sum, row) => sum + Math.min(0, row[categoryIndex] ?? 0), 0))
  ));

  let minValue = Number.POSITIVE_INFINITY;
  let maxValue = Number.NEGATIVE_INFINITY;
  if (isStacked) {
    for (let categoryIndex = 0; categoryIndex < categoryCount; categoryIndex += 1) {
      let positive = 0;
      let negative = 0;
      for (let seriesIndex = 0; seriesIndex < seriesCount; seriesIndex += 1) {
        const value = matrix[seriesIndex][categoryIndex] ?? 0;
        if (value >= 0) {
          positive += value;
        } else {
          negative += value;
        }
      }
      maxValue = Math.max(maxValue, positive);
      minValue = Math.min(minValue, negative);
    }
  } else {
    for (const row of matrix) {
      for (const value of row) {
        maxValue = Math.max(maxValue, value);
        minValue = Math.min(minValue, value);
      }
    }
  }
  if (!Number.isFinite(minValue) || !Number.isFinite(maxValue)) {
    minValue = 0;
    maxValue = 1;
  }

  const valueDomain = resolveAxisDomainWithChartOverrides(
    chart.valueAxis,
    minValue,
    maxValue,
    true
  );
  minValue = valueDomain.min;
  maxValue = valueDomain.max;

  const categoryBandPadding = isHistogramLike
    ? { inner: 0, outer: 0 }
    : resolveCategoryBandPadding(chart.gapWidth);
  const categoryScale = scaleBand<string>()
    .domain(categories)
    .range(isHorizontal ? [plot.top, plot.top + plot.height] : [plot.left, plot.left + plot.width])
    .paddingInner(categoryBandPadding.inner)
    .paddingOuter(categoryBandPadding.outer);

  const seriesScale = scaleBand<string>()
    .domain(Array.from({ length: seriesCount }, (_, index) => String(index)))
    .range([0, categoryScale.bandwidth()])
    .paddingInner(isHistogramLike ? 0 : 0.16)
    .paddingOuter(isHistogramLike ? 0 : 0.08);
  const usesSeriesDepthAxis = (
    chart.is3d === true
    && !isHorizontal
    && !isStacked
    && seriesCount > 1
    && chart.seriesAxis != null
  );
  const depthLaneSpread = usesSeriesDepthAxis
    ? clamp(1.26 + seriesCount * 0.16, 1.4, 1.96)
    : 1;
  const depthGridSpanX = usesSeriesDepthAxis && frameOffsets
    ? frameOffsets.depthX * depthLaneSpread
    : 0;
  const depthGridSpanY = usesSeriesDepthAxis && frameOffsets
    ? frameOffsets.depthY * depthLaneSpread
    : 0;
  const depthSlotX = usesSeriesDepthAxis ? depthGridSpanX / Math.max(1, seriesCount) : 0;
  const depthSlotY = usesSeriesDepthAxis ? depthGridSpanY / Math.max(1, seriesCount) : 0;
  const depthBarX = usesSeriesDepthAxis
    ? Math.sign(depthSlotX || frameOffsets?.depthX || 1) * Math.max(8, Math.abs(depthSlotX) * 1.14)
    : frameOffsets?.depthX;
  const depthBarY = usesSeriesDepthAxis
    ? Math.sign(depthSlotY || frameOffsets?.depthY || -1) * Math.max(5.5, Math.abs(depthSlotY) * 1.12)
    : frameOffsets?.depthY;

  const shouldReverseValueAxis = chart.valueAxis?.orientation === "maxMin";
  const valueScale = scaleLinear()
    .domain([minValue, maxValue])
    .range(
      isHorizontal
        ? (shouldReverseValueAxis ? [plot.left + plot.width, plot.left] : [plot.left, plot.left + plot.width])
        : (shouldReverseValueAxis ? [plot.top, plot.top + plot.height] : [plot.top + plot.height, plot.top])
    );

  const ticks = isPercentStacked
    ? buildNumericTickValues(minValue, maxValue, chart.valueAxis?.majorUnit ?? 20)
    : valueDomain.ticks;
  const categoryPositions = categories.map((category) => {
    const bandStart = categoryScale(category) ?? 0;
    return bandStart + categoryScale.bandwidth() / 2;
  });

  const bars: BarRect[] = [];
  const positive = Array.from({ length: categoryCount }, () => 0);
  const negative = Array.from({ length: categoryCount }, () => 0);

  for (let seriesIndex = 0; seriesIndex < seriesCount; seriesIndex += 1) {
    for (let categoryIndex = 0; categoryIndex < categoryCount; categoryIndex += 1) {
      const value = matrix[seriesIndex][categoryIndex] ?? 0;
      const rawValue = rawMatrix[seriesIndex][categoryIndex] ?? value;
      const pointIndex = shouldReverseCategories ? (categoryCount - 1 - categoryIndex) : categoryIndex;
      const pointStyle = chart.series[seriesIndex]?.dataPointStyles?.find((entry) => entry.index === pointIndex);
      let start = 0;
      let end = value;
      if (isStacked) {
        if (value >= 0) {
          start = positive[categoryIndex];
          positive[categoryIndex] += value;
          end = positive[categoryIndex];
        } else {
          start = negative[categoryIndex];
          negative[categoryIndex] += value;
          end = negative[categoryIndex];
        }
      }

      const colors = chartSeriesBarColors(chart, seriesIndex, rawValue, negativeFillMode);
      const isInvertedNegative = rawValue < 0 && chart.series[seriesIndex]?.invertIfNegative === true;
      if (!(isInvertedNegative)) {
        if (variesBarColorsByCategory(chart)) {
          colors.fill = chartPointColor(chart, pointIndex, seriesIndex);
          colors.stroke = colors.fill;
        }
        if (pointStyle?.color) {
          colors.fill = pointStyle.color;
          if (!pointStyle.lineColor) {
            colors.stroke = pointStyle.color;
          }
        }
        if (pointStyle?.lineColor) {
          colors.stroke = pointStyle.lineColor;
        }
      }
      if (isInvertedNegative && isPercentStacked && chart.is3d) {
        colors.fill = "#ffffff";
      }
      const category = categories[categoryIndex];
      const categoryStart = categoryScale(category) ?? 0;
      const barThickness = usesSeriesDepthAxis
        ? Math.max(7, categoryScale.bandwidth() * 0.4)
        : isHistogramLike
          ? categoryScale.bandwidth()
        : isStacked
          ? categoryScale.bandwidth()
          : seriesScale.bandwidth();
      const barOffset = usesSeriesDepthAxis
        ? Math.max(0, (categoryScale.bandwidth() - barThickness) * 0.5)
        : isHistogramLike
          ? 0
        : isStacked
          ? 0
          : (seriesScale(String(seriesIndex)) ?? 0);
      const depthOffsetX = usesSeriesDepthAxis ? depthSlotX * seriesIndex : 0;
      const depthOffsetY = usesSeriesDepthAxis ? depthSlotY * seriesIndex : 0;
      const depthOrder = depthOffsetX * (frameOffsets?.depthX ?? 0) + depthOffsetY * (frameOffsets?.depthY ?? 0);
      const shapeTaper = normalized3dShape === "pyramid"
        ? 0.94
        : normalized3dShape === "cone"
          ? 1
          : 0;
      let topScale = 1;
      let bottomScale = 1;
      if (chart.is3d && shapeTaper > 0) {
        const scaleAt = (ratio: number) => {
          const clampedRatio = clamp(ratio, 0, 1);
          const taperedRatio = normalized3dShape === "cone"
            ? Math.pow(clampedRatio, isHorizontal ? 1.85 : 1.08)
            : clampedRatio;
          return clamp(1 - (taperedRatio * shapeTaper), normalized3dShape === "cone" ? 0 : 0.04, 1);
        };
        if (rawValue >= 0) {
          const total = Math.max(1e-6, isStacked ? positiveTotals[categoryIndex] : Math.abs(rawValue));
          bottomScale = scaleAt(isStacked ? start / total : 0);
          topScale = scaleAt(isStacked ? end / total : 1);
        } else {
          const total = Math.max(1e-6, isStacked ? negativeTotals[categoryIndex] : Math.abs(rawValue));
          topScale = scaleAt(isStacked ? Math.abs(start) / total : 0);
          bottomScale = scaleAt(isStacked ? Math.abs(end) / total : 1);
        }
      }

      if (isHorizontal) {
        const x1 = valueScale(start);
        const x2 = valueScale(end);
        const maxPositive = positiveTotals[categoryIndex] ?? 0;
        const maxNegative = negativeTotals[categoryIndex] ?? 0;
        bars.push({
          capEnd: !isStacked || (rawValue >= 0
            ? Math.abs(end - maxPositive) < 1e-6
            : Math.abs(end - maxNegative) < 1e-6),
          capStart: !isStacked || Math.abs(start) < 1e-6,
          bottomScale,
          categoryIndex,
          color: colors.fill,
          depthOrder,
          depthOffsetX,
          depthOffsetY,
          height: Math.max(1, barThickness),
          isHorizontal: true,
          key: `bar-${seriesIndex}-${categoryIndex}`,
          left: Math.min(x1, x2),
          depthX: usesSeriesDepthAxis ? depthBarX : frameOffsets?.depthX,
          depthY: usesSeriesDepthAxis ? depthBarY : frameOffsets?.depthY,
          shape3d: normalized3dShape,
          seriesIndex,
          stroke: colors.stroke,
          strokeWidth: colors.strokeWidth,
          topScale,
          top: categoryStart + barOffset,
          invertedNegative: isInvertedNegative,
          value: rawValue,
          width: Math.max(1, Math.abs(x2 - x1))
        });
      } else {
        const y1 = valueScale(start);
        const y2 = valueScale(end);
        const maxPositive = positiveTotals[categoryIndex] ?? 0;
        const maxNegative = negativeTotals[categoryIndex] ?? 0;
        bars.push({
          capEnd: !isStacked || (rawValue >= 0
            ? Math.abs(end - maxPositive) < 1e-6
            : Math.abs(end - maxNegative) < 1e-6),
          capStart: !isStacked || Math.abs(start) < 1e-6,
          bottomScale,
          categoryIndex,
          color: colors.fill,
          depthOrder,
          depthOffsetX,
          depthOffsetY,
          height: Math.max(1, Math.abs(y2 - y1)),
          isHorizontal: false,
          key: `bar-${seriesIndex}-${categoryIndex}`,
          left: categoryStart + barOffset,
          depthX: usesSeriesDepthAxis ? depthBarX : frameOffsets?.depthX,
          depthY: usesSeriesDepthAxis ? depthBarY : frameOffsets?.depthY,
          shape3d: normalized3dShape,
          seriesIndex,
          stroke: colors.stroke,
          strokeWidth: colors.strokeWidth,
          topScale,
          top: Math.min(y1, y2),
          invertedNegative: isInvertedNegative,
          value: rawValue,
          width: Math.max(1, barThickness)
        });
      }
    }
  }

  const renderedBars = chart.is3d && isStacked
    ? bars.slice().sort((left, right) => {
        if (left.categoryIndex !== right.categoryIndex) {
          return left.categoryIndex - right.categoryIndex;
        }
        const leftNegative = left.value < 0 ? 0 : 1;
        const rightNegative = right.value < 0 ? 0 : 1;
        if (leftNegative !== rightNegative) {
          return leftNegative - rightNegative;
        }
        if (isHorizontal) {
          if (left.left !== right.left) {
            return left.left - right.left;
          }
          return left.seriesIndex - right.seriesIndex;
        }
        if (left.top !== right.top) {
          return right.top - left.top;
        }
        return left.seriesIndex - right.seriesIndex;
      })
    : bars;
  const sortedBars = usesSeriesDepthAxis
    ? renderedBars.slice().sort((left, right) => {
        if ((left.depthOrder ?? 0) !== (right.depthOrder ?? 0)) {
          return (right.depthOrder ?? 0) - (left.depthOrder ?? 0);
        }
        if (left.categoryIndex !== right.categoryIndex) {
          return left.categoryIndex - right.categoryIndex;
        }
        return left.top - right.top;
      })
    : renderedBars;

  const useVertical3dGradient = chart.is3d && !isHorizontal;
  const gradientByColor = new Map<string, string>();
  if (useVertical3dGradient) {
    for (const bar of renderedBars) {
      if (!bar.color || bar.color === "none" || bar.color.startsWith("url(")) {
        continue;
      }
      if (!gradientByColor.has(bar.color)) {
        gradientByColor.set(
          bar.color,
          `bar3d-front-${chart.id}-${gradientByColor.size}`.replace(/[^A-Za-z0-9_-]/g, "-")
        );
      }
      bar.gradientId = gradientByColor.get(bar.color);
    }
  }

  const axisNode = renderCartesianAxes(
    chart,
    palette,
    plot,
    isHorizontal,
    categories,
    categoryPositions,
    ticks,
    (value) => valueScale(value),
    isPercentStacked ? formatPercentTickValue : formatTickValue
  );
  const depthAxisColor = chart.axisLineColor ?? chart.chartAreaBorderColor ?? palette.border;
  const depthGridColor = lightenColor(depthAxisColor, 0.48);
  const frameDepthX = usesSeriesDepthAxis ? depthGridSpanX : (frameOffsets?.depthX ?? 0);
  const frameDepthY = usesSeriesDepthAxis ? depthGridSpanY : (frameOffsets?.depthY ?? 0);
  const depthAxisNode = usesSeriesDepthAxis && frameOffsets ? (
    <g>
      {Array.from({ length: categoryCount + 1 }, (_, boundaryIndex) => {
        const x = plot.left + ((boundaryIndex / Math.max(1, categoryCount)) * plot.width);
        return (
          <line
            key={`bar3d-floor-x-${boundaryIndex}`}
            stroke={depthGridColor}
            strokeWidth={0.9}
            x1={x}
            x2={x + depthGridSpanX}
            y1={plot.top + plot.height}
            y2={plot.top + plot.height + depthGridSpanY}
          />
        );
      })}
      {Array.from({ length: seriesCount + 1 }, (_, boundaryIndex) => {
        const ratio = boundaryIndex / Math.max(1, seriesCount);
        const offsetX = depthGridSpanX * ratio;
        const offsetY = depthGridSpanY * ratio;
        return (
          <line
            key={`bar3d-floor-z-${boundaryIndex}`}
            stroke={depthGridColor}
            strokeWidth={0.9}
            x1={plot.left + offsetX}
            x2={plot.left + plot.width + offsetX}
            y1={plot.top + plot.height + offsetY}
            y2={plot.top + plot.height + offsetY}
          />
        );
      })}
      {chart.series.map((series, seriesIndex) => {
        const ratio = (seriesIndex + 0.5) / Math.max(1, seriesCount);
        const x = plot.left + plot.width + depthGridSpanX * ratio;
        const y = plot.top + plot.height + depthGridSpanY * ratio;
        return (
          <text
            key={`bar3d-ser-label-${seriesIndex}`}
            fill={resolveChartAxisTextColor(chart)}
            fontSize={10}
            textAnchor="start"
            x={x + 7}
            y={y + 3}
          >
            {series.name ?? `Series ${seriesIndex + 1}`}
          </text>
        );
      })}
    </g>
  ) : null;

  const frameNode = chart.is3d && frameOffsets ? (
    <g>
      <polygon
        fill={lightenColor(chart.chartAreaFillColor ?? palette.surface, 0.05)}
        points={`${plot.left},${plot.top + plot.height} ${plot.left + plot.width},${plot.top + plot.height} ${plot.left + plot.width + frameDepthX},${plot.top + plot.height + frameDepthY} ${plot.left + frameDepthX},${plot.top + plot.height + frameDepthY}`}
        stroke={lightenColor(chart.axisLineColor ?? palette.border, 0.2)}
        strokeWidth={1}
      />
      <polygon
        fill={lightenColor(chart.chartAreaFillColor ?? palette.surface, 0.12)}
        points={`${plot.left},${plot.top} ${plot.left + plot.width},${plot.top} ${plot.left + plot.width + frameDepthX},${plot.top + frameDepthY} ${plot.left + frameDepthX},${plot.top + frameDepthY}`}
        stroke={lightenColor(chart.axisLineColor ?? palette.border, 0.2)}
        strokeWidth={1}
      />
    </g>
  ) : null;

  return (
    <g>
      {useVertical3dGradient && gradientByColor.size > 0 ? (
        <defs>
          {Array.from(gradientByColor.entries()).map(([color, id]) => (
            <linearGradient id={id} key={id} x1="0%" x2="0%" y1="0%" y2="100%">
              <stop offset="0%" stopColor={lightenColor(color, 0.28)} />
              <stop offset="42%" stopColor={lightenColor(color, 0.12)} />
              <stop offset="100%" stopColor={darkenColor(color, 0.1)} />
            </linearGradient>
          ))}
        </defs>
      ) : null}
      {axisNode}
      {frameNode}
      {depthAxisNode}
      {chart.is3d
        ? sortedBars.map((bar) => (
            <g key={`bar-hit-${bar.key}`} {...barChartElementDataProps(bar.seriesIndex, bar.categoryIndex)}>
              {renderExtrudedRect(bar)}
            </g>
          ))
        : renderedBars.map((bar) => (
            <rect
              {...barChartElementDataProps(bar.seriesIndex, bar.categoryIndex)}
              key={bar.key}
              fill={bar.color}
              height={bar.height}
              stroke={isHistogramLike ? "none" : bar.stroke}
              strokeWidth={isHistogramLike ? 0 : bar.strokeWidth}
              width={bar.width}
              x={bar.left}
              y={bar.top}
            />
          ))}
      {renderedBars.map((bar) => (
        isSelectedChartPoint(selectedChartElement, chart.id, bar.seriesIndex, bar.categoryIndex)
        || (
          isSelectedChartSeries(selectedChartElement, chart.id, bar.seriesIndex)
          && selectedChartElement?.kind !== "point"
        )
          ? renderSelectionRectHandles(`bar-selection-${bar.key}`, bar.left, bar.top, bar.width, bar.height)
          : null
      ))}
    </g>
  );
}

function renderLineOrAreaChart(chart: XlsxChart, palette: ChartRendererPalette, layout: ChartLayout, chartType: string, selectedChartElement?: XlsxChartElementSelection | null) {
  if (chart.series.length === 0) {
    return null;
  }
  const categories = getCategoryLabels(chart);
  if (categories.length === 0) {
    return null;
  }
  const plot = layout.plot;
  const displayBlanksAs = chart.displayBlanksAs;
  const isAreaChart = chartType === "Area" || chartType === "AreaStacked" || chartType === "AreaPercentStacked";
  const isStackedArea = chartType === "AreaStacked" || chartType === "AreaPercentStacked";
  const isStackedLine = chartType === "LineStacked" || chartType === "LinePercentStacked";
  const isStackedSeries = isStackedArea || isStackedLine;
  const isPercentStackedSeries = chartType === "AreaPercentStacked" || chartType === "LinePercentStacked";
  const resolvedValuesBySeries = chart.series.map((series) => (
    categories.map((_, categoryIndex) => resolveRenderableSeriesValue(series.values[categoryIndex], displayBlanksAs))
  ));

  type SeriesPoint = {
    defined: boolean;
    y: number | null;
    y0: number | null;
    y1: number | null;
  };

  const stackedPointsBySeries: SeriesPoint[][] = isStackedSeries
    ? (() => {
        const positive = Array.from({ length: categories.length }, () => 0);
        const negative = Array.from({ length: categories.length }, () => 0);
        const categoryTotals = isPercentStackedSeries
          ? categories.map((_, categoryIndex) => {
              const total = resolvedValuesBySeries.reduce((sum, seriesValues) => (
                sum + (seriesValues[categoryIndex] ?? 0)
              ), 0);
              return Math.abs(total) < 1e-9 ? 1 : total;
            })
          : null;
        return chart.series.map((_, seriesIndex) => (
          categories.map((_, categoryIndex) => {
            const rawValue = resolvedValuesBySeries[seriesIndex]?.[categoryIndex] ?? null;
            if (rawValue == null) {
              return { defined: false, y: null, y0: null, y1: null };
            }
            const value = isPercentStackedSeries
              ? (rawValue / (categoryTotals?.[categoryIndex] ?? 1)) * 100
              : rawValue;
            if (value >= 0) {
              const start = positive[categoryIndex];
              const end = start + value;
              positive[categoryIndex] = end;
              return { defined: true, y: end, y0: start, y1: end };
            }
            const start = negative[categoryIndex];
            const end = start + value;
            negative[categoryIndex] = end;
            return { defined: true, y: end, y0: start, y1: end };
          })
        ));
      })()
    : chart.series.map((_, seriesIndex) => (
        categories.map((_, categoryIndex) => {
          const value = resolvedValuesBySeries[seriesIndex]?.[categoryIndex] ?? null;
          return {
            defined: value != null,
            y: value,
            y0: null,
            y1: value
          };
        })
      ));

  const stackedExtents = stackedPointsBySeries
    .flatMap((seriesPoints) => (
      seriesPoints.flatMap((point) => (
        point.defined
          ? [point.y0, point.y1].filter((value): value is number => value != null && Number.isFinite(value))
          : []
      ))
    ));
  if (stackedExtents.length === 0) {
    return null;
  }
  let minValue = Math.min(...stackedExtents);
  let maxValue = Math.max(...stackedExtents);
  if (isAreaChart) {
    minValue = Math.min(0, minValue);
  } else if ((chartType === "Line" || isStackedLine) && chart.valueAxis?.crosses === "autoZero") {
    minValue = Math.min(0, minValue);
  }
  if (isPercentStackedSeries) {
    const explicitMin = typeof chart.valueAxis?.min === "number" && Number.isFinite(chart.valueAxis.min)
      ? Number(chart.valueAxis.min)
      : Math.min(0, minValue);
    const explicitMax = typeof chart.valueAxis?.max === "number" && Number.isFinite(chart.valueAxis.max)
      ? Number(chart.valueAxis.max)
      : Math.max(100, maxValue);
    minValue = explicitMin;
    maxValue = explicitMax <= explicitMin ? explicitMin + 1 : explicitMax;
  } else {
    const valueDomain = resolveAxisDomainWithChartOverrides(
      chart.valueAxis,
      minValue,
      maxValue,
      isAreaChart || chart.valueAxis?.crosses === "autoZero"
    );
    minValue = valueDomain.min;
    maxValue = valueDomain.max;
  }

  const xScale = scalePoint<string>()
    .domain(categories)
    .range([plot.left, plot.left + plot.width]);
  const yScale = scaleLinear()
    .domain([minValue, maxValue])
    .range([plot.top + plot.height, plot.top]);

  const ticks = isPercentStackedSeries
    ? buildNumericTickValues(minValue, maxValue, chart.valueAxis?.majorUnit ?? 20)
    : resolveAxisDomainWithChartOverrides(
        chart.valueAxis,
        minValue,
        maxValue,
        isAreaChart || chart.valueAxis?.crosses === "autoZero"
      ).ticks;
  const categoryPositions = categories.map((category) => xScale(category) ?? plot.left);

  const curve: CurveFactory = curveLinear;
  const areaBaseline = yScale(Math.max(minValue, 0));
  const plotPointsBySeries = chart.series.map((_, seriesIndex) => (
    categories.map((category, categoryIndex) => {
      const point = stackedPointsBySeries[seriesIndex]?.[categoryIndex] ?? {
        defined: false,
        y: null,
        y0: null,
        y1: null
      };
      return {
        defined: point.defined,
        x: xScale(category) ?? plot.left,
        y: point.y,
        y0: point.y0,
        y1: point.y1
      };
    })
  ));

  if (chart.is3d) {
    return renderLineOrAreaChart3d(
      chart,
      palette,
      layout,
      categories,
      stackedPointsBySeries,
      minValue,
      maxValue,
      isAreaChart,
      isStackedSeries,
      selectedChartElement
    );
  }

  const rawRecord = chart.raw && typeof chart.raw === "object"
    ? chart.raw as Record<string, unknown>
    : null;
  const dropLinesRecord = rawRecord?.dropLines && typeof rawRecord.dropLines === "object"
    ? rawRecord.dropLines as Record<string, unknown>
    : null;
  const highLowLinesRecord = rawRecord?.highLowLines && typeof rawRecord.highLowLines === "object"
    ? rawRecord.highLowLines as Record<string, unknown>
    : null;
  const upDownBarsRecord = rawRecord?.upDownBars && typeof rawRecord.upDownBars === "object"
    ? rawRecord.upDownBars as Record<string, unknown>
    : null;
  const dropLineColor = normalizeRendererHexColor((dropLinesRecord?.shapeProperties as Record<string, unknown> | undefined)?.lineColorHex)
    ?? lightenColor(chart.axisLineColor ?? chart.chartAreaBorderColor ?? palette.border, 0.45);
  const highLowLineColor = normalizeRendererHexColor((highLowLinesRecord?.shapeProperties as Record<string, unknown> | undefined)?.lineColorHex)
    ?? chart.axisLineColor
    ?? chart.chartAreaBorderColor
    ?? palette.border;
  const upDownGapWidth = typeof upDownBarsRecord?.gapWidth === "number" && Number.isFinite(upDownBarsRecord.gapWidth)
    ? upDownBarsRecord.gapWidth
    : 150;
  const categoryStep = categoryPositions.length >= 2
    ? Math.min(...categoryPositions.slice(1).map((position, index) => Math.abs(position - (categoryPositions[index] ?? 0))).filter((value) => value > 0))
    : plot.width;
  const upDownBarWidth = clamp(categoryStep * (1 - resolveCategoryBandPadding(upDownGapWidth).inner) * 0.82, 3, 28);

  return (
    <g>
      {renderCartesianAxes(
        chart,
        palette,
        plot,
        false,
        categories,
        categoryPositions,
        ticks,
        (value) => yScale(value)
      )}
      {!isAreaChart && !isStackedSeries && highLowLinesRecord
        ? categories.map((_, categoryIndex) => {
            const yValues = plotPointsBySeries
              .map((seriesPoints) => seriesPoints[categoryIndex]?.y)
              .filter((value): value is number => value != null && Number.isFinite(value));
            if (yValues.length < 2) {
              return null;
            }
            const x = categoryPositions[categoryIndex] ?? plot.left;
            return (
              <line
                key={`high-low-${categoryIndex}`}
                stroke={highLowLineColor}
                strokeWidth={1}
                x1={x}
                x2={x}
                y1={yScale(Math.min(...yValues))}
                y2={yScale(Math.max(...yValues))}
              />
            );
          })
        : null}
      {!isAreaChart && !isStackedSeries && upDownBarsRecord && plotPointsBySeries.length >= 2
        ? categories.map((_, categoryIndex) => {
            const first = plotPointsBySeries[0]?.[categoryIndex];
            const second = plotPointsBySeries[plotPointsBySeries.length - 1]?.[categoryIndex];
            if (!first || !second || first.y == null || second.y == null) {
              return null;
            }
            const top = yScale(Math.max(first.y, second.y));
            const bottom = yScale(Math.min(first.y, second.y));
            const isUpBar = second.y >= first.y;
            const fill = isUpBar ? darkenColor(highLowLineColor, 0.12) : "#c0504d";
            return (
              <rect
                key={`up-down-${categoryIndex}`}
                fill={fill}
                fillOpacity={0.92}
                height={Math.max(1, bottom - top)}
                stroke={darkenColor(fill, 0.18)}
                strokeWidth={1}
                width={upDownBarWidth}
                x={(categoryPositions[categoryIndex] ?? plot.left) - upDownBarWidth / 2}
                y={top}
              />
            );
          })
        : null}
      {chart.series.map((series, seriesIndex) => {
        const points = plotPointsBySeries[seriesIndex] ?? [];
        const lineStrokePoints = displayBlanksAs === "span"
          ? points.filter((point) => point.y != null)
          : points;
        const linePath = d3Line<{ x: number; y: number | null }>()
          .defined((point) => point.y != null)
          .x((point) => point.x)
          .y((point) => yScale(point.y ?? 0))
          .curve(curve)(lineStrokePoints) ?? "";

        const areaPath = isAreaChart
          ? d3Area<{ x: number; y: number | null; y0: number | null; y1: number | null }>()
            .defined((point) => (
              isStackedSeries
                ? point.y0 != null && point.y1 != null
                : point.y != null
            ))
            .x((point) => point.x)
            .y0((point) => (
              isStackedSeries
                ? yScale(point.y0 ?? 0)
                : areaBaseline
            ))
            .y1((point) => yScale((isStackedSeries ? point.y1 : point.y) ?? 0))
            .curve(curve)(points) ?? ""
          : "";

        const seriesFillColor = typeof series.shapeProperties?.xmlFillColor === "string"
          ? series.shapeProperties.xmlFillColor
          : chartSeriesColor(chart, seriesIndex);
        return (
          <g key={`line-series-${seriesIndex}`}>
            {!isAreaChart && !isStackedSeries && dropLinesRecord
              ? points.map((point, pointIndex) => (
                  point.y == null ? null : (
                    <line
                      key={`drop-line-${seriesIndex}-${pointIndex}`}
                      stroke={dropLineColor}
                      strokeWidth={1}
                      x1={point.x}
                      x2={point.x}
                      y1={yScale(point.y)}
                      y2={areaBaseline}
                    />
                  )
                ))
              : null}
            {isAreaChart ? (
              <path
                {...chartElementDataProps(seriesIndex)}
                d={areaPath}
                fill={seriesFillColor}
                fillOpacity={1}
                stroke="none"
              />
            ) : null}
            <path
              {...chartElementDataProps(seriesIndex)}
              d={linePath}
              fill="none"
              stroke={series.shapeProperties?.xmlLineHidden === true ? "none" : chartSeriesStrokeColor(chart, seriesIndex)}
              strokeLinecap="round"
              strokeLinejoin="round"
              strokeWidth={Math.max(1.5, series.lineWidthPx ?? 2)}
            />
            {points.map((point, pointIndex) => (
              point.y == null || normalizeChartMarkerSymbol(series.markerSymbol) === "none" ? null : (
                <path
                  {...chartElementDataProps(seriesIndex, pointIndex)}
                  key={`line-marker-${seriesIndex}-${pointIndex}`}
                  d={markerSymbolPath(normalizeChartMarkerSymbol(series.markerSymbol), markerRadiusPx(series.markerSize))}
                  transform={`translate(${point.x}, ${yScale(point.y)})`}
                  fill={series.markerColor ?? chartSeriesColor(chart, seriesIndex)}
                  stroke={series.markerLineColor ?? chart.chartAreaFillColor ?? chartSeriesStrokeColor(chart, seriesIndex)}
                  strokeWidth={1}
                />
              )
            ))}
            {points.map((point, pointIndex) => (
              point.y != null && (
                isSelectedChartPoint(selectedChartElement, chart.id, seriesIndex, pointIndex)
                || (
                  isSelectedChartSeries(selectedChartElement, chart.id, seriesIndex)
                  && selectedChartElement?.kind !== "point"
                )
              )
                ? isAreaChart
                  ? renderSelectionPointHandles(`area-selection-${seriesIndex}-${pointIndex}`, [{ x: point.x, y: yScale(point.y) }])
                  : renderSelectionRectHandles(
                      `line-selection-${seriesIndex}-${pointIndex}`,
                      point.x - 4,
                      yScale(point.y) - 4,
                      8,
                      8
                    )
                : null
            ))}
          </g>
        );
      })}
    </g>
  );
}

function renderComboChart(chart: XlsxChart, palette: ChartRendererPalette, layout: ChartLayout, selectedChartElement?: XlsxChartElementSelection | null) {
  const groups = buildComboGroups(chart);
  const columnGroup = groups.find((group) => group.chartType.startsWith("Column"));
  const lineGroup = groups.find((group) => group.chartType.startsWith("Line"));
  if (!columnGroup || !lineGroup) {
    return null;
  }

  const primaryChart: XlsxChart = {
    ...chart,
    categoryAxis: columnGroup.categoryAxis ?? chart.categoryAxis,
    chartType: columnGroup.chartType,
    dataLabels: undefined,
    is3d: columnGroup.is3d ?? false,
    series: columnGroup.series,
    valueAxis: columnGroup.valueAxis ?? chart.valueAxis
  };
  const categories = getCategoryLabels(primaryChart);
  if (categories.length === 0) {
    return null;
  }

  const plot = layout.plot;
  const groupGapWidth = typeof columnGroup.raw?.gapWidth === "number" && Number.isFinite(columnGroup.raw.gapWidth)
    ? columnGroup.raw.gapWidth
    : columnGroup.gapWidth ?? chart.gapWidth;
  const histogramColumns = columnGroup.series.some((series) => isHistogramLikeSeries(series));
  const categoryBandPadding = histogramColumns
    ? { inner: 0, outer: 0 }
    : resolveCategoryBandPadding(groupGapWidth);
  const categoryScale = scaleBand<string>()
    .domain(categories)
    .range([plot.left, plot.left + plot.width])
    .paddingInner(categoryBandPadding.inner)
    .paddingOuter(categoryBandPadding.outer);
  const seriesScale = scaleBand<string>()
    .domain(Array.from({ length: columnGroup.series.length }, (_, index) => String(index)))
    .range([0, categoryScale.bandwidth()])
    .paddingInner(histogramColumns ? 0 : 0.16)
    .paddingOuter(histogramColumns ? 0 : 0.08);
  const categoryPositions = categories.map((category) => (
    (categoryScale(category) ?? plot.left) + categoryScale.bandwidth() / 2
  ));
  const resolveGlobalSeriesIndex = (series: XlsxChartSeries) => {
    const index = chart.series.findIndex((candidate) => candidate.id === series.id);
    return index >= 0 ? index : 0;
  };

  const primaryValues = columnGroup.series.flatMap((series) => (
    series.values.map((value) => safeNumber(value)).filter((value): value is number => value != null)
  ));
  const secondaryValues = lineGroup.series.flatMap((series) => (
    series.values.map((value) => safeNumber(value)).filter((value): value is number => value != null)
  ));
  if (primaryValues.length === 0 || secondaryValues.length === 0) {
    return null;
  }

  const primaryDomain = resolveNumericAxisDomain(
    typeof columnGroup.valueAxis?.min === "number" ? columnGroup.valueAxis.min : Math.min(...primaryValues),
    typeof columnGroup.valueAxis?.max === "number" ? columnGroup.valueAxis.max : Math.max(...primaryValues),
    columnGroup.valueAxis?.majorUnit,
    true
  );
  const secondaryDomain = resolveNumericAxisDomain(
    typeof lineGroup.valueAxis?.min === "number" ? lineGroup.valueAxis.min : Math.min(...secondaryValues),
    typeof lineGroup.valueAxis?.max === "number" ? lineGroup.valueAxis.max : Math.max(...secondaryValues),
    lineGroup.valueAxis?.majorUnit,
    false
  );
  const primaryScale = scaleLinear()
    .domain([primaryDomain.min, primaryDomain.max])
    .range([plot.top + plot.height, plot.top]);
  const secondaryScale = scaleLinear()
    .domain([secondaryDomain.min, secondaryDomain.max])
    .range([plot.top + plot.height, plot.top]);
  const axisColor = chart.axisLineColor ?? chart.chartAreaBorderColor ?? palette.border;
  const labelColor = resolveChartAxisTextColor(chart);

  return (
    <g>
      {renderCartesianAxes(
        primaryChart,
        palette,
        plot,
        false,
        categories,
        categoryPositions,
        primaryDomain.ticks,
        (value) => primaryScale(value)
      )}
      {secondaryDomain.ticks.map((tick) => {
        const y = secondaryScale(tick);
        return (
          <text
            key={`combo-secondary-tick-${tick}`}
            fill={labelColor}
            fontSize={10}
            textAnchor="start"
            x={plot.left + plot.width + 6}
            y={y + 3}
          >
            {formatTickValue(tick)}
          </text>
        );
      })}
      <line
        stroke={axisColor}
        strokeWidth={1.2}
        x1={plot.left + plot.width}
        x2={plot.left + plot.width}
        y1={plot.top}
        y2={plot.top + plot.height}
      />
      {columnGroup.series.flatMap((series, seriesIndex) => categories.map((category, categoryIndex) => {
        const globalSeriesIndex = resolveGlobalSeriesIndex(series);
        const value = safeNumber(series.values[categoryIndex]) ?? 0;
        const categoryStart = categoryScale(category) ?? plot.left;
        const barWidth = Math.max(1, histogramColumns ? categoryScale.bandwidth() : seriesScale.bandwidth());
        const x = categoryStart + (histogramColumns ? 0 : (seriesScale(String(seriesIndex)) ?? 0));
        const y = primaryScale(Math.max(0, value));
        const zeroY = primaryScale(0);
        const height = Math.max(1, Math.abs(zeroY - primaryScale(value)));
        const selected = isSelectedChartPoint(selectedChartElement, chart.id, globalSeriesIndex, categoryIndex)
          || (
            isSelectedChartSeries(selectedChartElement, chart.id, globalSeriesIndex)
            && selectedChartElement?.kind !== "point"
          );
        return (
          <React.Fragment key={`combo-bar-${seriesIndex}-${categoryIndex}`}>
            <rect
              {...barChartElementDataProps(globalSeriesIndex, categoryIndex)}
              fill={series.color ?? series.lineColor ?? chartSeriesColor(primaryChart, seriesIndex)}
              height={height}
              stroke={selected ? "#64748b" : (histogramColumns ? "none" : (series.lineColor ?? series.color ?? chartSeriesStrokeColor(primaryChart, seriesIndex)))}
              strokeWidth={selected ? 2 : (histogramColumns ? 0 : 1)}
              width={barWidth}
              x={x}
              y={Math.min(y, zeroY)}
            />
            {selected
              ? renderSelectionRectHandles(
                  `combo-bar-selection-${globalSeriesIndex}-${categoryIndex}`,
                  x,
                  Math.min(y, zeroY),
                  barWidth,
                  height
                )
              : null}
          </React.Fragment>
        );
      }))}
      {lineGroup.series.map((series, seriesIndex) => {
        const globalSeriesIndex = resolveGlobalSeriesIndex(series);
        const points = categories.map((category, categoryIndex) => ({
          x: (categoryScale(category) ?? plot.left) + categoryScale.bandwidth() / 2,
          y: safeNumber(series.values[categoryIndex])
        }));
        const lineStrokePoints = chart.displayBlanksAs === "span"
          ? points.filter((point) => point.y != null)
          : points;
        const path = d3Line<{ x: number; y: number | null }>()
          .defined((point) => point.y != null)
          .x((point) => point.x)
          .y((point) => secondaryScale(point.y ?? 0))
          .curve(curveLinear)(lineStrokePoints) ?? "";
        return (
          <g key={`combo-line-${seriesIndex}`}>
            <path
              {...chartElementDataProps(globalSeriesIndex)}
              d={path}
              fill="none"
              stroke={series.lineColor ?? series.color ?? chartSeriesStrokeColor(chart, columnGroup.series.length + seriesIndex)}
              strokeLinecap="round"
              strokeLinejoin="round"
              strokeWidth={Math.max(1.5, series.lineWidthPx ?? 2)}
            />
            {points.map((point, pointIndex) => (
              point.y == null || normalizeChartMarkerSymbol(series.markerSymbol) === "none" ? null : (
                <path
                  {...chartElementDataProps(globalSeriesIndex, pointIndex)}
                  d={markerSymbolPath(normalizeChartMarkerSymbol(series.markerSymbol), markerRadiusPx(series.markerSize))}
                  fill={series.markerColor ?? series.color ?? series.lineColor ?? chartSeriesColor(chart, columnGroup.series.length + seriesIndex)}
                  key={`combo-line-marker-${seriesIndex}-${pointIndex}`}
                  stroke={series.markerLineColor ?? chart.chartAreaFillColor ?? "#ffffff"}
                  strokeWidth={1}
                  transform={`translate(${point.x}, ${secondaryScale(point.y)})`}
                />
              )
            ))}
            {points.map((point, pointIndex) => (
              point.y != null && isSelectedChartPointOrSeries(selectedChartElement, chart.id, globalSeriesIndex, pointIndex)
                ? renderSelectionRectHandles(
                    `combo-line-selection-${globalSeriesIndex}-${pointIndex}`,
                    point.x - 4,
                    secondaryScale(point.y) - 4,
                    8,
                    8
                  )
                : null
            ))}
          </g>
        );
      })}
    </g>
  );
}

function renderScatterChart(chart: XlsxChart, palette: ChartRendererPalette, layout: ChartLayout, smooth: boolean, selectedChartElement?: XlsxChartElementSelection | null) {
  const plot = layout.plot;
  const rawRecord = chart.raw && typeof chart.raw === "object"
    ? chart.raw as Record<string, unknown>
    : null;
  const scatterStyle = typeof chart.scatterStyle === "string"
    ? chart.scatterStyle
    : rawRecord && typeof rawRecord.scatterStyle === "string"
      ? rawRecord.scatterStyle
      : undefined;
  const styleDrawsLine = scatterStyle
    ? scatterStyle === "line" || scatterStyle === "lineMarker" || scatterStyle === "smooth" || scatterStyle === "smoothMarker"
    : chart.chartType !== "Scatter";
  const styleShowsMarkers = scatterStyle
    ? scatterStyle === "marker" || scatterStyle === "lineMarker" || scatterStyle === "smoothMarker"
    : true;
  const styleUsesSmoothCurve = scatterStyle
    ? scatterStyle === "smooth" || scatterStyle === "smoothMarker"
    : smooth;
  const pointsBySeries = chart.series.map((series) => {
    const yLength = series.values.length;
    const directXValues = Array.from({ length: yLength }, (_, index) => coerceLooseNumber(series.categories[index]));

    let resolvedXValues = directXValues;
    let usesSyntheticIndex = false;
    if (directXValues.filter((value) => value != null).length < Math.max(2, yLength - 1)) {
      const categories = series.categories ?? [];
      const pairedXValues = Array.from({ length: yLength }, (_, index) => {
        const candidates = [
          categories[index],
          categories[index * 2],
          categories[index * 2 + 1],
          categories[index + 1]
        ];
        for (const candidate of candidates) {
          const parsed = coerceLooseNumber(candidate);
          if (parsed != null) {
            return parsed;
          }
        }
        return null;
      });
      if (pairedXValues.filter((value) => value != null).length >= 2) {
        resolvedXValues = pairedXValues;
      } else {
        const numericPool = categories
          .map((entry) => coerceLooseNumber(entry))
          .filter((value): value is number => value != null);
        if (numericPool.length >= 2) {
          resolvedXValues = Array.from(
            { length: yLength },
            (_, index) => numericPool[index] ?? numericPool[numericPool.length - 1] ?? null
          );
        } else {
          usesSyntheticIndex = true;
          resolvedXValues = Array.from({ length: yLength }, (_, index) => index + 1);
        }
      }
    }

    const points = series.values.map((value, index) => {
      const x = coerceLooseNumber(resolvedXValues[index]);
      const y = safeNumber(value);
      return x == null || y == null ? null : { x, y };
    }).filter((point): point is { x: number; y: number } => point != null);
    return {
      pointCount: yLength,
      points,
      series,
      usesSyntheticIndex
    };
  });

  const allX = pointsBySeries.flatMap((series) => series.points.map((point) => point.x));
  const allY = pointsBySeries.flatMap((series) => series.points.map((point) => point.y));
  if (allX.length === 0 || allY.length === 0) {
    return null;
  }

  const hasExplicitMinX = typeof chart.categoryAxis?.min === "number" && Number.isFinite(chart.categoryAxis.min);
  const hasExplicitMaxX = typeof chart.categoryAxis?.max === "number" && Number.isFinite(chart.categoryAxis.max);
  const hasExplicitMinY = typeof chart.valueAxis?.min === "number" && Number.isFinite(chart.valueAxis.min);
  const hasExplicitMaxY = typeof chart.valueAxis?.max === "number" && Number.isFinite(chart.valueAxis.max);
  const hasSyntheticIndexAxis = pointsBySeries.some((series) => series.usesSyntheticIndex);
  const syntheticPointCount = Math.max(0, ...pointsBySeries.map((series) => series.pointCount));

  let minX = hasExplicitMinX ? Number(chart.categoryAxis?.min) : Math.min(...allX);
  let maxX = hasExplicitMaxX ? Number(chart.categoryAxis?.max) : Math.max(...allX);
  let minY = hasExplicitMinY ? Number(chart.valueAxis?.min) : Math.min(...allY);
  let maxY = hasExplicitMaxY ? Number(chart.valueAxis?.max) : Math.max(...allY);

  if (!hasExplicitMinX) {
    minX = hasSyntheticIndexAxis
      ? 0
      : chart.categoryAxis?.crosses === "autoZero"
        ? Math.min(0, minX)
        : minX;
  }
  if (!hasExplicitMaxX) {
    maxX = hasSyntheticIndexAxis
      ? Math.max(maxX, syntheticPointCount + 1)
      : chart.categoryAxis?.crosses === "autoZero"
        ? Math.max(0, maxX)
        : maxX;
  }
  if (!hasExplicitMinY && chart.valueAxis?.crosses === "autoZero") {
    minY = Math.min(0, minY);
  }
  if (!hasExplicitMaxY && chart.valueAxis?.crosses === "autoZero") {
    maxY = Math.max(0, maxY);
  }

  if (maxX <= minX) {
    maxX = minX + 1;
  }
  if (maxY <= minY) {
    maxY = minY + 1;
  }

  const xStep = typeof chart.categoryAxis?.majorUnit === "number" && chart.categoryAxis.majorUnit > 0
    ? chart.categoryAxis.majorUnit
    : buildNiceStep(minX, maxX, 5);
  const yStep = typeof chart.valueAxis?.majorUnit === "number" && chart.valueAxis.majorUnit > 0
    ? chart.valueAxis.majorUnit
    : buildNiceStep(minY, maxY, 5);

  if (!hasExplicitMinX && !hasSyntheticIndexAxis && Number.isFinite(xStep) && xStep > 0) {
    minX = Math.floor(minX / xStep) * xStep;
  }
  if (!hasExplicitMaxX && !hasSyntheticIndexAxis && Number.isFinite(xStep) && xStep > 0) {
    maxX = Math.ceil(maxX / xStep) * xStep;
  }
  if (!hasExplicitMinY && Number.isFinite(yStep) && yStep > 0) {
    minY = Math.floor(minY / yStep) * yStep;
  }
  if (!hasExplicitMaxY && Number.isFinite(yStep) && yStep > 0) {
    maxY = Math.ceil(maxY / yStep) * yStep;
  }

  const safeMaxX = maxX <= minX ? minX + 1 : maxX;
  const safeMaxY = maxY <= minY ? minY + 1 : maxY;

  const xScale = scaleLinear().domain([minX, safeMaxX]).range([plot.left, plot.left + plot.width]);
  const yScale = scaleLinear().domain([minY, safeMaxY]).range([plot.top + plot.height, plot.top]);
  const xMajorUnit = typeof chart.categoryAxis?.majorUnit === "number" && chart.categoryAxis.majorUnit > 0
    ? chart.categoryAxis.majorUnit
    : hasSyntheticIndexAxis
      ? 1
      : undefined;
  const yMajorUnit = typeof chart.valueAxis?.majorUnit === "number" && chart.valueAxis.majorUnit > 0
    ? chart.valueAxis.majorUnit
    : buildNiceStep(minY, safeMaxY, 6);
  const xTicks = buildNumericTickValues(minX, safeMaxX, xMajorUnit);
  const yTicks = buildNumericTickValues(minY, safeMaxY, yMajorUnit);
  const axisColor = chart.axisLineColor ?? chart.chartAreaBorderColor ?? palette.border;
  const labelColor = resolveChartAxisTextColor(chart);

  const hasZeroX = minX <= 0 && safeMaxX >= 0;
  const hasZeroY = minY <= 0 && safeMaxY >= 0;

  return (
    <g>
      {xTicks.map((tick) => (
        <g key={`scatter-x-${tick}`}>
          <line
            stroke={lightenColor(axisColor, 0.7)}
            strokeWidth={1}
            x1={xScale(tick)}
            x2={xScale(tick)}
            y1={plot.top}
            y2={plot.top + plot.height}
          />
          <text fill={labelColor} fontSize={10} textAnchor="middle" x={xScale(tick)} y={plot.top + plot.height + 14}>
            {formatTickValue(tick)}
          </text>
        </g>
      ))}
      {yTicks.map((tick) => (
        <g key={`scatter-y-${tick}`}>
          <line
            stroke={lightenColor(axisColor, 0.7)}
            strokeWidth={1}
            x1={plot.left}
            x2={plot.left + plot.width}
            y1={yScale(tick)}
            y2={yScale(tick)}
          />
          <text fill={labelColor} fontSize={10} textAnchor="end" x={plot.left - 6} y={yScale(tick) + 3}>
            {formatTickValue(tick)}
          </text>
        </g>
      ))}
      {hasZeroY ? (
        <line
          stroke={axisColor}
          strokeWidth={1.2}
          x1={plot.left}
          x2={plot.left + plot.width}
          y1={yScale(0)}
          y2={yScale(0)}
        />
      ) : null}
      {hasZeroX ? (
        <line
          stroke={axisColor}
          strokeWidth={1.2}
          x1={xScale(0)}
          x2={xScale(0)}
          y1={plot.top}
          y2={plot.top + plot.height}
        />
      ) : null}
      {pointsBySeries.map((seriesPoints, seriesIndex) => {
        const series = seriesPoints.series;
        const markerSize = markerRadiusPx(series.markerSize);
        const markerPath = markerSymbolPath(
          normalizeChartMarkerSymbol(series.markerSymbol),
          markerSize
        );
        const markerFill = series.lineColor ?? series.markerColor ?? series.color ?? chartSeriesStrokeColor(chart, seriesIndex);
        const shouldDrawLine = styleDrawsLine
          && series.shapeProperties?.xmlLineHidden !== true
          && seriesPoints.points.length > 1;
        const shouldDrawMarkers = styleShowsMarkers
          && markerPath.length > 0;
        const lineCurve = smooth || styleUsesSmoothCurve || series.smooth === true
          ? curveCatmullRom.alpha(0.5)
          : curveLinear;
        const linePath = shouldDrawLine
          ? d3Line<{ x: number; y: number }>()
              .x((point) => xScale(point.x))
              .y((point) => yScale(point.y))
              .curve(lineCurve)(seriesPoints.points) ?? ""
          : "";

        return (
          <g key={`scatter-series-${seriesIndex}`}>
            {shouldDrawLine && linePath.length > 0 ? (
              <path
                {...chartElementDataProps(seriesIndex)}
                d={linePath}
                fill="none"
                stroke={series.lineColor ?? chartSeriesStrokeColor(chart, seriesIndex)}
                strokeLinecap="round"
                strokeLinejoin="round"
                strokeWidth={Math.max(1.2, series.lineWidthPx ?? 2)}
              />
            ) : null}
            {seriesPoints.points.map((point, pointIndex) => {
              if (!shouldDrawMarkers) {
                return null;
              }
              return (
                <g
                  key={`scatter-point-${seriesIndex}-${pointIndex}`}
                  transform={`translate(${xScale(point.x)}, ${yScale(point.y)})`}
                >
                  <path
                    {...chartElementDataProps(seriesIndex, pointIndex, { selectionMode: "seriesFirst" })}
                    d={markerPath}
                    fill={markerFill}
                    stroke="none"
                    strokeWidth={0}
                  />
                </g>
              );
            })}
            {seriesPoints.points.map((point, pointIndex) => (
              isSelectedChartPoint(selectedChartElement, chart.id, seriesIndex, pointIndex)
              || (
                isSelectedChartSeries(selectedChartElement, chart.id, seriesIndex)
                && selectedChartElement?.kind !== "point"
              )
                ? renderSelectionPointHandles(
                    `scatter-selection-${seriesIndex}-${pointIndex}`,
                    [{ x: xScale(point.x), y: yScale(point.y) }]
                  )
                : null
            ))}
          </g>
        );
      })}
    </g>
  );
}

function renderBubbleChart(chart: XlsxChart, palette: ChartRendererPalette, layout: ChartLayout, selectedChartElement?: XlsxChartElementSelection | null) {
  const plot = layout.plot;
  const pointsBySeries = chart.series.map((series) => (
    series.values.map((value, index) => {
      const x = safeNumber(series.categories[index]);
      const y = safeNumber(value);
      const bubble = safeNumber(series.bubbleSizes?.[index]);
      return x == null || y == null ? null : { bubble: bubble ?? 1, index, x, y };
    }).filter((point): point is { bubble: number; index: number; x: number; y: number } => point != null)
  ));

  const allX = pointsBySeries.flatMap((points) => points.map((point) => point.x));
  const allY = pointsBySeries.flatMap((points) => points.map((point) => point.y));
  const allBubble = pointsBySeries.flatMap((points) => points.map((point) => point.bubble));
  if (allX.length === 0 || allY.length === 0) {
    return null;
  }
  const minX = Math.min(...allX);
  const maxX = Math.max(...allX);
  const minY = Math.min(...allY);
  const maxY = Math.max(...allY);
  const bubbleMagnitudes = allBubble.map((value) => Math.sqrt(Math.max(0, value)));
  const minBubbleMagnitude = Math.min(...bubbleMagnitudes);
  const maxBubbleMagnitude = Math.max(...bubbleMagnitudes);

  const bubbleScaleFactor = clamp((chart.bubbleScale ?? 100) / 100, 0.2, 4);
  const minRadius = 4;
  const maxRadius = Math.max(minRadius + 2, 12 * bubbleScaleFactor);
  const radiusScale = scaleLinear()
    .domain([minBubbleMagnitude, maxBubbleMagnitude <= minBubbleMagnitude ? minBubbleMagnitude + 1 : maxBubbleMagnitude])
    .range([minRadius, maxRadius]);
  const safeMaxX = maxX <= minX ? minX + 1 : maxX;
  const safeMaxY = maxY <= minY ? minY + 1 : maxY;
  const xSpan = safeMaxX - minX;
  const ySpan = safeMaxY - minY;
  const xPad = Math.max(xSpan * 0.04, (xSpan * maxRadius) / Math.max(1, plot.width));
  const yPad = Math.max(ySpan * 0.06, (ySpan * maxRadius) / Math.max(1, plot.height));
  const xDomain = resolveAxisDomainWithChartOverrides(
    chart.categoryAxis,
    minX - xPad,
    safeMaxX + xPad,
    chart.categoryAxis?.crosses === "autoZero"
  );
  const yDomain = resolveAxisDomainWithChartOverrides(
    chart.valueAxis,
    minY - yPad,
    safeMaxY + yPad,
    chart.valueAxis?.crosses === "autoZero"
  );
  const xScale = scaleLinear().domain([xDomain.min, xDomain.max]).range([plot.left, plot.left + plot.width]);
  const yScale = scaleLinear().domain([yDomain.min, yDomain.max]).range([plot.top + plot.height, plot.top]);

  const xTicks = xDomain.ticks;
  const yTicks = yDomain.ticks;
  const axisColor = chart.axisLineColor ?? chart.chartAreaBorderColor ?? palette.border;
  const labelColor = resolveChartAxisTextColor(chart);
  const isBubble3d = chart.bubble3d === true || chart.is3d === true;
  const labelsEnabled = Boolean(
    chart.dataLabels?.showCategoryName
    || chart.dataLabels?.showSeriesName
    || chart.dataLabels?.showValue
    || chart.dataLabels?.showBubbleSize
    || chart.dataLabels?.showPercent
  );
  const bubbleTotals = pointsBySeries.map((points) => (
    points.reduce((sum, point) => sum + Math.abs(point.bubble), 0)
  ));

  return (
    <g>
      {xTicks.map((tick) => (
        <g key={`bubble-x-${tick}`}>
          <line
            stroke={lightenColor(axisColor, 0.72)}
            strokeWidth={1}
            x1={xScale(tick)}
            x2={xScale(tick)}
            y1={plot.top}
            y2={plot.top + plot.height}
          />
          <text fill={labelColor} fontSize={10} textAnchor="middle" x={xScale(tick)} y={plot.top + plot.height + 14}>
            {formatTickValue(tick)}
          </text>
        </g>
      ))}
      {yTicks.map((tick) => (
        <g key={`bubble-y-${tick}`}>
          <line
            stroke={lightenColor(axisColor, 0.72)}
            strokeWidth={1}
            x1={plot.left}
            x2={plot.left + plot.width}
            y1={yScale(tick)}
            y2={yScale(tick)}
          />
          <text fill={labelColor} fontSize={10} textAnchor="end" x={plot.left - 6} y={yScale(tick) + 3}>
            {formatTickValue(tick)}
          </text>
        </g>
      ))}
      <line stroke={axisColor} strokeWidth={1.2} x1={plot.left} x2={plot.left + plot.width} y1={plot.top + plot.height} y2={plot.top + plot.height} />
      <line stroke={axisColor} strokeWidth={1.2} x1={plot.left} x2={plot.left} y1={plot.top} y2={plot.top + plot.height} />
      {pointsBySeries.map((points, seriesIndex) => (
        <g key={`bubble-series-${seriesIndex}`}>
          {isBubble3d ? (
            <defs>
              <radialGradient id={`bubble3d-grad-${chart.id}-${seriesIndex}`} cx="35%" cy="30%" r="70%">
                <stop offset="0%" stopColor={lightenColor(chart.series[seriesIndex]?.color ?? chart.series[seriesIndex]?.lineColor ?? chartSeriesColor(chart, seriesIndex), 0.42)} />
                <stop offset="58%" stopColor={chart.series[seriesIndex]?.color ?? chart.series[seriesIndex]?.lineColor ?? chartSeriesColor(chart, seriesIndex)} />
                <stop offset="100%" stopColor={darkenColor(chart.series[seriesIndex]?.color ?? chart.series[seriesIndex]?.lineColor ?? chartSeriesColor(chart, seriesIndex), 0.18)} />
              </radialGradient>
            </defs>
          ) : null}
          {[...points]
            .sort((left, right) => {
              if (left.bubble !== right.bubble) {
                return right.bubble - left.bubble;
              }
              return left.index - right.index;
            })
            .map((point) => {
            const series = chart.series[seriesIndex];
            const baseColor = series?.color ?? series?.lineColor ?? chartSeriesColor(chart, seriesIndex);
            const radius = radiusScale(Math.sqrt(Math.max(0, point.bubble)));
            const pieces: string[] = [];
            if (chart.dataLabels?.showSeriesName && series?.name) {
              pieces.push(series.name);
            }
            if (chart.dataLabels?.showCategoryName) {
              pieces.push(formatTickValue(point.x));
            }
            if (chart.dataLabels?.showValue) {
              pieces.push(formatTickValue(point.y));
            }
            if (chart.dataLabels?.showBubbleSize) {
              pieces.push(formatTickValue(point.bubble));
            }
            if (chart.dataLabels?.showPercent) {
              pieces.push(`${Math.round((Math.abs(point.bubble) / Math.max(1, bubbleTotals[seriesIndex] ?? 1)) * 100)}%`);
            }
            return (
              <g key={`bubble-${seriesIndex}-${point.index}`}>
                <circle
                  {...chartElementDataProps(seriesIndex, point.index, { selectionMode: "seriesFirst" })}
                  cx={xScale(point.x)}
                  cy={yScale(point.y)}
                  fill={isBubble3d ? `url(#bubble3d-grad-${chart.id}-${seriesIndex})` : baseColor}
                  fillOpacity={isBubble3d ? 0.98 : 0.78}
                  r={radius}
                  stroke={darkenColor(baseColor, 0.18)}
                  strokeWidth={isBubble3d ? 1.2 : 1}
                />
                {isSelectedChartPoint(selectedChartElement, chart.id, seriesIndex, point.index)
                || (
                  isSelectedChartSeries(selectedChartElement, chart.id, seriesIndex)
                  && selectedChartElement?.kind !== "point"
                )
                  ? renderSelectionRectHandles(
                      `bubble-selection-${seriesIndex}-${point.index}`,
                      xScale(point.x) - radius,
                      yScale(point.y) - radius,
                      radius * 2,
                      radius * 2
                    )
                  : null}
                {isBubble3d ? (
                  <ellipse
                    cx={xScale(point.x) - radius * 0.16}
                    cy={yScale(point.y) - radius * 0.22}
                    fill="#ffffff"
                    opacity={0.22}
                    rx={Math.max(1.5, radius * 0.34)}
                    ry={Math.max(1, radius * 0.2)}
                  />
                ) : null}
                {labelsEnabled && pieces.length > 0 ? (
                  <text
                    fill={resolveChartTextColor(chart)}
                    fontSize={10}
                    textAnchor="start"
                    x={xScale(point.x) + radius + 4}
                    y={yScale(point.y) + 3}
                  >
                    {pieces.join(", ")}
                  </text>
                ) : null}
              </g>
            );
          })}
        </g>
      ))}
    </g>
  );
}

function renderRadarChart(chart: XlsxChart, palette: ChartRendererPalette, layout: ChartLayout, selectedChartElement?: XlsxChartElementSelection | null) {
  if (chart.series.length === 0) {
    return null;
  }
  const categories = getCategoryLabels(chart);
  if (categories.length === 0) {
    return null;
  }
  const plot = layout.plot;
  const centerX = plot.left + plot.width * 0.5;
  const centerY = plot.top + plot.height * 0.52;
  const radius = Math.max(22, Math.min(plot.width, plot.height) * 0.38);
  const values = chart.series.flatMap((series) => (
    series.values
      .slice(0, categories.length)
      .map((value) => safeNumber(value))
      .filter((value): value is number => value != null)
  ));
  const hasExplicitMin = typeof chart.valueAxis?.min === "number" && Number.isFinite(chart.valueAxis.min);
  const hasExplicitMax = typeof chart.valueAxis?.max === "number" && Number.isFinite(chart.valueAxis.max);
  let minValue = hasExplicitMin
    ? Number(chart.valueAxis?.min)
    : Math.min(0, ...(values.length ? values : [0]));
  let maxValue = hasExplicitMax
    ? Number(chart.valueAxis?.max)
    : Math.max(1, ...(values.length ? values : [1]));
  if (chart.valueAxis?.crosses === "autoZero") {
    minValue = Math.min(0, minValue);
  }
  const safeMax = maxValue <= minValue ? minValue + 1 : maxValue;
  const span = Math.max(1e-6, safeMax - minValue);
  const candidateMajorUnit = typeof chart.valueAxis?.majorUnit === "number" && chart.valueAxis.majorUnit > 0
    ? chart.valueAxis.majorUnit
    : undefined;
  const preferredMajorUnit = candidateMajorUnit && span / candidateMajorUnit >= 3
    ? candidateMajorUnit
    : buildNiceStep(minValue, safeMax, 6);
  const ticks = buildNumericTickValues(minValue, safeMax, preferredMajorUnit);
  const axisColor = chart.axisLineColor ?? chart.chartAreaBorderColor ?? palette.border;
  const labelColor = resolveChartAxisTextColor(chart);
  const filled = chart.radarStyle === "filled";
  const showSpokes = chart.categoryAxis?.majorGridlines === true;

  const angleAt = (index: number) => (Math.PI * 2 * index) / categories.length - Math.PI / 2;
  const radialPoint = (index: number, valueRatio: number) => {
    const angle = angleAt(index);
    const r = radius * valueRatio;
    return {
      x: centerX + Math.cos(angle) * r,
      y: centerY + Math.sin(angle) * r
    };
  };

  return (
    <g>
      {ticks.map((tick, ringIndex) => {
        const ratio = (tick - minValue) / (safeMax - minValue);
        const ringPoints = categories.map((_, categoryIndex) => radialPoint(categoryIndex, ratio));
        const ringPath = d3Line<{ x: number; y: number }>()
          .x((point) => point.x)
          .y((point) => point.y)
          .curve(curveLinearClosed)(ringPoints) ?? "";
        return (
          <g key={`radar-ring-${ringIndex}`}>
            <path
              d={ringPath}
              fill={ringIndex % 2 === 0 ? "transparent" : lightenColor(chart.chartAreaFillColor ?? palette.surface, 0.04)}
              stroke={lightenColor(axisColor, 0.5)}
              strokeWidth={1}
            />
            <text fill={labelColor} fontSize={9} x={centerX + 4} y={centerY - radius * ratio + 3}>
              {formatTickValue(tick)}
            </text>
          </g>
        );
      })}
      {categories.map((category, categoryIndex) => {
        const edge = radialPoint(categoryIndex, 1);
        return (
          <g key={`radar-axis-${categoryIndex}`}>
            {showSpokes ? (
              <line stroke={lightenColor(axisColor, 0.52)} strokeWidth={1} x1={centerX} x2={edge.x} y1={centerY} y2={edge.y} />
            ) : null}
            <text
              fill={labelColor}
              fontSize={10}
              textAnchor="middle"
              x={centerX + (edge.x - centerX) * 1.1}
              y={centerY + (edge.y - centerY) * 1.1}
            >
              {category}
            </text>
          </g>
        );
      })}
      {chart.series.map((series, seriesIndex) => {
        const points = categories.map((_, categoryIndex) => {
          const rawValue = safeNumber(series.values[categoryIndex]);
          if (rawValue == null) {
            return {
              defined: false,
              pointIndex: categoryIndex,
              ...radialPoint(categoryIndex, 0)
            };
          }
          const ratio = clamp((rawValue - minValue) / (safeMax - minValue), 0, 1);
          return {
            defined: true,
            pointIndex: categoryIndex,
            ...radialPoint(categoryIndex, ratio)
          };
        });
        const definedPoints = points.filter((point) => point.defined);
        if (definedPoints.length === 0) {
          return null;
        }
        const hasGap = points.some((point) => !point.defined);
        const polygon = d3Line<{ defined: boolean; x: number; y: number }>()
          .defined((point) => point.defined)
          .x((point) => point.x)
          .y((point) => point.y)
          .curve(hasGap ? curveLinear : curveLinearClosed)(points) ?? "";
        const color = chartSeriesColor(chart, seriesIndex);
        const markerSymbol = normalizeChartMarkerSymbol(series.markerSymbol);
        const markerSize = markerRadiusPx(series.markerSize);
        const markerPath = markerSymbolPath(markerSymbol, markerSize);
	        const showMarkers = markerSymbol !== "none" && markerPath.length > 0;
	        return (
	          <g key={`radar-series-${seriesIndex}`}>
	            {definedPoints.length >= 2 ? (
	              <path
	                {...chartElementDataProps(seriesIndex)}
	                d={polygon}
	                fill={filled && !hasGap && definedPoints.length >= 3 ? color : "none"}
	                fillOpacity={filled ? 0.26 : 0}
	                stroke={chartSeriesStrokeColor(chart, seriesIndex)}
	                strokeWidth={1.8}
	              />
	            ) : null}
	            {showMarkers
	              ? definedPoints.map((point) => (
	                  <g
	                    key={`radar-point-${seriesIndex}-${point.pointIndex}`}
	                    transform={`translate(${point.x}, ${point.y})`}
	                  >
	                    <path
	                      {...chartElementDataProps(seriesIndex, point.pointIndex)}
	                      d={markerPath}
	                      fill={series.markerColor ?? color}
	                      stroke={series.markerLineColor ?? chart.chartAreaFillColor ?? palette.surface}
	                      strokeWidth={1}
	                    />
	                  </g>
	                ))
	              : null}
	            {definedPoints.map((point) => (
	              isSelectedChartPoint(selectedChartElement, chart.id, seriesIndex, point.pointIndex)
	              || (
	                isSelectedChartSeries(selectedChartElement, chart.id, seriesIndex)
	                && selectedChartElement?.kind !== "point"
	              )
	                ? renderSelectionPointHandles(
	                    `radar-selection-${seriesIndex}-${point.pointIndex}`,
	                      [{ x: point.x, y: point.y }]
	                    )
	                : null
	            ))}
	          </g>
	        );
      })}
    </g>
  );
}

function renderPieChart(chart: XlsxChart, palette: ChartRendererPalette, layout: ChartLayout, chartType: string, selectedChartElement?: XlsxChartElementSelection | null) {
  const pieSeriesIndex = selectPrimaryPieSeriesIndex(chart);
  const pieSeries = chart.series[pieSeriesIndex];
  const pieData = buildPieEntries(chart, pieSeriesIndex);
  if (pieData.length === 0) {
    return null;
  }

  const legendOnRight = layout.legendPosition === "right";
  const centerX = legendOnRight ? layout.plot.left + layout.plot.width * 0.42 : layout.plot.left + layout.plot.width * 0.5;
  const centerY = layout.plot.top + layout.plot.height * 0.54;
  const outerRadius = Math.max(16, Math.min(layout.plot.width, layout.plot.height) * (chartType === "Doughnut" ? 0.32 : 0.38));
  const innerRatio = chartType === "Doughnut" ? clamp((chart.holeSize ?? 56) / 100, 0.1, 0.9) : 0;
  const innerRadius = outerRadius * innerRatio;
  const startAngle = ((chart.firstSliceAngle ?? 0) * Math.PI) / 180;
  const arcs = d3Pie<{ color: string; index: number; label: string; value: number }>()
    .value((entry) => entry.value)
    .sort(null)
    .startAngle(startAngle)
    .endAngle(startAngle + Math.PI * 2)(pieData);

  const arcPath = d3Arc<typeof arcs[number]>()
    .innerRadius(innerRadius)
    .outerRadius(outerRadius);

  const isPie3d = chartType === "Pie3D" || (chartType === "PieExploded" && chart.is3d === true);
  const rotX = chart.view3d?.rotX ?? 20;
  const perspective = chart.view3d?.perspective ?? 30;
  const tiltFromRotX = clamp(0.14 + (rotX / 90) * 0.72, 0.12, 0.92);
  const perspectiveCompression = clamp(1 - (perspective / 180), 0.55, 1);
  const tilt = isPie3d ? clamp(tiltFromRotX * perspectiveCompression, 0.14, 0.78) : 1;
  const depthScale = (chart.view3d?.depthPercent ?? 100) / 100;
  const depth = isPie3d ? Math.max(9, outerRadius * 0.28 * depthScale) : 0;
  const dataLabelsEnabled = Boolean(chart.dataLabels?.showCategoryName || chart.dataLabels?.showPercent || chart.dataLabels?.showValue);
  const total = pieData.reduce((sum, entry) => sum + entry.value, 0);
  const shadowId = `pie3d-shadow-${chart.id}`.replace(/[^A-Za-z0-9_-]/g, "-");
  const baseShadowId = `pie3d-base-shadow-${chart.id}`.replace(/[^A-Za-z0-9_-]/g, "-");
  const sliceSeparatorColor = chart.chartAreaFillColor ?? "#ffffff";
  const labelBounds = {
    bottom: layout.height - 6,
    left: 6,
    right: layout.width - 6,
    top: layout.plot.top + 8
  };
  const pointLabelByIndex = new Map((chart.dataLabels?.pointLabels ?? []).map((label) => [label.index, label]));
  const centerValuePointLabel = chartType === "Doughnut"
    ? (chart.dataLabels?.pointLabels ?? []).find((label) => (
      label.deleted !== true
      && label.showValue === true
      && label.showCategoryName !== true
      && label.showPercent !== true
      && label.showSeriesName !== true
      && label.showBubbleSize !== true
    )) ?? null
    : null;
  const centerValueEntry = centerValuePointLabel
    ? pieData.find((entry) => entry.index === centerValuePointLabel.index) ?? null
    : null;
  const shouldRenderCenterValue = chartType === "Doughnut" && centerValueEntry != null;
  const centerValueLooksPercent = centerValueEntry != null
    && total > 0
    && total <= 1.0000001
    && centerValueEntry.value >= 0
    && centerValueEntry.value <= 1.0000001;
  const centerValueText = centerValueEntry == null
    ? ""
    : centerValueLooksPercent
      ? `${Math.round(centerValueEntry.value * 100)}%`
      : formatTickValue(centerValueEntry.value);
  const centerValueFontSize = shouldRenderCenterValue
    ? Math.max(
        14,
        Math.min(
          Math.max(20, innerRadius * 0.72),
          (centerValuePointLabel?.fontSizePt ?? 28) * 0.85
        )
      )
    : 0;

  const resolveSliceExplosion = (pointIndex: number) => {
    const pointStyle = pieSeries?.dataPointStyles?.find((entry) => entry.index === pointIndex);
    const seriesExplosion = typeof pieSeries?.shapeProperties?.xmlExplosion === "number"
      ? pieSeries.shapeProperties.xmlExplosion
      : 0;
    const rawExplosion = Math.max(0, pointStyle?.explosion ?? seriesExplosion ?? 0);
    if (rawExplosion <= 0) {
      return 0;
    }
    // OOXML explosion values are percent-like offsets of the pie radius.
    return outerRadius * clamp(rawExplosion / 100, 0, 4);
  };
  const hasExplodedSlices = arcs.some((arc) => resolveSliceExplosion(arc.data.index) > 0);
  const sliceSeparatorWidth = hasExplodedSlices ? 2 : 1.2;

  return (
    <g>
      {isPie3d ? (
        <defs>
          <filter id={shadowId} x="-40%" y="-40%" width="180%" height="200%">
            <feDropShadow dx="1.2" dy="3.6" floodColor="#000000" floodOpacity="0.28" stdDeviation="2.8" />
          </filter>
          <filter id={baseShadowId} x="-50%" y="-50%" width="220%" height="220%">
            <feGaussianBlur stdDeviation="3.2" />
          </filter>
        </defs>
      ) : null}
      {isPie3d ? (
        <ellipse
          cx={centerX}
          cy={centerY + depth + 3}
          fill="#000000"
          filter={`url(#${baseShadowId})`}
          opacity={0.14}
          rx={outerRadius * 1.02}
          ry={outerRadius * tilt * 0.68}
        />
      ) : null}
      {isPie3d
          ? arcs.map((arc) => {
              const midAngle = (arc.startAngle + arc.endAngle) / 2;
              const explosion = resolveSliceExplosion(arc.data.index);
              const explodeX = Math.sin(midAngle) * explosion;
              const explodeY = -Math.cos(midAngle) * explosion * tilt;
              const sidePaths = buildPieOuterWallPath(
                centerX,
                centerY,
                outerRadius,
                tilt,
                depth,
                arc.startAngle,
                arc.endAngle,
                explosion > 0
              );
            if (sidePaths.length === 0) {
              return null;
            }
            return (
              <g key={`pie-side-${arc.data.index}`} transform={`translate(${explodeX}, ${explodeY})`}>
                {sidePaths.map((sidePath, sideIndex) => (
                  <path
                    d={sidePath}
                    fill={darkenColor(arc.data.color, 0.34)}
                    key={`pie-side-path-${arc.data.index}-${sideIndex}`}
                    stroke={darkenColor(arc.data.color, 0.5)}
                    strokeWidth={0.8}
                  />
                ))}
                {explosion > 0 ? (() => {
                  const startWall = buildPieRadialWallPath(
                    centerX,
                    centerY,
                    outerRadius,
                    tilt,
                    depth,
                    arc.startAngle,
                    true
                  );
                  const endWall = buildPieRadialWallPath(
                    centerX,
                    centerY,
                    outerRadius,
                    tilt,
                    depth,
                    arc.endAngle,
                    true
                  );
                  return (
                    <>
                      {startWall ? (
                        <path
                          d={startWall}
                          fill={darkenColor(arc.data.color, 0.26)}
                          stroke={darkenColor(arc.data.color, 0.44)}
                          strokeWidth={0.8}
                        />
                      ) : null}
                      {endWall ? (
                        <path
                          d={endWall}
                          fill={darkenColor(arc.data.color, 0.2)}
                          stroke={darkenColor(arc.data.color, 0.4)}
                          strokeWidth={0.8}
                        />
                      ) : null}
                    </>
                  );
                })() : null}
              </g>
            );
          })
        : null}
      {arcs.map((arc) => {
        const explosion = resolveSliceExplosion(arc.data.index);
        const midAngle = (arc.startAngle + arc.endAngle) / 2;
        const explodeX = Math.sin(midAngle) * explosion;
        const explodeY = -Math.cos(midAngle) * explosion * (isPie3d ? tilt : 1);
        const isSliceSelected = isSelectedChartPoint(selectedChartElement, chart.id, pieSeriesIndex, arc.data.index)
          || (
            isSelectedChartSeries(selectedChartElement, chart.id, pieSeriesIndex)
            && selectedChartElement?.kind !== "point"
          );
        const sliceHandlePoints = isSliceSelected
          ? [
              pieEllipsePoint(centerX, centerY, outerRadius, isPie3d ? tilt : 1, arc.startAngle),
              pieEllipsePoint(centerX, centerY, outerRadius, isPie3d ? tilt : 1, midAngle),
              pieEllipsePoint(centerX, centerY, outerRadius, isPie3d ? tilt : 1, arc.endAngle),
              innerRadius > 0
                ? pieEllipsePoint(centerX, centerY, innerRadius, isPie3d ? tilt : 1, midAngle)
                : { x: centerX, y: centerY }
            ].map((point) => ({
              x: point.x + explodeX,
              y: point.y + explodeY
            }))
          : [];
        const labelRadius = outerRadius + (chartType === "PieExploded" ? 8 : 12);
        const labelX = centerX + Math.sin(midAngle) * labelRadius + explodeX;
        const labelY = centerY - Math.cos(midAngle) * labelRadius * (isPie3d ? tilt : 1) + explodeY;
        const pieces: string[] = [];
        if (chart.dataLabels?.showCategoryName && arc.data.label.trim().length > 0) {
          pieces.push(arc.data.label);
        }
        const pointLabel = pointLabelByIndex.get(arc.data.index);
        const showValue = pointLabel?.showValue ?? chart.dataLabels?.showValue;
        const showPercent = pointLabel?.showPercent ?? chart.dataLabels?.showPercent;
        if (showValue) {
          pieces.push(formatTickValue(arc.data.value));
        }
        if (showPercent) {
          pieces.push(`${Math.round((arc.data.value / Math.max(1, total)) * 100)}%`);
        }
        const labelText = pieces.join(", ");
        const truncatedLabelText = truncateSvgText(labelText, Math.max(48, layout.width * 0.42), 10);
        const approxLabelWidth = Math.max(12, truncatedLabelText.length * 5.6);
        let labelAnchor: "end" | "start" = labelX >= centerX ? "start" : "end";
        let resolvedLabelX = labelX;
        if (labelAnchor === "start" && resolvedLabelX + approxLabelWidth > labelBounds.right) {
          labelAnchor = "end";
        }
        if (labelAnchor === "end" && resolvedLabelX - approxLabelWidth < labelBounds.left) {
          labelAnchor = "start";
        }
        resolvedLabelX = clamp(resolvedLabelX, labelBounds.left, labelBounds.right);
        const resolvedLabelY = clamp(labelY, labelBounds.top, labelBounds.bottom);
        return (
          <React.Fragment key={`pie-top-${arc.data.index}`}>
            <g transform={`translate(${explodeX}, ${explodeY})`}>
              <path
                {...chartElementDataProps(pieSeriesIndex, arc.data.index)}
                d={(arcPath(arc) ?? "")}
                fill={arc.data.color}
                filter={isPie3d ? `url(#${shadowId})` : undefined}
                stroke={isSliceSelected ? "#64748b" : sliceSeparatorColor}
                strokeWidth={isSliceSelected ? Math.max(2, sliceSeparatorWidth) : sliceSeparatorWidth}
                transform={`translate(${centerX}, ${centerY})${isPie3d ? ` scale(1, ${tilt})` : ""}`}
              />
            </g>
            {isSliceSelected
              ? renderSelectionPointHandles(`pie-selection-${arc.data.index}`, sliceHandlePoints)
              : null}
            {dataLabelsEnabled && truncatedLabelText.length > 0 ? (
              <text
                fill={resolveChartTextColor(chart)}
                fontSize={10}
                textAnchor={labelAnchor}
                x={resolvedLabelX}
                y={resolvedLabelY}
              >
                {truncatedLabelText}
              </text>
            ) : null}
          </React.Fragment>
        );
      })}
      {shouldRenderCenterValue ? (
        <text
          fill={chart.textColor ?? chart.titleColor ?? DEFAULT_CHART_TEXT_COLOR}
          fontSize={centerValueFontSize}
          fontWeight={700}
          textAnchor="middle"
          x={centerX}
          y={centerY + centerValueFontSize * 0.34}
        >
          {centerValueText}
        </text>
      ) : null}
    </g>
  );
}

function renderBarOfPieChart(chart: XlsxChart, palette: ChartRendererPalette, layout: ChartLayout, selectedChartElement?: XlsxChartElementSelection | null) {
  const pieSeriesIndex = selectPrimaryPieSeriesIndex(chart);
  const pieSeries = chart.series[pieSeriesIndex];
  const categories = getCategoryLabels(chart);
  const values = pieSeries?.values.map((value) => Math.max(0, safeNumber(value) ?? 0)) ?? [];
  if (values.length === 0) {
    return null;
  }
  const raw = (chart.raw ?? {}) as Record<string, unknown>;
  const ofPieType = raw.ofPieType === "pie" ? "pie" : "bar";
  const splitPos = typeof raw.splitPos === "number" ? raw.splitPos : 0;
  let secondaryIndices = values
    .map((value, index) => ({ index, value }))
    .filter(({ value }) => value <= splitPos)
    .map(({ index }) => index);
  if (secondaryIndices.length === 0) {
    secondaryIndices = values
      .map((value, index) => ({ index, value }))
      .sort((left, right) => left.value - right.value)
      .slice(0, Math.min(2, values.length))
      .map(({ index }) => index);
  }
  const secondarySet = new Set(secondaryIndices);
  const secondaryTotal = secondaryIndices.reduce((sum, index) => sum + (values[index] ?? 0), 0);
  const primaryData = values.flatMap((value, index) => secondarySet.has(index)
    ? []
    : [{ color: chartPointColor(chart, index, pieSeriesIndex), index, label: categories[index], value }]);
  if (secondaryTotal > 0) {
    primaryData.push({
      color: chartPointColor(chart, secondaryIndices[0] ?? 0, pieSeriesIndex),
      index: secondaryIndices[0] ?? 0,
      label: "Other",
      value: secondaryTotal
    });
  }

  const pieCenterX = layout.plot.left + layout.plot.width * 0.28;
  const pieCenterY = layout.plot.top + layout.plot.height * 0.55;
  const pieRadius = Math.max(16, Math.min(layout.plot.height, layout.plot.width * 0.45) * 0.3);
  const arc = d3Arc<{ endAngle: number; startAngle: number }>().innerRadius(0).outerRadius(pieRadius);
  const pieArcs = d3Pie<{ color: string; index: number; label: string; value: number }>()
    .value((entry) => entry.value)
    .sort(null)
    .startAngle(((90 - (chart.firstSliceAngle ?? 0)) * Math.PI) / 180)
    .endAngle(((90 - (chart.firstSliceAngle ?? 0)) * Math.PI) / 180 + Math.PI * 2)(primaryData);

  const secondaryLabels = secondaryIndices.map((index) => categories[index] ?? "");
  const secondaryValues = secondaryIndices.map((index) => values[index] ?? 0);
  const connectorTargetX = layout.plot.left + layout.plot.width * 0.69;
  const secondaryData = secondaryIndices.map((index) => ({
    color: chartPointColor(chart, index, pieSeriesIndex),
    index,
    label: categories[index] ?? "",
    value: values[index] ?? 0
  }));
  const secondaryCenterX = layout.plot.left + layout.plot.width * 0.79;
  const secondaryCenterY = pieCenterY;
  const secondaryRadius = pieRadius * clamp(((typeof raw.secondPieSize === "number" ? raw.secondPieSize : 100) / 100), 0.55, 1.5);
  const secondaryArc = d3Arc<{ endAngle: number; startAngle: number }>().innerRadius(0).outerRadius(secondaryRadius);
  const secondaryPieArcs = d3Pie<{ color: string; index: number; label: string; value: number }>()
    .value((entry) => entry.value)
    .sort(null)
    .startAngle(((90 - (chart.firstSliceAngle ?? 0)) * Math.PI) / 180)
    .endAngle(((90 - (chart.firstSliceAngle ?? 0)) * Math.PI) / 180 + Math.PI * 2)(secondaryData);
  const stackedBarLeft = layout.plot.left + layout.plot.width * 0.72;
  const stackedBarWidth = Math.max(20, layout.plot.width * 0.13);
  const stackedBarTop = layout.plot.top + 16;
  const stackedBarHeight = Math.max(28, layout.plot.height - 32);
  const secondaryTotalSafe = Math.max(1e-6, secondaryTotal);
  let stackCursor = stackedBarTop;

  return (
    <g>
        {pieArcs.map((entry, index) => {
          const selected = isSelectedChartPointOrSeries(selectedChartElement, chart.id, pieSeriesIndex, entry.data.index);
          const midAngle = (entry.startAngle + entry.endAngle) / 2;
          const dot = pieEllipsePoint(pieCenterX, pieCenterY, pieRadius * 0.62, 1, midAngle);
          return (
            <React.Fragment key={`bar-of-pie-main-${index}`}>
              <path
                {...chartElementDataProps(pieSeriesIndex, entry.data.index)}
                d={arc(entry) ?? ""}
                fill={entry.data.color}
                stroke={selected ? "#64748b" : chart.chartAreaFillColor ?? palette.surface}
                strokeWidth={selected ? 2 : 1}
                transform={`translate(${pieCenterX}, ${pieCenterY})`}
              />
              {selected
                ? renderSelectionPointHandles(`bar-of-pie-main-selection-${index}`, [dot])
                : null}
            </React.Fragment>
          );
        })}
      <line
        stroke={chart.chartAreaBorderColor ?? palette.border}
        strokeWidth={1}
        x1={pieCenterX + pieRadius}
        x2={connectorTargetX}
        y1={pieCenterY - pieRadius * 0.4}
        y2={layout.plot.top + 10}
      />
      <line
        stroke={chart.chartAreaBorderColor ?? palette.border}
        strokeWidth={1}
        x1={pieCenterX + pieRadius}
        x2={connectorTargetX}
        y1={pieCenterY + pieRadius * 0.4}
        y2={layout.plot.top + layout.plot.height - 10}
      />
      {ofPieType === "pie"
          ? secondaryPieArcs.map((entry, index) => {
            const selected = isSelectedChartPointOrSeries(selectedChartElement, chart.id, pieSeriesIndex, entry.data.index);
            const midAngle = (entry.startAngle + entry.endAngle) / 2;
            const dot = pieEllipsePoint(secondaryCenterX, secondaryCenterY, secondaryRadius * 0.62, 1, midAngle);
            return (
              <React.Fragment key={`bar-of-pie-secondary-pie-${index}`}>
                <path
                  {...chartElementDataProps(pieSeriesIndex, entry.data.index)}
                  d={secondaryArc(entry) ?? ""}
                  fill={entry.data.color}
                  stroke={selected ? "#64748b" : chart.chartAreaFillColor ?? palette.surface}
                  strokeWidth={selected ? 2 : 1}
                  transform={`translate(${secondaryCenterX}, ${secondaryCenterY})`}
                />
                {selected
                  ? renderSelectionPointHandles(`bar-of-pie-secondary-pie-selection-${index}`, [dot])
                  : null}
              </React.Fragment>
            );
          })
        : secondaryData.map((entry, index) => {
          const segmentHeight = index === secondaryData.length - 1
            ? Math.max(1, stackedBarTop + stackedBarHeight - stackCursor)
            : Math.max(1, (entry.value / secondaryTotalSafe) * stackedBarHeight);
          const y = stackCursor;
          stackCursor += segmentHeight;
          const selected = isSelectedChartPointOrSeries(selectedChartElement, chart.id, pieSeriesIndex, entry.index);
          return (
            <g key={`bar-of-pie-secondary-bar-${index}`}>
                <rect
                  {...chartElementDataProps(pieSeriesIndex, entry.index)}
                  fill={entry.color}
                  height={segmentHeight}
                  stroke={selected ? "#64748b" : chart.chartAreaFillColor ?? palette.surface}
                  strokeWidth={selected ? 2 : 1}
                width={stackedBarWidth}
                x={stackedBarLeft}
                y={y}
              />
              {selected
                ? renderSelectionRectHandles(
                    `bar-of-pie-secondary-bar-selection-${index}`,
                    stackedBarLeft,
                    y,
                    stackedBarWidth,
                    segmentHeight
                  )
                : null}
              <text
                fill={resolveChartAxisTextColor(chart)}
                fontSize={10}
                textAnchor="start"
                x={stackedBarLeft + stackedBarWidth + 6}
                y={y + segmentHeight * 0.5 + 3}
              >
                {entry.label}
              </text>
            </g>
          );
        })}
    </g>
  );
}

function renderSurfaceChart(chart: XlsxChart, palette: ChartRendererPalette, layout: ChartLayout, selectedChartElement?: XlsxChartElementSelection | null) {
  const plot = resolveSurfacePlotRect(chart, layout);
  const categories = getCategoryLabels(chart);
  const rows = chart.series.length;
  const cols = Math.max(
    categories.length,
    chart.series.reduce((max, series) => Math.max(max, series.values.length), 0)
  );
  if (rows === 0 || cols === 0) {
    return null;
  }
  const matrix = chart.series.map((series) => (
    Array.from({ length: cols }, (_, columnIndex) => safeNumber(series.values[columnIndex]))
  ));
  const domain = getSurfaceDomain(chart);
  if (!domain) {
    return null;
  }
  const isContour = isContourSurfaceChart(chart);
  const minValue = domain.minValue;
  const safeMax = domain.safeMax;
  const wallFill = chart.backWall?.fillColor ?? "#d9d9df";
  const wallLineColor = chart.backWall?.lineColor ?? chart.sideWall?.lineColor ?? chart.floor?.lineColor ?? (chart.axisLineColor ?? lightenColor(resolveSurfaceBaseColor(chart, palette), 0.4));

  if (isContour) {
    const faces: React.ReactNode[] = [];
    const lines: React.ReactNode[] = [];
    const inset = chart.valueAxis?.crossBetween !== "midCat" ? plot.width / (2 * cols) : 0;
    const dataWidth = plot.width - 2 * inset;
    for (let rowIndex = 0; rowIndex < rows - 1; rowIndex += 1) {
      for (let columnIndex = 0; columnIndex < cols - 1; columnIndex += 1) {
        const p00 = matrix[rowIndex]?.[columnIndex];
        const p10 = matrix[rowIndex]?.[columnIndex + 1];
        const p01 = matrix[rowIndex + 1]?.[columnIndex];
        const p11 = matrix[rowIndex + 1]?.[columnIndex + 1];
        if (p00 == null || p10 == null || p01 == null || p11 == null) {
          continue;
        }
        const x0 = plot.left + inset + columnIndex / (cols - 1) * dataWidth;
        const x1 = plot.left + inset + (columnIndex + 1) / (cols - 1) * dataWidth;
        const y0 = plot.top + plot.height - rowIndex / (rows - 1) * plot.height;
        const y1 = plot.top + plot.height - (rowIndex + 1) / (rows - 1) * plot.height;
        const triangles = triangulateSurfaceQuad([
          { value: p00, x: x0, y: y0 },
          { value: p10, x: x1, y: y0 },
          { value: p11, x: x1, y: y1 },
          { value: p01, x: x0, y: y1 }
        ], chart.valueAxis);
        triangles.forEach((triangle, triangleIndex) => {
          const world = triangle.map(point => ({
            x: (point.x - plot.left - inset) / dataWidth,
            y: (1 - normalizeSurfaceValue(point.value, domain.minValue, domain.safeMax, chart.valueAxis)) * surfaceHeightRatio(chart.view3d?.heightPercent, layout.plot),
            z: (point.y - plot.top) / plot.height * ((chart.view3d?.depthPercent ?? 100) / 100)
          }));
          const normal = transformSurfaceNormal(surfaceNormal(world[0], world[1], world[2]), chart.view3d?.rotX ?? 90, chart.view3d?.rotY ?? 0);
          for (let band = 0; band < domain.ticks.length - 1; band += 1) {
            const lower = domain.ticks[band];
            const upper = domain.ticks[band + 1];
            const polygon = clipSurfaceBand(clipSurfaceBand(triangle, lower, true, chart.valueAxis, band === 0), upper, false, chart.valueAxis);
            if (polygon.length < 3) {
              continue;
            }
            const key = `${rowIndex}-${columnIndex}-${triangleIndex}-${band}`;
            const midpoint = (lower + upper) / 2;
            if (!chart.wireframe) {
              faces.push(<polygon key={`surface-band-${key}`} fill={chart.surfaceMaterial === "flat" ? resolveSurfaceBandColor(chart, palette, domain, midpoint) : shadeSurfaceMatteColor(resolveSurfaceBandColor(chart, palette, domain, midpoint), normal)} points={polygon.map((point) => `${point.x},${point.y}`).join(" ")} />);
            }
            const widths = chart.raw?.bandFormatLineWidths as Record<string, number> | undefined;
            const lineWidth = widths?.[band] ?? chart.series[0]?.lineWidthPx ?? 1;
            const edges = surfaceBandEdges(polygon, triangle);
            lines.push(
              <path
                key={`surface-edge-${key}`}
                d={edges.map(([from, to]) => `M ${from.x} ${from.y} L ${to.x} ${to.y}`).join(" ")}
                fill="none"
                stroke={resolveSurfaceWireframeColor(chart, palette, domain, midpoint)}
                strokeWidth={lineWidth}
              />
            );
          }
        });
      }
    }
    return (
      <g>
        <rect
          fill={chart.raw?.plotAreaNoFill === true ? "transparent" : typeof chart.raw?.plotAreaFillColor === "string" ? chart.raw.plotAreaFillColor : wallFill}
          x={layout.plot.left}
          y={layout.plot.top}
          width={layout.plot.width}
          height={layout.plot.height}
          stroke={typeof chart.raw?.plotAreaLineColor === "string" ? chart.raw.plotAreaLineColor : wallLineColor}
        />
        <rect fill="none" height={plot.height} stroke={wallLineColor} width={plot.width} x={plot.left} y={plot.top} />
        {faces}
        {lines}
      </g>
    );
  }

  const rotX = clamp(chart.view3d?.rotX ?? (chart.wireframe ? 90 : 25), -88, 88) * (Math.PI / 180);
  const rotY = clamp(chart.view3d?.rotY ?? (chart.wireframe ? 0 : 30), -88, 88) * (Math.PI / 180);
  const usePerspective = chart.view3d?.rAngAx === false;
  const perspectiveStrength = clamp(
    (chart.view3d?.perspective ?? (usePerspective ? 30 : 0)) / 100,
    0,
    1
  );
  const depthScale = clamp((chart.view3d?.depthPercent ?? 100) / 100, 0.2, 4);
  const heightExtent = 2 * surfaceHeightRatio(chart.view3d?.heightPercent, layout.plot);
  const cosX = Math.cos(rotX);
  const sinX = Math.sin(rotX);
  const cosY = Math.cos(rotY);
  const sinY = Math.sin(rotY);

  type SurfacePoint = { depth: number; hasValue: boolean; value: number; x: number; y: number };
  const projectRawPoint = (columnIndex: number, rowIndex: number, value: number | null): SurfacePoint => {
    const hasValue = value != null;
    const normalizedX = cols <= 1 ? 0 : ((columnIndex / (cols - 1)) - 0.5) * 2;
    const normalizedY = hasValue ? -(normalizeSurfaceValue(value, minValue, safeMax, chart.valueAxis) - 0.5) * heightExtent : heightExtent / 2;
    const normalizedZ = rows <= 1 ? 0 : ((rowIndex / (rows - 1)) - 0.5) * 2 * depthScale;
    const x1 = normalizedX * cosY + normalizedZ * sinY;
    const z1 = -normalizedX * sinY + normalizedZ * cosY;
    const y1 = normalizedY * cosX - z1 * sinX;
    const z2 = normalizedY * sinX + z1 * cosX;
    const perspective = usePerspective ? 1 / Math.max(0.18, 1 + z2 * (0.24 + perspectiveStrength * 0.5)) : 1;
    return { depth: z2, hasValue, value: value ?? minValue, x: x1 * perspective, y: y1 * perspective };
  };
  const rawPoints = matrix.map((row, rowIndex) => row.map((value, columnIndex) => projectRawPoint(columnIndex, rowIndex, value)));

  const bounds = rawPoints.flat();
  const minX = Math.min(...bounds.map((point) => point.x));
  const maxX = Math.max(...bounds.map((point) => point.x));
  const minY = Math.min(...bounds.map((point) => point.y));
  const maxY = Math.max(...bounds.map((point) => point.y));
  const scale = Math.min(
    layout.plot.width / Math.max(0.25, maxX - minX),
    layout.plot.height / Math.max(0.25, maxY - minY)
  ) * 0.82;
  const centerX = layout.plot.left + layout.plot.width / 2;
  const centerY = layout.plot.top + layout.plot.height / 2;
  const centerRawX = (minX + maxX) / 2;
  const centerRawY = (minY + maxY) / 2;

  const projectPoint = (point: {x: number; y: number; value: number}) => {
    const projected = projectRawPoint(point.x, point.y, point.value);
    return { ...projected, x: centerX + (projected.x - centerRawX) * scale, y: centerY + (projected.y - centerRawY) * scale };
  };
  const faces: Array<{ depth: number; node: React.ReactNode }> = [];
  for (let rowIndex = 0; rowIndex < rows - 1; rowIndex += 1) {
    for (let columnIndex = 0; columnIndex < cols - 1; columnIndex += 1) {
      const values = [matrix[rowIndex][columnIndex], matrix[rowIndex][columnIndex + 1], matrix[rowIndex + 1][columnIndex + 1], matrix[rowIndex + 1][columnIndex]];
      if (values.some(value => value == null)) continue;
      const triangles = triangulateSurfaceQuad([
        { value: values[0]!, x: columnIndex, y: rowIndex },
        { value: values[1]!, x: columnIndex + 1, y: rowIndex },
        { value: values[2]!, x: columnIndex + 1, y: rowIndex + 1 },
        { value: values[3]!, x: columnIndex, y: rowIndex + 1 }
      ], chart.valueAxis);
      triangles.forEach((triangle, triangleIndex) => {
        const world = triangle.map(point => ({
          x: point.x / Math.max(1, cols - 1),
          y: (1 - normalizeSurfaceValue(point.value, minValue, safeMax, chart.valueAxis)) * surfaceHeightRatio(chart.view3d?.heightPercent, layout.plot),
          z: (1 - point.y / Math.max(1, rows - 1)) * depthScale
        }));
        const normal = transformSurfaceNormal(surfaceNormal(world[0], world[1], world[2]), chart.view3d?.rotX ?? 25, chart.view3d?.rotY ?? 30);
        for (let band = 0; band < domain.ticks.length - 1; band += 1) {
          const lower = domain.ticks[band];
          const upper = domain.ticks[band + 1];
          const polygon = clipSurfaceBand(clipSurfaceBand(triangle, lower, true, chart.valueAxis, band === 0), upper, false, chart.valueAxis);
          if (polygon.length < 3) continue;
          const projected = polygon.map(projectPoint);
          const midpoint = (lower + upper) / 2;
          const widths = chart.raw?.bandFormatLineWidths as Record<string, number> | undefined;
          const edges = surfaceBandEdges(polygon, triangle).map(([from, to]) => [projectPoint(from), projectPoint(to)]);
          faces.push({
            depth: projected.reduce((sum, point) => sum + point.depth, 0) / projected.length,
            node: <g key={`surface-band-${rowIndex}-${columnIndex}-${triangleIndex}-${band}`}>
              {!chart.wireframe && <polygon fill={chart.surfaceMaterial === "flat" ? resolveSurfaceBandColor(chart, palette, domain, midpoint) : shadeSurfaceMatteColor(resolveSurfaceBandColor(chart, palette, domain, midpoint), normal)} points={projected.map(point => `${point.x},${point.y}`).join(" ")} />}
              <path fill="none" stroke={resolveSurfaceWireframeColor(chart, palette, domain, midpoint)} strokeWidth={widths?.[band] ?? chart.series[0]?.lineWidthPx ?? 1} d={edges.map(([from, to]) => `M ${from.x} ${from.y} L ${to.x} ${to.y}`).join(" ")} />
            </g>
          });
        }
      });
    }
  }
  faces.sort((left, right) => left.depth - right.depth);
  return <g>{faces.map(face => face.node)}</g>;
}

function renderSurfaceHitOverlay(
  chart: XlsxChart,
  layout: ChartLayout,
  selectedChartElement?: XlsxChartElementSelection | null
) {
  const plot = resolveSurfacePlotRect(chart, layout);
  const categories = getCategoryLabels(chart);
  const rows = chart.series.length;
  const cols = Math.max(
    categories.length,
    chart.series.reduce((max, series) => Math.max(max, series.values.length), 0)
  );
  if (rows === 0 || cols === 0) {
    return null;
  }

  const cellWidth = plot.width / cols;
  const cellHeight = plot.height / rows;
  return (
    <g>
      {chart.series.flatMap((_, seriesIndex) => (
        Array.from({ length: cols }, (_, pointIndex) => {
          const left = plot.left + pointIndex * cellWidth;
          const top = plot.top + seriesIndex * cellHeight;
          const selected = isSelectedChartPointOrSeries(selectedChartElement, chart.id, seriesIndex, pointIndex);
          return (
            <React.Fragment key={`surface-hit-${seriesIndex}-${pointIndex}`}>
              <rect
                {...chartElementDataProps(seriesIndex, pointIndex)}
                fill="transparent"
                height={cellHeight}
                stroke="none"
                width={cellWidth}
                x={left}
                y={top}
              />
              {selected
                ? renderSelectionRectHandles(
                    `surface-selection-${seriesIndex}-${pointIndex}`,
                    left,
                    top,
                    cellWidth,
                    cellHeight
                  )
                : null}
            </React.Fragment>
          );
        })
      ))}
    </g>
  );
}

function indexByName(series: XlsxChart["series"], matcher: RegExp) {
  const index = series.findIndex((entry) => matcher.test((entry.name ?? "").toLowerCase()));
  return index >= 0 ? index : null;
}

function renderStockChart(chart: XlsxChart, palette: ChartRendererPalette, layout: ChartLayout, selectedChartElement?: XlsxChartElementSelection | null) {
  return renderLineOrAreaChart(chart, palette, layout, "Stock", selectedChartElement);
}

function renderWaterfallChart(chart: XlsxChart, palette: ChartRendererPalette, layout: ChartLayout, selectedChartElement?: XlsxChartElementSelection | null) {
  const stages = buildChartStages(chart);
  if (stages.length === 0) {
    return null;
  }

  const plot = layout.plot;
  const labels = stages.map((stage) => stage.label);
  const xScale = scaleBand<string>()
    .domain(labels)
    .range([plot.left, plot.left + plot.width])
    .paddingInner(0.34)
    .paddingOuter(0.18);

  const bars: Array<ChartStage & { end: number; index: number; start: number }> = [];
  let runningTotal = 0;
  stages.forEach((stage, index) => {
    const start = stage.isSubtotal ? 0 : runningTotal;
    const end = stage.isSubtotal ? runningTotal : runningTotal + stage.value;
    if (!stage.isSubtotal) {
      runningTotal = end;
    }
    bars.push({ ...stage, end, index, start });
  });

  const extents = bars.flatMap((bar) => [0, bar.start, bar.end]);
  let minValue = Math.min(...extents);
  let maxValue = Math.max(...extents);
  if (maxValue <= minValue) {
    maxValue = minValue + 1;
  }

  const yScale = scaleLinear()
    .domain([minValue, maxValue])
    .range([plot.top + plot.height, plot.top]);
  const ticks = buildNumericTickValues(minValue, maxValue, chart.valueAxis?.majorUnit);
  const axisColor = chart.axisLineColor ?? chart.chartAreaBorderColor ?? palette.border;
  const labelColor = resolveChartAxisTextColor(chart);

  return (
    <g>
      {ticks.map((tick) => (
        <g key={`waterfall-y-${tick}`}>
          <line
            stroke={lightenColor(axisColor, 0.72)}
            strokeWidth={1}
            x1={plot.left}
            x2={plot.left + plot.width}
            y1={yScale(tick)}
            y2={yScale(tick)}
          />
          <text fill={labelColor} fontSize={10} textAnchor="end" x={plot.left - 6} y={yScale(tick) + 3}>
            {formatTickValue(tick)}
          </text>
        </g>
      ))}
      <line stroke={axisColor} strokeWidth={1.2} x1={plot.left} x2={plot.left} y1={plot.top} y2={plot.top + plot.height} />
      <line stroke={axisColor} strokeWidth={1.2} x1={plot.left} x2={plot.left + plot.width} y1={yScale(0)} y2={yScale(0)} />
      {bars.map((bar, index) => {
        const bandLeft = xScale(bar.label) ?? plot.left;
        const bandWidth = xScale.bandwidth();
        const topValue = Math.max(bar.start, bar.end);
        const bottomValue = Math.min(bar.start, bar.end);
        const top = yScale(topValue);
        const height = Math.max(1, yScale(bottomValue) - top);
        const fill = bar.isSubtotal
          ? darkenColor(bar.color, 0.18)
          : bar.value >= 0
            ? bar.color
            : lightenColor(bar.color, 0.22);
        const connectorStart = index > 0 ? bars[index - 1] : null;
        const connectorY = connectorStart ? yScale(connectorStart.end) : 0;
        const selected = isSelectedChartPointOrSeries(selectedChartElement, chart.id, 0, index);

        return (
          <g key={`waterfall-bar-${index}`}>
            {connectorStart ? (
              <line
                stroke={lightenColor(axisColor, 0.35)}
                strokeDasharray="3 3"
                strokeWidth={1}
                x1={(xScale(connectorStart.label) ?? plot.left) + xScale.bandwidth()}
                x2={bandLeft}
                y1={connectorY}
                y2={connectorY}
              />
            ) : null}
              <rect
                {...chartElementDataProps(0, index)}
                fill={fill}
                rx={2}
                ry={2}
                stroke={selected ? "#64748b" : darkenColor(fill, 0.22)}
                strokeWidth={selected ? 2 : 1}
              x={bandLeft}
              y={top}
              width={bandWidth}
              height={height}
            />
            {selected
              ? renderSelectionRectHandles(`waterfall-selection-${index}`, bandLeft, top, bandWidth, height)
              : null}
            <text
              fill={labelColor}
              fontSize={10}
              textAnchor="middle"
              x={bandLeft + bandWidth * 0.5}
              y={plot.top + plot.height + 14}
            >
              {bar.label}
            </text>
            <text
              fill={labelColor}
              fontSize={10}
              textAnchor="middle"
              x={bandLeft + bandWidth * 0.5}
              y={top - 6}
            >
              {formatTickValue(bar.end)}
            </text>
          </g>
        );
      })}
    </g>
  );
}

function renderFunnelChart(chart: XlsxChart, palette: ChartRendererPalette, layout: ChartLayout, selectedChartElement?: XlsxChartElementSelection | null) {
  const stages = buildChartStages(chart)
    .map((stage) => ({ ...stage, value: Math.max(0, stage.value) }))
    .filter((stage) => stage.value > 0);
  if (stages.length === 0) {
    return null;
  }

  const plot = layout.plot;
  const maxValue = Math.max(...stages.map((stage) => stage.value));
  const sectionHeight = plot.height / stages.length;
  const centerX = plot.left + plot.width * 0.5;
  const labelColor = resolveChartTextColor(chart);

  return (
    <g>
      {stages.map((stage, index) => {
        const stageWidth = (stage.value / maxValue) * plot.width;
        const topY = plot.top + index * sectionHeight;
        const stageHeight = Math.max(6, sectionHeight - 2);
        const left = centerX - stageWidth * 0.5;
        const fill = stage.isSubtotal ? darkenColor(stage.color, 0.14) : stage.color;
        const labelFitsInside = stageWidth > 90;
        const selected = isSelectedChartPointOrSeries(selectedChartElement, chart.id, 0, index);

        return (
          <g key={`funnel-stage-${index}`}>
              <rect
                {...chartElementDataProps(0, index)}
                fill={fill}
                height={stageHeight}
              rx={0}
              ry={0}
                stroke={selected ? "#64748b" : darkenColor(fill, 0.2)}
                strokeWidth={selected ? 2 : 1}
              width={stageWidth}
              x={left}
              y={topY}
            />
            {selected
              ? renderSelectionRectHandles(`funnel-selection-${index}`, left, topY, stageWidth, stageHeight)
              : null}
            <text
              fill={labelColor}
              fontSize={10}
              textAnchor={labelFitsInside ? "middle" : "start"}
              x={labelFitsInside ? centerX : centerX + stageWidth * 0.5 + 8}
              y={topY + sectionHeight * 0.5}
            >
              {`${stage.label} ${formatTickValue(stage.value)}`}
            </text>
          </g>
        );
      })}
    </g>
  );
}

function renderSunburstChart(chart: XlsxChart, palette: ChartRendererPalette, layout: ChartLayout, selectedChartElement?: XlsxChartElementSelection | null) {
  const hierarchyData = buildHierarchyData(chart);
  if (!hierarchyData) {
    return null;
  }

  const root = d3Hierarchy(hierarchyData)
    .sum((node) => node.children && node.children.length > 0 ? 0 : Math.max(0.0001, node.value ?? 0))
    .sort((left, right) => (right.value ?? 0) - (left.value ?? 0));
  const partitioned = d3Partition<ChartHierarchyDatum>()
    .size([Math.PI * 2, root.height + 1])(root);
  const plot = layout.plot;
  const radius = Math.max(24, Math.min(plot.width, plot.height) * 0.5);
  const holeRadius = radius * 0.16;
  const centerX = plot.left + plot.width * 0.5;
  const centerY = plot.top + plot.height * 0.5;
  const ringSpan = Math.max(1, partitioned.height || 1);
  const arcBuilder = d3Arc<HierarchyRectangularNode<ChartHierarchyDatum>>()
    .startAngle((node) => node.x0)
    .endAngle((node) => node.x1)
    .padAngle(0.005)
    .padRadius(radius)
    .innerRadius((node) => holeRadius + ((node.y0 - 1) / ringSpan) * (radius - holeRadius))
    .outerRadius((node) => holeRadius + ((node.y1 - 1) / ringSpan) * (radius - holeRadius) - 1);

  return (
    <g transform={`translate(${centerX}, ${centerY})`}>
      {partitioned.descendants().filter((node) => node.depth > 0).map((node, index) => {
        const path = arcBuilder(node);
        if (!path) {
          return null;
        }
        const labelAngle = (node.x0 + node.x1) * 0.5 - Math.PI * 0.5;
        const labelRadius = holeRadius + (((node.y0 + node.y1) * 0.5 - 1) / ringSpan) * (radius - holeRadius);
        const labelX = Math.cos(labelAngle) * labelRadius;
        const labelY = Math.sin(labelAngle) * labelRadius;
        const arcSpan = node.x1 - node.x0;
        const canShowLabel = arcSpan * labelRadius > 26;
        const fill = resolveHierarchyNodeColor(chart, node);
        const selected = isSelectedChartPointOrSeries(selectedChartElement, chart.id, 0, index);

        return (
          <g key={`sunburst-node-${index}`}>
              <path
                {...chartElementDataProps(0, index)}
                d={path}
                fill={fill}
                stroke={selected ? "#64748b" : palette.surface}
                strokeWidth={selected ? 2 : 1}
            />
            {selected
              ? renderSelectionPointHandles(`sunburst-selection-${index}`, [{ x: labelX, y: labelY }])
              : null}
            {canShowLabel ? (
              <text
                fill={darkenColor(fill, 0.65)}
                fontSize={9}
                textAnchor="middle"
                transform={`translate(${labelX}, ${labelY}) rotate(${(labelAngle * 180) / Math.PI})`}
              >
                {node.data.name}
              </text>
            ) : null}
          </g>
        );
      })}
      <circle fill={chart.chartAreaFillColor ?? palette.surface} r={holeRadius - 2} />
    </g>
  );
}

function renderTreemapChart(chart: XlsxChart, palette: ChartRendererPalette, layout: ChartLayout, selectedChartElement?: XlsxChartElementSelection | null) {
  const hierarchyData = buildHierarchyData(chart);
  if (!hierarchyData) {
    return null;
  }

  const root = d3Hierarchy(hierarchyData)
    .sum((node) => node.children && node.children.length > 0 ? 0 : Math.max(0.0001, node.value ?? 0))
    .sort((left, right) => (right.value ?? 0) - (left.value ?? 0));
  const treemapRoot = d3Treemap<ChartHierarchyDatum>()
    .size([layout.plot.width, layout.plot.height])
    .paddingInner(2)
    .paddingOuter(1)
    .round(true)
    .tile(excelTreemapTile)(root);

  return (
    <g transform={`translate(${layout.plot.left}, ${layout.plot.top})`}>
      {treemapRoot.leaves().map((leaf, index) => {
        const fill = resolveTreemapNodeColor(chart, leaf);
        const width = Math.max(0, leaf.x1 - leaf.x0);
        const height = Math.max(0, leaf.y1 - leaf.y0);
        const canShowLabel = width > 48 && height > 22;
        const selected = isSelectedChartPointOrSeries(selectedChartElement, chart.id, 0, index);

        return (
          <g key={`treemap-leaf-${index}`}>
              <rect
                {...chartElementDataProps(0, index)}
                fill={fill}
                rx={3}
                ry={3}
                stroke={selected ? "#64748b" : palette.surface}
                strokeWidth={selected ? 2 : 1}
              x={leaf.x0}
              y={leaf.y0}
              width={width}
              height={height}
            />
            {selected
              ? renderSelectionRectHandles(`treemap-selection-${index}`, leaf.x0, leaf.y0, width, height)
              : null}
            {canShowLabel ? (
              <>
                <text
                  fill={darkenColor(fill, 0.68)}
                  fontSize={10}
                  fontWeight={600}
                  x={leaf.x0 + 8}
                  y={leaf.y0 + 14}
                >
                  {leaf.data.name}
                </text>
                <text
                  fill={darkenColor(fill, 0.54)}
                  fontSize={9}
                  x={leaf.x0 + 8}
                  y={leaf.y0 + 28}
                >
                  {formatTickValue(leaf.value ?? 0)}
                </text>
              </>
            ) : null}
          </g>
        );
      })}
    </g>
  );
}

function renderBoxWhiskerChart(chart: XlsxChart, palette: ChartRendererPalette, layout: ChartLayout, selectedChartElement?: XlsxChartElementSelection | null) {
  const visibleSeries = chart.series.filter((series) => series.hidden !== true);
  if (visibleSeries.length === 0) {
    return null;
  }

  const seriesStats = visibleSeries.map((series, index) => ({
    color: series.color ?? chartSeriesColor(chart, typeof series.formatIdx === "number" ? series.formatIdx : index),
    label: normalizeCategoryLabel(series.name) || `Series ${index + 1}`,
    lineColor: series.lineColor ?? series.color ?? chartSeriesColor(chart, typeof series.formatIdx === "number" ? series.formatIdx : index),
    series,
    stats: computeBoxWhiskerStats(series),
    visibility: resolveBoxWhiskerVisibility(series)
  })).filter((entry): entry is {
    color: string;
    label: string;
    lineColor: string;
    series: XlsxChartSeries;
    stats: BoxWhiskerStats;
    visibility: ReturnType<typeof resolveBoxWhiskerVisibility>;
  } => entry.stats != null);

  if (seriesStats.length === 0) {
    return null;
  }

  const allValues = seriesStats.flatMap((entry) => [
    entry.stats.min,
    entry.stats.lowerWhisker,
    entry.stats.q1,
    entry.stats.median,
    entry.stats.q3,
    entry.stats.upperWhisker,
    entry.stats.max
  ]);
  const minValue = Math.min(...allValues);
  const maxValue = Math.max(...allValues);
  const valueDomain = resolveAxisDomainWithChartOverrides(chart.valueAxis, minValue, maxValue, false);
  const yScale = scaleLinear()
    .domain([valueDomain.min, valueDomain.max])
    .range([layout.plot.top + layout.plot.height, layout.plot.top]);
  const xScale = scalePoint<number>()
    .domain(seriesStats.map((_, index) => index))
    .range([layout.plot.left + 24, layout.plot.left + layout.plot.width - 24]);
  const gap = seriesStats.length > 1
    ? Math.abs((xScale(1) ?? 0) - (xScale(0) ?? 0))
    : layout.plot.width * 0.5;
  const boxWidth = clamp(gap * 0.34, 20, 54);
  const capWidth = Math.max(8, boxWidth * 0.52);
  const axisColor = chart.axisLineColor ?? chart.chartAreaBorderColor ?? palette.border;
  const labelColor = resolveChartAxisTextColor(chart);
  const meanLinePoints = seriesStats
    .filter((entry) => entry.visibility.meanLine)
    .map((entry, index) => `${xScale(index) ?? layout.plot.left},${yScale(entry.stats.mean)}`);

  return (
    <g>
      {valueDomain.ticks.map((tick, index) => {
        const y = yScale(tick);
        return (
          <g key={`box-whisker-grid-${index}`}>
            <line
              stroke={lightenColor(axisColor, 0.22)}
              strokeWidth={0.8}
              x1={layout.plot.left}
              x2={layout.plot.left + layout.plot.width}
              y1={y}
              y2={y}
            />
            <text
              fill={labelColor}
              fontSize={10}
              textAnchor="end"
              x={layout.plot.left - 6}
              y={y + 3}
            >
              {formatTickValue(tick)}
            </text>
          </g>
        );
      })}
      <line
        stroke={axisColor}
        strokeWidth={1}
        x1={layout.plot.left}
        x2={layout.plot.left}
        y1={layout.plot.top}
        y2={layout.plot.top + layout.plot.height}
      />
      <line
        stroke={axisColor}
        strokeWidth={1}
        x1={layout.plot.left}
        x2={layout.plot.left + layout.plot.width}
        y1={layout.plot.top + layout.plot.height}
        y2={layout.plot.top + layout.plot.height}
      />
      {meanLinePoints.length >= 2 ? (
        <polyline
          fill="none"
          points={meanLinePoints.join(" ")}
          stroke={darkenColor(axisColor, 0.2)}
          strokeDasharray="4 3"
          strokeWidth={1}
        />
      ) : null}
      {seriesStats.map((entry, index) => {
        const x = xScale(index) ?? layout.plot.left;
        const boxTop = yScale(entry.stats.q3);
        const boxBottom = yScale(entry.stats.q1);
        const medianY = yScale(entry.stats.median);
        const lowerWhiskerY = yScale(entry.stats.lowerWhisker);
        const upperWhiskerY = yScale(entry.stats.upperWhisker);
        const meanY = yScale(entry.stats.mean);
        const visiblePoints = entry.visibility.nonoutliers
          ? entry.stats.visiblePoints
          : [];
        const outliers = entry.visibility.outliers
          ? entry.stats.outliers
          : [];
        const selected = isSelectedChartPointOrSeries(selectedChartElement, chart.id, index, index);

        return (
            <g {...chartElementDataProps(index)} key={`box-whisker-series-${index}`}>
            <line
              stroke={entry.lineColor}
              strokeWidth={1.5}
              x1={x}
              x2={x}
              y1={upperWhiskerY}
              y2={boxTop}
            />
            <line
              stroke={entry.lineColor}
              strokeWidth={1.5}
              x1={x}
              x2={x}
              y1={boxBottom}
              y2={lowerWhiskerY}
            />
            <line
              stroke={entry.lineColor}
              strokeWidth={1.5}
              x1={x - capWidth * 0.5}
              x2={x + capWidth * 0.5}
              y1={upperWhiskerY}
              y2={upperWhiskerY}
            />
            <line
              stroke={entry.lineColor}
              strokeWidth={1.5}
              x1={x - capWidth * 0.5}
              x2={x + capWidth * 0.5}
              y1={lowerWhiskerY}
              y2={lowerWhiskerY}
            />
              <rect
                {...chartElementDataProps(index, index)}
                fill={lightenColor(entry.color, 0.35)}
              fillOpacity={0.72}
              height={Math.max(1, boxBottom - boxTop)}
                stroke={selected ? "#64748b" : entry.lineColor}
                strokeWidth={selected ? 2 : 1.5}
              width={boxWidth}
              x={x - boxWidth * 0.5}
              y={boxTop}
            />
            {selected
              ? renderSelectionRectHandles(
                  `box-whisker-selection-${index}`,
                  x - boxWidth * 0.5,
                  boxTop,
                  boxWidth,
                  Math.max(1, boxBottom - boxTop)
                )
              : null}
            <line
              stroke={darkenColor(entry.lineColor, 0.15)}
              strokeWidth={2}
              x1={x - boxWidth * 0.5}
              x2={x + boxWidth * 0.5}
              y1={medianY}
              y2={medianY}
            />
            {entry.visibility.meanMarker ? (
              <>
                <line
                  stroke={darkenColor(entry.lineColor, 0.18)}
                  strokeWidth={1.2}
                  x1={x - 4}
                  x2={x + 4}
                  y1={meanY - 4}
                  y2={meanY + 4}
                />
                <line
                  stroke={darkenColor(entry.lineColor, 0.18)}
                  strokeWidth={1.2}
                  x1={x - 4}
                  x2={x + 4}
                  y1={meanY + 4}
                  y2={meanY - 4}
                />
              </>
            ) : null}
            {visiblePoints.map((value, pointIndex) => {
              const y = yScale(value);
              const jitter = ((pointIndex % 7) - 3) * (boxWidth / 16);
              return (
                  <circle
                    {...chartElementDataProps(index, pointIndex)}
                    key={`box-whisker-visible-${index}-${pointIndex}`}
                  cx={x + jitter}
                  cy={y}
                  fill={entry.color}
                  fillOpacity={0.45}
                  r={2}
                  stroke="none"
                />
              );
            })}
            {outliers.map((value, pointIndex) => {
              const y = yScale(value);
              return (
                  <circle
                    {...chartElementDataProps(index, pointIndex)}
                    key={`box-whisker-outlier-${index}-${pointIndex}`}
                  cx={x}
                  cy={y}
                  fill="#ffffff"
                  r={3}
                  stroke={entry.lineColor}
                  strokeWidth={1.2}
                />
              );
            })}
            <text
              fill={labelColor}
              fontSize={10}
              textAnchor="middle"
              x={x}
              y={layout.plot.top + layout.plot.height + 14}
            >
              {entry.label}
            </text>
          </g>
        );
      })}
    </g>
  );
}

function renderRegionMapChart(
  chart: XlsxChart,
  palette: ChartRendererPalette,
  layout: ChartLayout,
  selectedChartElement: XlsxChartElementSelection | null | undefined,
  { geoIdentity, geoMercator, geoNaturalEarth1, geoPath, resolveRegionMapFeatureSet, resolveRegionMapFeature, getRegionMapBaseFeatures }: RegionMapResources
) {
  const primarySeriesIndex = Math.max(0, chart.series.findIndex((series) => series.hidden !== true));
  const primarySeries = chart.series[primarySeriesIndex] ?? null;
  if (!primarySeries) {
    return null;
  }

  const labels = getCategoryLabels(chart);
  const layoutProperties = resolveRegionMapLayoutProperties(primarySeries);
  const geography = layoutProperties?.geography && typeof layoutProperties.geography === "object"
    ? layoutProperties.geography as Record<string, unknown>
    : null;
  const featureSet = resolveRegionMapFeatureSet(labels, geography);
  const viewedRegionType = typeof geography?.viewedRegionType === "string"
    ? geography.viewedRegionType
    : null;
  const projectionType = typeof geography?.projectionType === "string"
    ? geography.projectionType
    : null;
  const projection = featureSet === "us-state"
    ? geoIdentity()
    : projectionType === "miller"
      ? geoMercator()
      : geoNaturalEarth1();
  const availableFeatures = getRegionMapBaseFeatures(featureSet);
  const categoricalValues = resolveRegionMapColorStrings(primarySeries);
  const categoricalColorByLabel = categoricalValues
    ? new Map(
        Array.from(new Set(categoricalValues.filter(Boolean))).map((value, index) => [
          value,
          chartPointColor(chart, index, primarySeriesIndex)
        ])
      )
    : null;
  const entries = labels.map((label, index) => {
    const feature = resolveRegionMapFeature(label, featureSet);
    const value = safeNumber(primarySeries.values[index]);
    const colorLabel = categoricalValues?.[index] || null;
    return {
      colorLabel,
      feature,
      key: normalizeCategoryLabel(label),
      label: normalizeCategoryLabel(label),
      value
    };
  }).filter((entry) => entry.feature != null && (entry.value != null || entry.colorLabel != null)) as Array<{
    colorLabel: string | null;
    feature: RegionMapFeature;
    key: string;
    label: string;
    value: number | null;
  }>;
  const fitToMatchedData = entries.length > 0 && (viewedRegionType === "dataOnly" || entries.length === 1);
  const fitFeatures = fitToMatchedData
    ? entries.map((entry) => entry.feature)
    : availableFeatures;
  projection.fitExtent(
    [
      [layout.plot.left + 8, layout.plot.top + 8],
      [layout.plot.left + layout.plot.width - 8, layout.plot.top + layout.plot.height - 8]
    ],
    {
      type: "FeatureCollection",
      features: fitFeatures
    } satisfies FeatureCollection<Geometry, { name?: string }>
  );
  const path = geoPath(projection);
  const baseFeatures = fitToMatchedData
    ? fitFeatures
    : availableFeatures;
  const noDataFill = resolveRegionMapNoDataColor(chart, primarySeriesIndex);
  const outlineColor = chart.chartAreaBorderColor && chart.chartAreaBorderColor !== "transparent"
    ? chart.chartAreaBorderColor
    : darkenColor(noDataFill, 0.22);
  const numericEntryValues = primarySeries.values
    .map(safeNumber)
    .filter((value): value is number => typeof value === "number" && Number.isFinite(value));
  const minValue = numericEntryValues.length > 0 ? Math.min(...numericEntryValues) : 0;
  const maxValue = numericEntryValues.length > 0 ? Math.max(...numericEntryValues) : 1;
  const showRegionLabels = layoutProperties?.regionLabelLayout !== "none";

  return (
    <g>
      {baseFeatures.map((feature, index) => {
        const d = path(feature);
        if (!d) {
          return null;
        }
        return (
          <path
            d={d}
            fill={noDataFill}
            key={`region-map-base-${feature.id ?? "feature"}-${index}`}
            stroke={outlineColor}
            strokeLinejoin="round"
            strokeWidth={0.6}
          />
        );
      })}
      {entries.map((entry, index) => {
        const d = path(entry.feature);
        if (!d) {
          return null;
        }
        const fill = entry.value != null
          ? (() => {
              const ratio = maxValue <= minValue ? 1 : (entry.value - minValue) / Math.max(1e-6, maxValue - minValue);
              return resolveRegionMapValueColor(chart, primarySeriesIndex, ratio);
            })()
          : categoricalColorByLabel?.get(entry.colorLabel ?? "") ?? resolveRegionMapDataColor(chart, primarySeriesIndex);
        const selected = isSelectedChartPointOrSeries(selectedChartElement, chart.id, primarySeriesIndex, index);
        const centroid = path.centroid(entry.feature);
        const canShowSelectionDot = Number.isFinite(centroid[0]) && Number.isFinite(centroid[1]);
        return (
          <React.Fragment key={`region-map-value-${entry.key || index}`}>
            <path
              {...chartElementDataProps(primarySeriesIndex, index)}
              d={d}
              fill={fill}
              stroke={selected ? "#64748b" : darkenColor(fill, 0.18)}
              strokeLinejoin="round"
              strokeWidth={selected ? 1.8 : 0.85}
            />
            {selected && canShowSelectionDot
              ? renderSelectionPointHandles(`region-map-selection-${index}`, [{ x: centroid[0], y: centroid[1] }])
              : null}
          </React.Fragment>
        );
      })}
      {showRegionLabels
        ? entries.map((entry, index) => {
            const bounds = path.bounds(entry.feature);
            const [[x0, y0], [x1, y1]] = bounds;
            const width = x1 - x0;
            const height = y1 - y0;
            if (!Number.isFinite(width) || !Number.isFinite(height) || width < 26 || height < 12) {
              return null;
            }
            const centroid = path.centroid(entry.feature);
            if (!Number.isFinite(centroid[0]) || !Number.isFinite(centroid[1])) {
              return null;
            }
            return (
              <text
                fill={resolveChartTextColor(chart)}
                fontSize={9}
                fontWeight={600}
                key={`region-map-label-${entry.key || index}`}
                textAnchor="middle"
                x={centroid[0]}
                y={centroid[1]}
              >
                {truncateSvgText(entry.label, Math.max(28, width - 4), 9)}
              </text>
            );
          })
        : null}
    </g>
  );
}

function renderUnsupported(chart: XlsxChart, palette: ChartRendererPalette, layout: ChartLayout, chartType: string) {
  return (
    <g>
      <rect
        fill={lightenColor(chart.chartAreaFillColor ?? palette.surface, 0.02)}
        height={layout.plot.height}
        stroke={lightenColor(chart.axisLineColor ?? palette.border, 0.2)}
        strokeDasharray="3 3"
        width={layout.plot.width}
        x={layout.plot.left}
        y={layout.plot.top}
      />
      <text
        fill={resolveChartMutedTextColor(chart)}
        fontSize={11}
        textAnchor="middle"
        x={layout.plot.left + layout.plot.width / 2}
        y={layout.plot.top + layout.plot.height / 2}
      >
        {`Unsupported chart type: ${chartType}`}
      </text>
    </g>
  );
}

function renderChartPlot(
  chart: XlsxChart,
  palette: ChartRendererPalette,
  layout: ChartLayout,
  chartType: string,
  selectedChartElement?: XlsxChartElementSelection | null
) {
  if (isComboChart(chart)) {
    return renderComboChart(chart, palette, layout, selectedChartElement) ?? renderUnsupported(chart, palette, layout, "Combo");
  }
  if (
    chartType === "ColumnClustered"
    || chartType === "ColumnStacked"
    || chartType === "ColumnPercentStacked"
    || chartType === "BarClustered"
    || chartType === "BarStacked"
    || chartType === "BarPercentStacked"
  ) {
    return renderBarChart(chart, palette, layout, chartType, selectedChartElement);
  }
  if (
    chartType === "Line"
    || chartType === "LineStacked"
    || chartType === "LinePercentStacked"
    || chartType === "Area"
    || chartType === "AreaStacked"
    || chartType === "AreaPercentStacked"
  ) {
    return renderLineOrAreaChart(chart, palette, layout, chartType, selectedChartElement);
  }
  if (chartType === "Scatter") {
    return renderScatterChart(chart, palette, layout, false, selectedChartElement);
  }
  if (chartType === "ScatterLines") {
    return renderScatterChart(chart, palette, layout, false, selectedChartElement);
  }
  if (chartType === "ScatterSmooth") {
    return renderScatterChart(chart, palette, layout, true, selectedChartElement);
  }
  if (chartType === "Bubble") {
    return renderBubbleChart(chart, palette, layout, selectedChartElement);
  }
  if (chartType === "Radar") {
    return renderRadarChart(chart, palette, layout, selectedChartElement);
  }
  if (chartType === "Pie" || chartType === "Pie3D" || chartType === "PieExploded" || chartType === "Doughnut") {
    return renderPieChart(chart, palette, layout, chartType, selectedChartElement);
  }
  if (chartType === "BarOfPie") {
    return renderBarOfPieChart(chart, palette, layout, selectedChartElement);
  }
  if (chartType === "Surface") {
    return renderSurfaceChart(chart, palette, layout, selectedChartElement);
  }
  if (chartType === "Stock") {
    return renderStockChart(chart, palette, layout, selectedChartElement);
  }
  if (chartType === "Waterfall") {
    return renderWaterfallChart(chart, palette, layout, selectedChartElement);
  }
  if (chartType === "Funnel") {
    return renderFunnelChart(chart, palette, layout, selectedChartElement);
  }
  if (chartType === "BoxWhisker") {
    return renderBoxWhiskerChart(chart, palette, layout, selectedChartElement);
  }
  if (chartType === "Sunburst") {
    return renderSunburstChart(chart, palette, layout, selectedChartElement);
  }
  if (chartType === "Treemap") {
    return renderTreemapChart(chart, palette, layout, selectedChartElement);
  }
  if (chartType === "RegionMap") {
    return <RegionMapPlot chart={chart} layout={layout} palette={palette} selectedChartElement={selectedChartElement} />;
  }
  return renderUnsupported(chart, palette, layout, chartType);
}

export const MemoChartSvg = React.memo(function MemoChartSvg({
  chart,
  onChartElementDoubleClick,
  onChartElementPointerDown,
  palette,
  rect,
  selectedChartElement
}: ChartSvgProps) {
  const renderChartType = normalizeRenderableChartType(chart);
  const legendItems = getLegendItems(chart, renderChartType, palette);
  const layout = buildLayout(chart, rect, legendItems);
  const chartRaw = chart.raw && typeof chart.raw === "object" ? chart.raw as Record<string, unknown> : null;
  const explicitNoFill = chartRaw?.chartAreaNoFill === true || chartRaw?.plotAreaNoFill === true;
  const background = chart.chartAreaFillColor ?? (explicitNoFill ? "transparent" : "#ffffff");
  const borderColor = chart.chartAreaBorderColor ?? "transparent";
  const normalizedBackground = background.trim().toLowerCase();
  const normalizedBorderColor = borderColor.trim().toLowerCase();
  const hideBackgroundRect = normalizedBackground === "transparent" && normalizedBorderColor === "transparent";
  const fontFamily = buildChartFontFamily(chart.fontFamily);
  const handlePointerDown = React.useCallback((event: React.PointerEvent<SVGSVGElement>) => {
    const selection = resolveChartSelectionFromTarget(chart, event.target, selectedChartElement);
    if (!selection) {
      return;
    }

    if (resolveChartElementTarget(event.target)) {
      event.stopPropagation();
    }
    onChartElementPointerDown?.(selection, event);
  }, [chart, onChartElementPointerDown, selectedChartElement]);
  const handleDoubleClick = React.useCallback((event: React.MouseEvent<SVGSVGElement>) => {
    const selection = resolveChartSelectionFromTarget(chart, event.target, selectedChartElement);
    if (!selection) {
      return;
    }

    if (resolveChartElementTarget(event.target)) {
      event.stopPropagation();
    }
    onChartElementDoubleClick?.(selection, event);
  }, [chart, onChartElementDoubleClick, selectedChartElement]);

  const renderSvgChart = (plot: React.ReactNode) => (
    <svg
      aria-label={chart.title ?? chart.name ?? "Chart"}
      onDoubleClick={handleDoubleClick}
      onPointerDown={handlePointerDown}
      role="img"
      style={{ display: "block", fontFamily, height: "100%", pointerEvents: "auto", width: "100%" }}
      viewBox={`0 0 ${layout.width} ${layout.height}`}
    >
      {hideBackgroundRect
        ? null
        : <rect fill={background} height={layout.height} stroke={borderColor} strokeWidth={1} width={layout.width} x={0} y={0} />}
      {renderTitle(chart, layout, palette)}
      {renderLegend(chart, layout, palette)}
      {plot}
    </svg>
  );
  if (renderChartType === "Surface") {
    const surfacePlot = renderSurfaceChart(chart, palette, layout, selectedChartElement);
    return (
      <SurfaceChartComposite
        loadingFallback={renderSvgChart(surfacePlot)}
        background={background}
        borderColor={borderColor}
        chart={chart}
        fallback={surfacePlot}
        fontFamily={fontFamily}
        layout={layout}
        onDoubleClick={handleDoubleClick}
        onPointerDown={handlePointerDown}
        overlay={
          <>
            {renderSurfaceHitOverlay(chart, layout, selectedChartElement)}
            {renderSurfaceAxes(chart, layout)}
            {renderTitle(chart, layout, palette)}
            {renderLegend(chart, layout, palette)}
          </>
        }
        palette={palette}
      />
    );
  }

  return renderSvgChart(renderChartPlot(chart, palette, layout, renderChartType, selectedChartElement));
}, (prev, next) => (
  prev.chart === next.chart
  && prev.onChartElementDoubleClick === next.onChartElementDoubleClick
  && prev.onChartElementPointerDown === next.onChartElementPointerDown
  && prev.palette === next.palette
  && prev.rect.height === next.rect.height
  && prev.rect.width === next.rect.width
  && prev.rect.left === next.rect.left
  && prev.rect.top === next.rect.top
  && prev.selectedChartElement === next.selectedChartElement
));
