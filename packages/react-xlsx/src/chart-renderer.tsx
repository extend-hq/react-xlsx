import { hierarchy as d3Hierarchy, partition as d3Partition, treemap as d3Treemap, treemapSquarify } from "d3-hierarchy";
import * as React from "react";
import { scaleBand, scaleLinear, scalePoint } from "d3-scale";
import {
  arc as d3Arc,
  area as d3Area,
  curveCatmullRom,
  curveLinear,
  curveLinearClosed,
  line as d3Line,
  pie as d3Pie,
  symbol as d3Symbol,
  symbolCircle,
  symbolCross,
  symbolDiamond,
  symbolSquare,
  symbolStar,
  symbolTriangle
} from "d3-shape";
import type { CurveFactory } from "d3-shape";
import { MemoSurfaceChartComposite } from "./surface-regl";
import type { HierarchyNode, HierarchyRectangularNode } from "d3-hierarchy";
import type { XlsxChart, XlsxChartAxis, XlsxChartSeries, XlsxChartTypeGroup, XlsxImageRect } from "./types";

type ChartRendererPalette = {
  border: string;
  mutedText: string;
  surface: string;
  text: string;
};

type ChartSvgProps = {
  chart: XlsxChart;
  palette: ChartRendererPalette;
  rect: XlsxImageRect;
};

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
  legendItems: LegendItem[];
  legendPosition: string | undefined;
  plot: PlotRect;
  titleHeight: number;
  width: number;
};

type BarRect = {
  categoryIndex: number;
  bottomScale?: number;
  color: string;
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

function findAxisForGroup(chart: XlsxChart, axisIds: number[], positionMatcher: (position: string | undefined) => boolean) {
  return chart.axes.find((axis) => axis.id != null && axisIds.includes(axis.id) && positionMatcher(axis.position))
    ?? null;
}

function buildComboGroups(chart: XlsxChart): ComboRenderableGroup[] {
  return (chart.typeGroups ?? []).map((group) => {
    const axisIds = group.axisIds ?? [];
    const categoryAxis = findAxisForGroup(chart, axisIds, (position) => position === "b" || position === "t")
      ?? chart.categoryAxis
      ?? null;
    const valueAxis = findAxisForGroup(chart, axisIds, (position) => position === "l" || position === "r")
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
  const rowCount = Math.max(0, ...chart.series.map((series) => series.values.length));

  for (let rowIndex = 0; rowIndex < rowCount; rowIndex += 1) {
    const path = chart.series
      .map((series) => normalizeCategoryLabel(series.categories[rowIndex] ?? series.values[rowIndex]))
      .filter((value) => value.length > 0);
    if (path.length === 0) {
      continue;
    }

    let current = root;
    path.forEach((part, levelIndex) => {
      current.children = current.children ?? [];
      let next = current.children.find((child) => child.name === part);
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

    current.value = (current.value ?? 0) + Math.max(0.0001, safeNumber(chart.series[chart.series.length - 1]?.values[rowIndex]) ?? 1);
  }

  return rootChildren.length > 0 ? root : null;
}

function resolveHierarchyNodeColor(chart: XlsxChart, node: HierarchyNode<ChartHierarchyDatum>) {
  const lineage = node.ancestors().reverse();
  const topLevel = lineage[1];
  const baseIndex = topLevel?.data.colorIndex ?? 0;
  const baseColor = chartSeriesColor(chart, baseIndex);
  const depth = Math.max(0, node.depth - 1);
  return depth === 0 ? baseColor : lightenColor(baseColor, clamp(depth * 0.16, 0, 0.5));
}

function resolveSurfaceBaseColor(chart: XlsxChart, palette: ChartRendererPalette) {
  return chart.chartColorPalette?.[0]
    ?? chart.series[0]?.color
    ?? chart.series[0]?.lineColor
    ?? chart.axisLineColor
    ?? chart.textColor
    ?? palette.text;
}

function normalizeBuiltinSurfaceStyleId(styleId: number | undefined) {
  if (typeof styleId !== "number" || !Number.isFinite(styleId)) {
    return null;
  }
  return styleId >= 100 ? styleId - 100 : styleId;
}

function hasExplicitSurfaceBaseColor(chart: XlsxChart) {
  const primarySeriesColor = normalizeRendererHexColor(chart.series[0]?.color ?? chart.series[0]?.lineColor);
  if (!primarySeriesColor) {
    return null;
  }
  const paletteColor = normalizeRendererHexColor(chart.chartColorPalette?.[0]);
  return paletteColor && paletteColor === primarySeriesColor ? null : primarySeriesColor;
}

function buildMonochromeSurfacePalette(baseColor: string, count: number) {
  if (count <= 3) {
    return [
      lightenColor(baseColor, 0.58),
      lightenColor(baseColor, 0.18),
      darkenColor(baseColor, 0.12)
    ];
  }
  return [
    lightenColor(baseColor, 0.62),
    lightenColor(baseColor, 0.34),
    baseColor,
    darkenColor(baseColor, 0.12),
    darkenColor(baseColor, 0.24)
  ];
}

function getBuiltinSurfacePalette(chart: XlsxChart) {
  const normalized = normalizeBuiltinSurfaceStyleId(chart.chartStyleId);
  const explicitBaseColor = hasExplicitSurfaceBaseColor(chart);
  if (normalized === 26) {
    return buildMonochromeSurfacePalette(explicitBaseColor ?? "#ff006e", 3);
  }
  if (normalized === 34 && explicitBaseColor) {
    return buildMonochromeSurfacePalette(explicitBaseColor, 3);
  }
  if (normalized === 34 || (chart.wireframe === true && normalized == null)) {
    return ["#5b9bd5", "#ed7d31", "#a5a5a5"];
  }
  if (normalized === 35 || normalized === 36 || (chart.wireframe !== true && normalized == null)) {
    return ["#2f5597", "#4472c4", "#5b9bd5", "#8faadc", "#b4c7e7"];
  }
  return null;
}

function getSurfaceBandCount(chart: XlsxChart) {
  const raw = chart.raw && typeof chart.raw === "object" ? chart.raw as Record<string, unknown> : null;
  const explicitBandCount = typeof raw?.bandFormatCount === "number" && Number.isFinite(raw.bandFormatCount)
    ? raw.bandFormatCount
    : null;
  if (explicitBandCount != null && explicitBandCount > 0) {
    return explicitBandCount;
  }
  const builtinPalette = getBuiltinSurfacePalette(chart);
  if (builtinPalette && builtinPalette.length > 0) {
    return builtinPalette.length;
  }
  if (chart.chartColorPalette && chart.chartColorPalette.length > 1) {
    return chart.chartColorPalette.length;
  }
  return chart.wireframe ? 3 : 5;
}

function getSurfaceColorStops(chart: XlsxChart, palette: ChartRendererPalette) {
  const explicitStops = (chart.chartColorPalette ?? []).filter((value): value is string => typeof value === "string" && value.length > 0);
  if (explicitStops.length >= 2) {
    return explicitStops;
  }
  const builtinPalette = getBuiltinSurfacePalette(chart);
  if (builtinPalette && builtinPalette.length >= 2) {
    return builtinPalette;
  }
  const baseColor = resolveSurfaceBaseColor(chart, palette);
  return [
    darkenColor(baseColor, 0.42),
    darkenColor(baseColor, 0.24),
    baseColor,
    lightenColor(baseColor, 0.18),
    lightenColor(baseColor, 0.34),
    lightenColor(baseColor, 0.5)
  ];
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
  const bandCount = Math.max(1, getSurfaceBandCount(chart));
  const spanBase = Math.max(1e-6, rawMax - Math.min(0, rawMin));
  const roughStep = spanBase / bandCount;
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
  return {
    maxValue,
    minValue,
    safeMax,
    ticks: ticks.length >= 2 ? ticks : [minValue, safeMax]
  };
}

function resolveSurfaceColor(chart: XlsxChart, palette: ChartRendererPalette, ratio: number) {
  const stops = getSurfaceColorStops(chart, palette);
  if (stops.length === 0) {
    return resolveSurfaceBaseColor(chart, palette);
  }
  if (stops.length === 1) {
    return stops[0];
  }
  const clamped = clamp(ratio, 0, 1) * (stops.length - 1);
  const lowerIndex = Math.floor(clamped);
  const upperIndex = Math.min(stops.length - 1, lowerIndex + 1);
  const mixRatio = clamped - lowerIndex;
  return mixRgbColor(stops[lowerIndex] ?? stops[0], stops[upperIndex] ?? stops[stops.length - 1], mixRatio);
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
    const ratio = (midpoint - domain.minValue) / Math.max(1e-6, domain.safeMax - domain.minValue);
    items.push({
      color: resolveSurfaceColor(chart, palette, ratio),
      label: `${formatTickValue(start)}-${formatTickValue(end)}`
    });
  }
  return items.reverse();
}

function renderSurfaceAxes(chart: XlsxChart, layout: ChartLayout) {
  const categories = getCategoryLabels(chart);
  const seriesLabels = chart.series.map((series, index) => normalizeCategoryLabel(series.name) || `Q${index + 1}`);
  const labelColor = resolveChartAxisTextColor(chart);
  const axisColor = chart.axisLineColor ?? chart.chartAreaBorderColor ?? "#888888";
  const rowCount = Math.max(1, seriesLabels.length);
  const columnCount = Math.max(1, categories.length);
  const columnPositions = categories.map((_, index) => (
    layout.plot.left + (columnCount <= 1 ? layout.plot.width / 2 : (index / (columnCount - 1)) * layout.plot.width)
  ));
  const rowPositions = seriesLabels.map((_, index) => (
    layout.plot.top + layout.plot.height - (rowCount <= 1 ? layout.plot.height / 2 : (index / (rowCount - 1)) * layout.plot.height)
  ));

  return (
    <g>
      <rect
        fill="none"
        height={layout.plot.height}
        stroke={lightenColor(axisColor, 0.18)}
        strokeWidth={0.8}
        width={layout.plot.width}
        x={layout.plot.left}
        y={layout.plot.top}
      />
      {categories.map((label, index) => (
        <text
          key={`surface-x-label-${index}`}
          fill={labelColor}
          fontSize={10}
          textAnchor="middle"
          x={columnPositions[index] ?? layout.plot.left}
          y={layout.plot.top + layout.plot.height + 14}
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
          x={layout.plot.left + layout.plot.width + 8}
          y={(rowPositions[index] ?? layout.plot.top) + 3}
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

function resolveStockRoleIndices(chart: XlsxChart) {
  const roles = {
    close: indexByName(chart.series, /close|last|end/),
    high: indexByName(chart.series, /high|max|top/),
    low: indexByName(chart.series, /low|min|bottom/),
    open: indexByName(chart.series, /open|start/),
    volume: indexByName(chart.series, /volume|vol/)
  };
  const usedIndices = new Set<number>(Object.values(roles).filter((value): value is number => value != null));
  const remainingIndices = chart.series.map((_, index) => index).filter((index) => !usedIndices.has(index));
  const canonicalOrder = chart.series.length >= 4
    ? (["open", "high", "low", "close"] as const)
    : (["high", "low", "close"] as const);
  canonicalOrder.forEach((role) => {
    if (roles[role] == null) {
      const nextIndex = remainingIndices.shift();
      if (nextIndex != null) {
        roles[role] = nextIndex;
      }
    }
  });
  return roles;
}

function resolveStockPalette(chart: XlsxChart, axisColor: string) {
  const raw = chart.raw && typeof chart.raw === "object" ? chart.raw as Record<string, unknown> : null;
  const highLowLines = raw?.highLowLines && typeof raw.highLowLines === "object"
    ? raw.highLowLines as Record<string, unknown>
    : null;
  const shapeProperties = highLowLines?.shapeProperties && typeof highLowLines.shapeProperties === "object"
    ? highLowLines.shapeProperties as Record<string, unknown>
    : null;
  const lineColor = normalizeRendererHexColor(shapeProperties?.lineColorHex) ?? axisColor ?? "#333333";
  const chartStyleId = typeof chart.chartStyleId === "number" ? chart.chartStyleId : null;
  const closeAccent = chartStyleId != null && chartStyleId >= 128 ? "#c0504d" : lineColor;
  const lowAccent = chartStyleId != null && chartStyleId >= 128 ? "#d9a3a0" : lightenColor(lineColor, 0.45);
  return {
    closeAccent,
    downFill: lightenColor(lineColor, 0.4),
    lineColor,
    lowAccent,
    openAccent: lightenColor(lineColor, 0.22),
    upFill: "#ffffff",
    volumeFill: lightenColor(lineColor, 0.26)
  };
}

function buildStockLegendItems(chart: XlsxChart, palette: ChartRendererPalette): LegendItem[] {
  const axisColor = chart.axisLineColor ?? chart.chartAreaBorderColor ?? palette.border;
  const roles = resolveStockRoleIndices(chart);
  const stockPalette = resolveStockPalette(chart, axisColor);
  return chart.series.map((series, index) => {
    let color = stockPalette.lineColor;
    if (roles.volume === index) {
      color = stockPalette.volumeFill;
    } else if (roles.open === index) {
      color = stockPalette.openAccent;
    } else if (roles.low === index) {
      color = stockPalette.lowAccent;
    } else if (roles.close === index) {
      color = stockPalette.closeAccent;
    }
    return {
      color,
      label: series.name ?? `Series ${index + 1}`
    };
  });
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
  const titleHeight = chart.title ? 24 : 8;
  const legendPosition = normalizeLegendPosition(chart.legend?.position);

  const legendVertical = legendItems.length > 0 && (legendPosition === "right" || legendPosition === "left");
  const legendHorizontal = legendItems.length > 0 && (legendPosition === "top" || legendPosition === "bottom");

  const plotLeft = 42 + (legendPosition === "left" ? 98 : 0);
  const plotRight = 16 + (legendPosition === "right" ? 98 : 0);
  const plotTop = titleHeight + (legendPosition === "top" ? 24 : 0);
  const plotBottom = 28 + (legendPosition === "bottom" ? 24 : 0);

  const plotWidth = Math.max(40, width - plotLeft - plotRight);
  const plotHeight = Math.max(40, height - plotTop - plotBottom);

  const compactWidth = width <= 280;
  const compactHeight = height <= 200;
  const finalLegendPosition = compactWidth || compactHeight
    ? (legendVertical ? "bottom" : legendPosition)
    : legendPosition;

  return {
    height,
    legendItems,
    legendPosition: finalLegendPosition,
    plot: {
      height: Math.max(40, height - plotTop - plotBottom),
      left: plotLeft,
      top: plotTop,
      width: Math.max(40, width - plotLeft - plotRight)
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
  const numeric = safeNumber(value);
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

function markerSymbolPath(symbol: string, size: number) {
  if (symbol === "none") {
    return "";
  }
  const symbolType = (() => {
    switch (symbol) {
      case "diamond":
        return symbolDiamond;
      case "square":
        return symbolSquare;
      case "triangle":
      case "triangle-up":
        return symbolTriangle;
      case "cross":
      case "plus":
        return symbolCross;
      case "star":
        return symbolStar;
      case "circle":
      default:
        return symbolCircle;
    }
  })();
  return d3Symbol().type(symbolType).size(size * size * Math.PI)() ?? "";
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
  endAngle: number
) {
  const ry = radius * tilt;
  return resolvePieFrontSegments(startAngle, endAngle).map(([segmentStart, segmentEnd]) => {
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
  angle: number
) {
  if (!isPieFrontFacingAngle(angle)) {
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
  return (
    <text
      fill={chart.titleColor ?? chart.textColor ?? DEFAULT_CHART_TEXT_COLOR}
      fontFamily={buildChartFontFamily(chart.titleFontFamily ?? chart.fontFamily)}
      fontSize={fontSize}
      fontWeight={600}
      textAnchor="middle"
      x={layout.width / 2}
      y={16}
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

  if (legendPos === "left" || legendPos === "right") {
    const x = legendPos === "right"
      ? layout.plot.left + layout.plot.width + 8
      : 8;
    const startY = layout.plot.top + 6;
    return (
      <g>
        {items.map((item, index) => {
          const y = startY + index * 18;
          return (
            <g key={`legend-${index}`} transform={`translate(${x}, ${y})`}>
              <rect fill={item.color} height={swatchSize} rx={1.2} ry={1.2} width={swatchSize} x={0} y={-7} />
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
          <g key={`legend-${index}`} transform={`translate(${cursorX}, ${rowY})`}>
            <rect fill={item.color} height={swatchSize} rx={1.2} ry={1.2} width={swatchSize} x={0} y={-7} />
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
  const depthX = bar.isHorizontal ? 10 : 9;
  const depthY = -7;
  const frontX = bar.left;
  const frontY = bar.top;
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

  const topFace = `${frontX},${frontY} ${frontX2},${frontY} ${frontX2 + sideDepthX},${frontY + depthY} ${frontX + sideDepthX},${frontY + depthY}`;
  const sideFace = `${sideAnchorX},${frontY} ${sideAnchorX},${frontY2} ${sideAnchorX + sideDepthX},${frontY2 + depthY} ${sideAnchorX + sideDepthX},${frontY + depthY}`;
  const frontFill = bar.gradientId ? `url(#${bar.gradientId})` : bar.color;
  const sideFill = bar.invertedNegative ? bar.color : darkenColor(bar.color, 0.22);
  const topFill = bar.invertedNegative ? lightenColor(bar.color, 0.04) : lightenColor(bar.color, 0.24);

  if (bar.isHorizontal || normalizedShape === "box") {
    return (
      <g key={`${bar.key}-3d`}>
        <polygon fill={sideFill} points={sideFace} stroke={bar.stroke} strokeWidth={Math.max(0.6, bar.strokeWidth * 0.65)} />
        <polygon fill={topFill} points={topFace} stroke={bar.stroke} strokeWidth={Math.max(0.6, bar.strokeWidth * 0.65)} />
        <rect fill={frontFill} height={frontH} stroke={bar.stroke} strokeWidth={bar.strokeWidth} width={frontW} x={frontX} y={frontY} />
      </g>
    );
  }

  if (normalizedShape === "cylinder") {
    const capRy = Math.max(2, Math.min(8, frontW * 0.18));
    const bodyPath = [
      `M ${toSvgNumber(frontX)} ${toSvgNumber(frontY + capRy)}`,
      `Q ${toSvgNumber(centerX)} ${toSvgNumber(frontY - capRy * 0.7)} ${toSvgNumber(frontX2)} ${toSvgNumber(frontY + capRy)}`,
      `L ${toSvgNumber(frontX2)} ${toSvgNumber(frontY2 - capRy)}`,
      `Q ${toSvgNumber(centerX)} ${toSvgNumber(frontY2 + capRy * 0.7)} ${toSvgNumber(frontX)} ${toSvgNumber(frontY2 - capRy)}`,
      "Z"
    ].join(" ");
    const cylinderSide = [
      `M ${toSvgNumber(frontX2)} ${toSvgNumber(frontY + capRy)}`,
      `L ${toSvgNumber(frontX2)} ${toSvgNumber(frontY2 - capRy)}`,
      `L ${toSvgNumber(frontX2 + sideDepthX)} ${toSvgNumber(frontY2 + depthY - capRy)}`,
      `L ${toSvgNumber(frontX2 + sideDepthX)} ${toSvgNumber(frontY + depthY + capRy)}`,
      "Z"
    ].join(" ");

    return (
      <g key={`${bar.key}-3d-cylinder`}>
        <path d={cylinderSide} fill={sideFill} stroke={bar.stroke} strokeWidth={Math.max(0.6, bar.strokeWidth * 0.65)} />
        <ellipse
          cx={frontX2 + sideDepthX * 0.5}
          cy={frontY + depthY + capRy}
          fill={topFill}
          rx={Math.max(1.5, frontW * 0.5)}
          ry={capRy}
          stroke={bar.stroke}
          strokeWidth={Math.max(0.6, bar.strokeWidth * 0.65)}
        />
        <path d={bodyPath} fill={frontFill} stroke={bar.stroke} strokeWidth={bar.strokeWidth} />
        <ellipse
          cx={centerX}
          cy={frontY + capRy}
          fill={topFill}
          rx={Math.max(1.5, frontW * 0.5)}
          ry={capRy}
          stroke={bar.stroke}
          strokeWidth={Math.max(0.6, bar.strokeWidth * 0.75)}
        />
      </g>
    );
  }

  const taperedTopFace = `${topLeft},${frontY} ${topRight},${frontY} ${topRight + sideDepthX},${frontY + depthY} ${topLeft + sideDepthX},${frontY + depthY}`;
  const taperedSideFace = `${topRight},${frontY} ${bottomRight},${frontY2} ${bottomRight + sideDepthX},${frontY2 + depthY} ${topRight + sideDepthX},${frontY + depthY}`;
  const pyramidFrontFace = `${topLeft},${frontY} ${topRight},${frontY} ${bottomRight},${frontY2} ${bottomLeft},${frontY2}`;
  const coneFrontFace = [
    `M ${toSvgNumber(topLeft)} ${toSvgNumber(frontY)}`,
    `L ${toSvgNumber(topRight)} ${toSvgNumber(frontY)}`,
    `C ${toSvgNumber(topRight + (bottomRight - topRight) * 0.18)} ${toSvgNumber(frontY + frontH * 0.32)} ${toSvgNumber(bottomRight)} ${toSvgNumber(frontY + frontH * 0.72)} ${toSvgNumber(bottomRight)} ${toSvgNumber(frontY2)}`,
    `L ${toSvgNumber(bottomLeft)} ${toSvgNumber(frontY2)}`,
    `C ${toSvgNumber(bottomLeft)} ${toSvgNumber(frontY + frontH * 0.72)} ${toSvgNumber(topLeft - (topLeft - bottomLeft) * 0.18)} ${toSvgNumber(frontY + frontH * 0.32)} ${toSvgNumber(topLeft)} ${toSvgNumber(frontY)}`,
    "Z"
  ].join(" ");

  return (
    <g key={`${bar.key}-3d-${normalizedShape}`}>
      <polygon fill={sideFill} points={taperedSideFace} stroke={bar.stroke} strokeWidth={Math.max(0.6, bar.strokeWidth * 0.65)} />
      <polygon fill={topFill} points={taperedTopFace} stroke={bar.stroke} strokeWidth={Math.max(0.6, bar.strokeWidth * 0.65)} />
      {normalizedShape === "cone" ? (
        <path d={coneFrontFace} fill={frontFill} stroke={bar.stroke} strokeWidth={bar.strokeWidth} />
      ) : (
        <polygon fill={frontFill} points={pyramidFrontFace} stroke={bar.stroke} strokeWidth={bar.strokeWidth} />
      )}
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
  isStackedSeries: boolean
) {
  const plot = layout.plot;
  const valueSpan = Math.max(1e-6, maxValue - minValue);
  const seriesCount = Math.max(1, chart.series.length);
  const categoryCount = Math.max(1, categories.length);
  const depthScale = clamp((chart.view3d?.depthPercent ?? 100) / 100, 0.5, 4);
  const halfDepth = (seriesCount <= 1 ? 0.7 : 0.9) * depthScale;
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
  const normalizeZ = (seriesIndex: number) => (
    seriesCount <= 1 || (isAreaChart && isStackedSeries)
      ? frontZ * 0.94
      : (((seriesIndex / (seriesCount - 1)) - 0.5) * (frontZ - backZ))
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
      const top = projectCartesian3dPoint(
        x,
        normalizeY(topValue),
        z,
        rotXRad,
        rotYRad,
        usePerspective,
        perspectiveStrength
      );
      const bottom = projectCartesian3dPoint(
        x,
        normalizeY(bottomValue),
        z,
        rotXRad,
        rotYRad,
        usePerspective,
        perspectiveStrength
      );
      return {
        bottom,
        defined: point.defined,
        depth: top.depth,
        top
      };
    });
  });

  const bounds = [
    ...cubeCorners,
    ...projectedSeries.flatMap((series) => series.flatMap((point) => point.defined ? [point.top, point.bottom] : []))
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
    averageDepth: series.reduce((sum, point) => sum + point.depth, 0) / Math.max(1, series.length),
    points: series.map((point) => ({
      bottom: toScreenPoint(point.bottom),
      defined: point.defined,
      top: toScreenPoint(point.top)
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
  const yTicks = buildNumericTickValues(minValue, maxValue, chart.valueAxis?.majorUnit);
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
        const strokeColor = chartSeriesStrokeColor(chart, seriesIndex);
        const fillColor = chartSeriesColor(chart, seriesIndex);
        const markerSymbol = normalizeChartMarkerSymbol(chart.series[seriesIndex]?.markerSymbol);
        const markerPath = markerSymbolPath(markerSymbol, Math.max(4, chart.series[seriesIndex]?.markerSize ?? 6) * 0.52);

        return (
          <g key={`line3d-series-${seriesIndex}`}>
            {isAreaChart && areaPoints.length >= 3 ? (
              <path
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
                      d={markerPath}
                      fill={chart.series[seriesIndex]?.markerColor ?? fillColor}
                      stroke={chart.series[seriesIndex]?.markerLineColor ?? chart.chartAreaFillColor ?? palette.surface}
                      strokeWidth={1}
                    />
                  </g>
                ))
              : null}
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

function renderBarChart(chart: XlsxChart, palette: ChartRendererPalette, layout: ChartLayout, chartType: string) {
  if (chart.series.length === 0) {
    return null;
  }
  const sourceCategories = getCategoryLabels(chart);
  if (sourceCategories.length === 0) {
    return null;
  }
  const isPercentStacked = chartType === "ColumnPercentStacked" || chartType === "BarPercentStacked";
  const isHorizontal = chartType === "BarClustered" || chartType === "BarStacked" || chartType === "BarPercentStacked";
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
  const plot = chart.is3d
    ? {
        ...layout.plot,
        height: Math.max(20, layout.plot.height - 12),
        width: Math.max(20, layout.plot.width - 14)
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
  const positiveTotals = Array.from({ length: categoryCount }, (_, categoryIndex) => (
    matrix.reduce((sum, row) => sum + Math.max(0, row[categoryIndex] ?? 0), 0)
  ));
  const negativeTotals = Array.from({ length: categoryCount }, (_, categoryIndex) => (
    Math.abs(matrix.reduce((sum, row) => sum + Math.min(0, row[categoryIndex] ?? 0), 0))
  ));

  if (isPercentStacked) {
    for (let categoryIndex = 0; categoryIndex < categoryCount; categoryIndex += 1) {
      const total = matrix.reduce((sum, row) => sum + Math.max(0, row[categoryIndex] ?? 0), 0);
      for (let seriesIndex = 0; seriesIndex < seriesCount; seriesIndex += 1) {
        matrix[seriesIndex][categoryIndex] = total > 0 ? ((matrix[seriesIndex][categoryIndex] ?? 0) / total) * 100 : 0;
      }
    }
  }

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

  minValue = Math.min(0, minValue);
  maxValue = Math.max(0, maxValue);
  if (typeof chart.valueAxis?.min === "number" && Number.isFinite(chart.valueAxis.min)) {
    minValue = chart.valueAxis.min;
  }
  if (typeof chart.valueAxis?.max === "number" && Number.isFinite(chart.valueAxis.max)) {
    maxValue = chart.valueAxis.max;
  }
  if (maxValue <= minValue) {
    maxValue = minValue + 1;
  }

  const categoryBandPadding = resolveCategoryBandPadding(chart.gapWidth);
  const categoryScale = scaleBand<string>()
    .domain(categories)
    .range(isHorizontal ? [plot.top, plot.top + plot.height] : [plot.left, plot.left + plot.width])
    .paddingInner(categoryBandPadding.inner)
    .paddingOuter(categoryBandPadding.outer);

  const seriesScale = scaleBand<string>()
    .domain(Array.from({ length: seriesCount }, (_, index) => String(index)))
    .range([0, categoryScale.bandwidth()])
    .paddingInner(0.16)
    .paddingOuter(0.08);

  const shouldReverseValueAxis = chart.valueAxis?.orientation === "maxMin";
  const valueScale = scaleLinear()
    .domain([minValue, maxValue])
    .range(
      isHorizontal
        ? (shouldReverseValueAxis ? [plot.left + plot.width, plot.left] : [plot.left, plot.left + plot.width])
        : (shouldReverseValueAxis ? [plot.top, plot.top + plot.height] : [plot.top + plot.height, plot.top])
    );

  const majorUnit = isPercentStacked
    ? chart.valueAxis?.majorUnit ?? 20
    : chart.valueAxis?.majorUnit;
  const ticks = buildNumericTickValues(minValue, maxValue, majorUnit);
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
      const barThickness = isStacked ? categoryScale.bandwidth() : seriesScale.bandwidth();
      const barOffset = isStacked ? 0 : (seriesScale(String(seriesIndex)) ?? 0);
      const shapeTaper = normalized3dShape === "pyramid"
        ? 0.94
        : normalized3dShape === "cone"
          ? 0.88
          : 0;
      let topScale = 1;
      let bottomScale = 1;
      if (chart.is3d && !isHorizontal && shapeTaper > 0) {
        const scaleAt = (ratio: number) => clamp(1 - clamp(ratio, 0, 1) * shapeTaper, 0.04, 1);
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
        bars.push({
          bottomScale,
          categoryIndex,
          color: colors.fill,
          height: Math.max(1, barThickness),
          isHorizontal: true,
          key: `bar-${seriesIndex}-${categoryIndex}`,
          left: Math.min(x1, x2),
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
        bars.push({
          bottomScale,
          categoryIndex,
          color: colors.fill,
          height: Math.max(1, Math.abs(y2 - y1)),
          isHorizontal: false,
          key: `bar-${seriesIndex}-${categoryIndex}`,
          left: categoryStart + barOffset,
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

  const frameNode = chart.is3d ? (
    <g>
      <polygon
        fill={lightenColor(chart.chartAreaFillColor ?? palette.surface, 0.05)}
        points={`${plot.left},${plot.top + plot.height} ${plot.left + plot.width},${plot.top + plot.height} ${plot.left + plot.width + 10},${plot.top + plot.height - 7} ${plot.left + 10},${plot.top + plot.height - 7}`}
        stroke={lightenColor(chart.axisLineColor ?? palette.border, 0.2)}
        strokeWidth={1}
      />
      <polygon
        fill={lightenColor(chart.chartAreaFillColor ?? palette.surface, 0.12)}
        points={`${plot.left},${plot.top} ${plot.left + plot.width},${plot.top} ${plot.left + plot.width + 10},${plot.top - 7} ${plot.left + 10},${plot.top - 7}`}
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
      {chart.is3d
        ? renderedBars.map((bar) => renderExtrudedRect(bar))
        : renderedBars.map((bar) => (
            <rect
              key={bar.key}
              fill={bar.color}
              height={bar.height}
              stroke={bar.stroke}
              strokeWidth={bar.strokeWidth}
              width={bar.width}
              x={bar.left}
              y={bar.top}
            />
          ))}
    </g>
  );
}

function renderLineOrAreaChart(chart: XlsxChart, palette: ChartRendererPalette, layout: ChartLayout, chartType: string) {
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
  const hasExplicitMin = typeof chart.valueAxis?.min === "number" && Number.isFinite(chart.valueAxis.min);
  const hasExplicitMax = typeof chart.valueAxis?.max === "number" && Number.isFinite(chart.valueAxis.max);
  if (isAreaChart) {
    minValue = Math.min(0, minValue);
  } else if ((chartType === "Line" || isStackedLine) && chart.valueAxis?.crosses === "autoZero") {
    minValue = Math.min(0, minValue);
  }
  if (isPercentStackedSeries) {
    minValue = hasExplicitMin ? Number(chart.valueAxis?.min) : Math.min(0, minValue);
    maxValue = hasExplicitMax ? Number(chart.valueAxis?.max) : Math.max(100, maxValue);
  }
  minValue = hasExplicitMin ? Number(chart.valueAxis?.min) : minValue;
  maxValue = hasExplicitMax ? Number(chart.valueAxis?.max) : maxValue;
  if (maxValue <= minValue) {
    maxValue = minValue + 1;
  }

  const xScale = scalePoint<string>()
    .domain(categories)
    .range([plot.left, plot.left + plot.width]);
  const yScale = scaleLinear()
    .domain([minValue, maxValue])
    .range([plot.top + plot.height, plot.top]);

  const ticks = buildNumericTickValues(
    minValue,
    maxValue,
    isPercentStackedSeries ? (chart.valueAxis?.majorUnit ?? 20) : chart.valueAxis?.majorUnit
  );
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
      isStackedSeries
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
            const second = plotPointsBySeries[1]?.[categoryIndex];
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
                d={areaPath}
                fill={seriesFillColor}
                fillOpacity={1}
                stroke="none"
              />
            ) : null}
            <path
              d={linePath}
              fill="none"
              stroke={chartSeriesStrokeColor(chart, seriesIndex)}
              strokeLinecap="round"
              strokeLinejoin="round"
              strokeWidth={Math.max(1.5, series.lineWidthPx ?? 2)}
            />
            {points.map((point, pointIndex) => (
              point.y == null ? null : (
                <circle
                  key={`line-marker-${seriesIndex}-${pointIndex}`}
                  cx={point.x}
                  cy={yScale(point.y)}
                  fill={series.markerColor ?? chartSeriesColor(chart, seriesIndex)}
                  r={Math.max(2, (series.markerSize ?? 6) * 0.25)}
                  stroke={series.markerLineColor ?? chart.chartAreaFillColor ?? chartSeriesStrokeColor(chart, seriesIndex)}
                  strokeWidth={1}
                />
              )
            ))}
          </g>
        );
      })}
    </g>
  );
}

function renderComboChart(chart: XlsxChart, palette: ChartRendererPalette, layout: ChartLayout) {
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
  const categoryBandPadding = resolveCategoryBandPadding(groupGapWidth);
  const categoryScale = scaleBand<string>()
    .domain(categories)
    .range([plot.left, plot.left + plot.width])
    .paddingInner(categoryBandPadding.inner)
    .paddingOuter(categoryBandPadding.outer);
  const seriesScale = scaleBand<string>()
    .domain(Array.from({ length: columnGroup.series.length }, (_, index) => String(index)))
    .range([0, categoryScale.bandwidth()])
    .paddingInner(0.16)
    .paddingOuter(0.08);
  const categoryPositions = categories.map((category) => (
    (categoryScale(category) ?? plot.left) + categoryScale.bandwidth() / 2
  ));

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
        const value = safeNumber(series.values[categoryIndex]) ?? 0;
        const categoryStart = categoryScale(category) ?? plot.left;
        const barWidth = Math.max(1, seriesScale.bandwidth());
        const x = categoryStart + (seriesScale(String(seriesIndex)) ?? 0);
        const y = primaryScale(Math.max(0, value));
        const zeroY = primaryScale(0);
        const height = Math.max(1, Math.abs(zeroY - primaryScale(value)));
        return (
          <rect
            key={`combo-bar-${seriesIndex}-${categoryIndex}`}
            fill={series.color ?? series.lineColor ?? chartSeriesColor(primaryChart, seriesIndex)}
            height={height}
            stroke={series.lineColor ?? series.color ?? chartSeriesStrokeColor(primaryChart, seriesIndex)}
            strokeWidth={1}
            width={barWidth}
            x={x}
            y={Math.min(y, zeroY)}
          />
        );
      }))}
      {lineGroup.series.map((series, seriesIndex) => {
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
              d={path}
              fill="none"
              stroke={series.lineColor ?? series.color ?? chartSeriesStrokeColor(chart, columnGroup.series.length + seriesIndex)}
              strokeLinecap="round"
              strokeLinejoin="round"
              strokeWidth={Math.max(1.5, series.lineWidthPx ?? 2)}
            />
            {points.map((point, pointIndex) => (
              point.y == null ? null : (
                <path
                  d={markerSymbolPath(normalizeChartMarkerSymbol(series.markerSymbol), Math.max(4, series.markerSize ?? 7)) || markerSymbolPath("circle", 7)}
                  fill={series.markerColor ?? series.color ?? series.lineColor ?? chartSeriesColor(chart, columnGroup.series.length + seriesIndex)}
                  key={`combo-line-marker-${seriesIndex}-${pointIndex}`}
                  stroke={series.markerLineColor ?? chart.chartAreaFillColor ?? "#ffffff"}
                  strokeWidth={1}
                  transform={`translate(${point.x}, ${secondaryScale(point.y)})`}
                />
              )
            ))}
          </g>
        );
      })}
    </g>
  );
}

function renderScatterChart(chart: XlsxChart, palette: ChartRendererPalette, layout: ChartLayout, smooth: boolean) {
  const plot = layout.plot;
  const rawRecord = chart.raw && typeof chart.raw === "object"
    ? chart.raw as Record<string, unknown>
    : null;
  const normalizedTitle = typeof chart.title === "string" ? chart.title.trim().toLowerCase() : "";
  const scatterStyle = typeof chart.scatterStyle === "string"
    ? chart.scatterStyle
    : rawRecord && typeof rawRecord.scatterStyle === "string"
      ? rawRecord.scatterStyle
      : undefined;
  const styleDrawsLine = scatterStyle
    ? scatterStyle === "line" || scatterStyle === "lineMarker" || scatterStyle === "smooth" || scatterStyle === "smoothMarker"
    : true;
  const styleShowsMarkers = scatterStyle
    ? scatterStyle === "marker" || scatterStyle === "lineMarker" || scatterStyle === "smoothMarker"
    : true;
  const styleUsesSmoothCurve = scatterStyle
    ? scatterStyle === "smooth" || scatterStyle === "smoothMarker"
    : smooth;
  const titleForcesMarkerOnly = normalizedTitle.includes("marker only");
  const titleForcesNoMarkers = normalizedTitle.includes("no markers");
  const titleForcesWithMarkers = normalizedTitle.includes("with markers");
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
        const markerSize = Math.max(5, series.markerSize ?? 7);
        const markerPath = markerSymbolPath(
          normalizeChartMarkerSymbol(series.markerSymbol),
          markerSize * 0.55
        );
        const markerFill = series.lineColor ?? series.markerColor ?? series.color ?? chartSeriesStrokeColor(chart, seriesIndex);
        const shouldDrawLine = !titleForcesMarkerOnly
          && styleDrawsLine
          && series.shapeProperties?.xmlLineHidden !== true
          && seriesPoints.points.length > 1;
        const shouldDrawMarkers = !titleForcesNoMarkers
          && (titleForcesMarkerOnly || titleForcesWithMarkers || styleShowsMarkers)
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
                    d={markerPath}
                    fill={markerFill}
                    stroke="none"
                    strokeWidth={0}
                  />
                </g>
              );
            })}
          </g>
        );
      })}
    </g>
  );
}

function renderBubbleChart(chart: XlsxChart, palette: ChartRendererPalette, layout: ChartLayout) {
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
  const paddedMinX = typeof chart.categoryAxis?.min === "number" && Number.isFinite(chart.categoryAxis.min)
    ? chart.categoryAxis.min
    : minX - xPad;
  const paddedMaxX = typeof chart.categoryAxis?.max === "number" && Number.isFinite(chart.categoryAxis.max)
    ? chart.categoryAxis.max
    : safeMaxX + xPad;
  const paddedMinY = typeof chart.valueAxis?.min === "number" && Number.isFinite(chart.valueAxis.min)
    ? chart.valueAxis.min
    : minY - yPad;
  const paddedMaxY = typeof chart.valueAxis?.max === "number" && Number.isFinite(chart.valueAxis.max)
    ? chart.valueAxis.max
    : safeMaxY + yPad;
  const xScale = scaleLinear().domain([paddedMinX, paddedMaxX <= paddedMinX ? paddedMinX + 1 : paddedMaxX]).range([plot.left, plot.left + plot.width]);
  const yScale = scaleLinear().domain([paddedMinY, paddedMaxY <= paddedMinY ? paddedMinY + 1 : paddedMaxY]).range([plot.top + plot.height, plot.top]);

  const xTicks = buildNumericTickValues(paddedMinX, paddedMaxX <= paddedMinX ? paddedMinX + 1 : paddedMaxX, chart.categoryAxis?.majorUnit);
  const yTicks = buildNumericTickValues(paddedMinY, paddedMaxY <= paddedMinY ? paddedMinY + 1 : paddedMaxY, chart.valueAxis?.majorUnit);
  const axisColor = chart.axisLineColor ?? chart.chartAreaBorderColor ?? palette.border;
  const labelColor = resolveChartAxisTextColor(chart);
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
                  cx={xScale(point.x)}
                  cy={yScale(point.y)}
                  fill={baseColor}
                  fillOpacity={0.78}
                  r={radius}
                  stroke={darkenColor(baseColor, 0.18)}
                  strokeWidth={1}
                />
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

function renderRadarChart(chart: XlsxChart, palette: ChartRendererPalette, layout: ChartLayout) {
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
              ...radialPoint(categoryIndex, 0)
            };
          }
          const ratio = clamp((rawValue - minValue) / (safeMax - minValue), 0, 1);
          return {
            defined: true,
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
        const markerSize = Math.max(4, series.markerSize ?? 6);
        const markerPath = markerSymbolPath(markerSymbol, markerSize * 0.5);
        const showMarkers = markerSymbol !== "none" && markerPath.length > 0;
        return (
          <g key={`radar-series-${seriesIndex}`}>
            {definedPoints.length >= 2 ? (
              <path
                d={polygon}
                fill={filled && !hasGap && definedPoints.length >= 3 ? color : "none"}
                fillOpacity={filled ? 0.26 : 0}
                stroke={chartSeriesStrokeColor(chart, seriesIndex)}
                strokeWidth={1.8}
              />
            ) : null}
            {showMarkers
              ? definedPoints.map((point, pointIndex) => (
                  <g
                    key={`radar-point-${seriesIndex}-${pointIndex}`}
                    transform={`translate(${point.x}, ${point.y})`}
                  >
                    <path
                      d={markerPath}
                      fill={series.markerColor ?? color}
                      stroke={series.markerLineColor ?? chart.chartAreaFillColor ?? palette.surface}
                      strokeWidth={1}
                    />
                  </g>
                ))
              : null}
          </g>
        );
      })}
    </g>
  );
}

function renderPieChart(chart: XlsxChart, palette: ChartRendererPalette, layout: ChartLayout, chartType: string) {
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

  const isPie3d = chartType === "Pie3D";
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
              arc.endAngle
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
                    arc.startAngle
                  );
                  const endWall = buildPieRadialWallPath(
                    centerX,
                    centerY,
                    outerRadius,
                    tilt,
                    depth,
                    arc.endAngle
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
                d={(arcPath(arc) ?? "")}
                fill={arc.data.color}
                stroke={sliceSeparatorColor}
                strokeWidth={sliceSeparatorWidth}
                transform={`translate(${centerX}, ${centerY})${isPie3d ? ` scale(1, ${tilt})` : ""}`}
                filter={isPie3d ? `url(#${shadowId})` : undefined}
              />
            </g>
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

function renderBarOfPieChart(chart: XlsxChart, palette: ChartRendererPalette, layout: ChartLayout) {
  const pieSeriesIndex = selectPrimaryPieSeriesIndex(chart);
  const pieSeries = chart.series[pieSeriesIndex];
  const categories = getCategoryLabels(chart);
  const values = pieSeries?.values.map((value) => Math.max(0, safeNumber(value) ?? 0)) ?? [];
  if (values.length === 0) {
    return null;
  }
  const raw = (chart.raw ?? {}) as Record<string, unknown>;
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
    : [{ color: chartPointColor(chart, index, pieSeriesIndex), label: categories[index], value }]);
  if (secondaryTotal > 0) {
    primaryData.push({
      color: chartPointColor(chart, secondaryIndices[0] ?? 0, pieSeriesIndex),
      label: "Other",
      value: secondaryTotal
    });
  }

  const pieCenterX = layout.plot.left + layout.plot.width * 0.28;
  const pieCenterY = layout.plot.top + layout.plot.height * 0.55;
  const pieRadius = Math.max(16, Math.min(layout.plot.height, layout.plot.width * 0.45) * 0.3);
  const arc = d3Arc<{ endAngle: number; startAngle: number }>().innerRadius(0).outerRadius(pieRadius);
  const pieArcs = d3Pie<{ color: string; label: string; value: number }>()
    .value((entry) => entry.value)
    .sort(null)
    .startAngle(((90 - (chart.firstSliceAngle ?? 0)) * Math.PI) / 180)
    .endAngle(((90 - (chart.firstSliceAngle ?? 0)) * Math.PI) / 180 + Math.PI * 2)(primaryData);

  const secondaryLabels = secondaryIndices.map((index) => categories[index] ?? "");
  const secondaryValues = secondaryIndices.map((index) => values[index] ?? 0);
  const barAreaLeft = layout.plot.left + layout.plot.width * 0.62;
  const barAreaWidth = layout.plot.width * 0.34;
  const barScale = scaleLinear()
    .domain([0, Math.max(1, ...secondaryValues)])
    .range([barAreaLeft, barAreaLeft + barAreaWidth]);
  const bandScale = scaleBand<string>()
    .domain(secondaryLabels)
    .range([layout.plot.top + 8, layout.plot.top + layout.plot.height - 8])
    .paddingInner(0.25)
    .paddingOuter(0.15);

  return (
    <g>
      {pieArcs.map((entry, index) => (
        <path
          key={`bar-of-pie-main-${index}`}
          d={arc(entry) ?? ""}
          fill={entry.data.color}
          stroke={chart.chartAreaFillColor ?? palette.surface}
          strokeWidth={1}
          transform={`translate(${pieCenterX}, ${pieCenterY})`}
        />
      ))}
      <line
        stroke={chart.chartAreaBorderColor ?? palette.border}
        strokeWidth={1}
        x1={pieCenterX + pieRadius}
        x2={barAreaLeft}
        y1={pieCenterY - pieRadius * 0.4}
        y2={layout.plot.top + 10}
      />
      <line
        stroke={chart.chartAreaBorderColor ?? palette.border}
        strokeWidth={1}
        x1={pieCenterX + pieRadius}
        x2={barAreaLeft}
        y1={pieCenterY + pieRadius * 0.4}
        y2={layout.plot.top + layout.plot.height - 10}
      />
      {secondaryLabels.map((label, index) => {
        const y = bandScale(label) ?? layout.plot.top;
        const barWidth = Math.max(1, (barScale(secondaryValues[index] ?? 0) - barAreaLeft));
        return (
          <g key={`bar-of-pie-secondary-${index}`}>
            <rect
              fill={chartPointColor(chart, secondaryIndices[index] ?? index, pieSeriesIndex)}
              height={bandScale.bandwidth()}
              width={barWidth}
              x={barAreaLeft}
              y={y}
            />
            <text
              fill={resolveChartAxisTextColor(chart)}
              fontSize={10}
              textAnchor="end"
              x={barAreaLeft - 4}
              y={y + bandScale.bandwidth() * 0.6}
            >
              {label}
            </text>
          </g>
        );
      })}
    </g>
  );
}

function renderSurfaceChart(chart: XlsxChart, palette: ChartRendererPalette, layout: ChartLayout) {
  const categories = getCategoryLabels(chart);
  const rows = chart.series.length;
  const cols = categories.length;
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
  const rawChartType = chart.raw && typeof chart.raw === "object" && typeof (chart.raw as Record<string, unknown>).xmlChartType === "string"
    ? (chart.raw as Record<string, unknown>).xmlChartType
    : "";
  const isContour = rawChartType === "surfaceChart";
  const minValue = domain.minValue;
  const safeMax = domain.safeMax;
  const rotX = clamp(chart.view3d?.rotX ?? (chart.wireframe ? 90 : 25), -88, 88) * (Math.PI / 180);
  const rotY = clamp(chart.view3d?.rotY ?? (chart.wireframe ? 0 : 30), -88, 88) * (Math.PI / 180);
  const usePerspective = chart.view3d?.rAngAx === false;
  const perspectiveStrength = clamp(
    (chart.view3d?.perspective ?? (usePerspective ? 30 : 0)) / 100,
    0,
    1
  );
  const depthScale = clamp((chart.view3d?.depthPercent ?? 100) / 100, 0.2, 4);
  const cosX = Math.cos(rotX);
  const sinX = Math.sin(rotX);
  const cosY = Math.cos(rotY);
  const sinY = Math.sin(rotY);

  type SurfacePoint = { depth: number; hasValue: boolean; value: number; x: number; y: number };
  const rawPoints: SurfacePoint[][] = Array.from({ length: rows }, (_, rowIndex) => (
    Array.from({ length: cols }, (_, columnIndex) => {
      const value = matrix[rowIndex][columnIndex];
      const hasValue = value != null;
      const normalizedX = cols <= 1 ? 0 : ((columnIndex / (cols - 1)) - 0.5) * 2;
      const normalizedY = rows <= 1 ? 0 : ((rowIndex / (rows - 1)) - 0.5) * 2;
      const normalizedZ = hasValue ? ((((value - minValue) / (safeMax - minValue)) - 0.5) * 1.8 * depthScale) : (-0.9 * depthScale);

      const x1 = normalizedX * cosY + normalizedZ * sinY;
      const z1 = -normalizedX * sinY + normalizedZ * cosY;
      const y1 = normalizedY * cosX - z1 * sinX;
      const z2 = normalizedY * sinX + z1 * cosX;
      const perspective = usePerspective
        ? 1 / Math.max(0.18, 1 + z2 * (0.24 + perspectiveStrength * 0.5))
        : 1;

      return {
        depth: z2,
        hasValue,
        value: value ?? minValue,
        x: x1 * perspective,
        y: y1 * perspective
      };
    })
  ));

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

  const points = rawPoints.map((row) => row.map((point) => ({
    ...point,
    x: centerX + (point.x - centerRawX) * scale,
    y: centerY + (point.y - centerRawY) * scale
  })));

  const baseColor = resolveSurfaceBaseColor(chart, palette);
  const axisColor = chart.axisLineColor ?? lightenColor(baseColor, 0.4);
  const wallFill = chart.backWall?.fillColor ?? "#d9d9df";
  const wallLineColor = chart.backWall?.lineColor ?? chart.sideWall?.lineColor ?? chart.floor?.lineColor ?? axisColor;

  type Quad = {
    color: string;
    depth: number;
    key: string;
    points: string;
    stroke: string;
  };
  const quads: Quad[] = [];
  for (let rowIndex = 0; rowIndex < rows - 1; rowIndex += 1) {
    for (let columnIndex = 0; columnIndex < cols - 1; columnIndex += 1) {
      const p00 = points[rowIndex][columnIndex];
      const p10 = points[rowIndex][columnIndex + 1];
      const p11 = points[rowIndex + 1][columnIndex + 1];
      const p01 = points[rowIndex + 1][columnIndex];
      if (!p00.hasValue || !p10.hasValue || !p11.hasValue || !p01.hasValue) {
        continue;
      }
      const averageValue = (p00.value + p10.value + p11.value + p01.value) / 4;
      const ratio = clamp((averageValue - minValue) / (safeMax - minValue), 0, 1);
      quads.push({
        color: resolveSurfaceColor(chart, palette, ratio),
        depth: (p00.depth + p10.depth + p11.depth + p01.depth) / 4,
        key: `surface-quad-${rowIndex}-${columnIndex}`,
        points: `${p00.x},${p00.y} ${p10.x},${p10.y} ${p11.x},${p11.y} ${p01.x},${p01.y}`,
        stroke: lightenColor(baseColor, 0.18)
      });
    }
  }
  quads.sort((left, right) => left.depth - right.depth);

  const buildSurfaceSegments = (segmentPoints: Array<{ depth: number; hasValue: boolean; value: number; x: number; y: number }>, keyPrefix: string) => (
    segmentPoints.slice(1).map((point, index) => {
      const previous = segmentPoints[index];
      if (!previous || !previous.hasValue || !point.hasValue) {
        return null;
      }
      const averageValue = (previous.value + point.value) / 2;
      const ratio = clamp((averageValue - minValue) / Math.max(1e-6, safeMax - minValue), 0, 1);
      const strokeColor = chart.wireframe
        ? resolveSurfaceColor(chart, palette, ratio)
        : darkenColor(resolveSurfaceColor(chart, palette, ratio), 0.1);
      return (
        <line
          key={`${keyPrefix}-${index}`}
          stroke={strokeColor}
          strokeWidth={chart.wireframe ? 1.8 : 0.8}
          x1={previous.x}
          x2={point.x}
          y1={previous.y}
          y2={point.y}
        />
      );
    })
  );

  return (
    <g>
      {isContour ? (
        <>
          <rect
            fill={wallFill}
            height={layout.plot.height}
            stroke={lightenColor(axisColor, 0.18)}
            strokeWidth={0.8}
            width={layout.plot.width}
            x={layout.plot.left}
            y={layout.plot.top}
          />
          {Array.from({ length: cols }, (_, columnIndex) => {
            const x = layout.plot.left + (cols <= 1 ? layout.plot.width / 2 : (columnIndex / (cols - 1)) * layout.plot.width);
            return (
              <line
                key={`surface-fallback-column-grid-${columnIndex}`}
                stroke={lightenColor(wallLineColor, 0.12)}
                strokeWidth={0.8}
                x1={x}
                x2={x}
                y1={layout.plot.top}
                y2={layout.plot.top + layout.plot.height}
              />
            );
          })}
          {Array.from({ length: rows }, (_, rowIndex) => {
            const y = layout.plot.top + layout.plot.height - (rows <= 1 ? layout.plot.height / 2 : (rowIndex / (rows - 1)) * layout.plot.height);
            return (
              <line
                key={`surface-fallback-row-grid-${rowIndex}`}
                stroke={lightenColor(wallLineColor, 0.12)}
                strokeWidth={0.8}
                x1={layout.plot.left}
                x2={layout.plot.left + layout.plot.width}
                y1={y}
                y2={y}
              />
            );
          })}
        </>
      ) : null}
      {chart.wireframe ? null : quads.map((quad) => (
        <polygon
          key={quad.key}
          fill={quad.color}
          fillOpacity={0.95}
          points={quad.points}
          stroke={quad.stroke}
          strokeWidth={0.7}
        />
      ))}
      {Array.from({ length: rows }, (_, rowIndex) => {
        const rowPoints = points[rowIndex];
        return (
          <g key={`surface-row-${rowIndex}`}>
            {buildSurfaceSegments(rowPoints, `surface-row-${rowIndex}`)}
          </g>
        );
      })}
      {Array.from({ length: cols }, (_, columnIndex) => {
        const columnPoints = Array.from({ length: rows }, (_, rowIndex) => points[rowIndex][columnIndex]);
        return (
          <g key={`surface-column-${columnIndex}`}>
            {buildSurfaceSegments(columnPoints, `surface-column-${columnIndex}`)}
          </g>
        );
      })}
      {chart.wireframe ? (
        <rect
          fill="none"
          height={layout.plot.height}
          stroke={lightenColor(axisColor, 0.18)}
          strokeWidth={0.8}
          width={layout.plot.width}
          x={layout.plot.left}
          y={layout.plot.top}
        />
      ) : null}
    </g>
  );
}

function indexByName(series: XlsxChart["series"], matcher: RegExp) {
  const index = series.findIndex((entry) => matcher.test((entry.name ?? "").toLowerCase()));
  return index >= 0 ? index : null;
}

function renderStockChart(chart: XlsxChart, palette: ChartRendererPalette, layout: ChartLayout) {
  const categories = getCategoryLabels(chart);
  if (categories.length === 0 || chart.series.length < 2) {
    return null;
  }

  const roles = resolveStockRoleIndices(chart);
  const highIndex = roles.high ?? 0;
  const lowIndex = roles.low ?? Math.min(1, chart.series.length - 1);
  const closeIndex = roles.close ?? Math.min(2, chart.series.length - 1);
  const openIndex = roles.open;
  const volumeIndex = roles.volume != null && ![highIndex, lowIndex, closeIndex, openIndex].includes(roles.volume)
    ? roles.volume
    : null;
  const high = chart.series[highIndex];
  const low = chart.series[lowIndex];
  const close = chart.series[closeIndex];
  const open = openIndex != null ? chart.series[openIndex] : null;
  const volume = volumeIndex != null ? chart.series[volumeIndex] : null;

  const points = categories
    .map((category, index) => {
      const highValue = safeNumber(high.values[index]);
      const lowValue = safeNumber(low.values[index]);
      const closeValue = safeNumber(close.values[index]);
      const openValue = open ? safeNumber(open.values[index]) : null;
      const volumeValue = volume ? safeNumber(volume.values[index]) : null;
      if (highValue == null || lowValue == null || closeValue == null) {
        return null;
      }
      return {
        category,
        close: closeValue,
        high: highValue,
        low: lowValue,
        open: openValue,
        volume: volumeValue
      };
    })
    .filter((entry): entry is { category: string; close: number; high: number; low: number; open: number | null; volume: number | null } => entry != null);
  if (points.length === 0) {
    return null;
  }

  const plot = layout.plot;
  const hasVolume = volume != null && points.some((entry) => (entry.volume ?? 0) > 0);
  const axisColor = chart.axisLineColor ?? chart.chartAreaBorderColor ?? palette.border;
  const stockPalette = resolveStockPalette(chart, axisColor);
  const rawMinValue = hasVolume
    ? Math.min(0, ...points.flatMap((entry) => [entry.low, entry.close, entry.open ?? entry.close, entry.volume ?? 0]))
    : Math.min(...points.map((entry) => entry.low));
  const rawMaxValue = hasVolume
    ? Math.max(...points.flatMap((entry) => [entry.high, entry.close, entry.open ?? entry.close, entry.volume ?? 0]))
    : Math.max(...points.map((entry) => entry.high));
  const resolvedDomain = resolveNumericAxisDomain(
    typeof chart.valueAxis?.min === "number" && Number.isFinite(chart.valueAxis.min) ? chart.valueAxis.min : rawMinValue,
    typeof chart.valueAxis?.max === "number" && Number.isFinite(chart.valueAxis.max) ? chart.valueAxis.max : rawMaxValue,
    chart.valueAxis?.majorUnit,
    hasVolume
  );
  const yScale = scaleLinear()
    .domain([resolvedDomain.min, resolvedDomain.max])
    .range([plot.top + plot.height, plot.top]);
  const xScale = scaleBand<string>()
    .domain(categories)
    .range([plot.left, plot.left + plot.width])
    .paddingInner(0.28)
    .paddingOuter(0.18);
  const ticks = resolvedDomain.ticks;
  const labelColor = resolveChartAxisTextColor(chart);
  const openTickColor = stockPalette.openAccent;
  const closeTickColor = openIndex != null ? stockPalette.lineColor : stockPalette.closeAccent;
  const labelStep = Math.max(1, Math.ceil(categories.length / Math.max(4, Math.floor(plot.width / 68))));
  const zeroY = yScale(Math.max(resolvedDomain.min, 0));

  return (
    <g>
      {ticks.map((tick) => (
        <g key={`stock-tick-${tick}`}>
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
      {categories.map((category, index) => {
        if (index % labelStep !== 0 && index !== categories.length - 1) {
          return null;
        }
        const x = (xScale(category) ?? plot.left) + xScale.bandwidth() * 0.5;
        return (
          <text
            key={`stock-cat-${category}`}
            fill={labelColor}
            fontSize={10}
            textAnchor="middle"
            x={x}
            y={plot.top + plot.height + 14}
          >
            {category}
          </text>
        );
      })}
      <line stroke={axisColor} strokeWidth={1.2} x1={plot.left} x2={plot.left} y1={plot.top} y2={plot.top + plot.height} />
      <line stroke={axisColor} strokeWidth={1.2} x1={plot.left} x2={plot.left + plot.width} y1={plot.top + plot.height} y2={plot.top + plot.height} />
      {points.map((entry, index) => {
        const x = (xScale(entry.category) ?? plot.left) + xScale.bandwidth() * 0.5;
        const previousClose = index > 0 ? points[index - 1]?.close ?? entry.close : entry.close;
        const isUp = entry.open != null
          ? entry.close >= entry.open
          : entry.close >= previousClose;
        const stroke = stockPalette.lineColor;
        const candleWidth = Math.max(5, xScale.bandwidth() * 0.56);
        const bodyLeft = x - candleWidth / 2;
        const openY = entry.open != null ? yScale(entry.open) : null;
        const closeY = yScale(entry.close);
        const highY = yScale(entry.high);
        const lowY = yScale(entry.low);
        return (
          <g key={`stock-point-${index}`}>
            {hasVolume && entry.volume != null ? (
              <rect
                fill={stockPalette.volumeFill}
                height={Math.max(1, zeroY - yScale(entry.volume))}
                opacity={0.94}
                stroke="none"
                width={Math.max(3, xScale.bandwidth() * 0.64)}
                x={x - Math.max(3, xScale.bandwidth() * 0.64) / 2}
                y={yScale(entry.volume)}
              />
            ) : null}
            <line stroke={stroke} strokeWidth={1.6} x1={x} x2={x} y1={highY} y2={lowY} />
            {entry.open != null && openY != null ? (
              Math.abs(openY - closeY) >= 1 ? (
                <rect
                  fill={isUp ? stockPalette.upFill : stockPalette.downFill}
                  height={Math.max(1.4, Math.abs(closeY - openY))}
                  stroke={stroke}
                  strokeWidth={1.3}
                  width={candleWidth}
                  x={bodyLeft}
                  y={Math.min(openY, closeY)}
                />
              ) : (
                <line stroke={stroke} strokeWidth={1.6} x1={bodyLeft} x2={bodyLeft + candleWidth} y1={closeY} y2={closeY} />
              )
            ) : (
              <line stroke={closeTickColor} strokeWidth={1.8} x1={x} x2={x + 7} y1={closeY} y2={closeY} />
            )}
            {entry.open != null && openY != null ? (
              <line stroke={openTickColor} strokeWidth={1.8} x1={x - 7} x2={x} y1={openY} y2={openY} />
            ) : null}
          </g>
        );
      })}
    </g>
  );
}

function renderWaterfallChart(chart: XlsxChart, palette: ChartRendererPalette, layout: ChartLayout) {
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
              fill={fill}
              rx={2}
              ry={2}
              stroke={darkenColor(fill, 0.22)}
              strokeWidth={1}
              x={bandLeft}
              y={top}
              width={bandWidth}
              height={height}
            />
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

function renderFunnelChart(chart: XlsxChart, palette: ChartRendererPalette, layout: ChartLayout) {
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
        const nextValue = stages[index + 1]?.value ?? 0;
        const topWidth = (stage.value / maxValue) * plot.width;
        const bottomWidth = (nextValue / maxValue) * plot.width;
        const topY = plot.top + index * sectionHeight;
        const bottomY = topY + sectionHeight - 2;
        const fill = stage.isSubtotal ? darkenColor(stage.color, 0.14) : stage.color;
        const path = [
          `M ${centerX - topWidth * 0.5} ${topY}`,
          `L ${centerX + topWidth * 0.5} ${topY}`,
          `L ${centerX + bottomWidth * 0.5} ${bottomY}`,
          `L ${centerX - bottomWidth * 0.5} ${bottomY}`,
          "Z"
        ].join(" ");
        const labelFitsInside = topWidth > 90;

        return (
          <g key={`funnel-stage-${index}`}>
            <path
              d={path}
              fill={fill}
              stroke={darkenColor(fill, 0.2)}
              strokeWidth={1}
            />
            <text
              fill={labelColor}
              fontSize={10}
              textAnchor={labelFitsInside ? "middle" : "start"}
              x={labelFitsInside ? centerX : centerX + topWidth * 0.5 + 8}
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

function renderSunburstChart(chart: XlsxChart, palette: ChartRendererPalette, layout: ChartLayout) {
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

        return (
          <g key={`sunburst-node-${index}`}>
            <path
              d={path}
              fill={fill}
              stroke={palette.surface}
              strokeWidth={1}
            />
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

function renderTreemapChart(chart: XlsxChart, palette: ChartRendererPalette, layout: ChartLayout) {
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
    .tile(treemapSquarify)(root);

  return (
    <g transform={`translate(${layout.plot.left}, ${layout.plot.top})`}>
      {treemapRoot.leaves().map((leaf, index) => {
        const fill = resolveHierarchyNodeColor(chart, leaf);
        const width = Math.max(0, leaf.x1 - leaf.x0);
        const height = Math.max(0, leaf.y1 - leaf.y0);
        const canShowLabel = width > 48 && height > 22;

        return (
          <g key={`treemap-leaf-${index}`}>
            <rect
              fill={fill}
              rx={3}
              ry={3}
              stroke={palette.surface}
              strokeWidth={1}
              x={leaf.x0}
              y={leaf.y0}
              width={width}
              height={height}
            />
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

function renderChartPlot(chart: XlsxChart, palette: ChartRendererPalette, layout: ChartLayout, chartType: string) {
  if (isComboChart(chart)) {
    return renderComboChart(chart, palette, layout) ?? renderUnsupported(chart, palette, layout, "Combo");
  }
  if (
    chartType === "ColumnClustered"
    || chartType === "ColumnStacked"
    || chartType === "ColumnPercentStacked"
    || chartType === "BarClustered"
    || chartType === "BarStacked"
    || chartType === "BarPercentStacked"
  ) {
    return renderBarChart(chart, palette, layout, chartType);
  }
  if (
    chartType === "Line"
    || chartType === "LineStacked"
    || chartType === "LinePercentStacked"
    || chartType === "Area"
    || chartType === "AreaStacked"
    || chartType === "AreaPercentStacked"
  ) {
    return renderLineOrAreaChart(chart, palette, layout, chartType);
  }
  if (chartType === "Scatter") {
    return renderScatterChart(chart, palette, layout, false);
  }
  if (chartType === "ScatterLines") {
    return renderScatterChart(chart, palette, layout, false);
  }
  if (chartType === "ScatterSmooth") {
    return renderScatterChart(chart, palette, layout, true);
  }
  if (chartType === "Bubble") {
    return renderBubbleChart(chart, palette, layout);
  }
  if (chartType === "Radar") {
    return renderRadarChart(chart, palette, layout);
  }
  if (chartType === "Pie" || chartType === "Pie3D" || chartType === "PieExploded" || chartType === "Doughnut") {
    return renderPieChart(chart, palette, layout, chartType);
  }
  if (chartType === "BarOfPie") {
    return renderBarOfPieChart(chart, palette, layout);
  }
  if (chartType === "Surface") {
    return renderSurfaceChart(chart, palette, layout);
  }
  if (chartType === "Stock") {
    return renderStockChart(chart, palette, layout);
  }
  if (chartType === "Waterfall") {
    return renderWaterfallChart(chart, palette, layout);
  }
  if (chartType === "Funnel") {
    return renderFunnelChart(chart, palette, layout);
  }
  if (chartType === "Sunburst") {
    return renderSunburstChart(chart, palette, layout);
  }
  if (chartType === "Treemap") {
    return renderTreemapChart(chart, palette, layout);
  }
  return renderUnsupported(chart, palette, layout, chartType);
}

export const MemoChartSvg = React.memo(function MemoChartSvg({ chart, palette, rect }: ChartSvgProps) {
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

  if (renderChartType === "Surface") {
    return (
      <MemoSurfaceChartComposite
        background={background}
        borderColor={borderColor}
        chart={chart}
        fallback={renderSurfaceChart(chart, palette, layout)}
        fontFamily={fontFamily}
        layout={layout}
        overlay={
          <>
            {renderSurfaceAxes(chart, layout)}
            {renderTitle(chart, layout, palette)}
            {renderLegend(chart, layout, palette)}
          </>
        }
        palette={palette}
      />
    );
  }

  return (
    <svg
      aria-label={chart.title ?? chart.name ?? "Chart"}
      role="img"
      style={{ display: "block", fontFamily, height: "100%", pointerEvents: "none", width: "100%" }}
      viewBox={`0 0 ${layout.width} ${layout.height}`}
    >
      {hideBackgroundRect
        ? null
        : <rect fill={background} height={layout.height} stroke={borderColor} strokeWidth={1} width={layout.width} x={0} y={0} />}
      {renderTitle(chart, layout, palette)}
      {renderLegend(chart, layout, palette)}
      {renderChartPlot(chart, palette, layout, renderChartType)}
    </svg>
  );
}, (prev, next) => (
  prev.chart === next.chart
  && prev.palette === next.palette
  && prev.rect.height === next.rect.height
  && prev.rect.width === next.rect.width
  && prev.rect.left === next.rect.left
  && prev.rect.top === next.rect.top
));
