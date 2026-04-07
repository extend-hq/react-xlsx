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
import type { XlsxChart, XlsxImageRect } from "./types";

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
  color: string;
  gradientId?: string;
  height: number;
  isHorizontal: boolean;
  key: string;
  left: number;
  seriesIndex: number;
  stroke: string;
  strokeWidth: number;
  value: number;
  width: number;
  top: number;
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

function chartPointColor(chart: XlsxChart, pointIndex: number) {
  const pointStyle = chart.series[0]?.dataPointStyles?.find((entry) => entry.index === pointIndex);
  if (pointStyle?.color) {
    return pointStyle.color;
  }
  const rawPoint = chart.series[0]?.dataPoints?.[pointIndex];
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
  return chartSeriesColor(chart, 0);
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
  const referenceCategoryCount = Math.max(
    0,
    ...chart.series.map((series) => Math.max(
      estimateReferencePointCount(series.categoriesRef?.formula),
      estimateReferencePointCount(series.valuesRef?.formula)
    ))
  );
  const categoryCount = Math.max(
    referenceCategoryCount,
    primaryCategories.length,
    ...chart.series.map((series) => Math.max(series.categories.length, series.values.length))
  );
  if (categoryCount <= 0) {
    return [];
  }
  return Array.from({ length: categoryCount }, (_, categoryIndex) => {
    const primary = primaryCategories[categoryIndex];
    if (primary != null) {
      return String(primary);
    }
    const fallback = chart.series
      .map((series) => series.categories[categoryIndex])
      .find((value) => value != null);
    return fallback == null ? "" : String(fallback);
  });
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

function buildPieEntries(chart: XlsxChart) {
  const categories = getCategoryLabels(chart);
  const values = chart.series[0]?.values ?? [];
  return values
    .map((rawValue, index) => ({
      color: chartPointColor(chart, index),
      index,
      label: normalizeCategoryLabel(categories[index]),
      value: Math.max(0, safeNumber(rawValue) ?? 0)
    }))
    .filter((entry) => entry.value > 0 && entry.label.trim().length > 0);
}

function getLegendItems(chart: XlsxChart, chartType: string): LegendItem[] {
  if (!chart.legend) {
    return [];
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
  const isXyLegend = chartType === "ScatterLines" || chartType === "ScatterSmooth" || chartType === "Bubble";
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

function formatTickValue(value: number) {
  if (Math.abs(value) >= 1000) {
    return value.toLocaleString(undefined, { maximumFractionDigits: 0 });
  }
  return value.toLocaleString(undefined, { maximumFractionDigits: 2 });
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

function renderTitle(chart: XlsxChart, layout: ChartLayout, palette: ChartRendererPalette) {
  if (!chart.title) {
    return null;
  }
  return (
    <text
      fill={chart.titleColor ?? chart.textColor ?? palette.text}
      fontSize={12}
      fontWeight={600}
      textAnchor="middle"
      x={layout.width / 2}
      y={16}
    >
      {chart.title}
    </text>
  );
}

function renderLegend(chart: XlsxChart, layout: ChartLayout, palette: ChartRendererPalette) {
  if (!chart.legend || layout.legendItems.length === 0) {
    return null;
  }
  const textColor = chart.axisLabelColor ?? chart.textColor ?? palette.text;
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
  mapValue: (value: number) => number
) {
  const axisColor = chart.axisLineColor ?? chart.chartAreaBorderColor ?? palette.border;
  const labelColor = chart.axisLabelColor ?? chart.textColor ?? palette.text;
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
                {formatTickValue(tick)}
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
              {formatTickValue(tick)}
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

  const topFace = `${frontX},${frontY} ${frontX2},${frontY} ${frontX2 + sideDepthX},${frontY + depthY} ${frontX + sideDepthX},${frontY + depthY}`;
  const sideFace = `${sideAnchorX},${frontY} ${sideAnchorX},${frontY2} ${sideAnchorX + sideDepthX},${frontY2 + depthY} ${sideAnchorX + sideDepthX},${frontY + depthY}`;
  const frontFill = bar.gradientId ? `url(#${bar.gradientId})` : bar.color;

  return (
    <g key={`${bar.key}-3d`}>
      <polygon fill={darkenColor(bar.color, 0.22)} points={sideFace} stroke={bar.stroke} strokeWidth={Math.max(0.6, bar.strokeWidth * 0.65)} />
      <polygon fill={lightenColor(bar.color, 0.24)} points={topFace} stroke={bar.stroke} strokeWidth={Math.max(0.6, bar.strokeWidth * 0.65)} />
      <rect fill={frontFill} height={frontH} stroke={bar.stroke} strokeWidth={bar.strokeWidth} width={frontW} x={frontX} y={frontY} />
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
  const isHorizontal = chartType === "BarClustered" || chartType === "BarPercentStacked";
  const shouldReverseCategories = isHorizontal && chart.categoryAxis?.orientation !== "maxMin";
  const categories = shouldReverseCategories ? sourceCategories.slice().reverse() : sourceCategories;
  const isStacked = chartType === "ColumnStacked" || chartType === "BarPercentStacked";
  const negativeFillMode: "chartArea" | "none" | "series" = chartType === "BarPercentStacked"
    ? "chartArea"
    : chartType === "ColumnClustered"
      ? "none"
      : "series";
  const categoryCount = categories.length;
  const seriesCount = chart.series.length;
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

  if (chartType === "BarPercentStacked") {
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

  const valueScale = scaleLinear()
    .domain([minValue, maxValue])
    .range(isHorizontal ? [plot.left, plot.left + plot.width] : [plot.top + plot.height, plot.top]);

  const majorUnit = chartType === "BarPercentStacked"
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
      const category = categories[categoryIndex];
      const categoryStart = categoryScale(category) ?? 0;
      const barThickness = isStacked ? categoryScale.bandwidth() : seriesScale.bandwidth();
      const barOffset = isStacked ? 0 : (seriesScale(String(seriesIndex)) ?? 0);

      if (isHorizontal) {
        const x1 = valueScale(start);
        const x2 = valueScale(end);
        bars.push({
          categoryIndex,
          color: colors.fill,
          height: Math.max(1, barThickness),
          isHorizontal: true,
          key: `bar-${seriesIndex}-${categoryIndex}`,
          left: Math.min(x1, x2),
          seriesIndex,
          stroke: colors.stroke,
          strokeWidth: colors.strokeWidth,
          top: categoryStart + barOffset,
          value: rawValue,
          width: Math.max(1, Math.abs(x2 - x1))
        });
      } else {
        const y1 = valueScale(start);
        const y2 = valueScale(end);
        bars.push({
          categoryIndex,
          color: colors.fill,
          height: Math.max(1, Math.abs(y2 - y1)),
          isHorizontal: false,
          key: `bar-${seriesIndex}-${categoryIndex}`,
          left: categoryStart + barOffset,
          seriesIndex,
          stroke: colors.stroke,
          strokeWidth: colors.strokeWidth,
          top: Math.min(y1, y2),
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
    (value) => valueScale(value)
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

  const allValues = chart.series.flatMap((series) => (
    series.values
      .slice(0, categories.length)
      .map((value) => resolveRenderableSeriesValue(value, displayBlanksAs))
      .filter((value): value is number => value != null)
  ));
  if (allValues.length === 0) {
    return null;
  }
  let minValue = Math.min(...allValues);
  let maxValue = Math.max(...allValues);
  const hasExplicitMin = typeof chart.valueAxis?.min === "number" && Number.isFinite(chart.valueAxis.min);
  const hasExplicitMax = typeof chart.valueAxis?.max === "number" && Number.isFinite(chart.valueAxis.max);
  if (chartType === "Area") {
    minValue = Math.min(0, minValue);
  } else if (chartType === "Line" && chart.valueAxis?.crosses === "autoZero") {
    minValue = Math.min(0, minValue);
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

  const ticks = buildNumericTickValues(minValue, maxValue, chart.valueAxis?.majorUnit);
  const categoryPositions = categories.map((category) => xScale(category) ?? plot.left);

  const curve: CurveFactory = curveLinear;
  const areaBaseline = yScale(Math.max(minValue, 0));

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
      {chart.series.map((series, seriesIndex) => {
        const points = categories.map((category, categoryIndex) => ({
          x: xScale(category) ?? plot.left,
          y: resolveRenderableSeriesValue(series.values[categoryIndex], displayBlanksAs)
        }));
        const linePath = d3Line<{ x: number; y: number | null }>()
          .defined((point) => point.y != null)
          .x((point) => point.x)
          .y((point) => yScale(point.y ?? 0))
          .curve(curve)(points) ?? "";

        const areaPath = chartType === "Area"
          ? d3Area<{ x: number; y: number | null }>()
            .defined((point) => point.y != null)
            .x((point) => point.x)
            .y0(areaBaseline)
            .y1((point) => yScale(point.y ?? 0))
            .curve(curve)(points) ?? ""
          : "";

        const seriesFillColor = typeof series.shapeProperties?.xmlFillColor === "string"
          ? series.shapeProperties.xmlFillColor
          : chartSeriesColor(chart, seriesIndex);
        return (
          <g key={`line-series-${seriesIndex}`}>
            {chartType === "Area" ? (
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

function renderScatterChart(chart: XlsxChart, palette: ChartRendererPalette, layout: ChartLayout, smooth: boolean) {
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
  const labelColor = chart.axisLabelColor ?? chart.textColor ?? palette.text;

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
        const shouldDrawLine = styleDrawsLine && series.shapeProperties?.xmlLineHidden !== true && seriesPoints.points.length > 1;
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
              if (!markerPath) {
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

  const xScale = scaleLinear().domain([minX, maxX <= minX ? minX + 1 : maxX]).range([plot.left, plot.left + plot.width]);
  const yScale = scaleLinear().domain([minY, maxY <= minY ? minY + 1 : maxY]).range([plot.top + plot.height, plot.top]);
  const radiusScale = scaleLinear()
    .domain([minBubbleMagnitude, maxBubbleMagnitude <= minBubbleMagnitude ? minBubbleMagnitude + 1 : maxBubbleMagnitude])
    .range([6, 18 + ((chart.bubbleScale ?? 100) * 0.18)]);

  const xTicks = buildNumericTickValues(minX, maxX <= minX ? minX + 1 : maxX, chart.categoryAxis?.majorUnit);
  const yTicks = buildNumericTickValues(minY, maxY <= minY ? minY + 1 : maxY, chart.valueAxis?.majorUnit);
  const axisColor = chart.axisLineColor ?? chart.chartAreaBorderColor ?? palette.border;
  const labelColor = chart.axisLabelColor ?? chart.textColor ?? palette.text;
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
          {points.map((point, pointIndex) => {
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
              <g key={`bubble-${seriesIndex}-${pointIndex}`}>
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
                    fill={chart.textColor ?? palette.text}
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
  const labelColor = chart.axisLabelColor ?? chart.textColor ?? palette.text;
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
  const pieData = buildPieEntries(chart);
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
  const tilt = isPie3d ? clamp(1 - (rotX / 90) * 1.1, 0.58, 0.92) : 1;
  const depthScale = (chart.view3d?.depthPercent ?? 100) / 100;
  const depth = isPie3d ? Math.max(8, outerRadius * 0.2 * depthScale) : 0;
  const dataLabelsEnabled = Boolean(chart.dataLabels?.showCategoryName || chart.dataLabels?.showPercent || chart.dataLabels?.showValue);
  const total = pieData.reduce((sum, entry) => sum + entry.value, 0);
  const clipId = `pie3d-front-${chart.id}`.replace(/[^A-Za-z0-9_-]/g, "-");
  const shadowId = `pie3d-shadow-${chart.id}`.replace(/[^A-Za-z0-9_-]/g, "-");

  const resolveSliceExplosion = (pointIndex: number) => {
    const pointStyle = chart.series[0]?.dataPointStyles?.find((entry) => entry.index === pointIndex);
    const seriesExplosion = typeof chart.series[0]?.shapeProperties?.xmlExplosion === "number"
      ? chart.series[0].shapeProperties.xmlExplosion
      : 0;
    const rawExplosion = Math.max(0, pointStyle?.explosion ?? seriesExplosion ?? 0);
    return outerRadius * clamp(rawExplosion / 100, 0, 0.65);
  };

  return (
    <g>
      {isPie3d ? (
        <defs>
          <filter id={shadowId} x="-40%" y="-40%" width="180%" height="200%">
            <feDropShadow dx="1.2" dy="3.6" floodColor="#000000" floodOpacity="0.28" stdDeviation="2.8" />
          </filter>
          <clipPath id={clipId}>
            <rect
              height={outerRadius * 2.6 + depth}
              width={outerRadius * 4.2}
              x={centerX - outerRadius * 2.1}
              y={centerY}
            />
          </clipPath>
        </defs>
      ) : null}
      {isPie3d ? (
        <g clipPath={`url(#${clipId})`}>
          {arcs.map((arc) => {
            const midAngle = (arc.startAngle + arc.endAngle) / 2;
            const explosion = resolveSliceExplosion(arc.data.index);
            const explodeX = Math.cos(midAngle) * explosion;
            const explodeY = Math.sin(midAngle) * explosion;
            return (
              <g key={`pie-side-${arc.data.index}`} transform={`translate(${explodeX}, ${explodeY})`}>
                <path
                  d={(arcPath(arc) ?? "")}
                  fill="none"
                  stroke={darkenColor(arc.data.color, 0.45)}
                  strokeLinecap="butt"
                  strokeLinejoin="round"
                  strokeWidth={Math.max(2, depth * 0.95)}
                  transform={`translate(${centerX}, ${centerY + depth * 0.5}) scale(1, ${tilt})`}
                />
              </g>
            );
          })}
        </g>
      ) : null}
      {isPie3d
        ? arcs.map((arc) => {
            const midAngle = (arc.startAngle + arc.endAngle) / 2;
            const explosion = resolveSliceExplosion(arc.data.index);
            const explodeX = Math.cos(midAngle) * explosion;
            const explodeY = Math.sin(midAngle) * explosion;
            return (
              <g key={`pie-base-${arc.data.index}`} transform={`translate(${explodeX}, ${explodeY})`}>
                <path
                  d={(arcPath(arc) ?? "")}
                  fill={darkenColor(arc.data.color, 0.3)}
                  transform={`translate(${centerX}, ${centerY + depth}) scale(1, ${tilt})`}
                />
              </g>
            );
          })
        : null}
      {arcs.map((arc) => {
        const explosion = resolveSliceExplosion(arc.data.index);
        const midAngle = (arc.startAngle + arc.endAngle) / 2;
        const explodeX = Math.cos(midAngle) * explosion;
        const explodeY = Math.sin(midAngle) * explosion * tilt;
        const labelRadius = outerRadius + 12;
        const labelX = centerX + Math.cos(midAngle) * labelRadius + explodeX;
        const labelY = centerY + Math.sin(midAngle) * labelRadius * tilt + explodeY;
        const pieces: string[] = [];
        if (chart.dataLabels?.showCategoryName) {
          pieces.push(arc.data.label);
        }
        if (chart.dataLabels?.showValue) {
          pieces.push(formatTickValue(arc.data.value));
        }
        if (chart.dataLabels?.showPercent) {
          pieces.push(`${Math.round((arc.data.value / Math.max(1, total)) * 100)}%`);
        }
        return (
          <g key={`pie-top-${arc.data.index}`} transform={`translate(${explodeX}, ${explodeY})`}>
            <path
              d={(arcPath(arc) ?? "")}
              fill={arc.data.color}
              stroke={chart.chartAreaFillColor ?? palette.surface}
              strokeWidth={1.2}
              transform={`translate(${centerX}, ${centerY})${isPie3d ? ` scale(1, ${tilt})` : ""}`}
              filter={isPie3d ? `url(#${shadowId})` : undefined}
            />
            {dataLabelsEnabled && pieces.length > 0 ? (
              <text
                fill={chart.textColor ?? palette.text}
                fontSize={10}
                textAnchor={labelX >= centerX ? "start" : "end"}
                x={labelX}
                y={labelY}
              >
                {pieces.join(", ")}
              </text>
            ) : null}
          </g>
        );
      })}
    </g>
  );
}

function renderBarOfPieChart(chart: XlsxChart, palette: ChartRendererPalette, layout: ChartLayout) {
  const categories = getCategoryLabels(chart);
  const values = chart.series[0]?.values.map((value) => Math.max(0, safeNumber(value) ?? 0)) ?? [];
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
    : [{ color: chartPointColor(chart, index), label: categories[index], value }]);
  if (secondaryTotal > 0) {
    primaryData.push({
      color: lightenColor(chartPointColor(chart, secondaryIndices[0] ?? 0), 0.2),
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
              fill={chartPointColor(chart, secondaryIndices[index] ?? index)}
              height={bandScale.bandwidth()}
              width={barWidth}
              x={barAreaLeft}
              y={y}
            />
            <text
              fill={chart.axisLabelColor ?? chart.textColor ?? palette.text}
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

function resolveSurfaceBaseColor(chart: XlsxChart, palette: ChartRendererPalette) {
  return chart.chartColorPalette?.[0]
    ?? chart.series[0]?.color
    ?? chart.series[0]?.lineColor
    ?? chart.axisLineColor
    ?? chart.textColor
    ?? palette.text;
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
  const numericValues = matrix.flatMap((row) => row.filter((value): value is number => value != null));
  if (numericValues.length === 0) {
    return null;
  }
  const minValue = typeof chart.valueAxis?.min === "number" && Number.isFinite(chart.valueAxis.min)
    ? chart.valueAxis.min
    : Math.min(...numericValues);
  const maxValue = typeof chart.valueAxis?.max === "number" && Number.isFinite(chart.valueAxis.max)
    ? chart.valueAxis.max
    : Math.max(...numericValues);
  const safeMax = maxValue <= minValue ? minValue + 1 : maxValue;

  const rotX = clamp(chart.view3d?.rotX ?? (chart.wireframe ? 68 : 35), -88, 88) * (Math.PI / 180);
  const rotY = clamp(chart.view3d?.rotY ?? 28, -88, 88) * (Math.PI / 180);
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
      const normalizedZ = hasValue ? (((value - minValue) / (safeMax - minValue)) - 0.5) * 1.8 : -0.9;

      const x1 = normalizedX * cosY + normalizedZ * sinY;
      const z1 = -normalizedX * sinY + normalizedZ * cosY;
      const y1 = normalizedY * cosX - z1 * sinX;
      const z2 = normalizedY * sinX + z1 * cosX;
      const perspective = 1 / Math.max(0.15, 1 + z2 * 0.38);

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
        color: mixRgbColor(darkenColor(baseColor, 0.28), lightenColor(baseColor, 0.45), ratio),
        depth: (p00.depth + p10.depth + p11.depth + p01.depth) / 4,
        key: `surface-quad-${rowIndex}-${columnIndex}`,
        points: `${p00.x},${p00.y} ${p10.x},${p10.y} ${p11.x},${p11.y} ${p01.x},${p01.y}`,
        stroke: lightenColor(baseColor, 0.18)
      });
    }
  }
  quads.sort((left, right) => left.depth - right.depth);

  return (
    <g>
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
        const path = d3Line<{ hasValue: boolean; x: number; y: number }>()
          .defined((point) => point.hasValue)
          .x((point) => point.x)
          .y((point) => point.y)
          .curve(curveLinear)(rowPoints) ?? "";
        return (
          <path
            key={`surface-row-${rowIndex}`}
            d={path}
            fill="none"
            stroke={chart.wireframe ? axisColor : lightenColor(baseColor, 0.15)}
            strokeWidth={chart.wireframe ? 2 : 0.8}
          />
        );
      })}
      {Array.from({ length: cols }, (_, columnIndex) => {
        const columnPoints = Array.from({ length: rows }, (_, rowIndex) => points[rowIndex][columnIndex]);
        const path = d3Line<{ hasValue: boolean; x: number; y: number }>()
          .defined((point) => point.hasValue)
          .x((point) => point.x)
          .y((point) => point.y)
          .curve(curveLinear)(columnPoints) ?? "";
        return (
          <path
            key={`surface-column-${columnIndex}`}
            d={path}
            fill="none"
            stroke={chart.wireframe ? axisColor : lightenColor(baseColor, 0.15)}
            strokeWidth={chart.wireframe ? 2 : 0.8}
          />
        );
      })}
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

  const highIndex = indexByName(chart.series, /high|max|top/) ?? 0;
  const lowIndex = indexByName(chart.series, /low|min|bottom/) ?? Math.min(1, chart.series.length - 1);
  const closeIndex = indexByName(chart.series, /close|last|end/) ?? Math.min(2, chart.series.length - 1);
  const high = chart.series[highIndex];
  const low = chart.series[lowIndex];
  const close = chart.series[closeIndex];

  const points = categories
    .map((category, index) => {
      const highValue = safeNumber(high.values[index]);
      const lowValue = safeNumber(low.values[index]);
      const closeValue = safeNumber(close.values[index]);
      if (highValue == null || lowValue == null || closeValue == null) {
        return null;
      }
      return {
        category,
        close: closeValue,
        high: highValue,
        low: lowValue
      };
    })
    .filter((entry): entry is { category: string; close: number; high: number; low: number } => entry != null);
  if (points.length === 0) {
    return null;
  }

  const plot = layout.plot;
  const minValue = Math.min(...points.map((entry) => entry.low ?? 0));
  const maxValue = Math.max(...points.map((entry) => entry.high ?? 0));
  const yScale = scaleLinear()
    .domain([minValue, maxValue <= minValue ? minValue + 1 : maxValue])
    .range([plot.top + plot.height, plot.top]);
  const xScale = scaleBand<string>()
    .domain(categories)
    .range([plot.left, plot.left + plot.width])
    .paddingInner(0.32)
    .paddingOuter(0.18);
  const ticks = buildNumericTickValues(minValue, maxValue <= minValue ? minValue + 1 : maxValue, chart.valueAxis?.majorUnit);
  const axisColor = chart.axisLineColor ?? chart.chartAreaBorderColor ?? palette.border;
  const labelColor = chart.axisLabelColor ?? chart.textColor ?? palette.text;

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
      {categories.map((category) => {
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
        const isUp = (entry.close ?? 0) >= previousClose;
        const stroke = isUp ? (chartSeriesColor(chart, closeIndex) ?? "#2f7d4d") : (chartSeriesColor(chart, lowIndex) ?? "#b03a2e");
        return (
          <g key={`stock-point-${index}`}>
            <line stroke={stroke} strokeWidth={1.8} x1={x} x2={x} y1={yScale(entry.high ?? 0)} y2={yScale(entry.low ?? 0)} />
            <line stroke={stroke} strokeWidth={1.8} x1={x} x2={x + 7} y1={yScale(entry.close ?? 0)} y2={yScale(entry.close ?? 0)} />
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
        fill={chart.textColor ?? palette.mutedText}
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
  if (chartType === "ColumnClustered" || chartType === "ColumnStacked" || chartType === "BarClustered" || chartType === "BarPercentStacked") {
    return renderBarChart(chart, palette, layout, chartType);
  }
  if (chartType === "Line" || chartType === "Area") {
    return renderLineOrAreaChart(chart, palette, layout, chartType);
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
  return renderUnsupported(chart, palette, layout, chartType);
}

export const MemoChartSvg = React.memo(function MemoChartSvg({ chart, palette, rect }: ChartSvgProps) {
  const renderChartType = normalizeRenderableChartType(chart);
  const legendItems = getLegendItems(chart, renderChartType);
  const layout = buildLayout(chart, rect, legendItems);
  const background = chart.chartAreaFillColor ?? palette.surface;
  const borderColor = chart.chartAreaBorderColor ?? lightenColor(palette.border, 0.2);

  return (
    <svg
      aria-label={chart.title ?? chart.name ?? "Chart"}
      role="img"
      style={{ display: "block", height: "100%", pointerEvents: "none", width: "100%" }}
      viewBox={`0 0 ${layout.width} ${layout.height}`}
    >
      <rect fill={background} height={layout.height} stroke={borderColor} strokeWidth={1} width={layout.width} x={0} y={0} />
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
