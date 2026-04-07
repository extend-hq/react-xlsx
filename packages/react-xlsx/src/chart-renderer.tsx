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
  color: string;
  height: number;
  isHorizontal: boolean;
  key: string;
  left: number;
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
  const palette = chart.chartColorPalette;
  if (palette && palette.length > 0) {
    const offset = chart.chartColorPaletteOffset ?? 0;
    return palette[(pointIndex + offset) % palette.length] ?? palette[pointIndex % palette.length];
  }
  return chartSeriesColor(chart, 0);
}

function chartSeriesBarColors(chart: XlsxChart, seriesIndex: number, value: number) {
  const series = chart.series[seriesIndex];
  const defaultFill = chartSeriesColor(chart, seriesIndex);
  const defaultStroke = chartSeriesStrokeColor(chart, seriesIndex);
  if (value < 0 && series?.invertIfNegative) {
    return {
      fill: series.negativeColor ?? chart.chartAreaFillColor ?? defaultFill,
      stroke: series.negativeLineColor ?? defaultStroke
    };
  }
  return {
    fill: defaultFill,
    stroke: defaultStroke
  };
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

function getCategoryLabels(chart: XlsxChart) {
  return chart.series[0]?.categories?.map((value) => (value == null ? "" : String(value))) ?? [];
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
      label: categories[index] ?? "",
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
          label,
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
  return chart.series.map((series, index) => ({
    color: chartSeriesColor(chart, index),
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
  for (let current = start; current <= maxValue + step * 0.75; current += step) {
    if (current >= minValue - step * 0.4) {
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

  if (legendPos === "left" || legendPos === "right") {
    const x = legendPos === "right"
      ? layout.plot.left + layout.plot.width + 8
      : 8;
    const startY = layout.plot.top + 6;
    return (
      <g>
        {items.map((item, index) => {
          const y = startY + index * 16;
          return (
            <g key={`legend-${index}`} transform={`translate(${x}, ${y})`}>
              <rect fill={item.color} height={10} rx={1.5} ry={1.5} width={10} x={0} y={-8} />
              <text fill={textColor} fontSize={10} x={14} y={0}>
                {item.label}
              </text>
            </g>
          );
        })}
      </g>
    );
  }

  const rowY = legendPos === "top" ? (layout.titleHeight + 12) : (layout.height - 8);
  const totalWidth = items.reduce((sum, item) => sum + 16 + Math.min(90, item.label.length * 5.4), 0);
  let cursorX = Math.max(8, (layout.width - totalWidth) / 2);
  return (
    <g>
      {items.map((item, index) => {
        const labelWidth = Math.min(90, item.label.length * 5.4);
        const node = (
          <g key={`legend-${index}`} transform={`translate(${cursorX}, ${rowY})`}>
            <rect fill={item.color} height={10} rx={1.5} ry={1.5} width={10} x={0} y={-8} />
            <text fill={textColor} fontSize={10} x={14} y={0}>
              {item.label}
            </text>
          </g>
        );
        cursorX += 16 + labelWidth;
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

  const sideAnchorX = bar.value >= 0 ? frontX2 : frontX;
  const sideDepthX = bar.value >= 0 ? depthX : -depthX;

  const topFace = `${frontX},${frontY} ${frontX2},${frontY} ${frontX2 + sideDepthX},${frontY + depthY} ${frontX + sideDepthX},${frontY + depthY}`;
  const sideFace = `${sideAnchorX},${frontY} ${sideAnchorX},${frontY2} ${sideAnchorX + sideDepthX},${frontY2 + depthY} ${sideAnchorX + sideDepthX},${frontY + depthY}`;

  return (
    <g key={`${bar.key}-3d`}>
      <polygon fill={darkenColor(bar.color, 0.22)} points={sideFace} />
      <polygon fill={lightenColor(bar.color, 0.24)} points={topFace} />
      <rect fill={bar.color} height={frontH} width={frontW} x={frontX} y={frontY} />
    </g>
  );
}

function renderBarChart(chart: XlsxChart, palette: ChartRendererPalette, layout: ChartLayout, chartType: string) {
  const categories = getCategoryLabels(chart);
  if (categories.length === 0 || chart.series.length === 0) {
    return null;
  }
  const isHorizontal = chartType === "BarClustered" || chartType === "BarPercentStacked";
  const isStacked = chartType === "ColumnStacked" || chartType === "BarPercentStacked";
  const categoryCount = categories.length;
  const seriesCount = chart.series.length;
  const plot = chart.is3d
    ? {
        ...layout.plot,
        height: Math.max(20, layout.plot.height - 12),
        width: Math.max(20, layout.plot.width - 14)
      }
    : layout.plot;

  const matrix = chart.series.map((series) => (
    Array.from({ length: categoryCount }, (_, categoryIndex) => safeNumber(series.values[categoryIndex]) ?? 0)
  ));
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

  if (chartType === "BarPercentStacked") {
    minValue = 0;
    maxValue = 100;
  } else {
    minValue = Math.min(0, minValue);
    maxValue = Math.max(0, maxValue);
    if (typeof chart.valueAxis?.min === "number" && Number.isFinite(chart.valueAxis.min)) {
      minValue = chart.valueAxis.min;
    }
    if (typeof chart.valueAxis?.max === "number" && Number.isFinite(chart.valueAxis.max)) {
      maxValue = chart.valueAxis.max;
    }
  }
  if (maxValue <= minValue) {
    maxValue = minValue + 1;
  }

  const categoryScale = scaleBand<string>()
    .domain(categories)
    .range(isHorizontal ? [plot.top, plot.top + plot.height] : [plot.left, plot.left + plot.width])
    .paddingInner(0.24)
    .paddingOuter(0.14);

  const seriesScale = scaleBand<string>()
    .domain(Array.from({ length: seriesCount }, (_, index) => String(index)))
    .range([0, categoryScale.bandwidth()])
    .paddingInner(0.16)
    .paddingOuter(0.08);

  const valueScale = scaleLinear()
    .domain([minValue, maxValue])
    .range(isHorizontal ? [plot.left, plot.left + plot.width] : [plot.top + plot.height, plot.top]);

  const ticks = buildNumericTickValues(minValue, maxValue, chart.valueAxis?.majorUnit);
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

      const colors = chartSeriesBarColors(chart, seriesIndex, rawValue);
      const category = categories[categoryIndex];
      const categoryStart = categoryScale(category) ?? 0;
      const barThickness = isStacked ? categoryScale.bandwidth() : seriesScale.bandwidth();
      const barOffset = isStacked ? 0 : (seriesScale(String(seriesIndex)) ?? 0);

      if (isHorizontal) {
        const x1 = valueScale(start);
        const x2 = valueScale(end);
        bars.push({
          color: colors.fill,
          height: Math.max(1, barThickness),
          isHorizontal: true,
          key: `bar-${seriesIndex}-${categoryIndex}`,
          left: Math.min(x1, x2),
          top: categoryStart + barOffset,
          value: rawValue,
          width: Math.max(1, Math.abs(x2 - x1))
        });
      } else {
        const y1 = valueScale(start);
        const y2 = valueScale(end);
        bars.push({
          color: colors.fill,
          height: Math.max(1, Math.abs(y2 - y1)),
          isHorizontal: false,
          key: `bar-${seriesIndex}-${categoryIndex}`,
          left: categoryStart + barOffset,
          top: Math.min(y1, y2),
          value: rawValue,
          width: Math.max(1, barThickness)
        });
      }
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
      {axisNode}
      {frameNode}
      {chart.is3d
        ? bars.map((bar) => renderExtrudedRect(bar))
        : bars.map((bar) => (
            <rect
              key={bar.key}
              fill={bar.color}
              height={bar.height}
              width={bar.width}
              x={bar.left}
              y={bar.top}
            />
          ))}
    </g>
  );
}

function renderLineOrAreaChart(chart: XlsxChart, palette: ChartRendererPalette, layout: ChartLayout, chartType: string) {
  const categories = getCategoryLabels(chart);
  if (categories.length === 0 || chart.series.length === 0) {
    return null;
  }
  const plot = layout.plot;

  const allValues = chart.series.flatMap((series) => series.values.map((value) => safeNumber(value)).filter((value): value is number => value != null));
  if (allValues.length === 0) {
    return null;
  }
  let minValue = Math.min(...allValues);
  let maxValue = Math.max(...allValues);
  if (chartType === "Area") {
    minValue = Math.min(0, minValue);
  }
  minValue = typeof chart.valueAxis?.min === "number" && Number.isFinite(chart.valueAxis.min) ? chart.valueAxis.min : minValue;
  maxValue = typeof chart.valueAxis?.max === "number" && Number.isFinite(chart.valueAxis.max) ? chart.valueAxis.max : maxValue;
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
          y: safeNumber(series.values[categoryIndex])
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
  const pointsBySeries = chart.series.map((series) => {
    const yLength = series.values.length;
    const directXValues = Array.from({ length: yLength }, (_, index) => coerceLooseNumber(series.categories[index]));

    let resolvedXValues = directXValues;
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
          resolvedXValues = Array.from({ length: yLength }, (_, index) => index + 1);
        }
      }
    }

    return series.values.map((value, index) => {
      const x = coerceLooseNumber(resolvedXValues[index]);
      const y = safeNumber(value);
      return x == null || y == null ? null : { x, y };
    }).filter((point): point is { x: number; y: number } => point != null);
  });

  const allX = pointsBySeries.flatMap((points) => points.map((point) => point.x));
  const allY = pointsBySeries.flatMap((points) => points.map((point) => point.y));
  if (allX.length === 0 || allY.length === 0) {
    return null;
  }

  const minX = typeof chart.categoryAxis?.min === "number" && Number.isFinite(chart.categoryAxis.min)
    ? chart.categoryAxis.min
    : Math.min(...allX);
  const maxX = typeof chart.categoryAxis?.max === "number" && Number.isFinite(chart.categoryAxis.max)
    ? chart.categoryAxis.max
    : Math.max(...allX);
  const minY = typeof chart.valueAxis?.min === "number" && Number.isFinite(chart.valueAxis.min)
    ? chart.valueAxis.min
    : Math.min(...allY);
  const maxY = typeof chart.valueAxis?.max === "number" && Number.isFinite(chart.valueAxis.max)
    ? chart.valueAxis.max
    : Math.max(...allY);

  const safeMaxX = maxX <= minX ? minX + 1 : maxX;
  const safeMaxY = maxY <= minY ? minY + 1 : maxY;

  const xScale = scaleLinear().domain([minX, safeMaxX]).range([plot.left, plot.left + plot.width]);
  const yScale = scaleLinear().domain([minY, safeMaxY]).range([plot.top + plot.height, plot.top]);
  const xTicks = buildNumericTickValues(minX, safeMaxX, chart.categoryAxis?.majorUnit);
  const yTicks = buildNumericTickValues(minY, safeMaxY, chart.valueAxis?.majorUnit);
  const axisColor = chart.axisLineColor ?? chart.chartAreaBorderColor ?? palette.border;
  const labelColor = chart.axisLabelColor ?? chart.textColor ?? palette.text;

  const curve = smooth ? curveCatmullRom.alpha(0.5) : curveLinear;

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
      <line
        stroke={axisColor}
        strokeWidth={1.2}
        x1={plot.left}
        x2={plot.left + plot.width}
        y1={yScale(0)}
        y2={yScale(0)}
      />
      <line
        stroke={axisColor}
        strokeWidth={1.2}
        x1={xScale(0)}
        x2={xScale(0)}
        y1={plot.top}
        y2={plot.top + plot.height}
      />
      {pointsBySeries.map((points, seriesIndex) => (
        <g key={`scatter-series-${seriesIndex}`}>
          {points.length > 1 ? (
            <path
              d={d3Line<{ x: number; y: number }>()
                .x((point) => xScale(point.x))
                .y((point) => yScale(point.y))
                .curve(curve)(points) ?? ""}
              fill="none"
              stroke={chartSeriesStrokeColor(chart, seriesIndex)}
              strokeWidth={Math.max(1.5, chart.series[seriesIndex]?.lineWidthPx ?? 2)}
            />
          ) : null}
          {points.map((point, pointIndex) => {
            const markerSize = Math.max(5, chart.series[seriesIndex]?.markerSize ?? 7);
            const markerPath = markerSymbolPath(
              normalizeChartMarkerSymbol(chart.series[seriesIndex]?.markerSymbol),
              markerSize * 0.55
            );
            return (
              <g
                key={`scatter-point-${seriesIndex}-${pointIndex}`}
                transform={`translate(${xScale(point.x)}, ${yScale(point.y)})`}
              >
                <path
                  d={markerPath}
                  fill={chart.series[seriesIndex]?.markerColor ?? chartSeriesColor(chart, seriesIndex)}
                  stroke={chart.series[seriesIndex]?.markerLineColor ?? chart.chartAreaFillColor ?? chartSeriesStrokeColor(chart, seriesIndex)}
                  strokeWidth={1}
                />
              </g>
            );
          })}
        </g>
      ))}
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
      return x == null || y == null ? null : { bubble: bubble ?? 1, x, y };
    }).filter((point): point is { bubble: number; x: number; y: number } => point != null)
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
  const minBubble = Math.min(...allBubble);
  const maxBubble = Math.max(...allBubble);

  const xScale = scaleLinear().domain([minX, maxX <= minX ? minX + 1 : maxX]).range([plot.left, plot.left + plot.width]);
  const yScale = scaleLinear().domain([minY, maxY <= minY ? minY + 1 : maxY]).range([plot.top + plot.height, plot.top]);
  const radiusScale = scaleLinear()
    .domain([minBubble, maxBubble <= minBubble ? minBubble + 1 : maxBubble])
    .range([6, 26 + ((chart.bubbleScale ?? 100) * 0.12)]);

  const xTicks = buildNumericTickValues(minX, maxX <= minX ? minX + 1 : maxX, chart.categoryAxis?.majorUnit);
  const yTicks = buildNumericTickValues(minY, maxY <= minY ? minY + 1 : maxY, chart.valueAxis?.majorUnit);
  const axisColor = chart.axisLineColor ?? chart.chartAreaBorderColor ?? palette.border;
  const labelColor = chart.axisLabelColor ?? chart.textColor ?? palette.text;

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
            const baseColor = chartSeriesColor(chart, seriesIndex);
            return (
              <circle
                key={`bubble-${seriesIndex}-${pointIndex}`}
                cx={xScale(point.x)}
                cy={yScale(point.y)}
                fill={lightenColor(baseColor, 0.15)}
                fillOpacity={0.8}
                r={radiusScale(point.bubble)}
                stroke={darkenColor(baseColor, 0.2)}
                strokeWidth={1.2}
              />
            );
          })}
        </g>
      ))}
    </g>
  );
}

function renderRadarChart(chart: XlsxChart, palette: ChartRendererPalette, layout: ChartLayout) {
  const categories = getCategoryLabels(chart);
  if (categories.length === 0 || chart.series.length === 0) {
    return null;
  }
  const plot = layout.plot;
  const centerX = plot.left + plot.width * 0.5;
  const centerY = plot.top + plot.height * 0.52;
  const radius = Math.max(22, Math.min(plot.width, plot.height) * 0.38);
  const values = chart.series.flatMap((series) => series.values.map((value) => safeNumber(value)).filter((value): value is number => value != null));
  const minValue = typeof chart.valueAxis?.min === "number" && Number.isFinite(chart.valueAxis.min)
    ? chart.valueAxis.min
    : Math.min(0, ...(values.length ? values : [0]));
  const maxValue = typeof chart.valueAxis?.max === "number" && Number.isFinite(chart.valueAxis.max)
    ? chart.valueAxis.max
    : Math.max(1, ...(values.length ? values : [1]));
  const safeMax = maxValue <= minValue ? minValue + 1 : maxValue;
  const ticks = buildNumericTickValues(minValue, safeMax, chart.valueAxis?.majorUnit);
  const axisColor = chart.axisLineColor ?? chart.chartAreaBorderColor ?? palette.border;
  const labelColor = chart.axisLabelColor ?? chart.textColor ?? palette.text;
  const filled = chart.radarStyle === "filled";

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
            <line stroke={lightenColor(axisColor, 0.52)} strokeWidth={1} x1={centerX} x2={edge.x} y1={centerY} y2={edge.y} />
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
          const value = safeNumber(series.values[categoryIndex]) ?? minValue;
          const ratio = clamp((value - minValue) / (safeMax - minValue), 0, 1);
          return radialPoint(categoryIndex, ratio);
        });
        const polygon = d3Line<{ x: number; y: number }>()
          .x((point) => point.x)
          .y((point) => point.y)
          .curve(curveLinearClosed)(points) ?? "";
        const color = chartSeriesColor(chart, seriesIndex);
        return (
          <g key={`radar-series-${seriesIndex}`}>
            <path
              d={polygon}
              fill={filled ? color : "none"}
              fillOpacity={filled ? 0.26 : 0}
              stroke={chartSeriesStrokeColor(chart, seriesIndex)}
              strokeWidth={1.8}
            />
            {points.map((point, pointIndex) => (
              <circle
                key={`radar-point-${seriesIndex}-${pointIndex}`}
                cx={point.x}
                cy={point.y}
                fill={color}
                r={2.6}
                stroke={chart.chartAreaFillColor ?? palette.surface}
                strokeWidth={1}
              />
            ))}
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

  const depth = chartType === "Pie3D" ? Math.max(8, outerRadius * 0.18) : 0;
  const dataLabelsEnabled = Boolean(chart.dataLabels?.showCategoryName || chart.dataLabels?.showPercent || chart.dataLabels?.showValue);
  const total = pieData.reduce((sum, entry) => sum + entry.value, 0);

  return (
    <g>
      {chartType === "Pie3D"
        ? arcs.map((arc) => (
            <path
              key={`pie-base-${arc.data.index}`}
              d={(arcPath(arc) ?? "")}
              fill={darkenColor(arc.data.color, 0.28)}
              transform={`translate(${centerX}, ${centerY + depth})`}
            />
          ))
        : null}
      {arcs.map((arc) => {
        const pointStyle = chart.series[0]?.dataPointStyles?.find((entry) => entry.index === arc.data.index);
        const explosion = chartType === "PieExploded" ? Math.max(0, pointStyle?.explosion ?? 0) : 0;
        const midAngle = (arc.startAngle + arc.endAngle) / 2;
        const explodeX = Math.cos(midAngle) * explosion;
        const explodeY = Math.sin(midAngle) * explosion;
        const labelRadius = outerRadius + 12;
        const labelX = centerX + Math.cos(midAngle) * labelRadius;
        const labelY = centerY + Math.sin(midAngle) * labelRadius;
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
              transform={`translate(${centerX}, ${centerY})`}
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
