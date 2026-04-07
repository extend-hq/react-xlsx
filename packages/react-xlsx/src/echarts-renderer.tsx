import * as React from "react";
import * as echarts from "echarts";
import "echarts-gl";
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

function buildLegend(chart: XlsxChart) {
  if (!chart.legend) {
    return undefined;
  }
  return {
    itemHeight: 10,
    itemWidth: 10,
    orient: chart.legend.position === "l" || chart.legend.position === "r" || chart.legend.position === "left" || chart.legend.position === "right"
      ? "vertical"
      : "horizontal",
    right: normalizeLegendPosition(chart.legend.position) === "right" ? 8 : undefined,
    left: normalizeLegendPosition(chart.legend.position) === "left" ? 8 : undefined,
    top: normalizeLegendPosition(chart.legend.position) === "top" ? 8 : undefined,
    bottom: normalizeLegendPosition(chart.legend.position) === "bottom" ? 8 : undefined,
    textStyle: {
      color: chart.axisLabelColor ?? chart.textColor
    }
  };
}

function buildTitle(chart: XlsxChart) {
  if (!chart.title) {
    return undefined;
  }
  return {
    left: "center",
    text: chart.title,
    textStyle: {
      color: chart.titleColor ?? chart.textColor,
      fontSize: 12,
      fontWeight: 600
    },
    top: 8
  };
}

function buildCommonOption(chart: XlsxChart, palette: ChartRendererPalette) {
  return {
    animation: false,
    backgroundColor: chart.chartAreaFillColor ?? palette.surface,
    legend: buildLegend(chart),
    textStyle: {
      color: chart.textColor ?? palette.text,
      fontFamily: "inherit"
    },
    title: buildTitle(chart),
    tooltip: {
      show: false,
      trigger: "item"
    }
  } as const;
}

function buildAxisStyle(color: string | undefined, border: string) {
  const stroke = color ?? border;
  return {
    axisLabel: {
      color: stroke,
      fontSize: 10
    },
    axisLine: {
      lineStyle: {
        color: stroke
      }
    },
    axisTick: {
      lineStyle: {
        color: stroke
      }
    },
    splitLine: {
      lineStyle: {
        color: lightenColor(stroke, 0.74),
        type: "solid"
      }
    }
  };
}

function normalizeRenderableChartType(chart: XlsxChart) {
  if (chart.chartType === "ScatterSmooth") {
    return "ScatterSmooth";
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
  return chart.chartType;
}

function buildCartesianOption(chart: XlsxChart, palette: ChartRendererPalette) {
  const renderChartType = normalizeRenderableChartType(chart);
  const categories = chart.series[0]?.categories?.map((value) => (value == null ? "" : String(value))) ?? [];
  const axisColor = chart.axisLineColor ?? chart.chartAreaBorderColor ?? palette.border;
  const labelColor = chart.axisLabelColor ?? chart.textColor ?? palette.text;
  const common = buildCommonOption(chart, palette);
  const seriesCount = Math.max(1, chart.series.length);
  const overlapRatio = Math.max(-1, Math.min(1, (chart.overlap ?? 0) / 100));
  const gapRatio = Math.max(0, (chart.gapWidth ?? 150) / 100);
  const clusterWidthFactor = 1 + (seriesCount - 1) * (1 - overlapRatio);
  const barWidthRatio = 1 / Math.max(1, clusterWidthFactor + gapRatio);
  const barCategoryGapRatio = gapRatio / Math.max(1, clusterWidthFactor + gapRatio);
  const grid = {
    bottom: 30,
    containLabel: true,
    left: 32,
    right: normalizeLegendPosition(chart.legend?.position) === "right" ? 96 : 20,
    top: chart.title ? 34 : 16
  };

  if (renderChartType === "ScatterLines" || renderChartType === "ScatterSmooth") {
    const series = chart.series.map((entry, seriesIndex) => {
      const rawX = entry.values.map((_, index) => (
        typeof entry.categories[index] === "number" ? entry.categories[index] as number : null
      ));
      const finiteX = rawX.filter((value): value is number => typeof value === "number" && Number.isFinite(value));
      const uniqueX = new Set(finiteX.map((value) => value.toFixed(6)));
      const useIndexed = finiteX.length !== entry.values.length || uniqueX.size <= 1;
      const data = entry.values.flatMap((value, index) => {
        if (typeof value !== "number") {
          return [];
        }
        return [[useIndexed ? index + 1 : (rawX[index] ?? index + 1), value]];
      });
      const lineColor = chartSeriesStrokeColor(chart, seriesIndex);
      const markerColor = entry.markerColor ?? chartSeriesColor(chart, seriesIndex);
      return [{
        data,
        itemStyle: {
          borderColor: entry.markerLineColor ?? chart.chartAreaFillColor ?? lineColor,
          borderWidth: 1,
          color: markerColor
        },
        lineStyle: {
          color: lineColor,
          opacity: renderChartType === "ScatterSmooth" ? 1 : 0,
          width: entry.lineWidthPx ?? 2
        },
        showSymbol: true,
        smooth: renderChartType === "ScatterSmooth" || entry.smooth === true,
        symbol: normalizeChartMarkerSymbol(entry.markerSymbol),
        symbolSize: Math.max(6, entry.markerSize ?? 5),
        type: renderChartType === "ScatterSmooth" ? "line" : "scatter"
      }];
    }).flat();

    return {
      ...common,
      grid,
      series,
      xAxis: {
        ...buildAxisStyle(labelColor, axisColor),
        max: chart.categoryAxis?.max,
        min: chart.categoryAxis?.min,
        scale: true,
        type: "value"
      },
      yAxis: {
        ...buildAxisStyle(labelColor, axisColor),
        max: chart.valueAxis?.max,
        min: chart.valueAxis?.min,
        type: "value"
      }
    };
  }

  if (renderChartType === "Radar") {
    const allValues = chart.series.flatMap((entry) => entry.values.filter((value): value is number => typeof value === "number"));
    const maxValue = chart.valueAxis?.max ?? Math.max(1, ...allValues);
    return {
      ...common,
      legend: {
        ...buildLegend(chart),
        right: normalizeLegendPosition(chart.legend?.position) === "right" ? 10 : undefined
      },
      radar: {
        axisLine: {
          lineStyle: {
            color: axisColor
          }
        },
        axisName: {
          color: labelColor
        },
        indicator: categories.map((name) => ({
          max: maxValue,
          min: chart.valueAxis?.min ?? 0,
          name
        })),
        shape: "polygon",
        splitLine: {
          lineStyle: {
            color: lightenColor(axisColor, 0.74)
          }
        },
        splitNumber: typeof chart.valueAxis?.majorUnit === "number" && chart.valueAxis.majorUnit > 0
          ? Math.max(1, Math.round(maxValue / chart.valueAxis.majorUnit))
          : 5,
        splitArea: {
          show: false
        }
      },
      series: [{
        data: chart.series.map((entry, index) => ({
          areaStyle: entry.shapeProperties?.xmlFillColor ? { color: entry.shapeProperties.xmlFillColor, opacity: 0.18 } : undefined,
          itemStyle: {
            color: chartSeriesColor(chart, index)
          },
          lineStyle: {
            color: chartSeriesStrokeColor(chart, index),
            width: entry.lineWidthPx ?? 2
          },
          name: entry.name ?? `Series ${index + 1}`,
          symbol: normalizeChartMarkerSymbol(entry.markerSymbol),
          symbolSize: entry.markerSize ?? 6,
          value: entry.values.map((value) => value ?? 0)
        })),
        type: "radar"
      }]
    };
  }

  const isHorizontal = renderChartType === "BarClustered" || renderChartType === "BarPercentStacked";
  const xAxis = isHorizontal
    ? {
        ...buildAxisStyle(labelColor, axisColor),
        max: renderChartType === "BarPercentStacked" ? 100 : chart.valueAxis?.max,
        min: renderChartType === "BarPercentStacked" ? 0 : chart.valueAxis?.min,
        type: "value"
      }
    : {
        ...buildAxisStyle(labelColor, axisColor),
        data: categories,
        type: "category"
      };
  const yAxis = isHorizontal
    ? {
        ...buildAxisStyle(labelColor, axisColor),
        data: categories,
        type: "category"
      }
    : {
        ...buildAxisStyle(labelColor, axisColor),
        max: chart.valueAxis?.max,
        min: chart.valueAxis?.min,
        type: "value"
      };

  const normalizePercentData = (categoryIndex: number) => {
    const total = chart.series.reduce((sum, entry) => sum + Math.max(0, typeof entry.values[categoryIndex] === "number" ? entry.values[categoryIndex] as number : 0), 0);
    return chart.series.map((entry) => {
      const value = entry.values[categoryIndex];
      return typeof value === "number" && total > 0 ? (Math.max(0, value) / total) * 100 : 0;
    });
  };

  const barSeries = chart.series.map((entry, seriesIndex) => {
    const baseColor = chartSeriesColor(chart, seriesIndex);
    const data = entry.values.map((rawValue, categoryIndex) => {
      const value = renderChartType === "BarPercentStacked"
        ? normalizePercentData(categoryIndex)[seriesIndex] ?? 0
        : (rawValue ?? 0);
      const colors = chartSeriesBarColors(chart, seriesIndex, rawValue ?? 0);
      const gradient = chart.is3d
        ? new echarts.graphic.LinearGradient(
            isHorizontal ? 0 : 0,
            isHorizontal ? 0 : 0,
            isHorizontal ? 1 : 0,
            isHorizontal ? 0 : 1,
            [
              { color: lightenColor(colors.fill, 0.18), offset: 0 },
              { color: colors.fill, offset: 0.6 },
              { color: darkenColor(colors.fill, 0.18), offset: 1 }
            ]
          )
        : colors.fill;
      return {
        itemStyle: {
          borderColor: colors.stroke,
          borderWidth: 1,
          color: gradient,
          opacity: 1,
          shadowBlur: chart.is3d ? 10 : 0,
          shadowColor: chart.is3d ? "rgba(0,0,0,0.18)" : undefined,
          shadowOffsetX: chart.is3d && isHorizontal ? 4 : 0,
          shadowOffsetY: chart.is3d && !isHorizontal ? 4 : 0
        },
        value
      };
    });
    return {
      barCategoryGap: `${Math.round(barCategoryGapRatio * 100)}%`,
      barGap: `${Math.round(-overlapRatio * 100)}%`,
      barMinHeight: 1,
      barWidth: `${Math.max(6, Math.round(barWidthRatio * 100))}%`,
      data,
      emphasis: {
        disabled: true
      },
      name: entry.name ?? `Series ${seriesIndex + 1}`,
      stack: renderChartType === "ColumnStacked" || renderChartType === "BarPercentStacked" ? "stack" : undefined,
      type: "bar"
    };
  });

  if (renderChartType === "Line") {
    return {
      ...common,
      grid,
      series: chart.series.map((entry, seriesIndex) => ({
        areaStyle: undefined,
        data: entry.values.map((value, index) => [categories[index] ?? index + 1, value]),
        itemStyle: {
          color: entry.markerColor ?? chartSeriesColor(chart, seriesIndex)
        },
        lineStyle: {
          color: chartSeriesStrokeColor(chart, seriesIndex),
          width: entry.lineWidthPx ?? 2
        },
        name: entry.name ?? `Series ${seriesIndex + 1}`,
        showSymbol: normalizeChartMarkerSymbol(entry.markerSymbol) !== "none",
        smooth: entry.smooth === true,
        symbol: normalizeChartMarkerSymbol(entry.markerSymbol),
        symbolSize: entry.markerSize ?? 5,
        type: "line"
      })),
      xAxis,
      yAxis
    };
  }

  if (renderChartType === "Area") {
    return {
      ...common,
      grid,
      series: chart.series.map((entry, seriesIndex) => ({
        areaStyle: {
          color: entry.shapeProperties?.xmlFillColor ?? chartSeriesColor(chart, seriesIndex),
          opacity: 1
        },
        data: entry.values.map((value, index) => [categories[index] ?? index + 1, value]),
        itemStyle: {
          color: chartSeriesColor(chart, seriesIndex)
        },
        lineStyle: {
          color: chartSeriesStrokeColor(chart, seriesIndex),
          opacity: entry.shapeProperties?.lineNoFill === true ? 0 : 1,
          width: entry.lineWidthPx ?? 2
        },
        name: entry.name ?? `Series ${seriesIndex + 1}`,
        showSymbol: normalizeChartMarkerSymbol(entry.markerSymbol) !== "none",
        symbol: normalizeChartMarkerSymbol(entry.markerSymbol),
        symbolSize: entry.markerSize ?? 5,
        type: "line"
      })),
      xAxis,
      yAxis
    };
  }

  return {
    ...common,
    grid,
    series: barSeries,
    xAxis,
    yAxis
  };
}

function buildPieOption(chart: XlsxChart, palette: ChartRendererPalette) {
  const renderChartType = normalizeRenderableChartType(chart);
  const categories = chart.series[0]?.categories ?? [];
  const values = chart.series[0]?.values ?? [];
  const data = values.map((value, index) => {
    const pointStyle = chart.series[0]?.dataPointStyles?.find((entry) => entry.index === index);
    return {
      itemStyle: {
        borderColor: pointStyle?.lineColor ?? chart.chartAreaFillColor ?? chart.chartAreaBorderColor ?? chartPointColor(chart, index),
        borderWidth: 1,
        color: chartPointColor(chart, index)
      },
      name: categories[index] == null ? "" : String(categories[index]),
      selected: renderChartType === "PieExploded" && ((pointStyle?.explosion ?? 0) > 0),
      selectedOffset: renderChartType === "PieExploded" ? Math.max(8, pointStyle?.explosion ?? 0) : 0,
      value: Math.max(0, value ?? 0)
    };
  });
  const common = buildCommonOption(chart, palette);
  const isDonut = renderChartType === "Doughnut";
  const is3d = renderChartType === "Pie3D";
  const inner = isDonut ? `${Math.max(0, Math.min(90, chart.holeSize ?? 56)) * 0.7}%` : "0%";
  const outer = isDonut ? "70%" : "76%";
  const labelEnabled = Boolean(chart.dataLabels?.showValue || chart.dataLabels?.showPercent || chart.dataLabels?.showCategoryName);
  const labelFormatterParts = [
    chart.dataLabels?.showCategoryName ? "{b}" : "",
    chart.dataLabels?.showValue ? "{c}" : "",
    chart.dataLabels?.showPercent ? "{d}%" : ""
  ].filter(Boolean);
  const topSeries = {
    avoidLabelOverlap: true,
    center: ["44%", "54%"],
    data,
    label: {
      color: chart.textColor ?? palette.text,
      formatter: labelEnabled ? labelFormatterParts.join(", ") : undefined,
      show: labelEnabled
    },
    labelLine: {
      show: labelEnabled
    },
    radius: [inner, outer],
    roseType: false,
    selectedMode: renderChartType === "PieExploded" ? "multiple" : false,
    startAngle: 90 - (chart.firstSliceAngle ?? 0),
    type: "pie"
  };
  if (!is3d) {
    return {
      ...common,
      series: [topSeries]
    };
  }

  const layeredSeries = Array.from({ length: 10 }, (_, depthIndex) => ({
    ...topSeries,
    animation: false,
    center: ["44%", `${54 + depthIndex * 0.7}%`],
    data: data.map((entry) => ({
      ...entry,
      itemStyle: {
        ...(entry.itemStyle ?? {}),
        borderWidth: 0,
        color: darkenColor(String(entry.itemStyle?.color ?? "#888888"), 0.22)
      }
    })),
    label: { show: false },
    labelLine: { show: false },
    silent: true,
    z: depthIndex
  }));

  return {
    ...common,
    series: [...layeredSeries, { ...topSeries, z: 20 }]
  };
}

function buildBarOfPieOption(chart: XlsxChart, palette: ChartRendererPalette) {
  const categories = chart.series[0]?.categories ?? [];
  const values = chart.series[0]?.values.map((value) => Math.max(0, value ?? 0)) ?? [];
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
    : [{
        itemStyle: { color: chartPointColor(chart, index) },
        name: categories[index] == null ? "" : String(categories[index]),
        value
      }]);
  if (secondaryTotal > 0) {
    primaryData.push({
      itemStyle: { color: lightenColor(chartPointColor(chart, secondaryIndices[0] ?? 0), 0.2) },
      name: "Other",
      value: secondaryTotal
    });
  }
  const common = buildCommonOption(chart, palette);
  return {
    ...common,
    grid: {
      bottom: 36,
      containLabel: true,
      left: "64%",
      right: 20,
      top: chart.title ? 34 : 16
    },
    graphic: secondaryIndices.length > 0 ? [
      {
        shape: { x1: "47%", x2: "63%", y1: "42%", y2: "34%" },
        style: { stroke: chart.chartAreaBorderColor ?? palette.border },
        type: "line"
      },
      {
        shape: { x1: "47%", x2: "63%", y1: "58%", y2: "70%" },
        style: { stroke: chart.chartAreaBorderColor ?? palette.border },
        type: "line"
      }
    ] : undefined,
    series: [
      {
        center: ["30%", "52%"],
        data: primaryData,
        label: { show: false },
        radius: "38%",
        startAngle: 90 - (chart.firstSliceAngle ?? 0),
        type: "pie"
      },
      {
        data: secondaryIndices.map((index) => ({
          itemStyle: { color: chartPointColor(chart, index) },
          name: categories[index] == null ? "" : String(categories[index]),
          value: values[index] ?? 0
        })),
        type: "bar",
        xAxisIndex: 0,
        yAxisIndex: 0
      }
    ],
    xAxis: {
      axisLabel: { show: false },
      axisLine: { show: false },
      axisTick: { show: false },
      splitLine: { show: false },
      type: "value"
    },
    yAxis: {
      axisLabel: { show: false },
      axisLine: { show: false },
      axisTick: { show: false },
      data: secondaryIndices.map((index) => categories[index] == null ? "" : String(categories[index])),
      type: "category"
    }
  };
}

function buildSurfaceOption(chart: XlsxChart, palette: ChartRendererPalette) {
  const categories = chart.series[0]?.categories?.map((value) => (value == null ? "" : String(value))) ?? [];
  const rowNames = chart.series.map((series, index) => series.name ?? `Series ${index + 1}`);
  const allValues = chart.series.flatMap((series) => series.values.filter((value): value is number => typeof value === "number"));
  const minValue = Math.min(...allValues);
  const maxValue = Math.max(...allValues);
  const paletteColors = chart.chartColorPalette && chart.chartColorPalette.length > 0
    ? chart.chartColorPalette
    : chart.series.map((_, index) => chartSeriesColor(chart, index));
  const data = chart.series.flatMap((series, rowIndex) => series.values.flatMap((value, columnIndex) => (
    typeof value === "number" ? [[columnIndex, rowIndex, value]] : []
  )));
  return {
    ...buildCommonOption(chart, palette),
    grid3D: {
      axisLine: {
        lineStyle: {
          color: chart.axisLineColor ?? palette.border
        }
      },
      axisPointer: {
        show: false
      },
      boxDepth: 80,
      boxHeight: 40,
      boxWidth: 120,
      environment: chart.chartAreaFillColor ?? palette.surface,
      light: {
        ambient: {
          intensity: 0.8
        },
        main: {
          intensity: 0.9
        }
      },
      viewControl: {
        alpha: chart.view3d?.rotX ?? 28,
        beta: chart.view3d?.rotY ?? 18,
        panSensitivity: 0,
        projection: chart.wireframe ? "orthographic" : "perspective",
        rotateSensitivity: 0,
        zoomSensitivity: 0
      }
    },
    visualMap: chart.wireframe ? undefined : {
      calculable: false,
      dimension: 2,
      inRange: {
        color: paletteColors
      },
      max: maxValue,
      min: minValue,
      show: false
    },
    xAxis3D: {
      axisLabel: { color: chart.axisLabelColor ?? palette.text },
      data: categories,
      type: "category"
    },
    yAxis3D: {
      axisLabel: { color: chart.axisLabelColor ?? palette.text },
      data: rowNames,
      type: "category"
    },
    zAxis3D: {
      axisLabel: { color: chart.axisLabelColor ?? palette.text },
      max: chart.valueAxis?.max,
      min: chart.valueAxis?.min,
      type: "value"
    },
    series: [{
      data,
      shading: chart.wireframe ? "color" : "lambert",
      type: "surface",
      wireframe: {
        lineStyle: {
          color: chart.axisLineColor ?? palette.border,
          width: chart.wireframe ? 1.2 : 0.4
        },
        show: chart.wireframe !== false
      }
    }]
  };
}

function buildOption(chart: XlsxChart, palette: ChartRendererPalette) {
  const type = normalizeRenderableChartType(chart);
  if (type === "Pie" || type === "PieExploded" || type === "Doughnut" || type === "Pie3D") {
    return buildPieOption(chart, palette);
  }
  if (type === "BarOfPie") {
    return buildBarOfPieOption(chart, palette);
  }
  if (type === "Surface") {
    return buildSurfaceOption(chart, palette);
  }
  return buildCartesianOption(chart, palette);
}

export const MemoChartSvg = React.memo(function MemoChartSvg({ chart, palette, rect }: ChartSvgProps) {
  const containerRef = React.useRef<HTMLDivElement | null>(null);
  const chartRef = React.useRef<echarts.EChartsType | null>(null);
  const option = React.useMemo(() => buildOption(chart, palette), [chart, palette]);
  const useCanvas = chart.is3d || chart.chartType === "Surface" || chart.chartType === "Pie3D";
  const renderer = useCanvas ? "canvas" : "svg";

  React.useLayoutEffect(() => {
    const element = containerRef.current;
    if (!element) {
      return;
    }

    const instance = echarts.init(element, undefined, {
      height: Math.max(60, Math.round(rect.height)),
      renderer,
      width: Math.max(80, Math.round(rect.width))
    });
    chartRef.current = instance;
    instance.setOption(option, true);

    return () => {
      chartRef.current?.dispose();
      chartRef.current = null;
    };
  }, [renderer]);

  React.useEffect(() => {
    const instance = chartRef.current;
    if (!instance) {
      return;
    }

    instance.resize({
      height: Math.max(60, Math.round(rect.height)),
      width: Math.max(80, Math.round(rect.width))
    });
    instance.setOption(option, true);
  }, [option, rect.height, rect.width]);

  return <div ref={containerRef} style={{ height: "100%", pointerEvents: "none", width: "100%" }} />;
}, (prev, next) => (
  prev.chart === next.chart
  && prev.palette === next.palette
  && prev.rect.height === next.rect.height
  && prev.rect.width === next.rect.width
  && prev.rect.left === next.rect.left
  && prev.rect.top === next.rect.top
));
