import * as React from "react";
import type { ChartSvgProps } from "./chart-renderer-types";
import type { XlsxChart, XlsxChartElementSelection } from "./types";

function resolveChartElementTarget(target: EventTarget | null) {
  if (!(target instanceof Element)) {
    return null;
  }
  if (target.closest("[data-chart-effect]")) {
    return null;
  }

  return target.closest("[data-xlsx-chart-series-index],[data-xlsx-chart-element-kind='legend']") as HTMLElement | SVGElement | null;
}

function resolveChartSelectionFromTarget(
  chart: XlsxChart,
  target: EventTarget | null,
  selectedChartElement: XlsxChartElementSelection | null | undefined
): XlsxChartElementSelection | null {
  const element = resolveChartElementTarget(target);
  if (!element) {
    return { chartId: chart.id, kind: "chart" };
  }
  if (element.dataset.xlsxChartElementKind === "legend") {
    return { chartId: chart.id, kind: "legend" };
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
  if (rawElementKind === "legendEntry" && !(selectedChartElement?.chartId === chart.id
    && (selectedChartElement.kind === "legend" || selectedChartElement.kind === "legendEntry"))) {
    return { chartId: chart.id, kind: "legend" };
  }
  const pointIndex = rawPointIndex == null || rawPointIndex === "" ? null : Number(rawPointIndex);
  const hasPoint = pointIndex != null && Number.isInteger(pointIndex) && pointIndex >= 0;
  const sameSelectedSeries = selectedChartElement?.chartId === chart.id
    && selectedChartElement.kind !== "chart"
    && selectedChartElement.kind !== "legend"
    && selectedChartElement.seriesIndex === seriesIndex;
  if (hasPoint && rawElementKind !== "legendEntry" && sameSelectedSeries) {
    return {
      chartId: chart.id,
      kind: "point",
      pointIndex,
      seriesId: series.id,
      seriesIndex
    };
  }
  if (rawElementKind === "legendEntry") {
    const rawLegendIndex = element.dataset.xlsxChartLegendIndex;
    const legendIndex = rawLegendIndex == null || rawLegendIndex === "" ? undefined : Number(rawLegendIndex);
    return {
      chartId: chart.id,
      kind: "legendEntry",
      seriesId: series.id,
      seriesIndex,
      ...(legendIndex != null && Number.isInteger(legendIndex) && legendIndex >= 0 ? { legendIndex } : {})
    };
  }

  return {
    chartId: chart.id,
    kind: "series",
    seriesId: series.id,
    seriesIndex
  };
}

export function useChartEvents({
  chart, onChartElementPointerDown, onChartElementDoubleClick, selectedChartElement
}: Pick<ChartSvgProps, "chart" | "onChartElementPointerDown" | "onChartElementDoubleClick" | "selectedChartElement">) {
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

  return { handlePointerDown, handleDoubleClick };
}
