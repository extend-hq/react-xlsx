import * as React from "react";
import type { ChartSvgProps } from "./chart-renderer";
import type { XlsxChart, XlsxChartElementSelection } from "./types";

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
    return { chartId: chart.id, kind: "chart" };
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

export function renderSelectionRectHandles(
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

export function renderSelectionPointHandles(
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
