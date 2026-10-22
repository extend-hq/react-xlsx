import type { XlsxChartAxis } from "./types";

export type SurfaceBandPoint = { x: number; y: number; value: number };

export function surfaceHeightRatio(heightPercent: number | undefined, plot: { width: number; height: number }) {
  return heightPercent != null && Number.isFinite(heightPercent) && heightPercent > 0
    ? heightPercent / 100
    : plot.height / Math.max(1, plot.width);
}

export function scaleSurfaceValue(value: number, axis?: XlsxChartAxis | null) {
  const scaled = axis?.logBase != null && axis.logBase > 1
    ? Math.log(Math.max(Number.MIN_VALUE, value)) / Math.log(axis.logBase)
    : value;
  return axis?.orientation === "maxMin" ? -scaled : scaled;
}

export function normalizeSurfaceValue(value: number, min: number, max: number, axis?: XlsxChartAxis | null) {
  const start = scaleSurfaceValue(min, axis);
  const end = scaleSurfaceValue(max, axis);
  const lower = Math.min(start, end);
  const span = Math.abs(end - start);
  return span > 0 ? (scaleSurfaceValue(value, axis) - lower) / span : 0.5;
}

export function triangulateSurfaceQuad<T extends { value: number }>(corners: [T, T, T, T], axis?: XlsxChartAxis | null): Array<[T, T, T]> {
  const heights = corners.map((point) => -scaleSurfaceValue(point.value, axis));
  return heights[0] * 0.5 + heights[2] * 0.5 > heights[1] * 0.5 + heights[3] * 0.5
    ? [[corners[3], corners[0], corners[1]], [corners[1], corners[2], corners[3]]]
    : [[corners[2], corners[3], corners[0]], [corners[0], corners[1], corners[2]]];
}

export function clipSurfaceBand(points: SurfaceBandPoint[], threshold: number, keepAbove: boolean, axis?: XlsxChartAxis | null, includeBoundary = true) {
  const result: SurfaceBandPoint[] = [];
  points.forEach((end, index) => {
    const start = points[(index + points.length - 1) % points.length];
    const startInside = keepAbove ? start.value >= threshold : start.value <= threshold;
    const endInside = keepAbove ? end.value >= threshold : end.value <= threshold;
    if (startInside !== endInside) {
      const startValue = scaleSurfaceValue(start.value, axis);
      const ratio = (scaleSurfaceValue(threshold, axis) - startValue) / (scaleSurfaceValue(end.value, axis) - startValue);
      result.push({ value: threshold, x: start.x + (end.x - start.x) * ratio, y: start.y + (end.y - start.y) * ratio });
    }
    if (endInside) {
      result.push(end);
    }
  });
  return !includeBoundary && result.every((point) => point.value === threshold) ? [] : result;
}

export function surfaceBandEdges(polygon: SurfaceBandPoint[], triangle: [SurfaceBandPoint, SurfaceBandPoint, SurfaceBandPoint]) {
  const start = triangle[2];
  const end = triangle[0];
  const dx = end.x - start.x;
  const dy = end.y - start.y;
  const tolerance = 1e-9 * Math.max(1, dx * dx + dy * dy);
  const onDiagonal = (point: SurfaceBandPoint) => Math.abs((point.x - start.x) * dy - (point.y - start.y) * dx) <= tolerance;
  return polygon.flatMap((from, index): Array<[SurfaceBandPoint, SurfaceBandPoint]> => {
    const to = polygon[(index + 1) % polygon.length];
    return (onDiagonal(from) && onDiagonal(to)) || (from.x === to.x && from.y === to.y)
      ? []
      : [[from, to]];
  });
}
