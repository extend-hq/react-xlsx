import createREGL from "regl";
import * as React from "react";
import type { XlsxChart } from "./types";

type SurfacePalette = {
  border: string;
  mutedText: string;
  surface: string;
  text: string;
};

type SurfaceLayout = {
  height: number;
  plot: {
    height: number;
    left: number;
    top: number;
    width: number;
  };
  width: number;
};

type SurfaceCompositeProps = {
  background: string;
  borderColor: string;
  chart: XlsxChart;
  fontFamily: string;
  layout: SurfaceLayout;
  overlay: React.ReactNode;
  palette: SurfacePalette;
  fallback: React.ReactNode;
};

type SurfaceDomain = {
  maxValue: number;
  minValue: number;
  safeMax: number;
  ticks: number[];
};

type SurfacePoint3D = {
  depth: number;
  worldX: number;
  worldY: number;
  worldZ: number;
  x: number;
  y: number;
};

type SurfaceVertex = {
  bandColor: string;
  point: SurfacePoint3D;
  value: number;
};

type SolidSurfaceTriangle = {
  color: string;
  points: [SurfacePoint3D, SurfacePoint3D, SurfacePoint3D];
};

type SurfaceScene = {
  fillColors: Float32Array;
  fillPositions: Float32Array;
  lineColors: Float32Array;
  linePositions: Float32Array;
  wallColors: Float32Array;
  wallPositions: Float32Array;
};

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
  const clamped = clamp(ratio, 0, 1);
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

function resolveSurfaceBaseColor(chart: XlsxChart, palette: SurfacePalette) {
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

function normalizeSurfaceRendererHexColor(value: unknown) {
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

function hasExplicitSurfaceBaseColor(chart: XlsxChart) {
  const primarySeriesColor = normalizeSurfaceRendererHexColor(chart.series[0]?.color ?? chart.series[0]?.lineColor);
  if (!primarySeriesColor) {
    return null;
  }
  const paletteColor = normalizeSurfaceRendererHexColor(chart.chartColorPalette?.[0]);
  return paletteColor && paletteColor === primarySeriesColor ? null : primarySeriesColor;
}

function buildMonochromeSurfacePalette(baseColor: string, count: number) {
  if (count <= 3) {
    return [
      lightenColor(baseColor, 0.22),
      baseColor,
      darkenColor(baseColor, 0.2)
    ];
  }
  return [
    lightenColor(baseColor, 0.3),
    lightenColor(baseColor, 0.14),
    baseColor,
    darkenColor(baseColor, 0.1),
    darkenColor(baseColor, 0.22)
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
    return ["#2f5597", "#4472c4", "#5b9bd5", "#8faadc", "#d9e2f3"];
  }
  return null;
}

function shouldPreferBuiltinSurfacePalette(chart: XlsxChart) {
  const normalized = normalizeBuiltinSurfaceStyleId(chart.chartStyleId);
  const rawChartType = chart.raw && typeof chart.raw === "object" && typeof (chart.raw as Record<string, unknown>).xmlChartType === "string"
    ? String((chart.raw as Record<string, unknown>).xmlChartType)
    : "";
  return (
    (rawChartType === "surfaceChart" || rawChartType === "surface3DChart")
    && (normalized === 26 || normalized === 34 || normalized === 35 || normalized === 36)
  );
}

function getSurfaceBandCount(chart: XlsxChart) {
  const raw = chart.raw && typeof chart.raw === "object" ? chart.raw as Record<string, unknown> : null;
  const explicitBandCount = typeof raw?.bandFormatCount === "number" && Number.isFinite(raw.bandFormatCount)
    ? raw.bandFormatCount
    : null;
  if (explicitBandCount != null && explicitBandCount > 0 && !(isContourSurfaceChart(chart) && chart.wireframe !== true)) {
    return explicitBandCount;
  }
  const builtinPalette = getBuiltinSurfacePalette(chart);
  if (shouldPreferBuiltinSurfacePalette(chart) && builtinPalette && builtinPalette.length > 0) {
    return builtinPalette.length;
  }
  if (chart.chartColorPalette && chart.chartColorPalette.length > 1) {
    return chart.chartColorPalette.length;
  }
  if (builtinPalette && builtinPalette.length > 0) {
    return builtinPalette.length;
  }
  return chart.wireframe ? 3 : 5;
}

function getSurfaceColorStops(chart: XlsxChart, palette: SurfacePalette) {
  const builtinPalette = getBuiltinSurfacePalette(chart);
  if (shouldPreferBuiltinSurfacePalette(chart) && builtinPalette && builtinPalette.length >= 2) {
    return builtinPalette;
  }
  const explicitStops = (chart.chartColorPalette ?? []).filter((value): value is string => typeof value === "string" && value.length > 0);
  if (explicitStops.length >= 2) {
    return explicitStops;
  }
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
  const magnitude = Math.pow(10, Math.floor(Math.log10(Math.max(roughStep, 1e-6))));
  const normalized = roughStep / magnitude;
  const niceStep = normalized <= 1 ? 1 : normalized <= 2 ? 2 : normalized <= 5 ? 5 : 10;
  const step = typeof chart.valueAxis?.majorUnit === "number" && chart.valueAxis.majorUnit > 0
    ? chart.valueAxis.majorUnit
    : niceStep * magnitude;
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
    ticks: [...new Set(ticks)].sort((left, right) => left - right)
  };
}

function resolveSurfaceColor(chart: XlsxChart, palette: SurfacePalette, ratio: number) {
  const stops = getSurfaceColorStops(chart, palette);
  if (stops.length === 0) {
    return resolveSurfaceBaseColor(chart, palette);
  }
  if (stops.length === 1) {
    return stops[0];
  }
  const clampedRatio = clamp(ratio, 0, 1) * (stops.length - 1);
  const lowerIndex = Math.floor(clampedRatio);
  const upperIndex = Math.min(stops.length - 1, lowerIndex + 1);
  const mixRatio = clampedRatio - lowerIndex;
  return mixRgbColor(stops[lowerIndex] ?? stops[0], stops[upperIndex] ?? stops[stops.length - 1], mixRatio);
}

function resolveSurfaceBandPaletteColor(chart: XlsxChart, palette: SurfacePalette, domain: SurfaceDomain, value: number) {
  const stops = getSurfaceColorStops(chart, palette);
  if (stops.length === 0) {
    return resolveSurfaceBaseColor(chart, palette);
  }
  const ticks = domain.ticks;
  for (let index = 0; index < ticks.length - 1; index += 1) {
    const end = ticks[index + 1] ?? domain.safeMax;
    if (value <= end || index === ticks.length - 2) {
      return stops[Math.min(stops.length - 1, Math.max(0, index))] ?? stops[stops.length - 1] ?? resolveSurfaceBaseColor(chart, palette);
    }
  }
  return stops[stops.length - 1] ?? resolveSurfaceBaseColor(chart, palette);
}

function resolveSurfaceBandColor(chart: XlsxChart, palette: SurfacePalette, domain: SurfaceDomain, value: number) {
  return resolveSurfaceBandPaletteColor(chart, palette, domain, value);
}

function getSurfaceWireframePalette(chart: XlsxChart, palette: SurfacePalette) {
  const raw = chart.raw && typeof chart.raw === "object" ? chart.raw as Record<string, unknown> : null;
  const explicit = Array.isArray(raw?.bandFormatLineColors)
    ? raw.bandFormatLineColors.filter((color): color is string => typeof color === "string" && color.length > 0)
    : [];
  if (explicit.length > 0) {
    return explicit;
  }
  return getSurfaceColorStops(chart, palette);
}

function resolveSurfaceWireframeColor(chart: XlsxChart, palette: SurfacePalette, domain: SurfaceDomain, value: number) {
  const stops = getSurfaceWireframePalette(chart, palette);
  if (stops.length === 0) {
    return resolveSurfaceBandPaletteColor(chart, palette, domain, value);
  }
  const ticks = domain.ticks;
  for (let index = 0; index < ticks.length - 1; index += 1) {
    const end = ticks[index + 1] ?? domain.safeMax;
    if (value <= end || index === ticks.length - 2) {
      return stops[Math.min(stops.length - 1, Math.max(0, index))] ?? resolveSurfaceBandPaletteColor(chart, palette, domain, value);
    }
  }
  return stops[stops.length - 1] ?? resolveSurfaceBandPaletteColor(chart, palette, domain, value);
}

function colorToRgba(color: string, alpha = 1) {
  const rgb = parseRgbColor(color);
  if (!rgb) {
    return [0.1, 0.1, 0.1, alpha] as const;
  }
  return [rgb.red / 255, rgb.green / 255, rgb.blue / 255, alpha] as const;
}

function shadeSurfaceColor(color: string, intensity: number) {
  const shaded = intensity >= 1
    ? lightenColor(color, clamp((intensity - 1) * 0.4, 0, 0.22))
    : darkenColor(color, clamp((1 - intensity) * 0.5, 0, 0.28));
  return colorToRgba(shaded, 1);
}

function bilinearInterpolate(p00: number, p10: number, p01: number, p11: number, u: number, v: number) {
  return (
    p00 * (1 - u) * (1 - v)
    + p10 * u * (1 - v)
    + p01 * (1 - u) * v
    + p11 * u * v
  );
}

function normalizeSurfaceX(column: number, cols: number) {
  return cols <= 1 ? 0 : ((column / (cols - 1)) - 0.5) * 2;
}

function normalizeSurfaceY(row: number, rows: number) {
  return rows <= 1 ? 0 : ((row / (rows - 1)) - 0.5) * 2;
}

function normalizeSurfaceZ(value: number, domain: SurfaceDomain, depthScale: number) {
  return -((((value - domain.minValue) / (domain.safeMax - domain.minValue)) - 0.5) * 1.8) * depthScale;
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
    ? 1 / Math.max(0.18, 1 + z2 * (0.24 + perspectiveStrength * 0.5))
    : 1;
  return {
    depth: z2,
    x: x1 * perspective,
    y: y1 * perspective
  };
}

function computeNormal(left: SurfacePoint3D, middle: SurfacePoint3D, right: SurfacePoint3D) {
  const ux = middle.worldX - left.worldX;
  const uy = middle.worldY - left.worldY;
  const uz = middle.worldZ - left.worldZ;
  const vx = right.worldX - left.worldX;
  const vy = right.worldY - left.worldY;
  const vz = right.worldZ - left.worldZ;
  const nx = uy * vz - uz * vy;
  const ny = uz * vx - ux * vz;
  const nz = ux * vy - uy * vx;
  const length = Math.hypot(nx, ny, nz) || 1;
  return {
    x: nx / length,
    y: ny / length,
    z: nz / length
  };
}

function withProjectedDepth(point: SurfacePoint3D, depth: number): SurfacePoint3D {
  return {
    ...point,
    depth
  };
}

function nudgeProjectedDepth(point: SurfacePoint3D, delta: number): SurfacePoint3D {
  return {
    ...point,
    depth: point.depth + delta
  };
}

function isContourSurfaceChart(chart: XlsxChart) {
  const rawChartType = chart.raw && typeof chart.raw === "object" && typeof chart.raw.xmlChartType === "string"
    ? chart.raw.xmlChartType
    : "";
  if (rawChartType === "surfaceChart") {
    return true;
  }
  if (rawChartType === "surface3DChart") {
    return false;
  }
  return chart.chartType === "Surface" && chart.is3d !== true;
}

function buildSurfaceMesh(chart: XlsxChart, palette: SurfacePalette, layout: SurfaceLayout): SurfaceScene | null {
  const rows = chart.series.length;
  const cols = chart.series.reduce((max, series) => Math.max(max, series.values.length), 0);
  if (rows < 2 || cols < 2) {
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
  const isWireframe = chart.wireframe === true;
  const fillVertices: SurfaceVertex[] = [];
  const wallTriangles: SolidSurfaceTriangle[] = [];
  const contourLines: Array<{ color: string; from: SurfacePoint3D; to: SurfacePoint3D }> = [];
  const meshLines: Array<{ color: string; from: SurfacePoint3D; to: SurfacePoint3D }> = [];
  const stepsPerCell = isContour ? 12 : 24;
  const rotX = clamp(chart.view3d?.rotX ?? (isWireframe ? 85 : 25), -88, 88) * (Math.PI / 180);
  const rotY = clamp(chart.view3d?.rotY ?? (isWireframe ? 0 : 30), -88, 88) * (Math.PI / 180);
  const usePerspective = !isContour && chart.view3d?.rAngAx === false;
  const perspectiveStrength = clamp((chart.view3d?.perspective ?? (usePerspective ? 30 : 0)) / 100, 0, 1);
  const depthScale = isContour ? 1 : clamp((chart.view3d?.depthPercent ?? 100) / 100, 0.2, 4);
  const usesFlatMaterial = chart.surfaceMaterial === "flat";
  const wallLineColor = chart.backWall?.lineColor ?? chart.sideWall?.lineColor ?? chart.floor?.lineColor ?? chart.axisLineColor ?? "#a6a6a6";
  const backWallFill = chart.backWall?.fillColor ?? "#d9d9df";
  const sideWallFill = chart.sideWall?.fillColor ?? "#d5d7dc";
  const floorFill = chart.floor?.fillColor ?? "#d0d2d8";
  const projectPoint = (column: number, row: number, value: number): SurfacePoint3D => {
    const worldX = normalizeSurfaceX(column, cols);
    if (isContour) {
      return {
        depth: 0,
        worldX,
        worldY: normalizeSurfaceY(row, rows),
        worldZ: 0,
        x: worldX,
        y: normalizeSurfaceY(row, rows)
      };
    }
    const worldY = normalizeSurfaceZ(value, domain, depthScale);
    const worldZ = normalizeSurfaceY(row, rows);
    const projected = projectCartesian3dPoint(worldX, worldY, worldZ, rotX, rotY, usePerspective, perspectiveStrength);
    return {
      depth: projected.depth,
      worldX,
      worldY,
      worldZ,
      x: projected.x,
      y: projected.y
    };
  };
  const addWallQuad = (a: SurfacePoint3D, b: SurfacePoint3D, c: SurfacePoint3D, d: SurfacePoint3D, color: string) => {
    wallTriangles.push(
      { color, points: [a, b, c] },
      { color, points: [a, c, d] }
    );
  };
  const addMeshLine = (from: SurfacePoint3D, to: SurfacePoint3D, color: string) => {
    meshLines.push({ color, from, to });
  };

  for (let rowIndex = 0; rowIndex < rows - 1; rowIndex += 1) {
    for (let columnIndex = 0; columnIndex < cols - 1; columnIndex += 1) {
      const p00 = matrix[rowIndex]?.[columnIndex];
      const p10 = matrix[rowIndex]?.[columnIndex + 1];
      const p01 = matrix[rowIndex + 1]?.[columnIndex];
      const p11 = matrix[rowIndex + 1]?.[columnIndex + 1];
      if (p00 == null || p10 == null || p01 == null || p11 == null) {
        continue;
      }
      for (let stepRow = 0; stepRow < stepsPerCell; stepRow += 1) {
        for (let stepColumn = 0; stepColumn < stepsPerCell; stepColumn += 1) {
          const u0 = stepColumn / stepsPerCell;
          const u1 = (stepColumn + 1) / stepsPerCell;
          const v0 = stepRow / stepsPerCell;
          const v1 = (stepRow + 1) / stepsPerCell;
          const sample = (u: number, v: number) => {
            const value = bilinearInterpolate(p00, p10, p01, p11, u, v);
            return {
              bandColor: resolveSurfaceBandColor(chart, palette, domain, value),
              point: isContour
                ? withProjectedDepth(projectPoint(columnIndex + u, rowIndex + v, value), isWireframe ? 0.12 : 0.08)
                : projectPoint(columnIndex + u, rowIndex + v, value),
              value
            };
          };
          const a = sample(u0, v0);
          const b = sample(u1, v0);
          const c = sample(u1, v1);
          const d = sample(u0, v1);
          if (!isWireframe) {
            const avg1 = (a.value + b.value + c.value) / 3;
            const avg2 = (a.value + c.value + d.value) / 3;
            fillVertices.push(
              { ...a, bandColor: resolveSurfaceBandColor(chart, palette, domain, avg1) },
              { ...b, bandColor: resolveSurfaceBandColor(chart, palette, domain, avg1) },
              { ...c, bandColor: resolveSurfaceBandColor(chart, palette, domain, avg1) },
              { ...a, bandColor: resolveSurfaceBandColor(chart, palette, domain, avg2) },
              { ...c, bandColor: resolveSurfaceBandColor(chart, palette, domain, avg2) },
              { ...d, bandColor: resolveSurfaceBandColor(chart, palette, domain, avg2) }
            );
          }
        }
      }

      const thresholds = domain.ticks.slice(1, -1);
      if (isContour && isWireframe) {
        const corners = [
          { value: p00, x: columnIndex, y: rowIndex },
          { value: p10, x: columnIndex + 1, y: rowIndex },
          { value: p11, x: columnIndex + 1, y: rowIndex + 1 },
          { value: p01, x: columnIndex, y: rowIndex + 1 }
        ];
        const edges: Array<[typeof corners[number], typeof corners[number]]> = [
          [corners[0], corners[1]],
          [corners[1], corners[2]],
          [corners[2], corners[3]],
          [corners[3], corners[0]]
        ];
        thresholds.forEach((threshold) => {
          const intersections: Array<{ x: number; y: number }> = [];
          edges.forEach(([start, end]) => {
            const delta = end.value - start.value;
            if (delta === 0) {
              return;
            }
            const crosses = (start.value < threshold && end.value > threshold) || (start.value > threshold && end.value < threshold);
            if (!crosses) {
              return;
            }
            const ratio = (threshold - start.value) / delta;
            intersections.push({
              x: start.x + (end.x - start.x) * ratio,
              y: start.y + (end.y - start.y) * ratio
            });
          });
          if (intersections.length === 2) {
            contourLines.push({
              color: darkenColor(resolveSurfaceBandColor(chart, palette, domain, threshold), 0.22),
              from: withProjectedDepth(projectPoint(intersections[0]?.x ?? columnIndex, intersections[0]?.y ?? rowIndex, threshold), 0.18),
              to: withProjectedDepth(projectPoint(intersections[1]?.x ?? columnIndex + 1, intersections[1]?.y ?? rowIndex + 1, threshold), 0.18)
            });
          } else if (intersections.length === 4) {
            const center = (p00 + p10 + p01 + p11) / 4;
            const pairings = center >= threshold
              ? [[0, 1], [2, 3]]
              : [[0, 3], [1, 2]];
            pairings.forEach(([leftIndex, rightIndex]) => {
              const leftPoint = intersections[leftIndex] ?? intersections[0];
              const rightPoint = intersections[rightIndex] ?? intersections[intersections.length - 1];
              if (!leftPoint || !rightPoint) {
                return;
              }
              contourLines.push({
                color: darkenColor(resolveSurfaceBandColor(chart, palette, domain, threshold), 0.22),
                from: withProjectedDepth(projectPoint(leftPoint.x, leftPoint.y, threshold), 0.18),
                to: withProjectedDepth(projectPoint(rightPoint.x, rightPoint.y, threshold), 0.18)
              });
            });
          }
        });
      } else if (!isContour) {
        const corners = [
          { value: p00, x: columnIndex, y: rowIndex },
          { value: p10, x: columnIndex + 1, y: rowIndex },
          { value: p11, x: columnIndex + 1, y: rowIndex + 1 },
          { value: p01, x: columnIndex, y: rowIndex + 1 }
        ];
        const edges: Array<[typeof corners[number], typeof corners[number]]> = [
          [corners[0], corners[1]],
          [corners[1], corners[2]],
          [corners[2], corners[3]],
          [corners[3], corners[0]]
        ];
        thresholds.forEach((threshold) => {
          const intersections: Array<{ x: number; y: number }> = [];
          edges.forEach(([start, end]) => {
            const delta = end.value - start.value;
            if (delta === 0) {
              return;
            }
            const crosses = (start.value < threshold && end.value > threshold) || (start.value > threshold && end.value < threshold);
            if (!crosses) {
              return;
            }
            const ratio = (threshold - start.value) / delta;
            intersections.push({
              x: start.x + (end.x - start.x) * ratio,
              y: start.y + (end.y - start.y) * ratio
            });
          });
          if (intersections.length === 2) {
            contourLines.push({
              color: darkenColor(resolveSurfaceBandPaletteColor(chart, palette, domain, threshold), 0.26),
              from: projectPoint(intersections[0]?.x ?? columnIndex, intersections[0]?.y ?? rowIndex, threshold),
              to: projectPoint(intersections[1]?.x ?? columnIndex + 1, intersections[1]?.y ?? rowIndex + 1, threshold)
            });
          } else if (intersections.length === 4) {
            const center = (p00 + p10 + p01 + p11) / 4;
            const pairings = center >= threshold
              ? [[0, 1], [2, 3]]
              : [[0, 3], [1, 2]];
            pairings.forEach(([leftIndex, rightIndex]) => {
              const leftPoint = intersections[leftIndex] ?? intersections[0];
              const rightPoint = intersections[rightIndex] ?? intersections[intersections.length - 1];
              if (!leftPoint || !rightPoint) {
                return;
              }
              contourLines.push({
                color: darkenColor(resolveSurfaceBandPaletteColor(chart, palette, domain, threshold), 0.26),
                from: projectPoint(leftPoint.x, leftPoint.y, threshold),
                to: projectPoint(rightPoint.x, rightPoint.y, threshold)
              });
            });
          }
        });
      }
    }
  }

  const columnSkip = Math.max(1, chart.categoryAxis?.tickMarkSkip ?? 1);
  const rowSkip = Math.max(1, chart.seriesAxis?.tickMarkSkip ?? 1);
  if (isContour) {
    const topLeft = withProjectedDepth(projectPoint(0, 0, domain.minValue), -0.12);
    const topRight = withProjectedDepth(projectPoint(cols - 1, 0, domain.minValue), -0.12);
    const bottomRight = withProjectedDepth(projectPoint(cols - 1, rows - 1, domain.minValue), -0.12);
    const bottomLeft = withProjectedDepth(projectPoint(0, rows - 1, domain.minValue), -0.12);
    if (chart.backWall?.hidden !== true && backWallFill !== "transparent") {
      addWallQuad(topLeft, topRight, bottomRight, bottomLeft, backWallFill);
    }
    for (let columnIndex = 0; columnIndex < cols; columnIndex += columnSkip) {
      addMeshLine(
        withProjectedDepth(projectPoint(columnIndex, 0, domain.minValue), 0.04),
        withProjectedDepth(projectPoint(columnIndex, rows - 1, domain.minValue), 0.04),
        wallLineColor
      );
    }
    for (let rowIndex = 0; rowIndex < rows; rowIndex += rowSkip) {
      addMeshLine(
        withProjectedDepth(projectPoint(0, rowIndex, domain.minValue), 0.04),
        withProjectedDepth(projectPoint(cols - 1, rowIndex, domain.minValue), 0.04),
        wallLineColor
      );
    }
  } else {
    const minPlaneValue = domain.minValue;
    const wallDepthOffset = -0.2;
    if (chart.floor?.hidden !== true && floorFill !== "transparent") {
      addWallQuad(
        nudgeProjectedDepth(projectPoint(0, 0, minPlaneValue), wallDepthOffset),
        nudgeProjectedDepth(projectPoint(cols - 1, 0, minPlaneValue), wallDepthOffset),
        nudgeProjectedDepth(projectPoint(cols - 1, rows - 1, minPlaneValue), wallDepthOffset),
        nudgeProjectedDepth(projectPoint(0, rows - 1, minPlaneValue), wallDepthOffset),
        floorFill
      );
    }
    if (chart.backWall?.hidden !== true && backWallFill !== "transparent") {
      addWallQuad(
        nudgeProjectedDepth(projectPoint(0, 0, domain.safeMax), wallDepthOffset),
        nudgeProjectedDepth(projectPoint(cols - 1, 0, domain.safeMax), wallDepthOffset),
        nudgeProjectedDepth(projectPoint(cols - 1, 0, minPlaneValue), wallDepthOffset),
        nudgeProjectedDepth(projectPoint(0, 0, minPlaneValue), wallDepthOffset),
        backWallFill
      );
    }
    if (chart.sideWall?.hidden !== true && sideWallFill !== "transparent") {
      addWallQuad(
        nudgeProjectedDepth(projectPoint(cols - 1, 0, domain.safeMax), wallDepthOffset),
        nudgeProjectedDepth(projectPoint(cols - 1, rows - 1, domain.safeMax), wallDepthOffset),
        nudgeProjectedDepth(projectPoint(cols - 1, rows - 1, minPlaneValue), wallDepthOffset),
        nudgeProjectedDepth(projectPoint(cols - 1, 0, minPlaneValue), wallDepthOffset),
        sideWallFill
      );
    }
    for (const tick of domain.ticks) {
      if (chart.backWall?.hidden !== true && wallLineColor !== "transparent") {
        addMeshLine(nudgeProjectedDepth(projectPoint(0, 0, tick), wallDepthOffset * 0.75), nudgeProjectedDepth(projectPoint(cols - 1, 0, tick), wallDepthOffset * 0.75), wallLineColor);
      }
      if (chart.sideWall?.hidden !== true && wallLineColor !== "transparent") {
        addMeshLine(nudgeProjectedDepth(projectPoint(cols - 1, 0, tick), wallDepthOffset * 0.75), nudgeProjectedDepth(projectPoint(cols - 1, rows - 1, tick), wallDepthOffset * 0.75), wallLineColor);
      }
    }
    for (let columnIndex = 0; columnIndex < cols; columnIndex += columnSkip) {
      if (chart.floor?.hidden !== true && wallLineColor !== "transparent") {
        addMeshLine(nudgeProjectedDepth(projectPoint(columnIndex, 0, minPlaneValue), wallDepthOffset * 0.75), nudgeProjectedDepth(projectPoint(columnIndex, rows - 1, minPlaneValue), wallDepthOffset * 0.75), wallLineColor);
      }
      if (chart.backWall?.hidden !== true && wallLineColor !== "transparent") {
        addMeshLine(nudgeProjectedDepth(projectPoint(columnIndex, 0, minPlaneValue), wallDepthOffset * 0.75), nudgeProjectedDepth(projectPoint(columnIndex, 0, domain.safeMax), wallDepthOffset * 0.75), wallLineColor);
      }
    }
    for (let rowIndex = 0; rowIndex < rows; rowIndex += rowSkip) {
      if (chart.floor?.hidden !== true && wallLineColor !== "transparent") {
        addMeshLine(nudgeProjectedDepth(projectPoint(0, rowIndex, minPlaneValue), wallDepthOffset * 0.75), nudgeProjectedDepth(projectPoint(cols - 1, rowIndex, minPlaneValue), wallDepthOffset * 0.75), wallLineColor);
      }
      if (chart.sideWall?.hidden !== true && wallLineColor !== "transparent") {
        addMeshLine(nudgeProjectedDepth(projectPoint(cols - 1, rowIndex, minPlaneValue), wallDepthOffset * 0.75), nudgeProjectedDepth(projectPoint(cols - 1, rowIndex, domain.safeMax), wallDepthOffset * 0.75), wallLineColor);
      }
    }
  }

  if (!isContour && isWireframe) {
    const addWireMeshLine = (
      fromColumn: number,
      fromRow: number,
      toColumn: number,
      toRow: number,
      fromValue: number,
      toValue: number,
      color: string
    ) => {
      const averageValue = (fromValue + toValue) * 0.5;
      addMeshLine(
        projectPoint(fromColumn, fromRow, fromValue),
        projectPoint(toColumn, toRow, toValue),
        color || resolveSurfaceWireframeColor(chart, palette, domain, averageValue)
      );
    };
    for (let rowIndex = 0; rowIndex < rows; rowIndex += 1) {
      const rowValues = matrix[rowIndex]?.filter((value): value is number => value != null) ?? [];
      const rowAverage = rowValues.length > 0
        ? rowValues.reduce((sum, value) => sum + value, 0) / rowValues.length
        : domain.minValue;
      const rowColor = chart.series[rowIndex]?.lineColor
        ?? chart.series[rowIndex]?.color
        ?? resolveSurfaceWireframeColor(chart, palette, domain, rowAverage);
      for (let columnIndex = 0; columnIndex < cols - 1; columnIndex += 1) {
        const left = matrix[rowIndex]?.[columnIndex];
        const right = matrix[rowIndex]?.[columnIndex + 1];
        if (left == null || right == null) {
          continue;
        }
        for (let step = 0; step < stepsPerCell; step += 1) {
          const u0 = step / stepsPerCell;
          const u1 = (step + 1) / stepsPerCell;
          const fromColumn = columnIndex + u0;
          const toColumn = columnIndex + u1;
          const fromValue = left + (right - left) * u0;
          const toValue = left + (right - left) * u1;
          addWireMeshLine(fromColumn, rowIndex, toColumn, rowIndex, fromValue, toValue, rowColor);
        }
      }
    }
    for (let columnIndex = 0; columnIndex < cols; columnIndex += 1) {
      const columnValues = Array.from({ length: rows }, (_, rowIndex) => matrix[rowIndex]?.[columnIndex])
        .filter((value): value is number => value != null);
      const columnAverage = columnValues.length > 0
        ? columnValues.reduce((sum, value) => sum + value, 0) / columnValues.length
        : domain.minValue;
      const columnColor = resolveSurfaceWireframeColor(chart, palette, domain, columnAverage);
      for (let rowIndex = 0; rowIndex < rows - 1; rowIndex += 1) {
        const top = matrix[rowIndex]?.[columnIndex];
        const bottom = matrix[rowIndex + 1]?.[columnIndex];
        if (top == null || bottom == null) {
          continue;
        }
        for (let step = 0; step < stepsPerCell; step += 1) {
          const v0 = step / stepsPerCell;
          const v1 = (step + 1) / stepsPerCell;
          const fromRow = rowIndex + v0;
          const toRow = rowIndex + v1;
          const fromValue = top + (bottom - top) * v0;
          const toValue = top + (bottom - top) * v1;
          addWireMeshLine(columnIndex, fromRow, columnIndex, toRow, fromValue, toValue, columnColor);
        }
      }
    }
  }

  const allPoints = [
    ...fillVertices.map((vertex) => vertex.point),
    ...wallTriangles.flatMap((triangle) => triangle.points),
    ...meshLines.flatMap((line) => [line.from, line.to]),
    ...contourLines.flatMap((line) => [line.from, line.to])
  ];
  if (allPoints.length === 0) {
    return null;
  }

  const minX = Math.min(...allPoints.map((point) => point.x));
  const maxX = Math.max(...allPoints.map((point) => point.x));
  const minY = Math.min(...allPoints.map((point) => point.y));
  const maxY = Math.max(...allPoints.map((point) => point.y));
  const minDepth = Math.min(...allPoints.map((point) => point.depth));
  const maxDepth = Math.max(...allPoints.map((point) => point.depth));
  const scale = Math.min(
    layout.plot.width / Math.max(0.25, maxX - minX),
    layout.plot.height / Math.max(0.25, maxY - minY)
  ) * 0.82;
  const centerX = layout.plot.left + layout.plot.width / 2;
  const centerY = layout.plot.top + layout.plot.height / 2;
  const centerRawX = (minX + maxX) / 2;
  const centerRawY = (minY + maxY) / 2;
  const toScreenX = (x: number) => centerX + (x - centerRawX) * scale;
  const toScreenY = (y: number) => centerY + (y - centerRawY) * scale;
  const toClipPoint = (point: SurfacePoint3D) => {
    const screenX = toScreenX(point.x);
    const screenY = toScreenY(point.y);
    return {
      x: (screenX / layout.width) * 2 - 1,
      y: 1 - (screenY / layout.height) * 2,
      z: clamp(((point.depth - minDepth) / Math.max(1e-6, maxDepth - minDepth)) * 2 - 1, -1, 1)
    };
  };

  const fillPositions: number[] = [];
  const fillColors: number[] = [];
  const wallPositions: number[] = [];
  const wallColors: number[] = [];
  for (const triangle of wallTriangles) {
    const rgba = colorToRgba(triangle.color, 1);
    for (const point of triangle.points) {
      const clip = toClipPoint(point);
      wallPositions.push(clip.x, clip.y, clip.z);
      wallColors.push(rgba[0], rgba[1], rgba[2], rgba[3] * 0.92);
    }
  }
  for (let index = 0; index < fillVertices.length; index += 3) {
    const left = fillVertices[index];
    const middle = fillVertices[index + 1];
    const right = fillVertices[index + 2];
    if (!left || !middle || !right) {
      continue;
    }
    const normal = computeNormal(left.point, middle.point, right.point);
    const lightDirection = { x: -0.32, y: 0.84, z: 0.44 };
    const rawLight = normal.x * lightDirection.x + normal.y * lightDirection.y + normal.z * lightDirection.z;
    const shaded = usesFlatMaterial
      ? colorToRgba(left.bandColor, 1)
      : shadeSurfaceColor(left.bandColor, clamp(0.8 + rawLight * 0.28, 0.52, 1.08));
    [left, middle, right].forEach((vertex) => {
      const clip = toClipPoint(vertex.point);
      fillPositions.push(clip.x, clip.y, clip.z);
      fillColors.push(shaded[0], shaded[1], shaded[2], shaded[3]);
    });
  }

  const linePositions: number[] = [];
  const lineColors: number[] = [];
  meshLines.forEach((line) => {
    const from = toClipPoint(line.from);
    const to = toClipPoint(line.to);
    const rgba = colorToRgba(line.color, 1);
    linePositions.push(from.x, from.y, from.z, to.x, to.y, to.z);
    lineColors.push(rgba[0], rgba[1], rgba[2], rgba[3], rgba[0], rgba[1], rgba[2], rgba[3]);
  });
  contourLines.forEach((line) => {
    const from = toClipPoint(line.from);
    const to = toClipPoint(line.to);
    const rgba = colorToRgba(line.color, 1);
    linePositions.push(from.x, from.y, from.z, to.x, to.y, to.z);
    lineColors.push(rgba[0], rgba[1], rgba[2], rgba[3], rgba[0], rgba[1], rgba[2], rgba[3]);
  });

  return {
    fillColors: new Float32Array(fillColors),
    fillPositions: new Float32Array(fillPositions),
    lineColors: new Float32Array(lineColors),
    linePositions: new Float32Array(linePositions),
    wallColors: new Float32Array(wallColors),
    wallPositions: new Float32Array(wallPositions)
  };
}

export const MemoSurfaceChartComposite = React.memo(function MemoSurfaceChartComposite({
  background,
  borderColor,
  chart,
  fontFamily,
  layout,
  overlay,
  palette,
  fallback
}: SurfaceCompositeProps) {
  const canvasRef = React.useRef<HTMLCanvasElement | null>(null);
  const [ready, setReady] = React.useState(false);
  const [failed, setFailed] = React.useState(false);
  const isContour = isContourSurfaceChart(chart);

  React.useEffect(() => {
    if (isContour) {
      setReady(false);
      setFailed(false);
      return;
    }
    const canvas = canvasRef.current;
    if (!canvas) {
      return;
    }
    let cancelled = false;
    let reglInstance: ReturnType<typeof createREGL> | null = null;

    const drawSurface = () => {
      try {
        const scene = buildSurfaceMesh(chart, palette, layout);
        if (!scene) {
          if (!cancelled) {
            setFailed(true);
            setReady(false);
          }
          return;
        }
        const dpr = window.devicePixelRatio || 1;
        canvas.width = Math.max(1, Math.round(layout.width * dpr));
        canvas.height = Math.max(1, Math.round(layout.height * dpr));
        reglInstance = createREGL({
          attributes: {
            alpha: true,
            antialias: true,
            depth: true,
            premultipliedAlpha: true
          },
          canvas
        });
        const fillCommand = scene.fillPositions.length > 0
          ? reglInstance({
              attributes: {
                color: scene.fillColors,
                position: scene.fillPositions
              },
              blend: {
                enable: true,
                func: {
                  dstRGB: "one minus src alpha",
                  dstAlpha: "one minus src alpha",
                  srcRGB: "src alpha",
                  srcAlpha: "src alpha"
                }
              },
              count: scene.fillPositions.length / 3,
              depth: {
                enable: true
              },
              frag: `
                precision mediump float;
                varying vec4 vColor;
                void main() {
                  gl_FragColor = vColor;
                }
              `,
              primitive: "triangles",
              vert: `
                precision mediump float;
                attribute vec3 position;
                attribute vec4 color;
                varying vec4 vColor;
                void main() {
                  gl_Position = vec4(position, 1.0);
                  vColor = color;
                }
              `
            })
          : null;
        const wallCommand = scene.wallPositions.length > 0
          ? reglInstance({
              attributes: {
                color: scene.wallColors,
                position: scene.wallPositions
              },
              blend: {
                enable: true,
                func: {
                  dstRGB: "one minus src alpha",
                  dstAlpha: "one minus src alpha",
                  srcRGB: "src alpha",
                  srcAlpha: "src alpha"
                }
              },
              count: scene.wallPositions.length / 3,
              depth: {
                enable: false
              },
              frag: `
                precision mediump float;
                varying vec4 vColor;
                void main() {
                  gl_FragColor = vColor;
                }
              `,
              primitive: "triangles",
              vert: `
                precision mediump float;
                attribute vec3 position;
                attribute vec4 color;
                varying vec4 vColor;
                void main() {
                  gl_Position = vec4(position, 1.0);
                  vColor = color;
                }
              `
            })
          : null;
        const lineCommand = scene.linePositions.length > 0
          ? reglInstance({
              attributes: {
                color: scene.lineColors,
                position: scene.linePositions
              },
              blend: {
                enable: true,
                func: {
                  dstRGB: "one minus src alpha",
                  dstAlpha: "one minus src alpha",
                  srcRGB: "src alpha",
                  srcAlpha: "src alpha"
                }
              },
              count: scene.linePositions.length / 3,
              depth: {
                enable: false
              },
              frag: `
                precision mediump float;
                varying vec4 vColor;
                void main() {
                  gl_FragColor = vColor;
                }
              `,
              lineWidth: 1,
              primitive: "lines",
              vert: `
                precision mediump float;
                attribute vec3 position;
                attribute vec4 color;
                varying vec4 vColor;
                void main() {
                  gl_Position = vec4(position, 1.0);
                  vColor = color;
                }
              `
            })
          : null;

        reglInstance.clear({ color: [0, 0, 0, 0], depth: 1 });
        wallCommand?.();
        fillCommand?.();
        lineCommand?.();
        if (!cancelled) {
          setReady(true);
          setFailed(false);
        }
      } catch {
        if (!cancelled) {
          setReady(false);
          setFailed(true);
        }
      }
    };

    setReady(false);
    setFailed(false);
    drawSurface();

    return () => {
      cancelled = true;
      reglInstance?.destroy();
    };
  }, [
    chart,
    isContour,
    layout,
    palette
  ]);

  const shouldRenderFallback = isContour || failed || !ready;

  return (
    <div
      aria-label={chart.title ?? chart.name ?? "Chart"}
      role="img"
      style={{
        background,
        border: borderColor.trim().toLowerCase() === "transparent" ? "none" : `1px solid ${borderColor}`,
        boxSizing: "border-box",
        fontFamily,
        height: "100%",
        overflow: "hidden",
        pointerEvents: "none",
        position: "relative",
        width: "100%"
      }}
    >
      <canvas
        ref={canvasRef}
        style={{
          display: shouldRenderFallback ? "none" : "block",
          height: "100%",
          inset: 0,
          position: "absolute",
          width: "100%"
        }}
      />
      <svg
        style={{ display: "block", height: "100%", inset: 0, position: "absolute", width: "100%" }}
        viewBox={`0 0 ${layout.width} ${layout.height}`}
      >
        {shouldRenderFallback ? fallback : null}
        {overlay}
      </svg>
    </div>
  );
}, (prev, next) => (
  prev.chart === next.chart
  && prev.palette === next.palette
  && prev.background === next.background
  && prev.borderColor === next.borderColor
  && prev.fontFamily === next.fontFamily
  && prev.layout.width === next.layout.width
  && prev.layout.height === next.layout.height
  && prev.layout.plot.left === next.layout.plot.left
  && prev.layout.plot.top === next.layout.plot.top
  && prev.layout.plot.width === next.layout.plot.width
  && prev.layout.plot.height === next.layout.plot.height
  && prev.overlay === next.overlay
  && prev.fallback === next.fallback
));
