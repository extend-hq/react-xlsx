import { getBuiltinChartColor, getBuiltinChartSurfaceLine, resolveBuiltinChartColor } from "./builtin-chart-styles";
import createREGL from "regl";
import * as React from "react";
import { shadeSurfaceMatteColor, surfaceNormal, transformSurfaceNormal } from "./surface-lighting";
import { clipSurfaceBand, groupSurfaceBandEdges, normalizeSurfaceValue, surfaceDataRatios, triangulateSurfaceQuad } from "./surface-geometry";
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
  onDoubleClick?: React.MouseEventHandler<SVGSVGElement>;
  onPointerDown?: React.PointerEventHandler<SVGSVGElement>;
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

function resolveSurfaceBaseColor(chart: XlsxChart, palette: Pick<SurfacePalette, "text">) {
  return chart.chartColorPalette?.[0]
    ?? chart.series[0]?.color
    ?? chart.series[0]?.lineColor
    ?? chart.axisLineColor
    ?? chart.textColor
    ?? palette.text;
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

function resolveExplicitSurfaceSeriesColor(chart: XlsxChart, kind: "fill" | "line") {
  const shape = chart.series[0]?.shapeProperties;
  if (kind === "fill" ? shape?.xmlFillHidden === true : shape?.xmlLineHidden === true) return "transparent";
  const color = kind === "fill"
    ? shape?.xmlFillColor ?? shape?.solidFillHex
    : shape?.xmlLineColor ?? shape?.lineColorHex;
  return normalizeSurfaceRendererHexColor(color);
}

function getSurfaceColorStops(chart: XlsxChart, palette: SurfacePalette) {
  const explicitFill = resolveExplicitSurfaceSeriesColor(chart, "fill");
  if (explicitFill) return [explicitFill];
  const colors = (chart.chartColorPalette ?? []).filter((value): value is string => typeof value === "string" && value.length > 0);
  return colors.length > 0 ? colors : [resolveSurfaceBaseColor(chart, palette)];
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
  const boundedTicks = [...new Set([...ticks.filter((tick) => tick < safeMax), safeMax])];
  return {
    maxValue,
    minValue,
    safeMax,
    ticks: boundedTicks.length >= 2 ? boundedTicks : [minValue, safeMax]
  };
}

function resolveSurfaceBandPaletteColor(chart: XlsxChart, palette: SurfacePalette, domain: SurfaceDomain, value: number) {
  const bandIndex = resolveSurfaceBandIndex(domain, value);
  const explicitColor = resolveSurfaceBandFormatColor(chart, "bandFormatColors", bandIndex);
  if (explicitColor) return explicitColor;
  const explicitFill = resolveExplicitSurfaceSeriesColor(chart, "fill");
  if (explicitFill) return explicitFill;
  const themeColors = chart.raw?.builtinThemeColors as Record<number, string> | undefined;
  const paletteCount = Math.trunc((domain.safeMax - domain.minValue) / (domain.ticks[1] - domain.ticks[0])) + 1;
  const builtinColor = themeColors ? resolveBuiltinChartColor(getBuiltinChartColor(chart.chartStyleId, bandIndex % paletteCount, paletteCount), themeColors) : undefined;
  if (builtinColor) return builtinColor;
  const stops = getSurfaceColorStops(chart, palette);
  if (stops.length === 0) {
    return resolveSurfaceBaseColor(chart, palette);
  }
  return stops[Math.min(stops.length - 1, Math.max(0, bandIndex))] ?? stops[stops.length - 1] ?? resolveSurfaceBaseColor(chart, palette);
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

function resolveSurfaceBandColor(chart: XlsxChart, palette: SurfacePalette, domain: SurfaceDomain, value: number) {
  return resolveSurfaceBandPaletteColor(chart, palette, domain, value);
}

function resolveSurfaceBandFormatColor(chart: XlsxChart, key: "bandFormatColors" | "bandFormatLineColors", index: number) {
  const colors = chart.raw?.[key];
  if (!colors || typeof colors !== "object" || !Object.prototype.hasOwnProperty.call(colors, index)) return null;
  const color = (colors as Record<number, unknown>)[index];
  return typeof color === "string" && color.length > 0 ? color : null;
}

function resolveSurfaceWireframeColor(chart: XlsxChart, palette: SurfacePalette, domain: SurfaceDomain, value: number) {
  const bandIndex = resolveSurfaceBandIndex(domain, value);
  const themeColors = chart.raw?.builtinThemeColors as Record<number, string> | undefined;
  const paletteCount = Math.trunc((domain.safeMax - domain.minValue) / (domain.ticks[1] - domain.ticks[0])) + 1;
  const builtinLine = themeColors ? resolveBuiltinChartColor(getBuiltinChartSurfaceLine(chart.chartStyleId, bandIndex, paletteCount), themeColors) : undefined;
  return resolveSurfaceBandFormatColor(chart, "bandFormatLineColors", bandIndex)
    ?? resolveExplicitSurfaceSeriesColor(chart, "line")
    ?? builtinLine
    ?? resolveSurfaceBandPaletteColor(chart, palette, domain, value);
}

function colorToRgba(color: string, alpha = 1) {
  if (color === "transparent") return [0, 0, 0, 0] as const;
  const rgb = parseRgbColor(color);
  if (!rgb) {
    return [0.1, 0.1, 0.1, alpha] as const;
  }
  return [rgb.red / 255, rgb.green / 255, rgb.blue / 255, alpha] as const;
}

function normalizeSurfaceX(column: number, cols: number) {
  return cols <= 1 ? 0 : ((column / (cols - 1)) - 0.5) * 2;
}

function normalizeSurfaceY(row: number, rows: number) {
  return rows <= 1 ? 0 : ((row / (rows - 1)) - 0.5) * 2;
}

function normalizeSurfaceZ(value: number, domain: SurfaceDomain, heightExtent: number, axis?: XlsxChart["valueAxis"]) {
  return -(normalizeSurfaceValue(value, domain.minValue, domain.safeMax, axis) - 0.5) * heightExtent;
}

function projectCartesian3dPoint(
  x: number,
  y: number,
  z: number,
  rotXRad: number,
  rotYRad: number,
  usePerspective: boolean,
  perspectiveStrength: number,
  rightAngleAxes: boolean
) {
  if (rightAngleAxes) {
    return { depth: z, x: x + z * Math.sin(rotYRad), y: y - z * Math.sin(rotXRad) };
  }
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
  const rotX = (chart.view3d?.rotX ?? 25) * (Math.PI / 180);
  const rotY = (chart.view3d?.rotY ?? 30) * (Math.PI / 180);
  const usePerspective = !isContour && chart.view3d?.rAngAx === false;
  const perspectiveStrength = clamp((chart.view3d?.perspective ?? (usePerspective ? 30 : 0)) / 100, 0, 1);
  const ratios = surfaceDataRatios(rows, cols, chart.valueAxis?.crossBetween !== "midCat", chart.view3d?.heightPercent, chart.view3d?.depthPercent, layout);
  const depthScale = ratios.depth;
  const heightExtent = 2 * ratios.height;
  const usesFlatMaterial = chart.surfaceMaterial === "flat";
  const wallLineColor = chart.valueAxis?.majorGridlineColor ?? chart.backWall?.lineColor ?? chart.sideWall?.lineColor ?? chart.floor?.lineColor ?? chart.axisLineColor ?? "#a6a6a6";
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
    const worldY = normalizeSurfaceZ(value, domain, heightExtent, chart.valueAxis);
    const worldZ = normalizeSurfaceY(row, rows) * depthScale;
    const projected = projectCartesian3dPoint(worldX, worldY, worldZ, rotX, rotY, usePerspective, perspectiveStrength, chart.view3d?.rAngAx !== false);
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
      const triangles = triangulateSurfaceQuad([
        { value: p00, x: columnIndex, y: rowIndex },
        { value: p10, x: columnIndex + 1, y: rowIndex },
        { value: p11, x: columnIndex + 1, y: rowIndex + 1 },
        { value: p01, x: columnIndex, y: rowIndex + 1 }
      ], chart.valueAxis);
      for (const triangle of triangles) {
        for (let band = 0; band < domain.ticks.length - 1; band += 1) {
          const lower = domain.ticks[band];
          const upper = domain.ticks[band + 1];
          const polygon = clipSurfaceBand(clipSurfaceBand(triangle, lower, true, chart.valueAxis, band === 0), upper, false, chart.valueAxis);
          if (polygon.length < 3) continue;
          const bandColor = resolveSurfaceBandColor(chart, palette, domain, (lower + upper) / 2);
          if (!isWireframe && bandColor !== "transparent") {
            for (let index = 1; index < polygon.length - 1; index += 1) {
              for (const vertex of [polygon[0], polygon[index], polygon[index + 1]]) {
                const point = projectPoint(vertex.x, vertex.y, vertex.value);
                fillVertices.push({ bandColor, point: isContour ? withProjectedDepth(point, 0.08) : point, value: vertex.value });
              }
            }
          }
          for (const group of groupSurfaceBandEdges(polygon, triangle, band, lower, isWireframe)) {
            const edgeMidpoint = (domain.ticks[group.band] + domain.ticks[group.band + 1]) / 2;
            const color = resolveSurfaceWireframeColor(chart, palette, domain, edgeMidpoint);
            if (color !== "transparent") {
              for (const [from, to] of group.edges) {
                contourLines.push({ color, from: projectPoint(from.x, from.y, from.value), to: projectPoint(to.x, to.y, to.value) });
              }
            }
          }
        }
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
    const wallDepthOffset = 0;
    const backRow = chart.view3d?.rAngAx !== false || Math.cos(rotX) * Math.cos(rotY) >= 0 ? rows - 1 : 0;
    const sideColumn = Math.sin(rotY) >= 0 ? 0 : cols - 1;
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
        nudgeProjectedDepth(projectPoint(0, backRow, domain.safeMax), wallDepthOffset),
        nudgeProjectedDepth(projectPoint(cols - 1, backRow, domain.safeMax), wallDepthOffset),
        nudgeProjectedDepth(projectPoint(cols - 1, backRow, minPlaneValue), wallDepthOffset),
        nudgeProjectedDepth(projectPoint(0, backRow, minPlaneValue), wallDepthOffset),
        backWallFill
      );
    }
    if (chart.sideWall?.hidden !== true && sideWallFill !== "transparent") {
      addWallQuad(
        nudgeProjectedDepth(projectPoint(sideColumn, 0, domain.safeMax), wallDepthOffset),
        nudgeProjectedDepth(projectPoint(sideColumn, rows - 1, domain.safeMax), wallDepthOffset),
        nudgeProjectedDepth(projectPoint(sideColumn, rows - 1, minPlaneValue), wallDepthOffset),
        nudgeProjectedDepth(projectPoint(sideColumn, 0, minPlaneValue), wallDepthOffset),
        sideWallFill
      );
    }
    for (const tick of chart.valueAxis?.majorGridlines ? domain.ticks : []) {
      if (chart.backWall?.hidden !== true && wallLineColor !== "transparent") {
        addMeshLine(nudgeProjectedDepth(projectPoint(0, backRow, tick), wallDepthOffset * 0.75), nudgeProjectedDepth(projectPoint(cols - 1, backRow, tick), wallDepthOffset * 0.75), wallLineColor);
      }
      if (chart.sideWall?.hidden !== true && wallLineColor !== "transparent") {
        addMeshLine(nudgeProjectedDepth(projectPoint(sideColumn, 0, tick), wallDepthOffset * 0.75), nudgeProjectedDepth(projectPoint(sideColumn, rows - 1, tick), wallDepthOffset * 0.75), wallLineColor);
      }
    }
    for (let columnIndex = 0; chart.categoryAxis?.majorGridlines && columnIndex < cols; columnIndex += columnSkip) {
      if (chart.floor?.hidden !== true && wallLineColor !== "transparent") {
        addMeshLine(nudgeProjectedDepth(projectPoint(columnIndex, 0, minPlaneValue), wallDepthOffset * 0.75), nudgeProjectedDepth(projectPoint(columnIndex, rows - 1, minPlaneValue), wallDepthOffset * 0.75), wallLineColor);
      }
      if (chart.backWall?.hidden !== true && wallLineColor !== "transparent") {
        addMeshLine(nudgeProjectedDepth(projectPoint(columnIndex, backRow, minPlaneValue), wallDepthOffset * 0.75), nudgeProjectedDepth(projectPoint(columnIndex, backRow, domain.safeMax), wallDepthOffset * 0.75), wallLineColor);
      }
    }
    for (let rowIndex = 0; chart.seriesAxis?.majorGridlines && rowIndex < rows; rowIndex += rowSkip) {
      if (chart.floor?.hidden !== true && wallLineColor !== "transparent") {
        addMeshLine(nudgeProjectedDepth(projectPoint(0, rowIndex, minPlaneValue), wallDepthOffset * 0.75), nudgeProjectedDepth(projectPoint(cols - 1, rowIndex, minPlaneValue), wallDepthOffset * 0.75), wallLineColor);
      }
      if (chart.sideWall?.hidden !== true && wallLineColor !== "transparent") {
        addMeshLine(nudgeProjectedDepth(projectPoint(sideColumn, rowIndex, minPlaneValue), wallDepthOffset * 0.75), nudgeProjectedDepth(projectPoint(sideColumn, rowIndex, domain.safeMax), wallDepthOffset * 0.75), wallLineColor);
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
  );
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
    const world = triangle.points.map(point => ({ x: point.worldX, y: point.worldY, z: -point.worldZ }));
    const normal = transformSurfaceNormal(surfaceNormal(world[0], world[1], world[2]), chart.view3d?.rotX ?? 25, chart.view3d?.rotY ?? 30, chart.view3d?.rAngAx !== false);
    const rgba = colorToRgba(usesFlatMaterial ? triangle.color : shadeSurfaceMatteColor(triangle.color, normal), 1);
    for (const point of triangle.points) {
      const clip = toClipPoint(point);
      wallPositions.push(clip.x, clip.y, clip.z);
      wallColors.push(rgba[0], rgba[1], rgba[2], rgba[3]);
    }
  }
  for (let index = 0; index < fillVertices.length; index += 3) {
    const left = fillVertices[index];
    const middle = fillVertices[index + 1];
    const right = fillVertices[index + 2];
    if (!left || !middle || !right) {
      continue;
    }
    const world = [left, middle, right].map(vertex => ({ x: vertex.point.worldX, y: vertex.point.worldY, z: -vertex.point.worldZ }));
    const normal = transformSurfaceNormal(surfaceNormal(world[0], world[1], world[2]), chart.view3d?.rotX ?? 25, chart.view3d?.rotY ?? 30, chart.view3d?.rAngAx !== false);
    const shaded = colorToRgba(usesFlatMaterial ? left.bandColor : shadeSurfaceMatteColor(left.bandColor, normal), 1);
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
    linePositions.push(from.x, from.y, from.z - 0.00001, to.x, to.y, to.z - 0.00001);
    lineColors.push(rgba[0], rgba[1], rgba[2], rgba[3], rgba[0], rgba[1], rgba[2], rgba[3]);
  });
  contourLines.forEach((line) => {
    const from = toClipPoint(line.from);
    const to = toClipPoint(line.to);
    const rgba = colorToRgba(line.color, 1);
    linePositions.push(from.x, from.y, from.z - 0.00001, to.x, to.y, to.z - 0.00001);
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
  onDoubleClick,
  onPointerDown,
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
                enable: true,
                func: "lequal"
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
                enable: true,
                func: "lequal"
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
        pointerEvents: "auto",
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
        onDoubleClick={onDoubleClick}
        onPointerDown={onPointerDown}
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
  && prev.onDoubleClick === next.onDoubleClick
  && prev.onPointerDown === next.onPointerDown
  && prev.layout.width === next.layout.width
  && prev.layout.height === next.layout.height
  && prev.layout.plot.left === next.layout.plot.left
  && prev.layout.plot.top === next.layout.plot.top
  && prev.layout.plot.width === next.layout.plot.width
  && prev.layout.plot.height === next.layout.plot.height
  && prev.overlay === next.overlay
  && prev.fallback === next.fallback
));
