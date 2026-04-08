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

type SurfaceScene = {
  fillColors: Float32Array;
  fillPositions: Float32Array;
  lineColors: Float32Array;
  linePositions: Float32Array;
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

function getSurfaceColorStops(chart: XlsxChart, palette: SurfacePalette) {
  const explicitStops = (chart.chartColorPalette ?? []).filter((value): value is string => typeof value === "string" && value.length > 0);
  if (explicitStops.length >= 2) {
    return explicitStops;
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
  const minValue = typeof chart.valueAxis?.min === "number" && Number.isFinite(chart.valueAxis.min)
    ? chart.valueAxis.min
    : Math.min(...numericValues);
  const maxValue = typeof chart.valueAxis?.max === "number" && Number.isFinite(chart.valueAxis.max)
    ? chart.valueAxis.max
    : Math.max(...numericValues);
  const safeMax = maxValue <= minValue ? minValue + 1 : maxValue;
  const span = safeMax - minValue;
  const roughStep = span / 5;
  const magnitude = Math.pow(10, Math.floor(Math.log10(Math.max(roughStep, 1e-6))));
  const normalized = roughStep / magnitude;
  const niceStep = normalized <= 1 ? 1 : normalized <= 2 ? 2 : normalized <= 5 ? 5 : 10;
  const step = niceStep * magnitude;
  const ticks: number[] = [minValue];
  let nextTick = Math.ceil(minValue / step) * step;
  while (nextTick < safeMax) {
    if (nextTick > minValue) {
      ticks.push(nextTick);
    }
    nextTick += step;
  }
  ticks.push(safeMax);
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

function resolveSurfaceBandColor(chart: XlsxChart, palette: SurfacePalette, domain: SurfaceDomain, value: number) {
  const ticks = domain.ticks;
  for (let index = 0; index < ticks.length - 1; index += 1) {
    const start = ticks[index] ?? domain.minValue;
    const end = ticks[index + 1] ?? domain.safeMax;
    if (value <= end || index === ticks.length - 2) {
      const midpoint = start + (end - start) * 0.5;
      const ratio = (midpoint - domain.minValue) / Math.max(1e-6, domain.safeMax - domain.minValue);
      return resolveSurfaceColor(chart, palette, ratio);
    }
  }
  return resolveSurfaceColor(chart, palette, 1);
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
  return ((((value - domain.minValue) / (domain.safeMax - domain.minValue)) - 0.5) * 1.8) * depthScale;
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
  const rawChartType = chart.raw && typeof chart.raw === "object" && typeof chart.raw.xmlChartType === "string"
    ? chart.raw.xmlChartType
    : "";
  const isContour = rawChartType === "surfaceChart";
  const isWireframe = chart.wireframe === true;
  const fillVertices: SurfaceVertex[] = [];
  const contourLines: Array<{ color: string; from: { x: number; y: number }; to: { x: number; y: number } }> = [];
  const meshLines: Array<{ color: string; from: SurfacePoint3D; to: SurfacePoint3D }> = [];
  const stepsPerCell = isContour ? 12 : 8;
  const rotX = clamp(chart.view3d?.rotX ?? (isWireframe ? 85 : 25), -88, 88) * (Math.PI / 180);
  const rotY = clamp(chart.view3d?.rotY ?? (isWireframe ? 0 : 30), -88, 88) * (Math.PI / 180);
  const usePerspective = !isContour && chart.view3d?.rAngAx === false;
  const perspectiveStrength = clamp((chart.view3d?.perspective ?? (usePerspective ? 30 : 0)) / 100, 0, 1);
  const depthScale = isContour ? 1 : clamp((chart.view3d?.depthPercent ?? 100) / 100, 0.2, 4);
  const projectPoint = (column: number, row: number, value: number): SurfacePoint3D => {
    const worldX = normalizeSurfaceX(column, cols);
    const worldY = normalizeSurfaceY(row, rows);
    if (isContour) {
      return {
        depth: 0,
        worldX,
        worldY,
        worldZ: 0,
        x: worldX,
        y: worldY
      };
    }
    const worldZ = normalizeSurfaceZ(value, domain, depthScale);
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
              point: projectPoint(columnIndex + u, rowIndex + v, value),
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

      if (isContour && isWireframe) {
        const thresholds = domain.ticks.slice(1, -1);
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
              from: intersections[0],
              to: intersections[1]
            });
          } else if (intersections.length === 4) {
            const center = (p00 + p10 + p01 + p11) / 4;
            const pairings = center >= threshold
              ? [[0, 1], [2, 3]]
              : [[0, 3], [1, 2]];
            pairings.forEach(([leftIndex, rightIndex]) => {
              contourLines.push({
                color: darkenColor(resolveSurfaceBandColor(chart, palette, domain, threshold), 0.22),
                from: intersections[leftIndex] ?? intersections[0],
                to: intersections[rightIndex] ?? intersections[intersections.length - 1]
              });
            });
          }
        });
      }
    }
  }

  if (!isContour && isWireframe) {
    const lineColor = chart.axisLineColor ?? darkenColor(resolveSurfaceBaseColor(chart, palette), 0.1);
    const addMeshLine = (fromColumn: number, fromRow: number, toColumn: number, toRow: number, fromValue: number, toValue: number) => {
      meshLines.push({
        color: lineColor,
        from: projectPoint(fromColumn, fromRow, fromValue),
        to: projectPoint(toColumn, toRow, toValue)
      });
    };
    for (let rowIndex = 0; rowIndex < rows; rowIndex += 1) {
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
          addMeshLine(fromColumn, rowIndex, toColumn, rowIndex, fromValue, toValue);
        }
      }
    }
    for (let columnIndex = 0; columnIndex < cols; columnIndex += 1) {
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
          addMeshLine(columnIndex, fromRow, columnIndex, toRow, fromValue, toValue);
        }
      }
    }
  }

  const allPoints = [
    ...fillVertices.map((vertex) => vertex.point),
    ...meshLines.flatMap((line) => [line.from, line.to]),
    ...contourLines.flatMap((line) => [
      { depth: 0, worldX: 0, worldY: 0, worldZ: 0, x: normalizeSurfaceX(line.from.x, cols), y: normalizeSurfaceY(line.from.y, rows) },
      { depth: 0, worldX: 0, worldY: 0, worldZ: 0, x: normalizeSurfaceX(line.to.x, cols), y: normalizeSurfaceY(line.to.y, rows) }
    ])
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
    const shaded = shadeSurfaceColor(left.bandColor, clamp(0.72 + rawLight * 0.35, 0.4, 1.18));
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
    const fromPoint: SurfacePoint3D = {
      depth: 0,
      worldX: 0,
      worldY: 0,
      worldZ: 0,
      x: normalizeSurfaceX(line.from.x, cols),
      y: normalizeSurfaceY(line.from.y, rows)
    };
    const toPoint: SurfacePoint3D = {
      depth: 0,
      worldX: 0,
      worldY: 0,
      worldZ: 0,
      x: normalizeSurfaceX(line.to.x, cols),
      y: normalizeSurfaceY(line.to.y, rows)
    };
    const from = toClipPoint(fromPoint);
    const to = toClipPoint(toPoint);
    const rgba = colorToRgba(line.color, 1);
    linePositions.push(from.x, from.y, 0, to.x, to.y, 0);
    lineColors.push(rgba[0], rgba[1], rgba[2], rgba[3], rgba[0], rgba[1], rgba[2], rgba[3]);
  });

  return {
    fillColors: new Float32Array(fillColors),
    fillPositions: new Float32Array(fillPositions),
    lineColors: new Float32Array(lineColors),
    linePositions: new Float32Array(linePositions)
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

  React.useEffect(() => {
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
                enable: true
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
    layout,
    palette
  ]);

  const shouldRenderFallback = failed || !ready;

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
