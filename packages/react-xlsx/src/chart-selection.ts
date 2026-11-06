import type { XlsxChartElementSelection } from "./types";

const svgNamespace = "http://www.w3.org/2000/svg";
const geometrySelector = "path,rect,circle,ellipse,line,polyline,polygon";
const ownerSelector = "[data-chart-series]";
const selectionColor = "#3276b5";
const selectionHandleFill = "#cce7fc";
type Point = { x: number; y: number };

function keepHandleScreenSized(handle: SVGElement, owner: SVGGraphicsElement, point: Point) {
  const screen = owner.getScreenCTM();
  if (!screen) return;
  const inverse = screen.inverse();
  const { a, b, c, d } = inverse;
  if (![a, b, c, d].every(Number.isFinite)) return;
  const x = point.x - a * point.x - c * point.y;
  const y = point.y - b * point.x - d * point.y;
  handle.setAttribute("transform", `matrix(${a} ${b} ${c} ${d} ${x} ${y})`);
}

function selectionBounds(node: SVGGraphicsElement) {
  const matrix = node.getScreenCTM();
  if (!matrix) return null;
  const inverse = matrix.inverse();
  const shapes = node.matches(geometrySelector + ",text,image")
    ? [node] : Array.from(node.querySelectorAll<SVGGraphicsElement>(geometrySelector + ",text,image"));
  let left = Infinity;
  let top = Infinity;
  let right = -Infinity;
  let bottom = -Infinity;
  for (const shape of shapes) {
    if (shape.closest("defs,[data-chart-effect]")) continue;
    try {
      const transform = shape.getScreenCTM();
      if (!transform) continue;
      const local = inverse.multiply(transform);
      const box = shape.getBBox();
      for (const [x, y] of [
        [box.x, box.y], [box.x + box.width, box.y],
        [box.x, box.y + box.height], [box.x + box.width, box.y + box.height]
      ]) {
        const p = new DOMPoint(x, y).matrixTransform(local);
        if (!Number.isFinite(p.x) || !Number.isFinite(p.y)) continue;
        left = Math.min(left, p.x);
        top = Math.min(top, p.y);
        right = Math.max(right, p.x);
        bottom = Math.max(bottom, p.y);
      }
    } catch { /* Hidden primitives have no usable bounds. */ }
  }
  if (!Number.isFinite(left)) return null;
  return { x: left, y: top, width: right - left, height: bottom - top };
}

function projectedOutline(node: SVGGraphicsElement): Point[] {
  const matrix = node.getScreenCTM();
  if (!matrix) return [];
  let inverse: DOMMatrix;
  try { inverse = matrix.inverse(); } catch { return []; }
  const points: Point[] = [];
  for (const shape of node.querySelectorAll<SVGGeometryElement>(geometrySelector)) {
    if (shape.closest("defs,[data-chart-effect]") || shape.getAttribute("fill") === "transparent") continue;
    const transform = shape.getScreenCTM();
    if (!transform) continue;
    const local = inverse.multiply(transform);
    const append = (x: number, y: number) => {
      const p = new DOMPoint(x, y).matrixTransform(local);
      if (Number.isFinite(p.x) && Number.isFinite(p.y)) points.push({ x: p.x, y: p.y });
    };
    if (shape instanceof SVGPolygonElement || shape instanceof SVGPolylineElement) {
      for (let i = 0; i < shape.points.numberOfItems; i++) {
        const p = shape.points.getItem(i);
        append(p.x, p.y);
      }
    } else {
      try {
        const length = shape.getTotalLength();
        const count = Math.min(128, Math.max(16, Math.ceil(length / 4)));
        for (let i = 0; i <= count; i++) {
          const p = shape.getPointAtLength(length * i / count);
          append(p.x, p.y);
        }
      } catch { /* A non-rendered primitive has no selection geometry. */ }
    }
  }
  const vertices = Array.from(new Map(points.map(p => [p.x.toFixed(3) + ":" + p.y.toFixed(3), p])).values());
  if (vertices.length <= 16) return vertices;
  vertices.sort((a, b) => a.x - b.x || a.y - b.y);
  const turn = (a: Point, b: Point, c: Point) => (b.x - a.x) * (c.y - a.y) - (b.y - a.y) * (c.x - a.x);
  const hull: Point[] = [];
  for (const p of vertices) {
    while (hull.length >= 2 && turn(hull[hull.length - 2], hull[hull.length - 1], p) <= 0.01) hull.pop();
    hull.push(p);
  }
  const lower = hull.length;
  for (let i = vertices.length - 2; i >= 0; i--) {
    const p = vertices[i];
    while (hull.length > lower && turn(hull[hull.length - 2], hull[hull.length - 1], p) <= 0.01) hull.pop();
    hull.push(p);
  }
  if (hull.length > 1) hull.pop();
  const stride = Math.max(1, Math.ceil(hull.length / 16));
  return hull.filter((_, i) => i % stride === 0);
}

function matchesSelection(node: Element, selected: Exclude<XlsxChartElementSelection, { kind: "chart" | "legend" }>) {
  if (Number(node.getAttribute("data-chart-series")) !== selected.seriesIndex) return false;
  const legend = node.getAttribute("data-chart-kind") === "legendEntry";
  if (selected.kind === "legendEntry") return legend && (selected.legendIndex == null
    || Number(node.getAttribute("data-xlsx-chart-legend-index")) === selected.legendIndex);
  if (legend) return false;
  return selected.kind !== "point" || (node.hasAttribute("data-chart-point")
    && Number(node.getAttribute("data-chart-point")) === selected.pointIndex);
}

function wholeLegendHandles(svg: SVGSVGElement) {
  const matrix = svg.getScreenCTM();
  if (!matrix) return null;
  const inverse = matrix.inverse();
  const points: Point[] = [];
  for (const node of svg.querySelectorAll<SVGGraphicsElement>("[data-xlsx-chart-element-kind='legend'],[data-xlsx-chart-element-kind='legendEntry']")) {
    try {
      const transform = node.getScreenCTM();
      if (!transform) continue;
      const local = inverse.multiply(transform);
      const box = selectionBounds(node);
      if (!box) continue;
      for (const [x, y] of [
        [box.x, box.y], [box.x + box.width, box.y],
        [box.x, box.y + box.height], [box.x + box.width, box.y + box.height]
      ]) {
        const p = new DOMPoint(x, y).matrixTransform(local);
        if (Number.isFinite(p.x) && Number.isFinite(p.y)) points.push(p);
      }
    } catch { /* Hidden legend entries have no usable bounds. */ }
  }
  if (!points.length) return null;
  const left = Math.min(...points.map(p => p.x)) - 2;
  const right = Math.max(...points.map(p => p.x)) + 2;
  const top = Math.min(...points.map(p => p.y)) - 2;
  const bottom = Math.max(...points.map(p => p.y)) + 2;
  const middleX = (left + right) / 2;
  const middleY = (top + bottom) / 2;
  const group = document.createElementNS(svgNamespace, "g");
  group.setAttribute("data-xlsx-chart-selection-handles", "legend");
  group.setAttribute("pointer-events", "none");
  const outline = document.createElementNS(svgNamespace, "rect");
  outline.setAttribute("x", String(left));
  outline.setAttribute("y", String(top));
  outline.setAttribute("width", String(right - left));
  outline.setAttribute("height", String(bottom - top));
  outline.setAttribute("fill", "none");
  outline.setAttribute("stroke", "#a6a6a6");
  outline.setAttribute("stroke-width", "0.75");
  outline.setAttribute("vector-effect", "non-scaling-stroke");
  group.append(outline);
  for (const [x, y] of [
    [left, top], [middleX, top], [right, top],
    [left, middleY], [right, middleY],
    [left, bottom], [middleX, bottom], [right, bottom]
  ]) {
    const handle = document.createElementNS(svgNamespace, "rect");
    handle.setAttribute("x", String(x - 3));
    handle.setAttribute("y", String(y - 3));
    handle.setAttribute("width", "6");
    handle.setAttribute("height", "6");
    handle.setAttribute("fill", "#ffffff");
    handle.setAttribute("stroke", "#a6a6a6");
    handle.setAttribute("stroke-width", "0.75");
    handle.setAttribute("vector-effect", "non-scaling-stroke");
    keepHandleScreenSized(handle, svg, { x, y });
    group.append(handle);
  }
  return group;
}

function isPointAnchor(node: Element, owner: Element) {
  return node.localName === "circle" && node.getAttribute("fill") === "transparent"
    && owner.getAttribute("data-chart-kind") !== "bubble";
}

function outlineGeometry(node: SVGGraphicsElement) {
  const outline = node.cloneNode(false) as SVGGraphicsElement;
  for (const { name } of Array.from(outline.attributes)) {
    if (name === "id" || name === "style" || name === "filter" || name.startsWith("data-")) outline.removeAttribute(name);
  }
  outline.setAttribute("data-xlsx-chart-selection-outline", "");
  outline.setAttribute("fill", "none");
  outline.setAttribute("stroke", selectionColor);
  outline.setAttribute("stroke-width", "1.25");
  outline.setAttribute("stroke-opacity", "1");
  outline.setAttribute("stroke-linejoin", "round");
  outline.setAttribute("stroke-linecap", "round");
  outline.removeAttribute("stroke-dasharray");
  outline.setAttribute("opacity", "1");
  outline.setAttribute("pointer-events", "none");
  outline.setAttribute("vector-effect", "non-scaling-stroke");
  return outline;
}

function pointHandles(node: SVGGraphicsElement, points: Point[]) {
  const group = document.createElementNS(svgNamespace, "g");
  group.setAttribute("data-xlsx-chart-selection-handles", "");
  group.setAttribute("pointer-events", "none");
  const clip = node.getAttribute("clip-path");
  if (clip) group.setAttribute("clip-path", clip);
  for (const point of points) {
    const handle = document.createElementNS(svgNamespace, "circle");
    handle.setAttribute("cx", String(point.x));
    handle.setAttribute("cy", String(point.y));
    handle.setAttribute("r", "3");
    handle.setAttribute("fill", selectionHandleFill);
    handle.setAttribute("stroke", selectionColor);
    handle.setAttribute("stroke-width", "1.1");
    handle.setAttribute("vector-effect", "non-scaling-stroke");
    keepHandleScreenSized(handle, node, point);
    group.append(handle);
  }
  return group;
}

function ownerHandles(node: SVGGraphicsElement, is3d: boolean) {
  const kind = node.getAttribute("data-chart-kind");
  const x = Number(node.getAttribute("data-chart-x"));
  const y = Number(node.getAttribute("data-chart-y"));
  if ((kind === "point" || kind === "surfacePoint")
    && node.hasAttribute("data-chart-x") && node.hasAttribute("data-chart-y")
    && Number.isFinite(x) && Number.isFinite(y)) {
    return pointHandles(node, [{ x, y }]);
  }
  if (is3d && (kind === "bar3D" || kind === "slice")) {
    const outline = projectedOutline(node);
    if (outline.length) return pointHandles(node, outline);
  }
  try {
    const box = selectionBounds(node);
    if (!box) return null;
    if (![box.x, box.y, box.width, box.height].every(Number.isFinite)
      || (box.width === 0 && box.height === 0)) return null;
    return pointHandles(node, [
      { x: box.x, y: box.y },
      { x: box.x + box.width, y: box.y },
      { x: box.x, y: box.y + box.height },
      { x: box.x + box.width, y: box.y + box.height }
    ]);
  } catch {
    return null;
  }
}

function placeHandles(svg: SVGSVGElement, node: SVGGraphicsElement, handles: SVGElement) {
  const rootMatrix = svg.getScreenCTM();
  const nodeMatrix = node.getScreenCTM();
  if (!rootMatrix || !nodeMatrix) return false;
  const matrix = rootMatrix.inverse().multiply(nodeMatrix);
  if (![matrix.a, matrix.b, matrix.c, matrix.d, matrix.e, matrix.f].every(Number.isFinite)) return false;
  handles.setAttribute("transform", `matrix(${matrix.a} ${matrix.b} ${matrix.c} ${matrix.d} ${matrix.e} ${matrix.f})`);
  svg.append(handles);
  return true;
}

export function applyChartSelection(
  svg: SVGSVGElement,
  chartId: string,
  selected: XlsxChartElementSelection | null | undefined,
  is3d = false
) {
  if (!selected || selected.chartId !== chartId || selected.kind === "chart") return;
  if (selected.kind === "legend") {
    const handles = wholeLegendHandles(svg);
    if (!handles) return;
    svg.append(handles);
    svg.setAttribute("data-xlsx-chart-selection", "legend");
    return () => {
      handles.remove();
      svg.removeAttribute("data-xlsx-chart-selection");
    };
  }
  const overlays: SVGElement[] = [];
  const owners = Array.from(svg.querySelectorAll<SVGGraphicsElement>(ownerSelector))
    .filter(node => matchesSelection(node, selected));
  const geometry = Array.from(svg.querySelectorAll<SVGGraphicsElement>(geometrySelector));
  const ownersWithGeometry = new Set<Element>();
  for (const node of geometry) {
    if (node.closest("defs,[data-chart-effect]")) continue;
    const owner = node.closest(ownerSelector);
    if (!owner || !matchesSelection(owner, selected)) continue;
    ownersWithGeometry.add(owner);
    if (is3d || owner.getAttribute("data-chart-kind") === "surfacePoint" || isPointAnchor(node, owner)) continue;
    const outline = outlineGeometry(node);
    // Keeping each outline beside its geometry preserves projection, clipping, and occlusion.
    node.after(outline);
    overlays.push(outline);
  }
  const handled = new Set<string>();
  for (const node of owners) {
    if (!ownersWithGeometry.has(node) && selected.kind !== "legendEntry") continue;
    const point = node.getAttribute("data-chart-point");
    const kind = node.getAttribute("data-chart-kind");
    if (point == null && selected.kind !== "legendEntry"
      && node.querySelector("[data-chart-point]")) continue;
    if ((kind === "point" || kind === "surfacePoint") && point != null) {
      const key = node.getAttribute("data-chart-series") + ":" + point;
      if (handled.has(key)) continue;
      handled.add(key);
    }
    const handles = ownerHandles(node, is3d);
    if (!handles) continue;
    if (!placeHandles(svg, node, handles)) continue;
    overlays.push(handles);
  }
  svg.setAttribute("data-xlsx-chart-selection", selected.kind);
  return () => {
    for (const overlay of overlays) overlay.remove();
    svg.removeAttribute("data-xlsx-chart-selection");
  };
}
