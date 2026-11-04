import * as React from "react";
import type { OfficeCharts } from "@extend-ai/rust-office-charts";
import { renderSelectionRectHandles, useChartEvents } from "./chart-interaction";
import type { ChartSvgProps } from "./chart-renderer";
import type { XlsxChart, XlsxChartRenderSource, XlsxChartElementSelection } from "./types";

const svgNamespace = "http://www.w3.org/2000/svg";
const allowedElements = new Set([
  "svg", "g", "defs", "clipPath", "rect", "circle", "ellipse", "path", "line", "polyline", "polygon",
  "text", "tspan", "linearGradient", "radialGradient", "stop", "pattern"
]);
const plainAttributes = new Set([
  "x", "y", "width", "height", "x1", "x2", "y1", "y2", "cx", "cy", "fx", "fy", "r", "rx", "ry", "fr",
  "viewBox", "transform", "d", "points", "font-family", "font-weight", "font-size", "text-anchor",
  "dominant-baseline", "stroke-linecap", "stroke-linejoin", "stroke-width", "stroke-dasharray",
  "opacity", "fill-opacity", "stroke-opacity", "fill-rule", "shape-rendering", "offset", "stop-opacity",
  "gradientUnits", "gradientTransform", "patternUnits", "patternContentUnits",
  "data-chart-kind", "data-chart-series", "data-chart-point", "data-chart-x", "data-chart-y"
]);

function safeChartSvg(markup: string, prefix: string) {
  const svg = new DOMParser().parseFromString(markup, "image/svg+xml").documentElement;
  if (svg.localName !== "svg" || svg.namespaceURI !== svgNamespace) return null;

  // Only local paint servers and inert geometry can enter the document.
  for (const node of [svg, ...Array.from(svg.querySelectorAll("*"))]) {
    if (node.namespaceURI !== svgNamespace || !allowedElements.has(node.localName)) return null;
    for (const { name, value } of Array.from(node.attributes)) {
      if (plainAttributes.has(name) || name === "xmlns") continue;
      if (name === "id" && /^[\w-]+$/.test(value)) {
        node.setAttribute(name, prefix + value);
        continue;
      }
      if (name === "href" && /^(linearGradient|radialGradient)$/.test(node.localName) && /^#[\w-]+$/.test(value)) {
        node.setAttribute(name, "#" + prefix + value.slice(1));
        continue;
      }
      if (["fill", "stroke", "stop-color", "clip-path"].includes(name)) {
        const local = /^url\(#([\w-]+)\)$/.exec(value);
        if (local) {
          node.setAttribute(name, "url(#" + prefix + local[1] + ")");
          continue;
        }
        if (/^(?:#[\da-f]{3,8}|[a-z]+)$/i.test(value)) continue;
      }
      node.removeAttribute(name);
    }

    if (node.hasAttribute("data-chart-series")) {
      node.setAttribute("data-xlsx-chart-series-index", node.getAttribute("data-chart-series")!);
      const point = node.getAttribute("data-chart-point");
      if (point != null) node.setAttribute("data-xlsx-chart-point-index", point);
      const kind = node.getAttribute("data-chart-kind");
      if (kind === "bar") node.setAttribute("data-xlsx-chart-selection-mode", "seriesFirst");
      if (kind === "legendEntry") node.setAttribute("data-xlsx-chart-element-kind", kind);
      const style = (node as SVGElement).style;
      style.cursor = "pointer";
      style.pointerEvents = "all";
    }
  }
  return svg;
}

function renderSharedChart(renderer: OfficeCharts, chart: XlsxChart, source: XlsxChartRenderSource, width: number, height: number, prefix: string) {
  try {
    const svg = renderer.renderChart(source.xml, {
      themeColors: source.themeColors,
      themeOverrideXml: source.themeOverrideXml,
      styleXml: source.styleXml,
      colorStyleXml: source.colorStyleXml,
      width, height,
      interactive: true,
      spreadsheetDefaults: true,
      fontFamily: chart.fontFamily,
      data: {
        title: chart.title,
        series: chart.series.map((series) => ({
          chartType: source.kind,
          name: series.name,
          points: Array.from({ length: Math.max(series.values.length, series.categories.length) }, (_, index) => {
            const value = series.values[index];
            const label = series.categories[index];
            return {
              index,
              value: value == null ? null : String(value),
              categories: series.categories.length || series.categoriesRef?.formula
                ? [{ level: 0, label: label == null ? "" : String(label) }]
                : []
            };
          })
        }))
      }
    });
    return svg ? safeChartSvg(svg, prefix) : null;
  } catch {
    return null;
  }
}

function selectionHandles(svg: Element, chartId: string, selected: XlsxChartElementSelection | null | undefined) {
  if (!selected || selected.chartId !== chartId || selected.kind === "chart") return null;
  return Array.from(svg.querySelectorAll("[data-chart-point]")).map((node, index) => {
    if (Number(node.getAttribute("data-chart-series")) !== selected.seriesIndex
      || (selected.kind === "point" && Number(node.getAttribute("data-chart-point")) !== selected.pointIndex)) return null;
    if (node.getAttribute("data-chart-kind") === "bar") {
      return renderSelectionRectHandles(String(index), Number(node.getAttribute("x")), Number(node.getAttribute("y")),
        Number(node.getAttribute("width")), Number(node.getAttribute("height")));
    }
    if (!node.hasAttribute("data-chart-x") || !node.hasAttribute("data-chart-y")) return null;
    return renderSelectionRectHandles(String(index), Number(node.getAttribute("data-chart-x")) - 4,
      Number(node.getAttribute("data-chart-y")) - 4, 8, 8);
  });
}

export function SharedChartSvg(props: ChartSvgProps & {
  source: XlsxChartRenderSource;
  renderer: OfficeCharts;
  fallback: React.ReactNode;
}) {
  const { chart, rect, renderer, source, selectedChartElement, fallback } = props;
  const id = React.useId();
  const width = Math.max(1, rect.width);
  const height = Math.max(1, rect.height);
  const svg = React.useMemo(() => renderSharedChart(renderer, chart, source, width, height,
    "oc-" + id.replace(/[^\w-]/g, "") + "-"), [renderer, chart, source, width, height, id]);
  const graphics = React.useRef<SVGGElement>(null);
  React.useLayoutEffect(() => {
    if (graphics.current && svg) graphics.current.replaceChildren(svg);
  }, [svg]);
  const handles = React.useMemo(() => svg && selectionHandles(svg, chart.id, selectedChartElement),
    [svg, chart.id, selectedChartElement]);
  const { handlePointerDown, handleDoubleClick } = useChartEvents(props);
  if (!svg) return <>{fallback}</>;
  return (
    <svg
      aria-label={chart.title ?? chart.name ?? "Chart"}
      data-xlsx-chart-renderer="office-charts"
      onDoubleClick={handleDoubleClick}
      onPointerDown={handlePointerDown}
      role="img"
      style={{ display: "block", height: "100%", pointerEvents: "auto", width: "100%" }}
      viewBox={`0 0 ${width} ${height}`}
    >
      <g ref={graphics} />
      {handles}
    </svg>
  );
}
