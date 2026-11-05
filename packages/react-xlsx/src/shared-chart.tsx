import * as React from "react";
import type { ChartData, OfficeCharts } from "./office-chart-types";
import { projectChartData } from "./chart-data-projection";
import { useChartEvents } from "./chart-interaction";
import { applyChartSelection } from "./chart-selection";
import { officeFontFallback } from "@extend-ai/office-fonts";
import type { ChartSvgProps } from "./chart-renderer-types";
import type { XlsxChart, XlsxChartRenderSource } from "./types";

const svgNamespace = "http://www.w3.org/2000/svg";
const allowedElements = new Set([
  "svg", "g", "defs", "clipPath", "rect", "circle", "ellipse", "path", "line", "polyline", "polygon",
  "title", "desc", "text", "tspan", "linearGradient", "radialGradient", "stop", "pattern", "image",
  "filter", "feGaussianBlur", "feFlood", "feComposite", "feMorphology", "feDiffuseLighting",
  "feSpecularLighting", "feDistantLight"
]);
const plainAttributes = new Set([
  "x", "y", "width", "height", "x1", "x2", "y1", "y2", "cx", "cy", "fx", "fy", "r", "rx", "ry", "fr",
  "viewBox", "preserveAspectRatio", "transform", "d", "points", "font-family", "font-weight", "font-style", "font-size", "font-style", "text-decoration", "text-anchor",
  "dominant-baseline", "stroke-linecap", "stroke-linejoin", "stroke-width", "stroke-dasharray",
  "opacity", "fill-opacity", "stroke-opacity", "fill-rule", "shape-rendering", "offset", "stop-opacity",
  "gradientUnits", "gradientTransform", "patternUnits", "patternContentUnits",
  "filterUnits", "in", "in2", "result", "stdDeviation", "flood-color", "flood-opacity", "operator", "color-interpolation-filters",
  "radius", "surfaceScale", "diffuseConstant", "lighting-color", "specularConstant", "specularExponent", "azimuth", "elevation", "k1", "k2", "k3", "k4",
  "pointer-events", "data-chart-kind", "data-chart-series", "data-chart-point", "data-chart-x", "data-chart-y", "data-chart-effect"
]);

function safeChartSvg(markup: string, prefix: string, seriesCount: number, seriesMap?: number[], selectableSeries = true, pointMaps?: number[][]) {
  const svg = new DOMParser().parseFromString(markup, "image/svg+xml").documentElement;
  if (svg.localName !== "svg" || svg.namespaceURI !== svgNamespace) return null;
  let legendIndex = 0;

  // Only local paint servers and inert geometry can enter the document.
  for (const node of [svg, ...Array.from(svg.querySelectorAll("*"))]) {
    if (node.namespaceURI !== svgNamespace || !allowedElements.has(node.localName)) return null;
    for (const { name, value } of Array.from(node.attributes)) {
      if (name === "font-family") {
        const families = value.split(",").map((family) => family.trim());
        const unquote = (family: string) => family.replace(/^["']|["']$/g, "");
        const fallback = officeFontFallback(unquote(families[0] ?? ""));
        if (fallback && !families.some((family) => unquote(family) === fallback)) {
          families.splice(1, 0, JSON.stringify(fallback));
          node.setAttribute(name, families.join(", "));
        }
        continue;
      }
      if (plainAttributes.has(name) || name === "xmlns") continue;
      if (name === "id" && /^[\w-]+$/.test(value)) {
        node.setAttribute(name, prefix + value);
        continue;
      }
      if (name === "href" && node.localName === "image"
        && /^data:image\/(?:png|jpeg|gif|webp);base64,[A-Za-z0-9+/]+={0,2}$/.test(value)) continue;
      if (name === "href" && /^(linearGradient|radialGradient)$/.test(node.localName) && /^#[\w-]+$/.test(value)) {
        node.setAttribute(name, "#" + prefix + value.slice(1));
        continue;
      }
      if (["fill", "stroke", "stop-color", "clip-path", "filter"].includes(name)) {
        const local = /^url\(#([\w-]+)\)$/.exec(value);
        if (local) {
          node.setAttribute(name, "url(#" + prefix + local[1] + ")");
          continue;
        }
        if (/^(?:#[\da-f]{3,8}|[a-z]+)$/i.test(value)) continue;
      }
      node.removeAttribute(name);
    }

    if (node.closest("[data-chart-effect]")) {
      node.removeAttribute("data-chart-series");
      node.removeAttribute("data-chart-point");
      node.removeAttribute("data-chart-kind");
      (node as SVGElement).style.pointerEvents = "none";
      continue;
    }
    if (node.getAttribute("data-chart-kind") === "legend" || node.getAttribute("data-chart-kind") === "legendEntry") {
      node.setAttribute("data-xlsx-chart-legend-index", String(legendIndex++));
    }
    if (node.getAttribute("data-chart-kind") === "legend") {
      node.setAttribute("data-xlsx-chart-element-kind", "legend");
      (node as SVGElement).style.cursor = "pointer";
      (node as SVGElement).style.pointerEvents = "all";
    }
    if (node.hasAttribute("data-chart-series")) {
      if (!selectableSeries && node.getAttribute("data-chart-kind") === "legendEntry") {
        node.setAttribute("data-xlsx-chart-element-kind", "legend");
        (node as SVGElement).style.cursor = "pointer";
        (node as SVGElement).style.pointerEvents = "all";
      }
      const sourceSeries = Number(node.getAttribute("data-chart-series"));
      const series = seriesMap ? seriesMap[sourceSeries] : sourceSeries;
      if (!selectableSeries || !Number.isInteger(sourceSeries) || sourceSeries < 0 || series == null
        || !Number.isInteger(series) || series < 0 || series >= seriesCount) {
        node.removeAttribute("data-chart-series");
        node.removeAttribute("data-chart-point");
        node.removeAttribute("data-chart-kind");
        continue;
      }
      node.setAttribute("data-chart-series", String(series));
      node.setAttribute("data-xlsx-chart-series-index", String(series));
      const point = node.getAttribute("data-chart-point");
      const originalPoint = point != null && point !== "" && Number.isInteger(Number(point)) && Number(point) >= 0
        ? pointMaps ? pointMaps[sourceSeries]?.[Number(point)] : Number(point) : undefined;
      if (originalPoint != null && Number.isInteger(originalPoint) && originalPoint >= 0) {
        node.setAttribute("data-chart-point", String(originalPoint));
        node.setAttribute("data-xlsx-chart-point-index", String(originalPoint));
      } else {
        node.removeAttribute("data-chart-point");
      }
      const kind = node.getAttribute("data-chart-kind");
      if (kind === "legendEntry") node.setAttribute("data-xlsx-chart-element-kind", kind);
      const style = (node as SVGElement).style;
      style.cursor = "pointer";
      style.pointerEvents = "all";
    }
  }
  return svg as unknown as SVGSVGElement;
}

function renderSharedChart(renderer: OfficeCharts, chart: XlsxChart, source: XlsxChartRenderSource, width: number, height: number, prefix: string) {
  try {
    const data: ChartData = {
      title: chart.title,
      series: (source.sourceSeries ?? chart.series).map((series, ordinal) => {
        const chartType = source.seriesTypes?.[ordinal] ?? source.kind;
        const xy = chartType === "scatter" || chartType === "bubble";
        const indexedX = source.indexedXSeries?.[ordinal] === true;
        const numericCategories = series.raw?.chartCategoryNumericValues as Array<number | null> | undefined;
        const useNumericCategories = source.dateCategorySeries?.[ordinal] === true || (xy && !indexedX);
        const hierarchy = series.raw?.chartExHierarchyCategories as string[][] | undefined;
        const colorStrings = series.raw?.chartExColorStrings as Array<string | null> | undefined;
        const classicHierarchy = series.raw?.chartHierarchyCategories as Array<Array<{ level: number; label: string }>> | undefined;
        return {
          chartType,
          name: series.name,
          points: Array.from({ length: Math.max(series.values.length, indexedX ? 0 : series.categories.length, series.bubbleSizes?.length ?? 0) }, (_, index) => {
            const value = series.values[index];
            const label = useNumericCategories && numericCategories?.[index] != null ? numericCategories[index] : series.categories[index];
            const bubble = series.bubbleSizes?.[index];
            return {
              index,
              value: xy ? undefined : colorStrings?.[index] ?? (value == null ? null : String(value)),
              x: indexedX ? String(index + 1) : xy && label != null ? String(label) : undefined,
              y: xy && value != null ? String(value) : undefined,
              bubbleSize: bubble == null ? undefined : String(bubble),
              categories: xy ? [] : hierarchy?.[index]?.length
                ? [...hierarchy[index]].reverse().map((label, level) => ({level, label}))
                : classicHierarchy?.[index]?.length ? classicHierarchy[index]
                : series.categories.length || series.categoriesRef?.formula
                  ? [{ level: 0, label: label == null ? "" : String(label) }]
                  : []
            };
          })
        };
      })
    };
    const projected = projectChartData(source.xml, data, (source.sourceSeries ?? chart.series)
      .map(series => series.raw?.chartExcludedPointIndices as number[] | undefined));
    const svg = renderer.renderChart(projected.xml, {
      themeColors: source.themeColors,
      themeXml: source.themeXml,
      themeOverrideXml: source.themeOverrideXml,
      styleXml: source.styleXml,
      colorStyleXml: source.colorStyleXml,
      images: source.images,
      width, height,
      interactive: true,
      spreadsheetDefaults: true,
      fontFamily: chart.fontFamily,
      data: projected.data
    });
    return svg ? safeChartSvg(svg, prefix, chart.series.length,
      source.sourceSeries?.map(series => chart.series.findIndex(candidate => candidate.id === series.id)),
      chart.chartType !== "Surface", projected.pointMaps) : null;
  } catch {
    return null;
  }
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
    if (!graphics.current || !svg) return;
    if (graphics.current.firstChild !== svg) graphics.current.replaceChildren(svg);
    let cleanup = applyChartSelection(svg, chart.id, selectedChartElement, chart.is3d);
    const viewport = graphics.current.ownerSVGElement;
    if (!viewport || !selectedChartElement || selectedChartElement.chartId !== chart.id
      || selectedChartElement.kind === "chart" || typeof ResizeObserver === "undefined") return cleanup;
    const observer = new ResizeObserver(() => {
      cleanup?.();
      cleanup = applyChartSelection(svg, chart.id, selectedChartElement, chart.is3d);
    });
    observer.observe(viewport);
    return () => {
      observer.disconnect();
      cleanup?.();
    };
  }, [svg, chart.id, chart.is3d, selectedChartElement]);
  const { handlePointerDown, handleDoubleClick } = useChartEvents(props);
  if (!svg) return <div data-xlsx-chart-fallback="unsupported" style={{width:"100%",height:"100%"}}>{fallback}</div>;
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
    </svg>
  );
}
