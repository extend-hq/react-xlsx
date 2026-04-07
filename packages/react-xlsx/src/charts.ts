import type { Workbook } from "@dukelib/sheets-wasm";
import { strFromU8, strToU8 } from "fflate";
import type { WorkbookImageAssets, WorkbookImageSheetOrigin } from "./images";
import type {
  XlsxChart,
  XlsxChartAxis,
  XlsxChartDataLabels,
  XlsxChartLegend,
  XlsxChartPointStyle,
  XlsxChartReference,
  XlsxChartSeries,
  XlsxChartsheet,
  XlsxImageAnchor,
  XlsxThemePalette,
  XlsxWorkbookTab
} from "./types";

const CHART_NS = "http://schemas.openxmlformats.org/drawingml/2006/chart";
const DRAWINGML_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";
const DRAWING_SPREADSHEET_NS = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
const PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships";
const CHART_STYLE_REL_TYPE = "http://schemas.microsoft.com/office/2011/relationships/chartStyle";
const CHART_COLOR_STYLE_REL_TYPE = "http://schemas.microsoft.com/office/2011/relationships/chartColorStyle";
const SERIES_COLORS = [
  "#4472c4",
  "#ed7d31",
  "#a5a5a5",
  "#ffc000",
  "#5b9bd5",
  "#70ad47",
  "#264478",
  "#9e480e",
  "#636363",
  "#997300"
];
const EMU_PER_PIXEL = 9525;
const THEME_COLOR_INDEX_BY_NAME: Record<string, number> = {
  accent1: 4,
  accent2: 5,
  accent3: 6,
  accent4: 7,
  accent5: 8,
  accent6: 9,
  dk1: 1,
  dk2: 3,
  folHlink: 11,
  hlink: 10,
  lt1: 0,
  lt2: 2,
  tx1: 1,
  tx2: 3,
  bg1: 0,
  bg2: 2
};

export type WorkbookChartOrigin = {
  anchorIndex: number;
  chartPath: string | null;
  drawingPath: string;
  workbookSheetIndex: number;
};

export type WorkbookChartAssets = {
  chartOriginsById: Map<string, WorkbookChartOrigin>;
  chartsByWorkbookSheetIndex: XlsxChart[][];
  chartsheets: XlsxChartsheet[];
  tabs: XlsxWorkbookTab[];
};

function clampUnitInterval(value: number) {
  return Math.max(0, Math.min(1, value));
}

function normalizeHexColor(value: string) {
  const hex = value.replace(/^#/, "");
  if (hex.length === 8) {
    return `#${hex.slice(2).toLowerCase()}`;
  }
  if (hex.length === 6) {
    return `#${hex.toLowerCase()}`;
  }
  return null;
}

function parseHexColor(color: string): [number, number, number] | null {
  const normalized = normalizeHexColor(color);
  if (!normalized) {
    return null;
  }
  const match = /^#([0-9a-f]{6})$/.exec(normalized);
  if (!match) {
    return null;
  }
  return [
    Number.parseInt(match[1].slice(0, 2), 16),
    Number.parseInt(match[1].slice(2, 4), 16),
    Number.parseInt(match[1].slice(4, 6), 16)
  ];
}

function rgbToHsl(red: number, green: number, blue: number): [number, number, number] {
  const normalizedRed = red / 255;
  const normalizedGreen = green / 255;
  const normalizedBlue = blue / 255;
  const max = Math.max(normalizedRed, normalizedGreen, normalizedBlue);
  const min = Math.min(normalizedRed, normalizedGreen, normalizedBlue);
  const lightness = (max + min) / 2;

  if (max === min) {
    return [0, 0, lightness];
  }

  const delta = max - min;
  const saturation = lightness > 0.5 ? delta / (2 - max - min) : delta / (max + min);
  let hue = 0;

  switch (max) {
    case normalizedRed:
      hue = (normalizedGreen - normalizedBlue) / delta + (normalizedGreen < normalizedBlue ? 6 : 0);
      break;
    case normalizedGreen:
      hue = (normalizedBlue - normalizedRed) / delta + 2;
      break;
    default:
      hue = (normalizedRed - normalizedGreen) / delta + 4;
      break;
  }

  return [hue / 6, saturation, lightness];
}

function hueToRgb(p: number, q: number, t: number) {
  let nextT = t;
  if (nextT < 0) {
    nextT += 1;
  }
  if (nextT > 1) {
    nextT -= 1;
  }
  if (nextT < 1 / 6) {
    return p + (q - p) * 6 * nextT;
  }
  if (nextT < 1 / 2) {
    return q;
  }
  if (nextT < 2 / 3) {
    return p + (q - p) * (2 / 3 - nextT) * 6;
  }
  return p;
}

function hslToRgb(hue: number, saturation: number, lightness: number): [number, number, number] {
  if (saturation === 0) {
    const gray = Math.round(lightness * 255);
    return [gray, gray, gray];
  }

  const q = lightness < 0.5 ? lightness * (1 + saturation) : lightness + saturation - lightness * saturation;
  const p = 2 * lightness - q;

  return [
    Math.round(hueToRgb(p, q, hue + 1 / 3) * 255),
    Math.round(hueToRgb(p, q, hue) * 255),
    Math.round(hueToRgb(p, q, hue - 1 / 3) * 255)
  ];
}

function rgbToHex(red: number, green: number, blue: number) {
  return `#${[red, green, blue]
    .map((channel) => Math.max(0, Math.min(255, Math.round(channel))).toString(16).padStart(2, "0"))
    .join("")}`;
}

function applyLightnessTransform(baseColor: string, modifier = 1, offset = 0) {
  const rgb = parseHexColor(baseColor);
  if (!rgb) {
    return normalizeHexColor(baseColor);
  }

  const [hue, saturation, lightness] = rgbToHsl(rgb[0], rgb[1], rgb[2]);
  const nextLightness = clampUnitInterval(lightness * modifier + offset);
  const [nextRed, nextGreen, nextBlue] = hslToRgb(hue, saturation, nextLightness);
  return rgbToHex(nextRed, nextGreen, nextBlue);
}

function resolveThemeColor(name: string | null, themePalette?: XlsxThemePalette | null) {
  if (!name) {
    return null;
  }
  const index = THEME_COLOR_INDEX_BY_NAME[name];
  return index === undefined ? null : themePalette?.colorsByIndex[index] ?? null;
}

function resolveChartColorNode(node: Element | null, themePalette?: XlsxThemePalette | null): string | null {
  if (!node) {
    return null;
  }

  let baseColor: string | null = null;
  if (node.localName === "srgbClr") {
    baseColor = normalizeHexColor(`#${node.getAttribute("val") ?? ""}`);
  } else if (node.localName === "schemeClr") {
    baseColor = resolveThemeColor(node.getAttribute("val"), themePalette);
  } else if (node.localName === "sysClr") {
    baseColor = normalizeHexColor(`#${node.getAttribute("lastClr") ?? ""}`);
  }

  if (!baseColor) {
    return null;
  }

  let lightnessModifier = 1;
  let lightnessOffset = 0;
  for (const transformNode of Array.from(node.childNodes).filter((child): child is Element => child.nodeType === Node.ELEMENT_NODE)) {
    const rawValue = Number(transformNode.getAttribute("val") ?? Number.NaN);
    if (!Number.isFinite(rawValue)) {
      continue;
    }
    if (transformNode.localName === "lumMod") {
      lightnessModifier *= rawValue / 100000;
    } else if (transformNode.localName === "lumOff") {
      lightnessOffset += rawValue / 100000;
    } else if (transformNode.localName === "tint") {
      lightnessOffset += (1 - lightnessOffset) * (rawValue / 100000);
    } else if (transformNode.localName === "shade") {
      lightnessModifier *= rawValue / 100000;
    }
  }

  return applyLightnessTransform(baseColor, lightnessModifier, lightnessOffset);
}

function resolveChartFillColor(shapeNode: Element | null, themePalette?: XlsxThemePalette | null) {
  if (!shapeNode || getFirstLocalChild(shapeNode, "noFill")) {
    return null;
  }
  const solidFill = getFirstLocalChild(shapeNode, "solidFill");
  if (!solidFill) {
    return null;
  }
  const colorNode = Array.from(solidFill.childNodes).find((child): child is Element => child.nodeType === Node.ELEMENT_NODE) ?? null;
  return resolveChartColorNode(colorNode, themePalette);
}

function resolveChartLineStyle(shapeNode: Element | null, themePalette?: XlsxThemePalette | null) {
  const lineNode = shapeNode?.localName === "ln" ? shapeNode : (shapeNode ? getFirstLocalChild(shapeNode, "ln") : null);
  if (!lineNode || getFirstLocalChild(lineNode, "noFill")) {
    return { color: null, widthPx: undefined };
  }

  const solidFill = getFirstLocalChild(lineNode, "solidFill");
  const colorNode = solidFill
    ? Array.from(solidFill.childNodes).find((child): child is Element => child.nodeType === Node.ELEMENT_NODE) ?? null
    : null;
  const widthValue = Number(lineNode.getAttribute("w") ?? Number.NaN);
  return {
    color: resolveChartColorNode(colorNode, themePalette),
    widthPx: Number.isFinite(widthValue) ? Math.max(1, widthValue / EMU_PER_PIXEL) : undefined
  };
}

function normalizeLegend(raw: unknown): XlsxChartLegend | null {
  if (!raw || typeof raw !== "object") {
    return null;
  }
  const legend = raw as Record<string, unknown>;
  return {
    overlay: typeof legend.overlay === "boolean" ? legend.overlay : undefined,
    position: typeof legend.position === "string" ? legend.position : undefined,
    raw: legend
  };
}

function normalizeLegendPosition(position: string | undefined) {
  if (!position) {
    return undefined;
  }
  switch (position) {
    case "bottom":
      return "b";
    case "left":
      return "l";
    case "right":
      return "r";
    case "top":
      return "t";
    default:
      return position;
  }
}

function readChartNumericAttribute(parent: Element | null, localName: string) {
  const node = parent ? getFirstLocalChild(parent, localName) : null;
  const value = Number(node?.getAttribute("val") ?? Number.NaN);
  return Number.isFinite(value) ? value : undefined;
}

function readChartRelationships(
  archive: Record<string, Uint8Array>,
  chartPath: string
) {
  const relsPath = normalizeArchivePath(`${dirname(chartPath)}/_rels/${chartPath.split("/").pop()}.rels`);
  const relsXml = readArchiveText(archive, relsPath);
  if (!relsXml) {
    return new Map<string, string>();
  }

  const relsDocument = parseXml(relsXml);
  if (!relsDocument) {
    return new Map<string, string>();
  }

  const relationships = new Map<string, string>();
  for (const relationshipNode of getLocalDescendants(relsDocument, "Relationship")) {
    const type = relationshipNode.getAttribute("Type");
    const target = relationshipNode.getAttribute("Target");
    if (!type || !target) {
      continue;
    }
    relationships.set(type, resolveRelationshipPath(relsPath, target));
  }

  return relationships;
}

function readChartColorPalette(
  archive: Record<string, Uint8Array>,
  colorStylePath: string | null | undefined,
  themePalette?: XlsxThemePalette | null
) {
  const colorStyleXml = readArchiveText(archive, colorStylePath);
  if (!colorStyleXml) {
    return [];
  }

  const colorStyleDocument = parseXml(colorStyleXml);
  if (!colorStyleDocument?.documentElement) {
    return [];
  }

  return Array.from(colorStyleDocument.documentElement.childNodes)
    .filter((child): child is Element => child.nodeType === Node.ELEMENT_NODE && child.localName !== "variation")
    .map((child) => resolveChartColorNode(child, themePalette))
    .filter((color): color is string => typeof color === "string");
}

function readChartStylePaletteOffset(
  archive: Record<string, Uint8Array>,
  stylePath: string | null | undefined
) {
  const styleXml = readArchiveText(archive, stylePath);
  if (!styleXml) {
    return undefined;
  }

  const styleDocument = parseXml(styleXml);
  if (!styleDocument) {
    return undefined;
  }

  const dataPointNode = getFirstLocalDescendant(styleDocument, "dataPoint");
  const fillRefNode = dataPointNode ? getFirstLocalChild(dataPointNode, "fillRef") : null;
  const index = Number(fillRefNode?.getAttribute("idx") ?? Number.NaN);
  return Number.isFinite(index) ? index : undefined;
}

function parseChartPointStyles(seriesNode: Element, themePalette?: XlsxThemePalette | null): XlsxChartPointStyle[] {
  const pointStyles: XlsxChartPointStyle[] = [];

  for (const dataPointNode of getLocalChildren(seriesNode, "dPt")) {
    const indexValue = readChartNumericAttribute(dataPointNode, "idx");
    if (indexValue === undefined) {
      continue;
    }
    const shapeProperties = getFirstLocalChild(dataPointNode, "spPr");
    const lineStyle = resolveChartLineStyle(shapeProperties, themePalette);
    pointStyles.push({
      color: resolveChartFillColor(shapeProperties, themePalette) ?? undefined,
      index: indexValue,
      lineColor: lineStyle.color ?? undefined
    });
  }

  return pointStyles;
}

function applyChartSeriesStyleFromXml(chart: XlsxChart, chartTypeNode: Element, themePalette?: XlsxThemePalette | null) {
  const seriesNodes = getLocalChildren(chartTypeNode, "ser");
  chart.series = chart.series.map((series, index) => {
    const seriesNode = seriesNodes[index];
    if (!seriesNode) {
      return series;
    }

    const shapeProperties = getFirstLocalChild(seriesNode, "spPr");
    const markerNode = getFirstLocalChild(seriesNode, "marker");
    const markerShapeProperties = getFirstLocalChild(markerNode ?? chartTypeNode, "spPr");
    const lineStyle = resolveChartLineStyle(shapeProperties, themePalette);
    const markerLineStyle = resolveChartLineStyle(markerShapeProperties, themePalette);
    const fillColor = resolveChartFillColor(shapeProperties, themePalette);
    const markerSize = readChartNumericAttribute(markerNode, "size");
    const markerSymbolNode = markerNode ? getFirstLocalChild(markerNode, "symbol") : null;
    const markerSymbol = markerSymbolNode?.getAttribute("val") ?? undefined;
    const pointStyles = parseChartPointStyles(seriesNode, themePalette);

    return {
      ...series,
      color: fillColor ?? lineStyle.color ?? series.color,
      dataPointStyles: pointStyles.length > 0 ? pointStyles : series.dataPointStyles,
      lineColor: lineStyle.color ?? fillColor ?? series.lineColor ?? series.color,
      lineWidthPx: lineStyle.widthPx ?? series.lineWidthPx,
      markerColor: resolveChartFillColor(markerShapeProperties, themePalette) ?? fillColor ?? lineStyle.color ?? undefined,
      markerLineColor: markerLineStyle.color ?? lineStyle.color ?? fillColor ?? undefined,
      markerSize: markerSize ?? series.markerSize,
      markerSymbol,
      shapeProperties: {
        ...series.shapeProperties,
        xmlFillColor: fillColor ?? undefined,
        xmlLineColor: lineStyle.color ?? undefined,
        xmlLineWidthPx: lineStyle.widthPx ?? undefined
      }
    };
  });
}

function applyChartStyleFromXml(
  chart: XlsxChart,
  chartPath: string | undefined,
  archive: Record<string, Uint8Array>,
  themePalette?: XlsxThemePalette | null
) {
  const chartXml = readArchiveText(archive, chartPath);
  if (!chartXml) {
    return;
  }

  const chartDocument = parseXml(chartXml);
  const chartNode = chartDocument ? getFirstLocalDescendant(chartDocument, "chart") : null;
  const plotAreaNode = chartNode ? getFirstLocalChild(chartNode, "plotArea") : null;
  const chartTypeNode = plotAreaNode
    ? getLocalChildren(plotAreaNode, "barChart")[0]
      ?? getLocalChildren(plotAreaNode, "lineChart")[0]
      ?? getLocalChildren(plotAreaNode, "pieChart")[0]
      ?? getLocalChildren(plotAreaNode, "doughnutChart")[0]
      ?? getLocalChildren(plotAreaNode, "scatterChart")[0]
      ?? getLocalChildren(plotAreaNode, "areaChart")[0]
      ?? getLocalChildren(plotAreaNode, "radarChart")[0]
      ?? null
    : null;

  if (!chartNode || !chartTypeNode) {
    return;
  }

  const legendNode = getFirstLocalChild(chartNode, "legend");
  const legendPosition = legendNode ? getFirstLocalChild(legendNode, "legendPos")?.getAttribute("val") ?? undefined : undefined;
  const legendOverlay = legendNode ? getFirstLocalChild(legendNode, "overlay")?.getAttribute("val") : undefined;

  chart.legend = legendNode ? {
    overlay: legendOverlay === "1",
    position: normalizeLegendPosition(legendPosition),
    raw: chart.legend?.raw
  } : chart.legend;
  chart.firstSliceAngle = readChartNumericAttribute(chartTypeNode, "firstSliceAng") ?? chart.firstSliceAngle;
  chart.holeSize = readChartNumericAttribute(chartTypeNode, "holeSize") ?? chart.holeSize;
  chart.radarStyle = getFirstLocalChild(chartTypeNode, "radarStyle")?.getAttribute("val") ?? chart.radarStyle;

  const relationships = chartPath ? readChartRelationships(archive, chartPath) : new Map<string, string>();
  chart.chartColorPalette = readChartColorPalette(archive, relationships.get(CHART_COLOR_STYLE_REL_TYPE), themePalette);
  chart.chartColorPaletteOffset = readChartStylePaletteOffset(archive, relationships.get(CHART_STYLE_REL_TYPE)) ?? chart.chartColorPaletteOffset;

  applyChartSeriesStyleFromXml(chart, chartTypeNode, themePalette);
}

function normalizeArchivePath(path: string) {
  return path.replace(/^\/+/, "").replace(/\\/g, "/");
}

function dirname(path: string) {
  const normalized = normalizeArchivePath(path);
  const index = normalized.lastIndexOf("/");
  return index >= 0 ? normalized.slice(0, index) : "";
}

function resolveRelationshipPath(basePath: string, target: string) {
  if (!target) {
    return "";
  }

  const normalizedTarget = target.replace(/\\/g, "/");
  if (normalizedTarget.startsWith("/")) {
    return normalizeArchivePath(normalizedTarget);
  }

  const segments = [...dirname(basePath).split("/").filter(Boolean), ...normalizedTarget.split("/").filter(Boolean)];
  const resolved: string[] = [];
  for (const segment of segments) {
    if (segment === ".") {
      continue;
    }
    if (segment === "..") {
      resolved.pop();
      continue;
    }
    resolved.push(segment);
  }

  return resolved.join("/");
}

function readArchiveText(archive: Record<string, Uint8Array>, path: string | null | undefined) {
  if (!path) {
    return null;
  }

  const entry = archive[normalizeArchivePath(path)];
  return entry ? strFromU8(entry) : null;
}

function parseXml(xml: string) {
  if (typeof DOMParser === "undefined") {
    return null;
  }

  try {
    return new DOMParser().parseFromString(xml, "application/xml");
  } catch {
    return null;
  }
}

function serializeXml(document: XMLDocument) {
  return new XMLSerializer().serializeToString(document);
}

function getLocalChildren(parent: ParentNode, localName: string) {
  return Array.from(parent.childNodes).filter(
    (node): node is Element => node.nodeType === Node.ELEMENT_NODE && (node as Element).localName === localName
  );
}

function getLocalDescendants(parent: ParentNode, localName: string) {
  return Array.from((parent as Element | Document).getElementsByTagName("*")).filter(
    (node) => node.localName === localName
  );
}

function getFirstLocalChild(parent: ParentNode, localName: string) {
  return getLocalChildren(parent, localName)[0] ?? null;
}

function getFirstLocalDescendant(parent: ParentNode, localName: string) {
  return getLocalDescendants(parent, localName)[0] ?? null;
}

function ensureChild(parent: Element, localName: string, namespace = parent.namespaceURI ?? CHART_NS, prefix = "c") {
  const existing = getFirstLocalChild(parent, localName);
  if (existing) {
    return existing;
  }

  const document = parent.ownerDocument;
  const node = document.createElementNS(namespace, `${prefix}:${localName}`);
  parent.appendChild(node);
  return node;
}

function setLeafValue(parent: Element, localName: string, value: string, namespace = parent.namespaceURI ?? CHART_NS, prefix = "c") {
  const node = ensureChild(parent, localName, namespace, prefix);
  node.textContent = value;
  return node;
}

function setBooleanValue(parent: Element, localName: string, value: boolean) {
  const node = ensureChild(parent, localName);
  node.setAttribute("val", value ? "1" : "0");
  return node;
}

function setNumericValue(parent: Element, localName: string, value: number) {
  const node = ensureChild(parent, localName);
  node.setAttribute("val", String(Math.round(value)));
  return node;
}

function unquoteSheetName(value: string) {
  const trimmed = value.trim();
  if (trimmed.startsWith("'") && trimmed.endsWith("'")) {
    return trimmed.slice(1, -1).replace(/''/g, "'");
  }
  return trimmed;
}

function splitSheetReference(reference: string) {
  let bangIndex = -1;
  let quoted = false;
  for (let index = 0; index < reference.length; index += 1) {
    const char = reference[index];
    if (char === "'") {
      quoted = !quoted;
    } else if (char === "!" && !quoted) {
      bangIndex = index;
      break;
    }
  }

  if (bangIndex < 0) {
    return null;
  }

  return {
    range: reference.slice(bangIndex + 1),
    sheetName: unquoteSheetName(reference.slice(0, bangIndex))
  };
}

function parseA1Cell(reference: string) {
  const match = /^\$?([A-Z]+)\$?(\d+)$/i.exec(reference.trim());
  if (!match) {
    return null;
  }

  let col = 0;
  for (const char of match[1].toUpperCase()) {
    col = col * 26 + (char.charCodeAt(0) - 64);
  }

  return {
    col: col - 1,
    row: Number(match[2]) - 1
  };
}

function parseA1Range(reference: string) {
  const [startRef, endRef = startRef] = reference.split(":");
  const start = parseA1Cell(startRef ?? "");
  const end = parseA1Cell(endRef ?? "");
  if (!start || !end) {
    return null;
  }

  return {
    end: {
      col: Math.max(start.col, end.col),
      row: Math.max(start.row, end.row)
    },
    start: {
      col: Math.min(start.col, end.col),
      row: Math.min(start.row, end.row)
    }
  };
}

function resolveReferenceSheet(workbook: Workbook, fallbackSheetIndex: number, formula?: string | null) {
  if (!formula) {
    return {
      range: null,
      sheet: workbook.getSheet(fallbackSheetIndex),
      sheetName: workbook.getSheet(fallbackSheetIndex)?.name ?? ""
    };
  }

  const split = splitSheetReference(formula);
  if (!split) {
    return {
      range: parseA1Range(formula),
      sheet: workbook.getSheet(fallbackSheetIndex),
      sheetName: workbook.getSheet(fallbackSheetIndex)?.name ?? ""
    };
  }

  try {
    return {
      range: parseA1Range(split.range),
      sheet: workbook.getSheetByName(split.sheetName),
      sheetName: split.sheetName
    };
  } catch {
    return {
      range: parseA1Range(split.range),
      sheet: workbook.getSheet(fallbackSheetIndex),
      sheetName: workbook.getSheet(fallbackSheetIndex)?.name ?? ""
    };
  }
}

function cellValueToNumber(value: unknown) {
  if (typeof value === "number" && Number.isFinite(value)) {
    return value;
  }
  if (value && typeof value === "object") {
    if ((value as { is_empty?: boolean }).is_empty) {
      return null;
    }
    const candidates: unknown[] = [];
    if (typeof (value as { asNumber?: () => unknown }).asNumber === "function") {
      candidates.push((value as { asNumber: () => unknown }).asNumber());
    }
    if (typeof (value as { toJs?: () => unknown }).toJs === "function") {
      candidates.push((value as { toJs: () => unknown }).toJs());
    }
    if (typeof (value as { asText?: () => unknown }).asText === "function") {
      candidates.push((value as { asText: () => unknown }).asText());
    }
    if (typeof (value as { toString?: () => unknown }).toString === "function") {
      candidates.push((value as { toString: () => unknown }).toString());
    }

    for (const candidate of candidates) {
      if (typeof candidate === "number" && Number.isFinite(candidate)) {
        return candidate;
      }
      if (typeof candidate === "string") {
        const parsed = Number(candidate.replace(/,/g, ""));
        if (Number.isFinite(parsed)) {
          return parsed;
        }
      }
    }
  }
  if (typeof value === "string") {
    const parsed = Number(value.replace(/,/g, ""));
    return Number.isFinite(parsed) ? parsed : null;
  }
  return null;
}

function cellValueToDisplay(value: unknown) {
  if (value === null || value === undefined) {
    return "";
  }
  if (typeof value === "string") {
    return value;
  }
  if (typeof value === "number" || typeof value === "boolean") {
    return String(value);
  }
  if (value && typeof value === "object") {
    if ((value as { is_empty?: boolean }).is_empty) {
      return "";
    }
    const candidates: unknown[] = [];
    if (typeof (value as { asText?: () => unknown }).asText === "function") {
      candidates.push((value as { asText: () => unknown }).asText());
    }
    if (typeof (value as { toJs?: () => unknown }).toJs === "function") {
      candidates.push((value as { toJs: () => unknown }).toJs());
    }
    if (typeof (value as { toString?: () => unknown }).toString === "function") {
      candidates.push((value as { toString: () => unknown }).toString());
    }

    for (const candidate of candidates) {
      if (candidate === null || candidate === undefined) {
        continue;
      }
      if (typeof candidate === "string") {
        return candidate;
      }
      return String(candidate);
    }
  }
  return String(value);
}

function resolveReferenceValues(
  workbook: Workbook,
  fallbackSheetIndex: number,
  reference: XlsxChartReference | null | undefined,
  mode: "category" | "value"
): Array<number | string | null> {
  if (!reference?.formula) {
    return reference?.values ?? [];
  }

  const resolved = resolveReferenceSheet(workbook, fallbackSheetIndex, reference.formula);
  if (!resolved.sheet || !resolved.range) {
    return reference.values ?? [];
  }

  const values: Array<number | string | null> = [];
  for (let row = resolved.range.start.row; row <= resolved.range.end.row; row += 1) {
    for (let col = resolved.range.start.col; col <= resolved.range.end.col; col += 1) {
      const calculated = typeof resolved.sheet.getCalculatedValueAt === "function"
        ? resolved.sheet.getCalculatedValueAt(row, col)
        : null;
      const formatted = typeof resolved.sheet.getFormattedValueAt === "function"
        ? resolved.sheet.getFormattedValueAt(row, col)
        : calculated;
      if (mode === "value") {
        values.push(cellValueToNumber(calculated ?? formatted));
      } else {
        const display = cellValueToDisplay(formatted ?? calculated);
        const numeric = cellValueToNumber(calculated ?? formatted);
        values.push(display.length > 0 ? display : (numeric !== null ? numeric : null));
      }
    }
  }

  return values;
}

function resolveSeriesName(workbook: Workbook, fallbackSheetIndex: number, rawName: unknown) {
  if (typeof rawName !== "string" || !rawName) {
    return undefined;
  }

  const split = splitSheetReference(rawName);
  if (!split) {
    return rawName;
  }

  const resolved = resolveReferenceSheet(workbook, fallbackSheetIndex, rawName);
  if (!resolved.sheet || !resolved.range) {
    return rawName;
  }

  const value = typeof resolved.sheet.getFormattedValueAt === "function"
    ? resolved.sheet.getFormattedValueAt(resolved.range.start.row, resolved.range.start.col)
    : null;
  const display = cellValueToDisplay(value);
  return display || rawName;
}

function normalizeChartReference(raw: unknown): XlsxChartReference | null {
  if (!raw || typeof raw !== "object") {
    return null;
  }

  const record = raw as Record<string, unknown>;
  return {
    formula: typeof record.formula === "string" ? record.formula : undefined,
    refType: typeof record.refType === "string" ? record.refType : undefined,
    values: Array.isArray(record.values) ? record.values as Array<number | string | null> : undefined
  };
}

function normalizeChartAxis(raw: unknown): XlsxChartAxis | null {
  if (!raw || typeof raw !== "object") {
    return null;
  }

  const axis = raw as Record<string, unknown>;
  const numberFormat = axis.numberFormat && typeof axis.numberFormat === "object"
    ? axis.numberFormat as Record<string, unknown>
    : null;

  return {
    crosses: typeof axis.crosses === "string" ? axis.crosses : undefined,
    crossBetween: typeof axis.crossBetween === "string" ? axis.crossBetween : undefined,
    delete: typeof axis.delete === "boolean" ? axis.delete : undefined,
    labelPosition: typeof axis.labelPosition === "string" ? axis.labelPosition : undefined,
    majorGridlines: typeof axis.majorGridlines === "boolean" ? axis.majorGridlines : undefined,
    majorTickMark: typeof axis.majorTickMark === "string" ? axis.majorTickMark : undefined,
    minorGridlines: typeof axis.minorGridlines === "boolean" ? axis.minorGridlines : undefined,
    minorTickMark: typeof axis.minorTickMark === "string" ? axis.minorTickMark : undefined,
    numberFormat: numberFormat ? {
      formatCode: typeof numberFormat.formatCode === "string" ? numberFormat.formatCode : undefined,
      sourceLinked: typeof numberFormat.sourceLinked === "boolean" ? numberFormat.sourceLinked : undefined
    } : undefined,
    position: typeof axis.position === "string" ? axis.position : undefined,
    raw: axis,
    shapeProperties: axis.shapeProperties && typeof axis.shapeProperties === "object"
      ? axis.shapeProperties as Record<string, unknown>
      : undefined
  };
}

function normalizeChartDataLabels(raw: unknown): XlsxChartDataLabels | null {
  if (!raw || typeof raw !== "object") {
    return null;
  }

  const labels = raw as Record<string, unknown>;
  return {
    raw: labels,
    showBubbleSize: typeof labels.showBubbleSize === "boolean" ? labels.showBubbleSize : undefined,
    showCategoryName: typeof labels.showCategoryName === "boolean" ? labels.showCategoryName : undefined,
    showLegendKey: typeof labels.showLegendKey === "boolean" ? labels.showLegendKey : undefined,
    showPercent: typeof labels.showPercent === "boolean" ? labels.showPercent : undefined,
    showSeriesName: typeof labels.showSeriesName === "boolean" ? labels.showSeriesName : undefined,
    showValue: typeof labels.showValue === "boolean" ? labels.showValue : undefined
  };
}

function normalizeChartAnchor(raw: unknown): XlsxImageAnchor {
  if (!raw || typeof raw !== "object") {
    return {
      kind: "two-cell",
      from: { col: 0, colOffsetEmu: 0, row: 0, rowOffsetEmu: 0 },
      to: { col: 8, colOffsetEmu: 0, row: 15, rowOffsetEmu: 0 }
    };
  }

  const anchor = raw as Record<string, unknown>;
  return {
    kind: "two-cell",
    from: {
      col: typeof anchor.fromCol === "number" ? anchor.fromCol : 0,
      colOffsetEmu: typeof anchor.fromColOffset === "number" ? anchor.fromColOffset : 0,
      row: typeof anchor.fromRow === "number" ? anchor.fromRow : 0,
      rowOffsetEmu: typeof anchor.fromRowOffset === "number" ? anchor.fromRowOffset : 0
    },
    to: {
      col: typeof anchor.toCol === "number" ? anchor.toCol : 8,
      colOffsetEmu: typeof anchor.toColOffset === "number" ? anchor.toColOffset : 0,
      row: typeof anchor.toRow === "number" ? anchor.toRow : 15,
      rowOffsetEmu: typeof anchor.toRowOffset === "number" ? anchor.toRowOffset : 0
    }
  };
}

function normalizeChartSeries(
  workbook: Workbook,
  workbookSheetIndex: number,
  chartId: string,
  raw: unknown,
  index: number
): XlsxChartSeries {
  const series = raw && typeof raw === "object" ? raw as Record<string, unknown> : {};
  const categoriesRef = normalizeChartReference(series.categories);
  const valuesRef = normalizeChartReference(series.values);

  return {
    categories: resolveReferenceValues(workbook, workbookSheetIndex, categoriesRef, "category"),
    categoriesRef,
    color: SERIES_COLORS[index % SERIES_COLORS.length],
    dataPoints: Array.isArray(series.dataPoints) ? series.dataPoints : [],
    dataPointStyles: undefined,
    id: `${chartId}-series-${index}`,
    invertIfNegative: typeof series.invertIfNegative === "boolean" ? series.invertIfNegative : undefined,
    lineColor: undefined,
    lineWidthPx: typeof series.shapeProperties === "object" && typeof (series.shapeProperties as Record<string, unknown>).lineWidth === "number"
      ? Math.max(1, Number((series.shapeProperties as Record<string, unknown>).lineWidth) / EMU_PER_PIXEL)
      : undefined,
    marker: series.marker && typeof series.marker === "object" ? series.marker as Record<string, unknown> : undefined,
    markerColor: undefined,
    markerLineColor: undefined,
    markerSize: series.marker && typeof series.marker === "object" && typeof (series.marker as Record<string, unknown>).size === "number"
      ? Number((series.marker as Record<string, unknown>).size)
      : undefined,
    markerSymbol: series.marker && typeof series.marker === "object" && typeof (series.marker as Record<string, unknown>).symbol === "string"
      ? String((series.marker as Record<string, unknown>).symbol)
      : undefined,
    name: resolveSeriesName(workbook, workbookSheetIndex, series.name),
    raw: series,
    shapeProperties: series.shapeProperties && typeof series.shapeProperties === "object"
      ? series.shapeProperties as Record<string, unknown>
      : undefined,
    smooth: typeof series.smooth === "boolean" ? series.smooth : undefined,
    values: resolveReferenceValues(workbook, workbookSheetIndex, valuesRef, "value").map((value) => (
      typeof value === "number" && Number.isFinite(value) ? value : null
    )),
    valuesRef
  };
}

function normalizeChartsheet(raw: unknown, index: number): XlsxChartsheet {
  const chartsheet = raw && typeof raw === "object" ? raw as Record<string, unknown> : {};
  return {
    chartIds: Array.isArray(chartsheet.chartIds) ? chartsheet.chartIds.filter((value): value is string => typeof value === "string") : [],
    chartPath: typeof chartsheet.chartPath === "string" ? chartsheet.chartPath : undefined,
    id: `chartsheet-${index}`,
    index,
    name: typeof chartsheet.name === "string" ? chartsheet.name : `Chart ${index + 1}`,
    raw: chartsheet,
    workbookSheetIndex: typeof chartsheet.workbookSheetIndex === "number" ? chartsheet.workbookSheetIndex : undefined
  };
}

function buildTabs(
  workbook: Workbook,
  chartsheets: XlsxChartsheet[],
  visibleSheetIndexByWorkbookSheetIndex: Map<number, number>
): XlsxWorkbookTab[] {
  const rawOrder = Array.isArray(workbook.sheetOrder) ? workbook.sheetOrder as Array<Record<string, unknown>> : [];
  if (rawOrder.length === 0) {
    return workbook.sheetNames.map((name, index) => ({
      id: `sheet-${index}`,
      index,
      kind: "sheet" as const,
      name,
      sheetIndex: visibleSheetIndexByWorkbookSheetIndex.get(index) ?? index,
      workbookSheetIndex: index
    }));
  }

  return rawOrder.flatMap<XlsxWorkbookTab>((entry, index) => {
    const slotType = typeof entry.slotType === "string" ? entry.slotType : "worksheet";
    const slotIndex = typeof entry.index === "number" ? entry.index : index;
    if (slotType === "chartsheet") {
      const chartsheet = chartsheets[slotIndex];
      return chartsheet ? [{
        chartsheetIndex: slotIndex,
        id: `chartsheet-${slotIndex}`,
        index,
        kind: "chartsheet" as const,
        name: chartsheet.name
      }] : [];
    }

    const worksheet = workbook.getSheet(slotIndex);
    if (worksheet.visibility !== "visible") {
      return [];
    }

    return [{
      id: `sheet-${slotIndex}`,
      index,
      kind: "sheet" as const,
      name: worksheet.name,
      sheetIndex: visibleSheetIndexByWorkbookSheetIndex.get(slotIndex) ?? slotIndex,
      workbookSheetIndex: slotIndex
    }];
  });
}

function collectChartOriginsForSheet(
  archive: Record<string, Uint8Array>,
  origin: WorkbookImageSheetOrigin | null
) {
  if (!origin) {
    return [] as WorkbookChartOrigin[];
  }

  const chartOrigins: WorkbookChartOrigin[] = [];

  for (const attachment of origin.attachments) {
    const drawingXml = readArchiveText(archive, attachment.drawingPath);
    const relsXml = readArchiveText(archive, attachment.drawingRelsPath);
    if (!drawingXml || !relsXml) {
      continue;
    }

    const drawingDocument = parseXml(drawingXml);
    const relsDocument = parseXml(relsXml);
    if (!drawingDocument || !relsDocument) {
      continue;
    }

    const relationships = new Map<string, string>();
    for (const node of getLocalDescendants(relsDocument, "Relationship")) {
      const id = node.getAttribute("Id");
      const target = node.getAttribute("Target");
      if (id && target) {
        relationships.set(id, resolveRelationshipPath(attachment.drawingRelsPath ?? attachment.drawingPath, target));
      }
    }

    const anchorNodes = getLocalChildren(drawingDocument.documentElement, "twoCellAnchor")
      .concat(getLocalChildren(drawingDocument.documentElement, "oneCellAnchor"))
      .concat(getLocalChildren(drawingDocument.documentElement, "absoluteAnchor"));

    let chartAnchorIndex = 0;
    for (const anchorNode of anchorNodes) {
      const graphicFrame = getFirstLocalChild(anchorNode, "graphicFrame");
      const chartNode = graphicFrame ? getFirstLocalDescendant(graphicFrame, "chart") : null;
      const relationshipId = chartNode?.getAttributeNS("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "id")
        ?? chartNode?.getAttribute("r:id")
        ?? chartNode?.getAttribute("id");
      if (!relationshipId) {
        continue;
      }

      chartOrigins.push({
        anchorIndex: chartAnchorIndex,
        chartPath: relationships.get(relationshipId) ?? null,
        drawingPath: attachment.drawingPath,
        workbookSheetIndex: origin.workbookSheetIndex
      });
      chartAnchorIndex += 1;
    }
  }

  return chartOrigins;
}

function applyChartOrigins(
  chartsByWorkbookSheetIndex: XlsxChart[][],
  chartOriginsById: Map<string, WorkbookChartOrigin>,
  archive: Record<string, Uint8Array>,
  sheetOrigins: Array<WorkbookImageSheetOrigin | null>
) {
  for (let workbookSheetIndex = 0; workbookSheetIndex < chartsByWorkbookSheetIndex.length; workbookSheetIndex += 1) {
    const charts = chartsByWorkbookSheetIndex[workbookSheetIndex] ?? [];
    const origins = collectChartOriginsForSheet(archive, sheetOrigins[workbookSheetIndex] ?? null);
    charts.forEach((chart, index) => {
      const origin = origins[index];
      if (!origin) {
        return;
      }
      chart.chartPath = origin.chartPath ?? undefined;
      chartOriginsById.set(chart.id, origin);
    });
  }
}

export function loadWorkbookChartAssets(
  workbook: Workbook,
  imageAssets: Pick<WorkbookImageAssets, "archive" | "sheetOrigins" | "themePalette"> | null,
  visibleSheetIndexByWorkbookSheetIndex: Map<number, number>
): WorkbookChartAssets {
  const chartsByWorkbookSheetIndex = Array.from({ length: workbook.sheetCount }, (_, workbookSheetIndex) => {
    const worksheet = workbook.getSheet(workbookSheetIndex);
    const rawCharts = Array.isArray(worksheet.charts) ? worksheet.charts : [];
    const visibleSheetIndex = visibleSheetIndexByWorkbookSheetIndex.get(workbookSheetIndex) ?? workbookSheetIndex;

    return rawCharts.map((rawChart, chartIndex) => {
      const chartId = `chart-${workbookSheetIndex}-${chartIndex}`;
      const chart = rawChart && typeof rawChart === "object" ? rawChart as Record<string, unknown> : {};
      return {
        anchor: normalizeChartAnchor(chart.anchor),
        autoTitleDeleted: typeof chart.autoTitleDeleted === "boolean" ? chart.autoTitleDeleted : undefined,
        axes: Array.isArray(chart.axes) ? chart.axes.map(normalizeChartAxis).filter((value): value is XlsxChartAxis => Boolean(value)) : [],
        categoryAxis: normalizeChartAxis(chart.categoryAxis),
        chartColorPalette: undefined,
        chartColorPaletteOffset: undefined,
        chartPath: undefined,
        chartType: typeof chart.chartType === "string" ? chart.chartType : "ColumnClustered",
        dataLabels: normalizeChartDataLabels(chart.dataLabels),
        displayBlanksAs: typeof chart.displayBlanksAs === "string" ? chart.displayBlanksAs : undefined,
        editable: true,
        firstSliceAngle: typeof chart.firstSliceAngle === "number" ? chart.firstSliceAngle : undefined,
        gapWidth: typeof chart.gapWidth === "number" ? chart.gapWidth : undefined,
        holeSize: typeof chart.holeSize === "number" ? chart.holeSize : undefined,
        id: chartId,
        is3d: typeof chart.is3d === "boolean" ? chart.is3d : undefined,
        legend: normalizeLegend(chart.legend)
          ? {
              ...normalizeLegend(chart.legend),
              position: normalizeLegendPosition(normalizeLegend(chart.legend)?.position)
            }
          : null,
        name: typeof chart.name === "string" ? chart.name : undefined,
        overlap: typeof chart.overlap === "number" ? chart.overlap : undefined,
        plotVisibleOnly: typeof chart.plotVisibleOnly === "boolean" ? chart.plotVisibleOnly : undefined,
        raw: chart,
        radarStyle: typeof chart.radarStyle === "string" ? chart.radarStyle : undefined,
        roundedCorners: typeof chart.roundedCorners === "boolean" ? chart.roundedCorners : undefined,
        series: Array.isArray(chart.series)
          ? chart.series.map((entry, seriesIndex) => normalizeChartSeries(workbook, workbookSheetIndex, chartId, entry, seriesIndex))
          : [],
        sheetIndex: visibleSheetIndex,
        showDlblsOverMax: typeof chart.showDlblsOverMax === "boolean" ? chart.showDlblsOverMax : undefined,
        title: typeof chart.title === "string" ? chart.title : undefined,
        typeGroups: Array.isArray(chart.typeGroups) ? chart.typeGroups : [],
        valueAxis: normalizeChartAxis(chart.valueAxis),
        varyColors: typeof chart.varyColors === "boolean" ? chart.varyColors : undefined,
        workbookSheetIndex,
        zIndex: 200 + chartIndex
      } satisfies XlsxChart;
    });
  });

  const chartsheets = Array.isArray(workbook.chartsheets)
    ? workbook.chartsheets.map((entry, index) => normalizeChartsheet(entry, index))
    : [];
  const tabs = buildTabs(workbook, chartsheets, visibleSheetIndexByWorkbookSheetIndex);
  const chartOriginsById = new Map<string, WorkbookChartOrigin>();

  if (imageAssets) {
    applyChartOrigins(chartsByWorkbookSheetIndex, chartOriginsById, imageAssets.archive, imageAssets.sheetOrigins);
    for (const charts of chartsByWorkbookSheetIndex) {
      for (const chart of charts) {
        applyChartStyleFromXml(chart, chart.chartPath, imageAssets.archive, imageAssets.themePalette);
      }
    }
  }

  return {
    chartOriginsById,
    chartsByWorkbookSheetIndex,
    chartsheets,
    tabs
  };
}

function getChartAnchorNodes(drawingDocument: XMLDocument) {
  return getLocalChildren(drawingDocument.documentElement, "twoCellAnchor")
    .concat(getLocalChildren(drawingDocument.documentElement, "oneCellAnchor"))
    .concat(getLocalChildren(drawingDocument.documentElement, "absoluteAnchor"))
    .filter((anchorNode) => {
      const graphicFrame = getFirstLocalChild(anchorNode, "graphicFrame");
      return Boolean(graphicFrame && getFirstLocalDescendant(graphicFrame, "chart"));
    });
}

function updateMarkerNode(markerNode: Element | null, marker: { col: number; colOffsetEmu: number; row: number; rowOffsetEmu: number }) {
  if (!markerNode) {
    return;
  }

  setLeafValue(markerNode, "col", String(Math.max(0, Math.round(marker.col))));
  setLeafValue(markerNode, "colOff", String(Math.max(0, Math.round(marker.colOffsetEmu))));
  setLeafValue(markerNode, "row", String(Math.max(0, Math.round(marker.row))));
  setLeafValue(markerNode, "rowOff", String(Math.max(0, Math.round(marker.rowOffsetEmu))));
}

function updateAnchorNode(anchorNode: Element, anchor: XlsxImageAnchor) {
  if (anchor.kind === "two-cell") {
    updateMarkerNode(getFirstLocalChild(anchorNode, "from"), anchor.from);
    updateMarkerNode(getFirstLocalChild(anchorNode, "to"), anchor.to);
    return;
  }

  if (anchor.kind === "one-cell") {
    updateMarkerNode(getFirstLocalChild(anchorNode, "from"), anchor.from);
    const ext = getFirstLocalChild(anchorNode, "ext");
    if (ext) {
      ext.setAttribute("cx", String(Math.max(0, Math.round(anchor.sizeEmu.cx))));
      ext.setAttribute("cy", String(Math.max(0, Math.round(anchor.sizeEmu.cy))));
    }
    return;
  }

  const pos = getFirstLocalChild(anchorNode, "pos");
  if (pos) {
    pos.setAttribute("x", String(Math.max(0, Math.round(anchor.positionEmu.x))));
    pos.setAttribute("y", String(Math.max(0, Math.round(anchor.positionEmu.y))));
  }
  const ext = getFirstLocalChild(anchorNode, "ext");
  if (ext) {
    ext.setAttribute("cx", String(Math.max(0, Math.round(anchor.sizeEmu.cx))));
    ext.setAttribute("cy", String(Math.max(0, Math.round(anchor.sizeEmu.cy))));
  }
}

function setChartTitle(chartNode: Element, value: string | undefined) {
  const existing = getFirstLocalChild(chartNode, "title");
  if (!value) {
    existing?.remove();
    return;
  }

  const titleNode = existing ?? chartNode.insertBefore(
    chartNode.ownerDocument.createElementNS(CHART_NS, "c:title"),
    chartNode.firstChild
  );
  while (titleNode.firstChild) {
    titleNode.removeChild(titleNode.firstChild);
  }
  const tx = titleNode.ownerDocument.createElementNS(CHART_NS, "c:tx");
  const rich = titleNode.ownerDocument.createElementNS(CHART_NS, "c:rich");
  const bodyPr = titleNode.ownerDocument.createElementNS(DRAWINGML_NS, "a:bodyPr");
  const lstStyle = titleNode.ownerDocument.createElementNS(DRAWINGML_NS, "a:lstStyle");
  const p = titleNode.ownerDocument.createElementNS(DRAWINGML_NS, "a:p");
  const r = titleNode.ownerDocument.createElementNS(DRAWINGML_NS, "a:r");
  const t = titleNode.ownerDocument.createElementNS(DRAWINGML_NS, "a:t");
  t.textContent = value;
  r.appendChild(t);
  p.appendChild(r);
  rich.append(bodyPr, lstStyle, p);
  tx.appendChild(rich);
  titleNode.appendChild(tx);
}

function setRefFormula(parent: Element, refNodeName: string, formula: string | undefined) {
  if (!formula) {
    return;
  }

  const refNode = ensureChild(parent, refNodeName);
  setLeafValue(refNode, "f", formula);
}

function updateSeriesNodes(chartTypeNode: Element, chart: Partial<XlsxChart>) {
  if (!chart.series) {
    return;
  }

  const seriesNodes = getLocalDescendants(chartTypeNode, "ser");
  chart.series.forEach((series, index) => {
    const seriesNode = seriesNodes[index];
    if (!seriesNode) {
      return;
    }

    if (series.name !== undefined) {
      const tx = ensureChild(seriesNode, "tx");
      const strRef = ensureChild(tx, "strRef");
      setLeafValue(strRef, "f", series.name);
    }
    if (series.categoriesRef?.formula) {
      const target = chart.chartType === "ScatterLines" ? ensureChild(seriesNode, "xVal") : ensureChild(seriesNode, "cat");
      setRefFormula(target, "strRef", series.categoriesRef.formula);
    }
    if (series.valuesRef?.formula) {
      const target = chart.chartType === "ScatterLines" ? ensureChild(seriesNode, "yVal") : ensureChild(seriesNode, "val");
      setRefFormula(target, "numRef", series.valuesRef.formula);
    }
    if (series.invertIfNegative !== undefined) {
      setBooleanValue(seriesNode, "invertIfNegative", series.invertIfNegative);
    }
    if (series.smooth !== undefined) {
      setBooleanValue(seriesNode, "smooth", series.smooth);
    }
  });
}

function updateAxisNode(axisNode: Element | null, axis: XlsxChartAxis | null | undefined) {
  if (!axisNode || !axis) {
    return;
  }

  if (axis.position) {
    setLeafValue(ensureChild(axisNode, "axPos"), "val", axis.position);
    getFirstLocalChild(axisNode, "axPos")?.setAttribute("val", axis.position);
  }
  if (axis.majorGridlines !== undefined) {
    const gridlines = getFirstLocalChild(axisNode, "majorGridlines");
    if (axis.majorGridlines && !gridlines) {
      axisNode.appendChild(axisNode.ownerDocument.createElementNS(CHART_NS, "c:majorGridlines"));
    } else if (!axis.majorGridlines) {
      gridlines?.remove();
    }
  }
  if (axis.minorGridlines !== undefined) {
    const gridlines = getFirstLocalChild(axisNode, "minorGridlines");
    if (axis.minorGridlines && !gridlines) {
      axisNode.appendChild(axisNode.ownerDocument.createElementNS(CHART_NS, "c:minorGridlines"));
    } else if (!axis.minorGridlines) {
      gridlines?.remove();
    }
  }
  if (axis.majorTickMark) {
    getFirstLocalChild(axisNode, "majorTickMark")?.setAttribute("val", axis.majorTickMark)
      ?? setBooleanValue(axisNode, "majorTickMark", false).setAttribute("val", axis.majorTickMark);
  }
  if (axis.minorTickMark) {
    getFirstLocalChild(axisNode, "minorTickMark")?.setAttribute("val", axis.minorTickMark)
      ?? setBooleanValue(axisNode, "minorTickMark", false).setAttribute("val", axis.minorTickMark);
  }
  if (axis.labelPosition) {
    getFirstLocalChild(axisNode, "tickLblPos")?.setAttribute("val", axis.labelPosition)
      ?? setBooleanValue(axisNode, "tickLblPos", false).setAttribute("val", axis.labelPosition);
  }
  if (axis.crosses) {
    getFirstLocalChild(axisNode, "crosses")?.setAttribute("val", axis.crosses)
      ?? setBooleanValue(axisNode, "crosses", false).setAttribute("val", axis.crosses);
  }
  if (axis.crossBetween) {
    getFirstLocalChild(axisNode, "crossBetween")?.setAttribute("val", axis.crossBetween)
      ?? setBooleanValue(axisNode, "crossBetween", false).setAttribute("val", axis.crossBetween);
  }
  if (axis.delete !== undefined) {
    setBooleanValue(axisNode, "delete", axis.delete);
  }
  if (axis.numberFormat?.formatCode) {
    const numFmt = ensureChild(axisNode, "numFmt");
    numFmt.setAttribute("formatCode", axis.numberFormat.formatCode);
    if (axis.numberFormat.sourceLinked !== undefined) {
      numFmt.setAttribute("sourceLinked", axis.numberFormat.sourceLinked ? "1" : "0");
    }
  }
}

function updateDataLabels(chartTypeNode: Element, labels: XlsxChartDataLabels | null | undefined) {
  if (!labels) {
    return;
  }

  const labelsNode = ensureChild(chartTypeNode, "dLbls");
  if (labels.showLegendKey !== undefined) {
    setBooleanValue(labelsNode, "showLegendKey", labels.showLegendKey);
  }
  if (labels.showValue !== undefined) {
    setBooleanValue(labelsNode, "showVal", labels.showValue);
  }
  if (labels.showCategoryName !== undefined) {
    setBooleanValue(labelsNode, "showCatName", labels.showCategoryName);
  }
  if (labels.showSeriesName !== undefined) {
    setBooleanValue(labelsNode, "showSerName", labels.showSeriesName);
  }
  if (labels.showPercent !== undefined) {
    setBooleanValue(labelsNode, "showPercent", labels.showPercent);
  }
  if (labels.showBubbleSize !== undefined) {
    setBooleanValue(labelsNode, "showBubbleSize", labels.showBubbleSize);
  }
}

export function updateWorkbookChartAnchor(
  imageAssets: Pick<WorkbookImageAssets, "archive">,
  chartAssets: WorkbookChartAssets,
  chartId: string,
  anchor: XlsxImageAnchor
) {
  const origin = chartAssets.chartOriginsById.get(chartId);
  if (!origin) {
    return false;
  }

  const drawingXml = readArchiveText(imageAssets.archive, origin.drawingPath);
  if (!drawingXml) {
    return false;
  }

  const drawingDocument = parseXml(drawingXml);
  if (!drawingDocument) {
    return false;
  }

  const anchorNode = getChartAnchorNodes(drawingDocument)[origin.anchorIndex];
  if (!anchorNode) {
    return false;
  }

  updateAnchorNode(anchorNode, anchor);
  imageAssets.archive[normalizeArchivePath(origin.drawingPath)] = strToU8(serializeXml(drawingDocument));
  return true;
}

export function updateWorkbookChartDefinition(
  imageAssets: Pick<WorkbookImageAssets, "archive">,
  chartAssets: WorkbookChartAssets,
  chartId: string,
  patch: Partial<XlsxChart>
) {
  const origin = chartAssets.chartOriginsById.get(chartId);
  if (!origin?.chartPath) {
    return false;
  }

  const chartXml = readArchiveText(imageAssets.archive, origin.chartPath);
  if (!chartXml) {
    return false;
  }

  const chartDocument = parseXml(chartXml);
  if (!chartDocument) {
    return false;
  }

  const chartNode = getFirstLocalDescendant(chartDocument, "chart");
  const plotAreaNode = chartNode ? getFirstLocalChild(chartNode, "plotArea") : null;
  const chartTypeNode = plotAreaNode
    ? getLocalChildren(plotAreaNode, "barChart")[0]
      ?? getLocalChildren(plotAreaNode, "lineChart")[0]
      ?? getLocalChildren(plotAreaNode, "pieChart")[0]
      ?? getLocalChildren(plotAreaNode, "doughnutChart")[0]
      ?? getLocalChildren(plotAreaNode, "scatterChart")[0]
      ?? getLocalChildren(plotAreaNode, "areaChart")[0]
      ?? getLocalChildren(plotAreaNode, "radarChart")[0]
      ?? null
    : null;
  if (!chartNode || !plotAreaNode || !chartTypeNode) {
    return false;
  }

  if (patch.title !== undefined) {
    setChartTitle(chartNode, patch.title);
  }
  if (patch.displayBlanksAs) {
    const node = ensureChild(chartNode, "dispBlanksAs");
    node.setAttribute("val", patch.displayBlanksAs);
  }
  if (patch.roundedCorners !== undefined) {
    setBooleanValue(chartNode, "roundedCorners", patch.roundedCorners);
  }
  if (patch.showDlblsOverMax !== undefined) {
    setBooleanValue(chartNode, "showDLblsOverMax", patch.showDlblsOverMax);
  }
  if (patch.varyColors !== undefined) {
    setBooleanValue(chartTypeNode, "varyColors", patch.varyColors);
  }
  if (patch.gapWidth !== undefined) {
    setNumericValue(chartTypeNode, "gapWidth", patch.gapWidth);
  }
  if (patch.overlap !== undefined) {
    const overlapNode = ensureChild(chartTypeNode, "overlap");
    overlapNode.setAttribute("val", String(Math.round(patch.overlap)));
  }
  if (patch.dataLabels) {
    updateDataLabels(chartTypeNode, patch.dataLabels);
  }
  updateSeriesNodes(chartTypeNode, patch);
  updateAxisNode(getLocalChildren(plotAreaNode, "catAx")[0] ?? getLocalChildren(plotAreaNode, "serAx")[0] ?? null, patch.categoryAxis);
  updateAxisNode(getLocalChildren(plotAreaNode, "valAx")[0] ?? null, patch.valueAxis);

  imageAssets.archive[normalizeArchivePath(origin.chartPath)] = strToU8(serializeXml(chartDocument));
  return true;
}
