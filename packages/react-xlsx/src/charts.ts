import type { Workbook } from "@dukelib/sheets-wasm";
import { strFromU8, strToU8 } from "fflate";
import type { WorkbookImageAssets, WorkbookImageSheetOrigin } from "./images";
import type {
  XlsxChart,
  XlsxChartAxis,
  XlsxChartDataLabels,
  XlsxChartLegend,
  XlsxChartPointDataLabel,
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
const CHART_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart";
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
  anchor: XlsxImageAnchor | null;
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

type ChartStyleAppearance = {
  axisLabelColor?: string;
  axisLineColor?: string;
  chartAreaBorderColor?: string;
  chartAreaFillColor?: string;
  chartAreaNoFill?: boolean;
  paletteOffset?: number;
  textColor?: string;
  titleColor?: string;
};

function clampUnitInterval(value: number) {
  return Math.max(0, Math.min(1, value));
}

function isElementNode(node: Node | ChildNode | null | undefined): node is Element {
  return node != null && node.nodeType === 1;
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

function resolveColorFromXmlFragment(fragment: string, themePalette?: XlsxThemePalette | null) {
  if (!fragment) {
    return undefined;
  }

  const srgbMatch = fragment.match(/<a:srgbClr\b[^>]*\bval="([0-9a-fA-F]{6,8})"/i);
  if (srgbMatch?.[1]) {
    return normalizeHexColor(srgbMatch[1]) ?? undefined;
  }

  const schemeMatch = fragment.match(/<a:schemeClr\b[^>]*\bval="([^"]+)"[^>]*>([\s\S]*?)<\/a:schemeClr>/i)
    ?? fragment.match(/<a:schemeClr\b[^>]*\bval="([^"]+)"[^>]*/i);
  if (!schemeMatch?.[1]) {
    return undefined;
  }

  const baseColor = resolveThemeColor(schemeMatch[1], themePalette);
  if (!baseColor) {
    return undefined;
  }

  const transforms = schemeMatch[2] ?? "";
  let lightnessModifier = 1;
  let lightnessOffset = 0;
  for (const match of transforms.matchAll(/<a:(lumMod|lumOff|tint|shade)\b[^>]*\bval="(-?\d+(?:\.\d+)?)"/gi)) {
    const transform = match[1]?.toLowerCase();
    const rawValue = Number(match[2] ?? Number.NaN);
    if (!transform || !Number.isFinite(rawValue)) {
      continue;
    }
    if (transform === "lummod") {
      lightnessModifier *= rawValue / 100000;
    } else if (transform === "lumoff") {
      lightnessOffset += rawValue / 100000;
    } else if (transform === "tint") {
      lightnessOffset += (1 - lightnessOffset) * (rawValue / 100000);
    } else if (transform === "shade") {
      lightnessModifier *= rawValue / 100000;
    }
  }

  return applyLightnessTransform(baseColor, lightnessModifier, lightnessOffset) ?? undefined;
}

function readHexColorFromXmlFragment(
  fragment: string,
  preferLine = false,
  themePalette?: XlsxThemePalette | null
) {
  const source = preferLine
    ? fragment.match(/<a:ln\b[\s\S]*?<\/a:ln>/i)?.[0] ?? ""
    : fragment.match(/<a:solidFill\b[\s\S]*?<\/a:solidFill>/i)?.[0] ?? "";
  return resolveColorFromXmlFragment(source, themePalette);
}

type FallbackSeriesStyle = {
  color?: string;
  lineColor?: string;
};

function parseFallbackSeriesStylesFromChartXml(
  chartXml: string,
  themePalette?: XlsxThemePalette | null
): FallbackSeriesStyle[] {
  const seriesBlocks = chartXml.match(/<c:ser\b[\s\S]*?<\/c:ser>/gi) ?? [];
  if (seriesBlocks.length === 0) {
    return [];
  }

  return seriesBlocks.map((seriesBlock) => {
    const shapeBlock = seriesBlock.match(/<c:spPr\b[\s\S]*?<\/c:spPr>/i)?.[0] ?? "";
    return {
      color: readHexColorFromXmlFragment(shapeBlock, false, themePalette),
      lineColor: readHexColorFromXmlFragment(shapeBlock, true, themePalette)
    };
  });
}

function parseFallbackPointStylesFromChartXml(
  chartXml: string,
  themePalette?: XlsxThemePalette | null
): XlsxChartPointStyle[][] {
  const chartDocument = parseXml(chartXml);
  if (chartDocument) {
    const parsedSeriesStyles = getLocalDescendants(chartDocument, "ser").map((seriesNode) => {
      const styles: XlsxChartPointStyle[] = [];
      for (const dataPointNode of getLocalChildren(seriesNode, "dPt")) {
        const indexValue = readChartNumericAttribute(dataPointNode, "idx");
        if (indexValue === undefined) {
          continue;
        }
        const shapeProperties = getFirstLocalChild(dataPointNode, "spPr");
        const lineStyle = resolveChartLineStyle(shapeProperties, themePalette);
        styles.push({
          color: resolveChartFillColor(shapeProperties, themePalette) ?? undefined,
          explosion: readChartNumericAttribute(dataPointNode, "explosion"),
          index: indexValue,
          lineColor: lineStyle.color ?? undefined
        });
      }
      return styles;
    });
    if (parsedSeriesStyles.some((styles) => styles.length > 0)) {
      return parsedSeriesStyles;
    }
  }

  const seriesBlocks = chartXml.match(/<c:ser\b[\s\S]*?<\/c:ser>/gi) ?? [];
  if (seriesBlocks.length === 0) {
    return [];
  }

  return seriesBlocks.map((seriesBlock) => {
    const pointBlocks = seriesBlock.match(/<c:dPt\b[\s\S]*?<\/c:dPt>/gi) ?? [];
    if (pointBlocks.length === 0) {
      return [];
    }

    const styles: XlsxChartPointStyle[] = [];
    for (const pointBlock of pointBlocks) {
      const indexMatch = pointBlock.match(/<c:idx\b[^>]*\bval="(-?\d+)"/i);
      const index = indexMatch?.[1] ? Number(indexMatch[1]) : Number.NaN;
      if (!Number.isFinite(index)) {
        continue;
      }
      const explosionMatch = pointBlock.match(/<c:explosion\b[^>]*\bval="(-?\d+(?:\.\d+)?)"/i);
      const explosionValue = explosionMatch?.[1] ? Number(explosionMatch[1]) : Number.NaN;
      styles.push({
        color: readHexColorFromXmlFragment(pointBlock, false, themePalette),
        explosion: Number.isFinite(explosionValue) ? explosionValue : undefined,
        index,
        lineColor: readHexColorFromXmlFragment(pointBlock, true, themePalette)
      });
    }

    return styles;
  });
}

function parseNumericPointCacheFromXmlFragment(fragment: string) {
  const pointMatches = Array.from(fragment.matchAll(/<c:pt\b[^>]*\bidx="(-?\d+)"[^>]*>[\s\S]*?<c:v>([^<]*)<\/c:v>[\s\S]*?<\/c:pt>/gi));
  if (pointMatches.length === 0) {
    return [];
  }

  const explicitPointCountMatch = fragment.match(/<c:ptCount\b[^>]*\bval="(\d+)"/i);
  const explicitPointCount = explicitPointCountMatch?.[1] ? Number(explicitPointCountMatch[1]) : Number.NaN;
  const maxIndex = pointMatches.reduce((max, match) => {
    const current = Number(match[1] ?? Number.NaN);
    return Number.isFinite(current) ? Math.max(max, current) : max;
  }, -1);
  const pointCount = Math.max(
    pointMatches.length,
    Number.isFinite(explicitPointCount) ? explicitPointCount : 0,
    maxIndex + 1
  );
  const values = Array.from({ length: pointCount }, () => null as number | null);

  for (const match of pointMatches) {
    const index = Number(match[1] ?? Number.NaN);
    const rawValue = (match[2] ?? "").trim();
    const numericValue = Number(rawValue);
    if (!Number.isFinite(index) || index < 0 || !Number.isFinite(numericValue)) {
      continue;
    }
    values[index] = numericValue;
  }

  return values;
}

function parseFallbackBubbleSizesFromChartXml(chartXml: string): Array<Array<number | null>> {
  const seriesBlocks = chartXml.match(/<c:ser\b[\s\S]*?<\/c:ser>/gi) ?? [];
  if (seriesBlocks.length === 0) {
    return [];
  }

  return seriesBlocks.map((seriesBlock) => {
    const bubbleSizeBlock = seriesBlock.match(/<c:bubbleSize\b[\s\S]*?<\/c:bubbleSize>/i)?.[0] ?? "";
    if (!bubbleSizeBlock) {
      return [];
    }

    return parseNumericPointCacheFromXmlFragment(bubbleSizeBlock);
  });
}

function decodeChartXmlText(value: string) {
  return value
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&amp;/g, "&")
    .replace(/&quot;/g, "\"")
    .replace(/&apos;/g, "'");
}

function normalizeChartTitleForMatch(value: string | null | undefined) {
  return (value ?? "").trim().replace(/\s+/g, " ").toLowerCase();
}

function extractChartTitleFromXml(chartXml: string): string | null {
  const match = chartXml.match(/<c:title\b[\s\S]*?<a:t>([\s\S]*?)<\/a:t>/i);
  if (!match?.[1]) {
    return null;
  }
  const decoded = decodeChartXmlText(match[1]).trim();
  return decoded.length > 0 ? decoded : null;
}

function resolveArchiveFallbackBubbleSizes(
  archive: Record<string, Uint8Array>,
  preferredTitle: string | undefined
) {
  const preferred = normalizeChartTitleForMatch(preferredTitle);
  let bestScore = Number.NEGATIVE_INFINITY;
  let bestCandidate: Array<Array<number | null>> = [];

  for (const [path, bytes] of Object.entries(archive)) {
    if (!/\/charts\/chart\d+\.xml$/i.test(path)) {
      continue;
    }
    const chartXml = strFromU8(bytes);
    if (!/<c:bubbleChart\b/i.test(chartXml)) {
      continue;
    }
    const candidateBubbleSizes = parseFallbackBubbleSizesFromChartXml(chartXml);
    const hasCandidateValues = candidateBubbleSizes.some((seriesValues) => seriesValues.some((value) => value != null));
    if (!hasCandidateValues) {
      continue;
    }

    let score = 0;
    const candidateTitle = normalizeChartTitleForMatch(extractChartTitleFromXml(chartXml));
    if (preferred.length > 0 && candidateTitle.length > 0 && preferred === candidateTitle) {
      score += 100;
    }
    if (bestCandidate.length === 0) {
      score += 1;
    }

    if (score > bestScore) {
      bestScore = score;
      bestCandidate = candidateBubbleSizes;
      if (score >= 100) {
        break;
      }
    }
  }

  return bestCandidate;
}

function parseChartTypeFromXml(chartXml: string) {
  if (/<c:pie3DChart\b/i.test(chartXml)) {
    return "pie3DChart";
  }
  if (/<c:pieChart\b/i.test(chartXml)) {
    return "pieChart";
  }
  if (/<c:doughnutChart\b/i.test(chartXml)) {
    return "doughnutChart";
  }
  if (/<c:ofPieChart\b/i.test(chartXml)) {
    return "ofPieChart";
  }
  return "";
}

function resolveArchiveFallbackPointStyles(
  archive: Record<string, Uint8Array>,
  preferredTitle: string | undefined,
  preferredChartXmlType: string | undefined,
  themePalette?: XlsxThemePalette | null
) {
  const preferred = normalizeChartTitleForMatch(preferredTitle);
  const preferredType = (preferredChartXmlType ?? "").trim();
  let bestScore = Number.NEGATIVE_INFINITY;
  let bestCandidate: XlsxChartPointStyle[][] = [];

  for (const [path, bytes] of Object.entries(archive)) {
    if (!/\/charts\/chart\d+\.xml$/i.test(path)) {
      continue;
    }
    const chartXml = strFromU8(bytes);
    const candidateType = parseChartTypeFromXml(chartXml);
    if (!candidateType) {
      continue;
    }
    const candidatePointStyles = parseFallbackPointStylesFromChartXml(chartXml, themePalette);
    const hasCandidateValues = candidatePointStyles.some((seriesStyles) => seriesStyles.some((style) => (
      (typeof style.color === "string" && style.color.length > 0)
      || typeof style.explosion === "number"
    )));
    if (!hasCandidateValues) {
      continue;
    }

    let score = 0;
    const candidateTitle = normalizeChartTitleForMatch(extractChartTitleFromXml(chartXml));
    if (preferred.length > 0 && candidateTitle.length > 0 && preferred === candidateTitle) {
      score += 100;
    }
    if (preferredType && candidateType === preferredType) {
      score += 20;
    }
    if (bestCandidate.length === 0) {
      score += 1;
    }

    if (score > bestScore) {
      bestScore = score;
      bestCandidate = candidatePointStyles;
      if (score >= 120) {
        break;
      }
    }
  }

  return bestCandidate;
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

function resolveThemeTypeface(typeface: string | null, themePalette?: XlsxThemePalette | null) {
  if (!typeface) {
    return null;
  }
  if (typeface === "+mn-lt" || typeface === "+mn-ea" || typeface === "+mn-cs") {
    return themePalette?.minorLatinFont ?? null;
  }
  if (typeface === "+mj-lt" || typeface === "+mj-ea" || typeface === "+mj-cs") {
    return themePalette?.majorLatinFont ?? null;
  }
  return typeface;
}

function readChartTextTypeface(textPropertiesNode: Element | null, themePalette?: XlsxThemePalette | null) {
  if (!textPropertiesNode) {
    return null;
  }
  const defaultRunProperties = getFirstLocalDescendant(textPropertiesNode, "defRPr")
    ?? getFirstLocalDescendant(textPropertiesNode, "rPr");
  if (!defaultRunProperties) {
    return null;
  }
  const typeface = getFirstLocalChild(defaultRunProperties, "latin")?.getAttribute("typeface")
    ?? getFirstLocalChild(defaultRunProperties, "ea")?.getAttribute("typeface")
    ?? getFirstLocalChild(defaultRunProperties, "cs")?.getAttribute("typeface")
    ?? null;
  const resolved = resolveThemeTypeface(typeface, themePalette)?.trim() ?? "";
  return resolved.length > 0 ? resolved : null;
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
  for (const transformNode of Array.from(node.childNodes).filter(isElementNode)) {
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

function isChartColorElement(node: Element | null | undefined): node is Element {
  return Boolean(node && (node.localName === "schemeClr" || node.localName === "srgbClr" || node.localName === "sysClr"));
}

function findFirstChartColorElement(node: Element | null) {
  if (!node) {
    return null;
  }
  if (isChartColorElement(node)) {
    return node;
  }

  for (const localName of ["srgbClr", "schemeClr", "sysClr"]) {
    for (const candidate of getLocalDescendants(node, localName)) {
      if (isChartColorElement(candidate)) {
        return candidate;
      }
    }
  }

  return null;
}

function resolveChartFillColor(shapeNode: Element | null, themePalette?: XlsxThemePalette | null) {
  if (!shapeNode || getFirstLocalChild(shapeNode, "noFill")) {
    return null;
  }
  const solidFill = getFirstLocalChild(shapeNode, "solidFill");
  if (solidFill) {
    const colorNode = findFirstChartColorElement(Array.from(solidFill.childNodes).find(isElementNode) ?? null);
    return resolveChartColorNode(colorNode, themePalette);
  }

  const gradientFill = getFirstLocalChild(shapeNode, "gradFill");
  const gradientStops = gradientFill
    ? getLocalDescendants(gradientFill, "gs")
        .map((stopNode) => ({
          colorNode: Array.from(stopNode.childNodes).find(isElementNode) ?? null,
          position: Number(stopNode.getAttribute("pos") ?? Number.NaN)
        }))
        .filter((stop) => Boolean(stop.colorNode))
    : [];
  if (gradientStops.length === 0) {
    return null;
  }

  gradientStops.sort((left, right) => {
    const leftPos = Number.isFinite(left.position) ? left.position : 0;
    const rightPos = Number.isFinite(right.position) ? right.position : 0;
    return leftPos - rightPos;
  });
  const midpointStop = gradientStops.find((stop) => Number.isFinite(stop.position) && stop.position >= 50000)
    ?? gradientStops[Math.floor(gradientStops.length / 2)]
    ?? gradientStops[0];
  return resolveChartColorNode(midpointStop.colorNode, themePalette);
}

function resolveChartLineStyle(shapeNode: Element | null, themePalette?: XlsxThemePalette | null) {
  const lineNode = shapeNode?.localName === "ln" ? shapeNode : (shapeNode ? getFirstLocalChild(shapeNode, "ln") : null);
  if (!lineNode) {
    return { color: null, hidden: false, widthPx: undefined };
  }
  if (getFirstLocalChild(lineNode, "noFill")) {
    return { color: null, hidden: true, widthPx: undefined };
  }

  const solidFill = getFirstLocalChild(lineNode, "solidFill");
  const colorNode = solidFill ? findFirstChartColorElement(Array.from(solidFill.childNodes).find(isElementNode) ?? null) : null;
  const widthValue = Number(lineNode.getAttribute("w") ?? Number.NaN);
  return {
    color: resolveChartColorNode(colorNode, themePalette),
    hidden: false,
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

function readChartBooleanAttribute(parent: Element | null, localName: string) {
  const node = parent ? getFirstLocalChild(parent, localName) : null;
  if (!node) {
    return undefined;
  }
  const rawValue = node.getAttribute("val");
  if (rawValue == null) {
    return true;
  }
  if (rawValue === "1" || rawValue === "true") {
    return true;
  }
  if (rawValue === "0" || rawValue === "false") {
    return false;
  }
  return undefined;
}

function readChartLabelFontSizePt(textPropertiesNode: Element | null) {
  if (!textPropertiesNode) {
    return undefined;
  }
  const runPropertiesNode = getFirstLocalDescendant(textPropertiesNode, "defRPr")
    ?? getFirstLocalDescendant(textPropertiesNode, "rPr");
  const rawSize = Number(runPropertiesNode?.getAttribute("sz") ?? Number.NaN);
  if (!Number.isFinite(rawSize) || rawSize <= 0) {
    return undefined;
  }
  return rawSize / 100;
}

function parseChartPointDataLabelsFromXml(labelsNode: Element): XlsxChartPointDataLabel[] {
  const fallbackFontSizePt = readChartLabelFontSizePt(getFirstLocalChild(labelsNode, "txPr"));
  const labels: XlsxChartPointDataLabel[] = [];
  for (const pointLabelNode of getLocalChildren(labelsNode, "dLbl")) {
    const index = readChartNumericAttribute(pointLabelNode, "idx");
    if (typeof index !== "number" || !Number.isFinite(index)) {
      continue;
    }

    const layoutNode = getFirstLocalChild(pointLabelNode, "layout");
    const manualLayoutNode = getFirstLocalChild(layoutNode, "manualLayout");
    labels.push({
      deleted: readChartBooleanAttribute(pointLabelNode, "delete"),
      fontSizePt: readChartLabelFontSizePt(getFirstLocalChild(pointLabelNode, "txPr")) ?? fallbackFontSizePt,
      index,
      showBubbleSize: readChartBooleanAttribute(pointLabelNode, "showBubbleSize"),
      showCategoryName: readChartBooleanAttribute(pointLabelNode, "showCatName"),
      showPercent: readChartBooleanAttribute(pointLabelNode, "showPercent"),
      showSeriesName: readChartBooleanAttribute(pointLabelNode, "showSerName"),
      showValue: readChartBooleanAttribute(pointLabelNode, "showVal"),
      x: readChartNumericAttribute(manualLayoutNode, "x"),
      y: readChartNumericAttribute(manualLayoutNode, "y")
    });
  }
  return labels;
}

function parseChartDataLabelsFromXml(labelsNode: Element | null): XlsxChartDataLabels | null {
  if (!labelsNode) {
    return null;
  }

  const pointLabels = parseChartPointDataLabelsFromXml(labelsNode);
  const labels: XlsxChartDataLabels = {
    pointLabels: pointLabels.length > 0 ? pointLabels : undefined,
    raw: {},
    showBubbleSize: readChartBooleanAttribute(labelsNode, "showBubbleSize"),
    showCategoryName: readChartBooleanAttribute(labelsNode, "showCatName"),
    showLegendKey: readChartBooleanAttribute(labelsNode, "showLegendKey"),
    showPercent: readChartBooleanAttribute(labelsNode, "showPercent"),
    showSeriesName: readChartBooleanAttribute(labelsNode, "showSerName"),
    showValue: readChartBooleanAttribute(labelsNode, "showVal")
  };
  const hasValue = (
    labels.showBubbleSize !== undefined
    || labels.showCategoryName !== undefined
    || labels.showLegendKey !== undefined
    || labels.showPercent !== undefined
    || (labels.pointLabels?.length ?? 0) > 0
    || labels.showSeriesName !== undefined
    || labels.showValue !== undefined
  );
  return hasValue ? labels : null;
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
    .filter((child): child is Element => isElementNode(child) && child.localName !== "variation")
    .map((child) => (
      resolveChartColorNode(child, themePalette)
      ?? resolveChartColorNode(findFirstChartColorElement(child), themePalette)
    ))
    .filter((color): color is string => typeof color === "string");
}

function readChartStyleAppearance(
  archive: Record<string, Uint8Array>,
  stylePath: string | null | undefined,
  themePalette?: XlsxThemePalette | null
): ChartStyleAppearance {
  const styleXml = readArchiveText(archive, stylePath);
  if (!styleXml) {
    return {};
  }

  const styleDocument = parseXml(styleXml);
  if (!styleDocument) {
    return {};
  }

  const dataPointNode = getFirstLocalDescendant(styleDocument, "dataPoint");
  const fillRefNode = dataPointNode ? getFirstLocalChild(dataPointNode, "fillRef") : null;
  const index = Number(fillRefNode?.getAttribute("idx") ?? Number.NaN);
  const chartAreaNode = getFirstLocalDescendant(styleDocument, "chartArea");
  const chartAreaShapeProperties = chartAreaNode ? getFirstLocalChild(chartAreaNode, "spPr") : null;
  const chartAreaFontRef = chartAreaNode ? getFirstLocalChild(chartAreaNode, "fontRef") : null;
  const chartAreaFontColor = chartAreaFontRef
    ? resolveChartColorNode(Array.from(chartAreaFontRef.childNodes).find(isElementNode) ?? null, themePalette)
    : null;
  const titleNode = getFirstLocalDescendant(styleDocument, "title");
  const titleFontRef = titleNode ? getFirstLocalChild(titleNode, "fontRef") : null;
  const titleColor = titleFontRef
    ? resolveChartColorNode(Array.from(titleFontRef.childNodes).find(isElementNode) ?? null, themePalette)
    : null;
  const axisStyleNode = getFirstLocalDescendant(styleDocument, "categoryAxis")
    ?? getFirstLocalDescendant(styleDocument, "valueAxis");
  const axisShapeProperties = axisStyleNode ? getFirstLocalChild(axisStyleNode, "spPr") : null;
  const axisFontRef = axisStyleNode ? getFirstLocalChild(axisStyleNode, "fontRef") : null;
  const chartAreaNoFill = chartAreaShapeProperties ? getFirstLocalChild(chartAreaShapeProperties, "noFill") != null : false;

  return {
    axisLabelColor: axisFontRef
      ? resolveChartColorNode(Array.from(axisFontRef.childNodes).find(isElementNode) ?? null, themePalette) ?? undefined
      : undefined,
    axisLineColor: resolveChartLineStyle(axisShapeProperties, themePalette).color ?? undefined,
    chartAreaBorderColor: resolveChartLineStyle(chartAreaShapeProperties, themePalette).color ?? undefined,
    chartAreaFillColor: resolveChartFillColor(chartAreaShapeProperties, themePalette) ?? undefined,
    chartAreaNoFill,
    paletteOffset: Number.isFinite(index) ? index : undefined,
    textColor: chartAreaFontColor ?? undefined,
    titleColor: titleColor ?? chartAreaFontColor ?? undefined
  };
}

function buildThemeSeriesPalette(themePalette?: XlsxThemePalette | null) {
  const themeColors = [4, 5, 6, 7, 8, 9]
    .map((index) => themePalette?.colorsByIndex[index] ?? null)
    .filter((color): color is string => Boolean(color));
  return themeColors.length > 0 ? themeColors : SERIES_COLORS;
}

function applyBuiltinChartDefaults(chart: XlsxChart, themePalette?: XlsxThemePalette | null) {
  const textColor = themePalette?.colorsByIndex[1] ?? themePalette?.colorsByIndex[3] ?? null;
  const minorTypeface = themePalette?.minorLatinFont?.trim() || undefined;
  const derivedAxisColor = textColor ? applyLightnessTransform(textColor, 0.35, 0.55) : null;
  const derivedBorderColor = textColor
    ? applyLightnessTransform(textColor, chart.is3d ? 0.28 : 0.22, chart.is3d ? 0.6 : 0.7)
    : null;
  chart.chartAreaBorderColor = chart.chartAreaBorderColor ?? derivedBorderColor ?? undefined;
  chart.textColor = chart.textColor ?? textColor ?? undefined;
  chart.titleColor = chart.titleColor ?? textColor ?? undefined;
  chart.axisLabelColor = chart.axisLabelColor ?? derivedAxisColor ?? textColor ?? undefined;
  chart.axisLineColor = chart.axisLineColor ?? derivedAxisColor ?? textColor ?? undefined;
  chart.fontFamily = chart.fontFamily ?? minorTypeface;
  chart.titleFontFamily = chart.titleFontFamily ?? chart.fontFamily ?? minorTypeface;

  const seriesPalette = chart.chartColorPalette && chart.chartColorPalette.length > 0
    ? chart.chartColorPalette
    : buildThemeSeriesPalette(themePalette);
  if (!chart.chartColorPalette || chart.chartColorPalette.length === 0) {
    chart.chartColorPalette = seriesPalette;
  }

  chart.series = chart.series.map((series, index) => {
    const fallbackColor = seriesPalette[index % seriesPalette.length];
    return {
      ...series,
      color: series.color ?? series.lineColor ?? fallbackColor,
      lineColor: series.lineColor ?? series.color ?? fallbackColor,
      markerColor: series.markerColor ?? series.color ?? series.lineColor ?? fallbackColor,
      markerLineColor: series.markerLineColor ?? series.lineColor ?? series.color ?? fallbackColor
    };
  });
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
      explosion: readChartNumericAttribute(dataPointNode, "explosion"),
      index: indexValue,
      lineColor: lineStyle.color ?? undefined
    });
  }

  return pointStyles;
}

function parseInvertNegativeStyle(seriesNode: Element, themePalette?: XlsxThemePalette | null) {
  const invertNode = getFirstLocalDescendant(seriesNode, "invertSolidFillFmt");
  const shapeProperties = invertNode ? getFirstLocalChild(invertNode, "spPr") : null;
  if (!shapeProperties) {
    return {
      color: undefined,
      lineColor: undefined
    };
  }

  const lineStyle = resolveChartLineStyle(shapeProperties, themePalette);
  return {
    color: resolveChartFillColor(shapeProperties, themePalette) ?? undefined,
    lineColor: lineStyle.color ?? undefined
  };
}

function parseChartCacheValues(
  parentNode: Element | null,
  cacheName: "numCache" | "strCache",
  mode: "category" | "value"
): Array<number | string | null> | null {
  if (!parentNode) {
    return null;
  }

  const referenceNode = getFirstLocalChild(parentNode, "numRef")
    ?? getFirstLocalChild(parentNode, "strRef")
    ?? parentNode;
  const cacheNode = getFirstLocalChild(referenceNode, cacheName);
  if (!cacheNode) {
    return null;
  }

  const pointCount = readChartNumericAttribute(cacheNode, "ptCount");
  const pointNodes = getLocalChildren(cacheNode, "pt")
    .map((pointNode) => {
      const rawIndex = Number(pointNode.getAttribute("idx") ?? Number.NaN);
      return {
        index: Number.isFinite(rawIndex) ? rawIndex : 0,
        value: getFirstLocalChild(pointNode, "v")?.textContent ?? ""
      };
    })
    .sort((left, right) => left.index - right.index);

  if (pointNodes.length === 0) {
    return null;
  }

  const maxIndex = pointNodes.reduce((max, point) => Math.max(max, point.index), 0);
  const targetLength = Math.max(
    pointNodes.length,
    Number.isFinite(pointCount ?? Number.NaN) ? Number(pointCount) : 0,
    maxIndex + 1
  );
  const values = Array.from({ length: targetLength }, () => null as number | string | null);
  for (const point of pointNodes) {
    if (mode === "value") {
      values[point.index] = cellValueToNumber(point.value);
    } else {
      values[point.index] = point.value.length > 0 ? point.value : null;
    }
  }
  return values;
}

function parseChartMultiLevelCacheValues(
  parentNode: Element | null,
  mode: "category" | "value"
): Array<number | string | null> | null {
  if (!parentNode) {
    return null;
  }

  const referenceNode = getFirstLocalChild(parentNode, "multiLvlStrRef") ?? parentNode;
  const cacheNode = getFirstLocalChild(referenceNode, "multiLvlStrCache");
  if (!cacheNode) {
    return null;
  }

  const levelNodes = getLocalChildren(cacheNode, "lvl");
  if (levelNodes.length === 0) {
    return null;
  }

  const pointCount = readChartNumericAttribute(cacheNode, "ptCount");
  const primaryLevelNode = mode === "category"
    ? levelNodes[levelNodes.length - 1] ?? levelNodes[0]
    : levelNodes[0];
  const pointNodes = getLocalChildren(primaryLevelNode, "pt")
    .map((pointNode) => {
      const rawIndex = Number(pointNode.getAttribute("idx") ?? Number.NaN);
      return {
        index: Number.isFinite(rawIndex) ? rawIndex : 0,
        value: getFirstLocalChild(pointNode, "v")?.textContent ?? ""
      };
    })
    .sort((left, right) => left.index - right.index);

  if (pointNodes.length === 0) {
    return null;
  }

  const maxIndex = pointNodes.reduce((max, point) => Math.max(max, point.index), 0);
  const targetLength = Math.max(
    pointNodes.length,
    Number.isFinite(pointCount ?? Number.NaN) ? Number(pointCount) : 0,
    maxIndex + 1
  );
  const values = Array.from({ length: targetLength }, () => null as number | string | null);
  for (const point of pointNodes) {
    if (mode === "value") {
      values[point.index] = cellValueToNumber(point.value);
      continue;
    }
    values[point.index] = point.value.length > 0 ? point.value : null;
  }
  return values;
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
    const seriesExplosion = readChartNumericAttribute(seriesNode, "explosion");
    const invertNegativeStyle = parseInvertNegativeStyle(seriesNode, themePalette);
    const invertIfNegative = readChartBooleanAttribute(seriesNode, "invertIfNegative");
    const isScatterChart = chart.chartType === "ScatterLines" || chart.chartType === "ScatterSmooth" || chart.chartType === "Bubble";
    const cachedCategories = isScatterChart
      ? (
          parseChartCacheValues(getFirstLocalChild(seriesNode, "xVal"), "numCache", "value")
          ?? parseChartMultiLevelCacheValues(getFirstLocalChild(seriesNode, "xVal"), "category")
        )
      : parseChartCacheValues(getFirstLocalChild(seriesNode, "cat"), "strCache", "category");
    const cachedValues = isScatterChart
      ? parseChartCacheValues(getFirstLocalChild(seriesNode, "yVal"), "numCache", "value")
      : parseChartCacheValues(getFirstLocalChild(seriesNode, "val"), "numCache", "value");
    const cachedBubbleSizes = chart.chartType === "Bubble"
      ? parseChartCacheValues(getFirstLocalChild(seriesNode, "bubbleSize"), "numCache", "value")
      : null;
    const existingShapeProperties = series.shapeProperties && typeof series.shapeProperties === "object"
      ? series.shapeProperties as Record<string, unknown>
      : null;
    const rawFillColor = typeof existingShapeProperties?.solidFillHex === "string"
      ? normalizeHexColor(existingShapeProperties.solidFillHex)
      : null;
    const rawLineColor = typeof existingShapeProperties?.lineColorHex === "string"
      ? normalizeHexColor(existingShapeProperties.lineColorHex)
      : null;
    const resolvedLineColor = lineStyle.hidden
      ? undefined
      : rawLineColor ?? lineStyle.color ?? rawFillColor ?? fillColor ?? series.lineColor ?? series.color;

    const hasCategoryReference = typeof series.categoriesRef?.formula === "string" && series.categoriesRef.formula.length > 0;
    const hasValueReference = typeof series.valuesRef?.formula === "string" && series.valuesRef.formula.length > 0;
    const hasBubbleSizeReference = typeof series.bubbleSizeRef?.formula === "string" && series.bubbleSizeRef.formula.length > 0;

    return {
      ...series,
      bubbleSizes: !hasBubbleSizeReference && cachedBubbleSizes
        ? cachedBubbleSizes.map((value) => (typeof value === "number" && Number.isFinite(value) ? value : null))
        : series.bubbleSizes,
      categories: !hasCategoryReference && cachedCategories ? cachedCategories : series.categories,
      color: rawFillColor ?? rawLineColor ?? fillColor ?? lineStyle.color ?? series.color,
      dataPointStyles: pointStyles.length > 0 ? pointStyles : series.dataPointStyles,
      lineColor: resolvedLineColor,
      lineWidthPx: lineStyle.hidden ? undefined : lineStyle.widthPx ?? series.lineWidthPx,
      markerColor: rawFillColor
        ?? rawLineColor
        ?? resolveChartFillColor(markerShapeProperties, themePalette)
        ?? fillColor
        ?? lineStyle.color
        ?? undefined,
      markerLineColor: rawLineColor
        ?? rawFillColor
        ?? markerLineStyle.color
        ?? lineStyle.color
        ?? fillColor
        ?? undefined,
      markerSize: markerSize ?? series.markerSize,
      markerSymbol,
      smooth: readChartBooleanAttribute(seriesNode, "smooth") ?? series.smooth,
      invertIfNegative: invertIfNegative ?? series.invertIfNegative,
      shapeProperties: {
        ...series.shapeProperties,
        xmlExplosion: seriesExplosion ?? undefined,
        xmlFillColor: fillColor ?? undefined,
        xmlLineHidden: lineStyle.hidden ? true : undefined,
        xmlLineColor: lineStyle.color ?? undefined,
        xmlLineWidthPx: lineStyle.widthPx ?? undefined,
        xmlNegativeFillColor: invertNegativeStyle.color ?? undefined,
        xmlNegativeLineColor: invertNegativeStyle.lineColor ?? undefined
      },
      negativeColor: invertNegativeStyle.color ?? series.negativeColor,
      negativeLineColor: invertNegativeStyle.lineColor ?? series.negativeLineColor,
      values: !hasValueReference && cachedValues
        ? cachedValues.map((value) => (typeof value === "number" && Number.isFinite(value) ? value : null))
        : series.values
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
  const fallbackPointStylesBySeries = parseFallbackPointStylesFromChartXml(chartXml, themePalette);
  const fallbackSeriesStyles = parseFallbackSeriesStylesFromChartXml(chartXml, themePalette);
  const fallbackBubbleSizesBySeries = parseFallbackBubbleSizesFromChartXml(chartXml);
  const applyFallbackSeriesStyles = () => {
    if (fallbackBubbleSizesBySeries.length > 0) {
      chart.series = chart.series.map((series, seriesIndex) => {
        const fallbackBubbleSizes = fallbackBubbleSizesBySeries[seriesIndex] ?? [];
        if (fallbackBubbleSizes.length === 0) {
          return series;
        }

        const currentNumericPointCount = (series.bubbleSizes ?? []).filter(
          (value): value is number => typeof value === "number" && Number.isFinite(value)
        ).length;
        const fallbackNumericPointCount = fallbackBubbleSizes.filter(
          (value): value is number => typeof value === "number" && Number.isFinite(value)
        ).length;

        if (currentNumericPointCount >= fallbackNumericPointCount) {
          return series;
        }

        return {
          ...series,
          bubbleSizes: fallbackBubbleSizes
        };
      });
    }

    if (fallbackPointStylesBySeries.length > 0) {
      chart.series = chart.series.map((series, seriesIndex) => {
        const fallbackPointStyles = fallbackPointStylesBySeries[seriesIndex] ?? [];
        if (fallbackPointStyles.length === 0) {
          return series;
        }

        const existingByIndex = new Map((series.dataPointStyles ?? []).map((entry) => [entry.index, entry]));
        for (const fallbackStyle of fallbackPointStyles) {
          const existing = existingByIndex.get(fallbackStyle.index);
          existingByIndex.set(fallbackStyle.index, {
            color: existing?.color ?? fallbackStyle.color,
            explosion: existing?.explosion ?? fallbackStyle.explosion,
            index: fallbackStyle.index,
            lineColor: existing?.lineColor ?? fallbackStyle.lineColor
          });
        }

        return {
          ...series,
          dataPointStyles: Array.from(existingByIndex.values()).sort((left, right) => left.index - right.index)
        };
      });
    }

    if (fallbackSeriesStyles.length > 0) {
      chart.series = chart.series.map((series, seriesIndex) => {
        const fallbackStyle = fallbackSeriesStyles[seriesIndex];
        if (!fallbackStyle) {
          return series;
        }
        const fallbackColor = fallbackStyle.color ?? fallbackStyle.lineColor;
        return {
          ...series,
          color: series.color ?? fallbackColor,
          lineColor: series.lineColor ?? fallbackStyle.lineColor ?? fallbackColor,
          markerColor: series.markerColor ?? fallbackColor ?? series.color,
          markerLineColor: series.markerLineColor ?? fallbackStyle.lineColor ?? fallbackColor ?? series.lineColor
        };
      });
    }
  };

  const chartDocument = parseXml(chartXml);
  const chartNode = chartDocument ? getFirstLocalDescendant(chartDocument, "chart") : null;
  const plotAreaNode = chartNode ? getFirstLocalChild(chartNode, "plotArea") : null;
  const styleIdNode = chartDocument?.documentElement ? getFirstLocalChild(chartDocument.documentElement, "style") : null;
  const chartTypeNode = plotAreaNode
    ? getLocalChildren(plotAreaNode, "barChart")[0]
      ?? getLocalChildren(plotAreaNode, "lineChart")[0]
      ?? getLocalChildren(plotAreaNode, "pieChart")[0]
      ?? getLocalChildren(plotAreaNode, "doughnutChart")[0]
      ?? getLocalChildren(plotAreaNode, "scatterChart")[0]
      ?? getLocalChildren(plotAreaNode, "areaChart")[0]
      ?? getLocalChildren(plotAreaNode, "radarChart")[0]
      ?? getLocalChildren(plotAreaNode, "bar3DChart")[0]
      ?? getLocalChildren(plotAreaNode, "pie3DChart")[0]
      ?? getLocalChildren(plotAreaNode, "ofPieChart")[0]
      ?? getLocalChildren(plotAreaNode, "bubbleChart")[0]
      ?? getLocalChildren(plotAreaNode, "surfaceChart")[0]
      ?? getLocalChildren(plotAreaNode, "surface3DChart")[0]
      ?? getLocalChildren(plotAreaNode, "stockChart")[0]
      ?? null
    : null;

  if (!chartNode || !chartTypeNode) {
    applyFallbackSeriesStyles();
    applyBuiltinChartDefaults(chart, themePalette);
    return;
  }
  const plotArea = plotAreaNode;
  if (!plotArea) {
    applyFallbackSeriesStyles();
    applyBuiltinChartDefaults(chart, themePalette);
    return;
  }

  switch (chartTypeNode.localName) {
    case "barChart":
    case "bar3DChart": {
      const grouping = getFirstLocalChild(chartTypeNode, "grouping")?.getAttribute("val");
      const barDir = getFirstLocalChild(chartTypeNode, "barDir")?.getAttribute("val");
      const isHorizontalBar = barDir === "bar";
      chart.is3d = chartTypeNode.localName === "bar3DChart" ? true : chart.is3d;
      if (grouping === "percentStacked") {
        chart.chartType = isHorizontalBar ? "BarPercentStacked" : "ColumnPercentStacked";
      } else if (grouping === "stacked") {
        chart.chartType = isHorizontalBar ? "BarStacked" : "ColumnStacked";
      } else {
        chart.chartType = isHorizontalBar ? "BarClustered" : "ColumnClustered";
      }
      break;
    }
    case "areaChart": {
      const grouping = getFirstLocalChild(chartTypeNode, "grouping")?.getAttribute("val");
      if (grouping === "stacked") {
        chart.chartType = "AreaStacked";
      } else if (grouping === "percentStacked") {
        chart.chartType = "AreaPercentStacked";
      } else {
        chart.chartType = "Area";
      }
      break;
    }
    case "pie3DChart":
      chart.chartType = "Pie3D";
      chart.is3d = true;
      break;
    case "ofPieChart":
      chart.chartType = "BarOfPie";
      break;
    case "surfaceChart":
      chart.chartType = "Surface";
      chart.is3d = true;
      break;
    case "surface3DChart":
      chart.chartType = "Surface";
      chart.is3d = true;
      break;
    case "bubbleChart":
      chart.chartType = "Bubble";
      break;
    default:
      break;
  }

  const legendNode = getFirstLocalChild(chartNode, "legend");
  const legendPosition = legendNode ? getFirstLocalChild(legendNode, "legendPos")?.getAttribute("val") ?? undefined : undefined;
  const legendOverlay = legendNode ? getFirstLocalChild(legendNode, "overlay")?.getAttribute("val") : undefined;

  chart.legend = legendNode ? {
    overlay: legendOverlay === "1",
    position: normalizeLegendPosition(legendPosition),
    raw: chart.legend?.raw
  } : chart.legend;
  const plotVisibleOnly = readChartBooleanAttribute(chartNode, "plotVisOnly");
  if (plotVisibleOnly !== undefined) {
    chart.plotVisibleOnly = plotVisibleOnly;
  }
  chart.displayBlanksAs = getFirstLocalChild(chartNode, "dispBlanksAs")?.getAttribute("val") ?? chart.displayBlanksAs;
  const styleId = Number(styleIdNode?.getAttribute("val") ?? Number.NaN);
  chart.chartStyleId = Number.isFinite(styleId) ? styleId : chart.chartStyleId;
  chart.firstSliceAngle = readChartNumericAttribute(chartTypeNode, "firstSliceAng") ?? chart.firstSliceAngle;
  chart.gapWidth = readChartNumericAttribute(chartTypeNode, "gapWidth") ?? chart.gapWidth;
  chart.overlap = readChartNumericAttribute(chartTypeNode, "overlap") ?? chart.overlap;
  chart.bubbleScale = readChartNumericAttribute(chartTypeNode, "bubbleScale") ?? chart.bubbleScale;
  const bubble3dNode = getFirstLocalChild(chartTypeNode, "bubble3D");
  chart.bubble3d = bubble3dNode
    ? bubble3dNode.getAttribute("val") !== "0"
    : chart.bubble3d;
  chart.holeSize = readChartNumericAttribute(chartTypeNode, "holeSize") ?? chart.holeSize;
  chart.radarStyle = getFirstLocalChild(chartTypeNode, "radarStyle")?.getAttribute("val") ?? chart.radarStyle;
  chart.scatterStyle = getFirstLocalChild(chartTypeNode, "scatterStyle")?.getAttribute("val") ?? chart.scatterStyle;
  const wireframeNode = getFirstLocalChild(chartTypeNode, "wireframe");
  chart.wireframe = wireframeNode
    ? wireframeNode.getAttribute("val") !== "0"
    : chart.wireframe;
  const chartTypeDataLabels = parseChartDataLabelsFromXml(getFirstLocalChild(chartTypeNode, "dLbls"));
  const firstSeriesNode = getLocalChildren(chartTypeNode, "ser")[0] ?? null;
  const seriesDataLabels = parseChartDataLabelsFromXml(getFirstLocalChild(firstSeriesNode, "dLbls"));
  chart.dataLabels = chartTypeDataLabels ?? seriesDataLabels ?? chart.dataLabels;
  chart.raw = {
    ...(chart.raw ?? {}),
    bubble3d: chart.bubble3d,
    grouping: getFirstLocalChild(chartTypeNode, "grouping")?.getAttribute("val") ?? undefined,
    ofPieType: getFirstLocalChild(chartTypeNode, "ofPieType")?.getAttribute("val") ?? undefined,
    shape: getFirstLocalChild(chartTypeNode, "shape")?.getAttribute("val") ?? undefined,
    secondPieSize: readChartNumericAttribute(chartTypeNode, "secondPieSize"),
    scatterStyle: chart.scatterStyle,
    splitPos: readChartNumericAttribute(chartTypeNode, "splitPos"),
    splitType: getFirstLocalChild(chartTypeNode, "splitType")?.getAttribute("val") ?? undefined,
    xmlChartType: chartTypeNode.localName
  };
  const view3dNode = getFirstLocalDescendant(chartNode, "view3D");
  if (view3dNode) {
    chart.view3d = {
      depthPercent: readChartNumericAttribute(view3dNode, "depthPercent"),
      perspective: readChartNumericAttribute(view3dNode, "perspective"),
      rAngAx: getFirstLocalChild(view3dNode, "rAngAx")?.getAttribute("val") === "1",
      rotX: readChartNumericAttribute(view3dNode, "rotX"),
      rotY: readChartNumericAttribute(view3dNode, "rotY")
    };
  }

  const relationships = chartPath ? readChartRelationships(archive, chartPath) : new Map<string, string>();
  chart.chartColorPalette = readChartColorPalette(archive, relationships.get(CHART_COLOR_STYLE_REL_TYPE), themePalette);
  const styleAppearance = readChartStyleAppearance(
    archive,
    relationships.get(CHART_STYLE_REL_TYPE),
    themePalette
  );
  chart.axisLabelColor = styleAppearance.axisLabelColor ?? chart.axisLabelColor;
  chart.axisLineColor = styleAppearance.axisLineColor ?? chart.axisLineColor;
  chart.chartAreaBorderColor = styleAppearance.chartAreaBorderColor ?? chart.chartAreaBorderColor;
  chart.chartAreaFillColor = styleAppearance.chartAreaFillColor ?? chart.chartAreaFillColor;
  chart.chartColorPaletteOffset = styleAppearance.paletteOffset ?? chart.chartColorPaletteOffset;
  chart.textColor = styleAppearance.textColor ?? chart.textColor;
  chart.titleColor = styleAppearance.titleColor ?? chart.titleColor;
  const chartTextTypeface = readChartTextTypeface(getFirstLocalChild(chartNode, "txPr"), themePalette);
  const titleTypeface = readChartTextTypeface(getFirstLocalDescendant(chartNode, "title"), themePalette);
  chart.fontFamily = chartTextTypeface ?? chart.fontFamily;
  chart.titleFontFamily = titleTypeface ?? chart.titleFontFamily ?? chart.fontFamily;

  const chartAreaShapeProperties = chartDocument?.documentElement
    ? getFirstLocalChild(chartDocument.documentElement, "spPr")
    : null;
  const plotAreaShapeProperties = getFirstLocalChild(plotArea, "spPr");
  const chartAreaNoFill = chartAreaShapeProperties ? getFirstLocalChild(chartAreaShapeProperties, "noFill") != null : false;
  const plotAreaNoFill = plotAreaShapeProperties ? getFirstLocalChild(plotAreaShapeProperties, "noFill") != null : false;
  chart.raw = {
    ...(chart.raw ?? {}),
    chartAreaNoFill: styleAppearance.chartAreaNoFill === true || chartAreaNoFill,
    plotAreaNoFill
  };
  if (chartAreaShapeProperties) {
    const chartAreaFillColor = resolveChartFillColor(chartAreaShapeProperties, themePalette);
    if (chartAreaFillColor) {
      chart.chartAreaFillColor = chartAreaFillColor;
    } else if (getFirstLocalChild(chartAreaShapeProperties, "noFill")) {
      chart.chartAreaFillColor = "transparent";
    }
    const chartAreaLineStyle = resolveChartLineStyle(chartAreaShapeProperties, themePalette);
    if (chartAreaLineStyle.hidden) {
      chart.chartAreaBorderColor = "transparent";
    } else if (chartAreaLineStyle.color) {
      chart.chartAreaBorderColor = chartAreaLineStyle.color;
    }
  }
  if (!chart.chartAreaFillColor && (styleAppearance.chartAreaNoFill === true || plotAreaNoFill)) {
    chart.chartAreaFillColor = "transparent";
  }
  const categoryAxisNodes = getLocalChildren(plotArea, "catAx");
  const valueAxisNodes = getLocalChildren(plotArea, "valAx");
  const isScatterLikeChart = chart.chartType === "ScatterLines" || chart.chartType === "ScatterSmooth" || chart.chartType === "Bubble";
  let categoryAxisNode = categoryAxisNodes[0] ?? null;
  let valueAxisNode = valueAxisNodes[0] ?? null;
  if (!categoryAxisNode && isScatterLikeChart && valueAxisNodes.length >= 2) {
    categoryAxisNode = valueAxisNodes.find((axisNode) => {
      const position = getFirstLocalChild(axisNode, "axPos")?.getAttribute("val");
      return position === "b" || position === "t";
    }) ?? valueAxisNodes[0];
    valueAxisNode = valueAxisNodes.find((axisNode) => {
      const position = getFirstLocalChild(axisNode, "axPos")?.getAttribute("val");
      return position === "l" || position === "r";
    }) ?? valueAxisNodes[1] ?? valueAxisNodes[0];
  }
  chart.categoryAxis = mergeChartAxis(chart.categoryAxis, readChartAxisFromXml(categoryAxisNode));
  chart.valueAxis = mergeChartAxis(chart.valueAxis, readChartAxisFromXml(valueAxisNode));
  chart.axes = chart.axes.length > 0
    ? chart.axes.map((axis, index) => (
      index === 0 && categoryAxisNode
        ? { ...axis, ...readChartAxisFromXml(categoryAxisNode) }
        : index === 1 && valueAxisNode
          ? { ...axis, ...readChartAxisFromXml(valueAxisNode) }
          : axis
    ))
    : chart.axes;

  applyChartSeriesStyleFromXml(chart, chartTypeNode, themePalette);
  applyFallbackSeriesStyles();
  if (chart.chartType === "Bubble") {
    const archiveFallbackBubbleSizes = resolveArchiveFallbackBubbleSizes(archive, chart.title);
    if (archiveFallbackBubbleSizes.length > 0) {
      chart.series = chart.series.map((series, seriesIndex) => {
        const pointCount = Math.max(series.values.length, series.categories.length);
        if (pointCount <= 1) {
          return series;
        }

        const numericBubbleCount = (series.bubbleSizes ?? []).filter(
          (value): value is number => typeof value === "number" && Number.isFinite(value)
        ).length;
        if (numericBubbleCount >= pointCount) {
          return series;
        }

        const fallbackCandidate = archiveFallbackBubbleSizes[seriesIndex] ?? archiveFallbackBubbleSizes[0] ?? [];
        const fallbackNumericCount = fallbackCandidate.filter(
          (value): value is number => typeof value === "number" && Number.isFinite(value)
        ).length;
        if (fallbackNumericCount < pointCount) {
          return series;
        }

        return {
          ...series,
          bubbleSizes: fallbackCandidate
        };
      });
    }
  }
  if (chart.chartType === "Pie" || chart.chartType === "Pie3D" || chart.chartType === "PieExploded" || chart.chartType === "Doughnut" || chart.chartType === "BarOfPie") {
    const needsPointColorFallback = chart.series.some((series) => {
      const pointCount = Math.max(series.values.length, series.categories.length);
      if (pointCount <= 0) {
        return false;
      }
      const coloredPointCount = (series.dataPointStyles ?? []).filter(
        (style) => typeof style.color === "string" && style.color.length > 0
      ).length;
      return coloredPointCount === 0;
    });
    if (needsPointColorFallback) {
      const archiveFallbackPointStyles = resolveArchiveFallbackPointStyles(
        archive,
        chart.title,
        chartTypeNode.localName,
        themePalette
      );
      if (archiveFallbackPointStyles.length > 0) {
        chart.series = chart.series.map((series, seriesIndex) => {
          const fallbackStyles = archiveFallbackPointStyles[seriesIndex] ?? archiveFallbackPointStyles[0] ?? [];
          if (fallbackStyles.length === 0) {
            return series;
          }
          const existingByIndex = new Map((series.dataPointStyles ?? []).map((entry) => [entry.index, entry]));
          for (const fallbackStyle of fallbackStyles) {
            const existing = existingByIndex.get(fallbackStyle.index);
            existingByIndex.set(fallbackStyle.index, {
              color: existing?.color ?? fallbackStyle.color,
              explosion: existing?.explosion ?? fallbackStyle.explosion,
              index: fallbackStyle.index,
              lineColor: existing?.lineColor ?? fallbackStyle.lineColor
            });
          }
          return {
            ...series,
            dataPointStyles: Array.from(existingByIndex.values()).sort((left, right) => left.index - right.index)
          };
        });
      }
    }
  }
  applyBuiltinChartDefaults(chart, themePalette);
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
  const normalizedBasePath = normalizeArchivePath(basePath);
  let baseDirectory = dirname(normalizedBasePath);
  if (normalizedBasePath.endsWith(".rels")) {
    const relsMarker = "/_rels/";
    const relsMarkerIndex = normalizedBasePath.lastIndexOf(relsMarker);
    if (relsMarkerIndex >= 0) {
      const ownerPrefix = normalizedBasePath.slice(0, relsMarkerIndex);
      const relFileName = normalizedBasePath.slice(relsMarkerIndex + relsMarker.length);
      const ownerFileName = relFileName.endsWith(".rels")
        ? relFileName.slice(0, -".rels".length)
        : relFileName;
      baseDirectory = dirname(`${ownerPrefix}/${ownerFileName}`);
    }
  }

  const segments = [...baseDirectory.split("/").filter(Boolean), ...normalizedTarget.split("/").filter(Boolean)];
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
    logBase: typeof axis.logBase === "number" ? axis.logBase : undefined,
    orientation: typeof axis.orientation === "string" ? axis.orientation : undefined,
    majorUnit: typeof axis.majorUnit === "number" ? axis.majorUnit : undefined,
    max: typeof axis.max === "number" ? axis.max : undefined,
    min: typeof axis.min === "number" ? axis.min : undefined,
    majorGridlines: typeof axis.majorGridlines === "boolean" ? axis.majorGridlines : undefined,
    majorTickMark: typeof axis.majorTickMark === "string" ? axis.majorTickMark : undefined,
    minorUnit: typeof axis.minorUnit === "number" ? axis.minorUnit : undefined,
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

function mergeChartAxis(target: XlsxChartAxis | null | undefined, patch: Partial<XlsxChartAxis> | null | undefined) {
  if (!patch) {
    return target ?? null;
  }
  return {
    ...(target ?? {}),
    ...patch
  };
}

function readChartAxisFromXml(axisNode: Element | null): Partial<XlsxChartAxis> | null {
  if (!axisNode) {
    return null;
  }

  const numFmt = getFirstLocalChild(axisNode, "numFmt");
  const scalingNode = getFirstLocalChild(axisNode, "scaling");
  return {
    crosses: getFirstLocalChild(axisNode, "crosses")?.getAttribute("val") ?? undefined,
    crossBetween: getFirstLocalChild(axisNode, "crossBetween")?.getAttribute("val") ?? undefined,
    delete: getFirstLocalChild(axisNode, "delete")?.getAttribute("val") === "1"
      ? true
      : getFirstLocalChild(axisNode, "delete")?.getAttribute("val") === "0"
        ? false
        : undefined,
    labelPosition: getFirstLocalChild(axisNode, "tickLblPos")?.getAttribute("val") ?? undefined,
    logBase: readChartNumericAttribute(getFirstLocalChild(axisNode, "scaling"), "logBase"),
    orientation: getFirstLocalChild(scalingNode ?? axisNode, "orientation")?.getAttribute("val") ?? undefined,
    majorGridlines: Boolean(getFirstLocalChild(axisNode, "majorGridlines")),
    majorTickMark: getFirstLocalChild(axisNode, "majorTickMark")?.getAttribute("val") ?? undefined,
    majorUnit: readChartNumericAttribute(axisNode, "majorUnit"),
    max: readChartNumericAttribute(scalingNode, "max"),
    min: readChartNumericAttribute(scalingNode, "min"),
    minorGridlines: Boolean(getFirstLocalChild(axisNode, "minorGridlines")),
    minorTickMark: getFirstLocalChild(axisNode, "minorTickMark")?.getAttribute("val") ?? undefined,
    minorUnit: readChartNumericAttribute(axisNode, "minorUnit"),
    numberFormat: numFmt
      ? {
          formatCode: numFmt.getAttribute("formatCode") ?? undefined,
          sourceLinked: numFmt.getAttribute("sourceLinked") === "1"
            ? true
            : numFmt.getAttribute("sourceLinked") === "0"
              ? false
              : undefined
        }
      : undefined,
    position: getFirstLocalChild(axisNode, "axPos")?.getAttribute("val") ?? undefined
  };
}

function normalizeChartDataLabels(raw: unknown): XlsxChartDataLabels | null {
  if (!raw || typeof raw !== "object") {
    return null;
  }

  const labels = raw as Record<string, unknown>;
  const pointLabels = Array.isArray(labels.pointLabels)
    ? (() => {
        const normalized: XlsxChartPointDataLabel[] = [];
        for (const entry of labels.pointLabels) {
          if (!entry || typeof entry !== "object") {
            continue;
          }
          const point = entry as Record<string, unknown>;
          const index = typeof point.index === "number" && Number.isFinite(point.index)
            ? point.index
            : null;
          if (index == null) {
            continue;
          }

          const nextPoint: XlsxChartPointDataLabel = { index };
          if (typeof point.deleted === "boolean") {
            nextPoint.deleted = point.deleted;
          }
          if (typeof point.fontSizePt === "number" && Number.isFinite(point.fontSizePt)) {
            nextPoint.fontSizePt = point.fontSizePt;
          }
          if (typeof point.showBubbleSize === "boolean") {
            nextPoint.showBubbleSize = point.showBubbleSize;
          }
          if (typeof point.showCategoryName === "boolean") {
            nextPoint.showCategoryName = point.showCategoryName;
          }
          if (typeof point.showPercent === "boolean") {
            nextPoint.showPercent = point.showPercent;
          }
          if (typeof point.showSeriesName === "boolean") {
            nextPoint.showSeriesName = point.showSeriesName;
          }
          if (typeof point.showValue === "boolean") {
            nextPoint.showValue = point.showValue;
          }
          if (typeof point.x === "number" && Number.isFinite(point.x)) {
            nextPoint.x = point.x;
          }
          if (typeof point.y === "number" && Number.isFinite(point.y)) {
            nextPoint.y = point.y;
          }
          normalized.push(nextPoint);
        }
        return normalized;
      })()
    : undefined;
  return {
    pointLabels: pointLabels && pointLabels.length > 0 ? pointLabels : undefined,
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
  const fromCol = typeof anchor.fromCol === "number" ? anchor.fromCol : 0;
  const fromColOffsetEmu = typeof anchor.fromColOffset === "number" ? anchor.fromColOffset : 0;
  const fromRow = typeof anchor.fromRow === "number" ? anchor.fromRow : 0;
  const fromRowOffsetEmu = typeof anchor.fromRowOffset === "number" ? anchor.fromRowOffset : 0;
  const rawToCol = typeof anchor.toCol === "number" ? anchor.toCol : null;
  const rawToColOffsetEmu = typeof anchor.toColOffset === "number" ? anchor.toColOffset : 0;
  const rawToRow = typeof anchor.toRow === "number" ? anchor.toRow : null;
  const rawToRowOffsetEmu = typeof anchor.toRowOffset === "number" ? anchor.toRowOffset : 0;
  const hasExplicitTo = rawToCol !== null && rawToRow !== null;
  const collapsedWidth = hasExplicitTo && (
    rawToCol < fromCol ||
    (rawToCol === fromCol && rawToColOffsetEmu <= fromColOffsetEmu)
  );
  const collapsedHeight = hasExplicitTo && (
    rawToRow < fromRow ||
    (rawToRow === fromRow && rawToRowOffsetEmu <= fromRowOffsetEmu)
  );
  const fallbackToCol = Math.max(fromCol + 8, 8);
  const fallbackToRow = Math.max(fromRow + 15, 15);

  return {
    kind: "two-cell",
    from: {
      col: fromCol,
      colOffsetEmu: fromColOffsetEmu,
      row: fromRow,
      rowOffsetEmu: fromRowOffsetEmu
    },
    to: {
      col: !hasExplicitTo || collapsedWidth ? fallbackToCol : rawToCol,
      colOffsetEmu: !hasExplicitTo || collapsedWidth ? 0 : rawToColOffsetEmu,
      row: !hasExplicitTo || collapsedHeight ? fallbackToRow : rawToRow,
      rowOffsetEmu: !hasExplicitTo || collapsedHeight ? 0 : rawToRowOffsetEmu
    }
  };
}

function parseMarkerNode(node: Element | null) {
  if (!node) {
    return null;
  }

  const col = Number(getFirstLocalChild(node, "col")?.textContent ?? Number.NaN);
  const row = Number(getFirstLocalChild(node, "row")?.textContent ?? Number.NaN);
  const colOffsetEmu = Number(getFirstLocalChild(node, "colOff")?.textContent ?? 0);
  const rowOffsetEmu = Number(getFirstLocalChild(node, "rowOff")?.textContent ?? 0);

  if (!Number.isFinite(col) || !Number.isFinite(row)) {
    return null;
  }

  return {
    col: Math.max(0, Math.round(col)),
    colOffsetEmu: Number.isFinite(colOffsetEmu) ? Math.max(0, Math.round(colOffsetEmu)) : 0,
    row: Math.max(0, Math.round(row)),
    rowOffsetEmu: Number.isFinite(rowOffsetEmu) ? Math.max(0, Math.round(rowOffsetEmu)) : 0
  };
}

function parseChartAnchorNode(anchorNode: Element): XlsxImageAnchor | null {
  if (anchorNode.localName === "twoCellAnchor") {
    const from = parseMarkerNode(getFirstLocalChild(anchorNode, "from"));
    const to = parseMarkerNode(getFirstLocalChild(anchorNode, "to"));
    return from && to ? { from, kind: "two-cell", to } : null;
  }

  if (anchorNode.localName === "oneCellAnchor") {
    const from = parseMarkerNode(getFirstLocalChild(anchorNode, "from"));
    const ext = getFirstLocalChild(anchorNode, "ext");
    const cx = Number(ext?.getAttribute("cx") ?? Number.NaN);
    const cy = Number(ext?.getAttribute("cy") ?? Number.NaN);
    return from && Number.isFinite(cx) && Number.isFinite(cy)
      ? {
          from,
          kind: "one-cell",
          sizeEmu: {
            cx: Math.max(0, Math.round(cx)),
            cy: Math.max(0, Math.round(cy))
          }
        }
      : null;
  }

  const pos = getFirstLocalChild(anchorNode, "pos");
  const ext = getFirstLocalChild(anchorNode, "ext");
  const x = Number(pos?.getAttribute("x") ?? Number.NaN);
  const y = Number(pos?.getAttribute("y") ?? Number.NaN);
  const cx = Number(ext?.getAttribute("cx") ?? Number.NaN);
  const cy = Number(ext?.getAttribute("cy") ?? Number.NaN);
  return Number.isFinite(x) && Number.isFinite(y) && Number.isFinite(cx) && Number.isFinite(cy)
    ? {
        kind: "absolute",
        positionEmu: {
          x: Math.round(x),
          y: Math.round(y)
        },
        sizeEmu: {
          cx: Math.max(0, Math.round(cx)),
          cy: Math.max(0, Math.round(cy))
        }
      }
    : null;
}

function isCollapsedChartAnchor(anchor: XlsxImageAnchor) {
  if (anchor.kind !== "two-cell") {
    return false;
  }

  const collapsedWidth = anchor.to.col < anchor.from.col
    || (anchor.to.col === anchor.from.col && anchor.to.colOffsetEmu <= anchor.from.colOffsetEmu);
  const collapsedHeight = anchor.to.row < anchor.from.row
    || (anchor.to.row === anchor.from.row && anchor.to.rowOffsetEmu <= anchor.from.rowOffsetEmu);
  return collapsedWidth || collapsedHeight;
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
  const shapeProperties = series.shapeProperties && typeof series.shapeProperties === "object"
    ? series.shapeProperties as Record<string, unknown>
    : undefined;
  const rawFillColor = typeof shapeProperties?.solidFillHex === "string"
    ? normalizeHexColor(shapeProperties.solidFillHex)
    : null;
  const rawLineColor = typeof shapeProperties?.lineColorHex === "string"
    ? normalizeHexColor(shapeProperties.lineColorHex)
    : null;
  const bubbleSizeRef = normalizeChartReference(series.bubbleSize ?? series.bubbleSizes ?? series.bubbles);

  return {
    bubbleSizeRef,
    bubbleSizes: resolveReferenceValues(workbook, workbookSheetIndex, bubbleSizeRef, "value").map((value) => (
      typeof value === "number" && Number.isFinite(value) ? value : null
    )),
    categories: resolveReferenceValues(workbook, workbookSheetIndex, categoriesRef, "category"),
    categoriesRef,
    color: rawFillColor ?? undefined,
    dataPoints: Array.isArray(series.dataPoints) ? series.dataPoints : [],
    dataPointStyles: undefined,
    id: `${chartId}-series-${index}`,
    invertIfNegative: typeof series.invertIfNegative === "boolean" ? series.invertIfNegative : undefined,
    lineColor: rawLineColor ?? rawFillColor ?? undefined,
    lineWidthPx: typeof shapeProperties?.lineWidth === "number"
      ? Math.max(1, Number(shapeProperties.lineWidth) / EMU_PER_PIXEL)
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
    negativeColor: undefined,
    negativeLineColor: undefined,
    raw: series,
    shapeProperties,
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

    const relationships = new Map<string, { target: string; type: string | null }>();
    for (const node of getLocalDescendants(relsDocument, "Relationship")) {
      const id = node.getAttribute("Id");
      const target = node.getAttribute("Target");
      const type = node.getAttribute("Type");
      if (id && target) {
        relationships.set(id, {
          target: resolveRelationshipPath(attachment.drawingRelsPath ?? attachment.drawingPath, target),
          type
        });
      }
    }

    const anchorNodes = Array.from(drawingDocument.documentElement.childNodes).filter(
      (node): node is Element => (
        node.nodeType === Node.ELEMENT_NODE
        && (
          (node as Element).localName === "twoCellAnchor"
          || (node as Element).localName === "oneCellAnchor"
          || (node as Element).localName === "absoluteAnchor"
        )
      )
    );

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
      const relationship = relationships.get(relationshipId);
      if (!relationship || relationship.type !== CHART_REL_TYPE) {
        continue;
      }

      chartOrigins.push({
        anchorIndex: chartAnchorIndex,
        anchor: parseChartAnchorNode(anchorNode),
        chartPath: relationship.target,
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
      if (origin.anchor && isCollapsedChartAnchor(chart.anchor)) {
        chart.anchor = origin.anchor;
      } else if (origin.anchor && chart.anchor.kind === "two-cell" && chart.anchor.from.col === 0 && chart.anchor.from.row === 0) {
        chart.anchor = origin.anchor;
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
      const rawSeries = Array.isArray(chart.series) ? chart.series : [];
      const chartLevelDataLabels = normalizeChartDataLabels(chart.dataLabels);
      const firstSeriesDataLabels = rawSeries.length > 0 && rawSeries[0] && typeof rawSeries[0] === "object"
        ? normalizeChartDataLabels((rawSeries[0] as Record<string, unknown>).dataLabels)
        : null;
      return {
        anchor: normalizeChartAnchor(chart.anchor),
        autoTitleDeleted: typeof chart.autoTitleDeleted === "boolean" ? chart.autoTitleDeleted : undefined,
        axes: Array.isArray(chart.axes) ? chart.axes.map(normalizeChartAxis).filter((value): value is XlsxChartAxis => Boolean(value)) : [],
        axisLabelColor: undefined,
        axisLineColor: undefined,
        categoryAxis: normalizeChartAxis(chart.categoryAxis),
        chartAreaBorderColor: undefined,
        chartAreaFillColor: undefined,
        chartColorPalette: undefined,
        chartColorPaletteOffset: undefined,
        chartPath: undefined,
        chartStyleId: undefined,
        chartType: typeof chart.chartType === "string" ? chart.chartType : "ColumnClustered",
        dataLabels: chartLevelDataLabels ?? firstSeriesDataLabels,
        displayBlanksAs: typeof chart.displayBlanksAs === "string" ? chart.displayBlanksAs : undefined,
        editable: true,
        firstSliceAngle: typeof chart.firstSliceAngle === "number" ? chart.firstSliceAngle : undefined,
        fontFamily: undefined,
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
        scatterStyle: typeof chart.scatterStyle === "string" ? chart.scatterStyle : undefined,
        roundedCorners: typeof chart.roundedCorners === "boolean" ? chart.roundedCorners : undefined,
        series: rawSeries.map((entry, seriesIndex) => normalizeChartSeries(workbook, workbookSheetIndex, chartId, entry, seriesIndex)),
        sheetIndex: visibleSheetIndex,
        showDlblsOverMax: typeof chart.showDlblsOverMax === "boolean" ? chart.showDlblsOverMax : undefined,
        bubbleScale: typeof chart.bubbleScale === "number" ? chart.bubbleScale : undefined,
        bubble3d: typeof chart.bubble3d === "boolean" ? chart.bubble3d : undefined,
        textColor: undefined,
        title: typeof chart.title === "string" ? chart.title : undefined,
        titleColor: undefined,
        titleFontFamily: undefined,
        typeGroups: Array.isArray(chart.typeGroups) ? chart.typeGroups : [],
        valueAxis: normalizeChartAxis(chart.valueAxis),
        varyColors: typeof chart.varyColors === "boolean" ? chart.varyColors : undefined,
        view3d: undefined,
        wireframe: typeof chart.wireframe === "boolean" ? chart.wireframe : undefined,
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
        applyBuiltinChartDefaults(chart, imageAssets.themePalette);
      }
    }
  } else {
    for (const charts of chartsByWorkbookSheetIndex) {
      for (const chart of charts) {
        applyBuiltinChartDefaults(chart, null);
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
  return Array.from(drawingDocument.documentElement.childNodes).filter(
    (node): node is Element => (
      node.nodeType === Node.ELEMENT_NODE
      && (
        (node as Element).localName === "twoCellAnchor"
        || (node as Element).localName === "oneCellAnchor"
        || (node as Element).localName === "absoluteAnchor"
      )
    )
  )
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
