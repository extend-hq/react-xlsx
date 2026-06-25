import * as React from "react";
import type { Workbook, Worksheet } from "@dukelib/sheets-wasm";
import {
  type VirtualItem,
  useVirtualizer
} from "@tanstack/react-virtual";
import { resolveWorkbookColor, resolveWorkbookFillStyle } from "./colors";
import { useXlsxViewerController, XlsxFileSizeLimitExceededError } from "./controller";
import { MemoChartSvg } from "./chart-renderer";
import {
  emuToPixels,
  resizeImageRect,
  resolveRenderedSheetAxisPixels,
  resolveSheetColumnWidthPixels,
  resolveSheetRowHeightPixels
} from "./images";
import type {
  XlsxChart,
  XlsxChartElementSelection,
  XlsxChartsheet,
  XlsxCellAddress,
  XlsxCellRange,
  XlsxFormControl,
  XlsxImage,
  XlsxImageRect,
  XlsxImageRenderProps,
  XlsxImageResizeHandlePosition,
  XlsxImageSelectionRenderProps,
  XlsxScrollerRenderProps,
  XlsxShape,
  XlsxSheetData,
  XlsxSheetThumbnail,
  XlsxThemePalette,
  XlsxTable,
  XlsxTableColumn,
  XlsxTableHeaderMenuRenderProps,
  UseXlsxViewerThumbnailsOptions,
  XlsxViewerCharts,
  XlsxViewerTables,
  XlsxViewerController,
  XlsxViewerEditing,
  XlsxViewerImages,
  XlsxViewerProps,
  XlsxViewerProviderProps,
  XlsxViewerSelection,
  XlsxViewerThumbnails,
  XlsxViewerZoom
} from "./types";

const DEFAULT_ROW_HEIGHT = 24;
const DEFAULT_COL_WIDTH = 80;
const HEADER_HEIGHT = 24;
const ROW_HEADER_WIDTH = 40;
const INTERNAL_CLIPBOARD_MIME = "application/x-react-xlsx-range+json";
const MIN_OPEN_GRID_ROWS = 200;
const MIN_OPEN_GRID_COLS = 50;
const OPEN_GRID_ROW_PADDING = 120;
const OPEN_GRID_COL_PADDING = 24;
const INITIAL_WORKER_GRID_ROWS = 600;
const INITIAL_WORKER_GRID_COLS = 120;
const WORKER_GRID_GROW_ROWS = 2000;
const WORKER_GRID_GROW_COLS = 120;
const WORKER_GRID_GROW_THRESHOLD_ROWS = 160;
const WORKER_GRID_GROW_THRESHOLD_COLS = 24;
const OPEN_GRID_ROW_GROWTH = 200;
const OPEN_GRID_COL_GROWTH = 24;
const OPEN_GRID_VERTICAL_EDGE_PX = 600;
const OPEN_GRID_HORIZONTAL_EDGE_PX = 480;
const SELECTION_DRAG_THRESHOLD_PX = 4;
const IMAGE_MIN_SIZE_PX = 16;
const IMAGE_HANDLE_SIZE_PX = 10;
const CANVAS_RESIZE_HIT_SLOP_PX = 8;
const CANVAS_VIEWPORT_OVERSCAN_PX = 480;
const CANVAS_SCROLL_BUFFER_PX = 480;
const CANVAS_DEFERRED_VIEWPORT_SYNC_THRESHOLD_PX = CANVAS_SCROLL_BUFFER_PX - 96;
const CANVAS_IMMEDIATE_VIEWPORT_SYNC_THRESHOLD_PX = CANVAS_SCROLL_BUFFER_PX * 1.5;
const CANVAS_DIRTY_CELL_CULL_MARGIN_PX = CANVAS_VIEWPORT_OVERSCAN_PX * 2;
const THUMBNAIL_DEFAULT_MAX_DIMENSION = 192;
const THUMBNAIL_FALLBACK_ROWS = 12;
const THUMBNAIL_FALLBACK_COLS = 8;
const THUMBNAIL_MAX_ROWS = 200;
const THUMBNAIL_MAX_COLS = 80;
const THUMBNAIL_MAX_SOURCE_HEIGHT_PX = 900;
const THUMBNAIL_MAX_SOURCE_WIDTH_PX = 1440;
const LIVE_ZOOM_COMMIT_IDLE_MS = 48;
const WHEEL_ZOOM_SENSITIVITY = 0.00025;
const WHEEL_LINE_DELTA_PX = 16;
const CHART_SOURCE_HIGHLIGHT_COLORS = ["#2563eb", "#dc2626", "#7c3aed", "#059669", "#ea580c", "#db2777"];
const CHART_SOURCE_HIGHLIGHT_COLOR_BY_ROLE = {
  bubble: CHART_SOURCE_HIGHLIGHT_COLORS[3] ?? "#059669",
  category: CHART_SOURCE_HIGHLIGHT_COLORS[2] ?? "#7c3aed",
  name: CHART_SOURCE_HIGHLIGHT_COLORS[0] ?? "#2563eb",
  point: CHART_SOURCE_HIGHLIGHT_COLORS[5] ?? "#db2777",
  value: CHART_SOURCE_HIGHLIGHT_COLORS[4] ?? "#ea580c"
};
const SHEET_SURFACE = "#ffffff";
const SHEET_GRIDLINE = "#d9d9d9";
const DRAWING_SELECTION_STROKE = "#64748b";
const DRAWING_SELECTION_HANDLE_FILL = "#ffffff";
const DRAWING_SELECTION_HANDLE_SHADOW = "rgba(15, 23, 42, 0.18)";
const DEFAULT_CELL_PADDING = "0 4px";
const IMAGE_HANDLE_POSITIONS: XlsxImageResizeHandlePosition[] = ["nw", "n", "ne", "e", "se", "s", "sw", "w"];
const CANVAS_CELL_STYLE_CACHE_LIMIT = 4096;
const CANVAS_TEXT_MEASURE_CACHE_LIMIT = 20000;
const CANVAS_TEXT_TRUNCATE_CACHE_LIMIT = 4096;
const CANVAS_TEXT_WRAP_CACHE_LIMIT = 2048;
const CANVAS_PATH2D_CACHE_LIMIT = 512;
const EMPTY_VIRTUAL_ITEMS: VirtualItem[] = [];

type SelectionDragState = {
  anchor: XlsxCellAddress;
  axis: "cell" | "column" | "row";
  contentScaleX: number;
  contentScaleY: number;
  committedOnPointerDown: boolean;
  didDrag: boolean;
  originCell: XlsxCellAddress;
  originOverlayRect: { height: number; left: number; top: number; width: number } | null;
  originContentX: number;
  originContentY: number;
  previewRange: XlsxCellRange;
  pointerId: number;
  startClientX: number;
  startClientY: number;
};

const IMAGE_HANDLE_CURSOR: Record<XlsxImageResizeHandlePosition, React.CSSProperties["cursor"]> = {
  e: "ew-resize",
  n: "ns-resize",
  ne: "nesw-resize",
  nw: "nwse-resize",
  s: "ns-resize",
  se: "nwse-resize",
  sw: "nesw-resize",
  w: "ew-resize"
};

const NUMERIC_LENGTH_STYLE_KEYS = new Set([
  "borderRadius",
  "borderTopLeftRadius",
  "borderTopRightRadius",
  "borderBottomLeftRadius",
  "borderBottomRightRadius",
  "bottom",
  "fontSize",
  "gap",
  "height",
  "left",
  "letterSpacing",
  "margin",
  "marginBottom",
  "marginLeft",
  "marginRight",
  "marginTop",
  "maxHeight",
  "maxWidth",
  "minHeight",
  "minWidth",
  "outlineOffset",
  "outlineWidth",
  "padding",
  "paddingBottom",
  "paddingLeft",
  "paddingRight",
  "paddingTop",
  "right",
  "textIndent",
  "top",
  "width"
]);

const canvasCellStyleCache = new Map<string, CanvasCellStyleCache>();
const canvasTextMeasureCache = new Map<string, number>();
const canvasTextTruncateCache = new Map<string, string>();
const canvasTextWrapCache = new Map<string, string[]>();
const canvasPath2DCache = new Map<string, Path2D>();

function rememberBoundedCacheValue<K, V>(cache: Map<K, V>, key: K, value: V, limit: number) {
  if (cache.size >= limit) {
    const oldestKey = cache.keys().next().value as K | undefined;
    if (oldestKey !== undefined) {
      cache.delete(oldestKey);
    }
  }
  cache.set(key, value);
  return value;
}

function roundCanvasCacheWidth(width: number) {
  return Math.round(width * 4) / 4;
}

function measureCanvasTextWidth(context: CanvasRenderingContext2D, text: string) {
  if (text.length === 0) {
    return 0;
  }

  const cacheKey = `${context.font}\u0000${text}`;
  const cached = canvasTextMeasureCache.get(cacheKey);
  if (cached !== undefined) {
    return cached;
  }

  return rememberBoundedCacheValue(
    canvasTextMeasureCache,
    cacheKey,
    context.measureText(text).width,
    CANVAS_TEXT_MEASURE_CACHE_LIMIT
  );
}

function drawCanvasTextDecorations(
  context: CanvasRenderingContext2D,
  {
    align,
    color,
    decoration,
    ellipsize = false,
    lineThroughY,
    maxWidth,
    text,
    textX,
    underlineY,
    zoomFactor
  }: {
    align: CanvasTextAlign;
    color: string;
    decoration: string | undefined;
    ellipsize?: boolean;
    lineThroughY: number;
    maxWidth?: number;
    text: string;
    textX: number;
    underlineY: number;
    zoomFactor: number;
  }
) {
  if (!decoration || text.length === 0) {
    return;
  }

  const measured = ellipsize && maxWidth !== undefined
    ? Math.min(maxWidth, measureCanvasTextWidth(context, text))
    : measureCanvasTextWidth(context, text);
  const startX = align === "right"
    ? textX - measured
    : align === "center"
      ? textX - (measured / 2)
      : textX;

  context.strokeStyle = color;
  context.lineWidth = Math.max(1, zoomFactor * 0.75);

  if (decoration.includes("underline")) {
    context.beginPath();
    context.moveTo(startX, underlineY);
    context.lineTo(startX + measured, underlineY);
    context.stroke();
  }

  if (decoration.includes("line-through")) {
    context.beginPath();
    context.moveTo(startX, lineThroughY);
    context.lineTo(startX + measured, lineThroughY);
    context.stroke();
  }
}

function getCachedCanvasPath2D(path: string) {
  if (typeof Path2D === "undefined") {
    return null;
  }

  const cached = canvasPath2DCache.get(path);
  if (cached) {
    return cached;
  }

  return rememberBoundedCacheValue(canvasPath2DCache, path, new Path2D(path), CANVAS_PATH2D_CACHE_LIMIT);
}

function scaleCssLengthExpression(value: string, scale: number) {
  if (scale === 1) {
    return value;
  }

  return value.replace(/(-?\d*\.?\d+)(px|pt)\b/g, (_, rawNumber: string, unit: string) => {
    const nextValue = Number.parseFloat(rawNumber);
    if (!Number.isFinite(nextValue)) {
      return `${rawNumber}${unit}`;
    }
    return `${nextValue * scale}${unit}`;
  });
}

function scaleCssProperties(style: React.CSSProperties, scale: number): React.CSSProperties {
  if (scale === 1) {
    return style;
  }

  const nextStyle: React.CSSProperties = {};
  Object.entries(style).forEach(([key, value]) => {
    if (typeof value === "string") {
      nextStyle[key as keyof React.CSSProperties] = scaleCssLengthExpression(value, scale) as never;
      return;
    }

    if (typeof value === "number" && NUMERIC_LENGTH_STYLE_KEYS.has(key)) {
      nextStyle[key as keyof React.CSSProperties] = (value * scale) as never;
      return;
    }

    nextStyle[key as keyof React.CSSProperties] = value as never;
  });

  return nextStyle;
}

function parseCanvasLength(value: React.CSSProperties[keyof React.CSSProperties] | undefined, fallback = 0) {
  if (typeof value === "number" && Number.isFinite(value)) {
    return value;
  }

  if (typeof value === "string") {
    const match = value.match(/-?\d*\.?\d+/);
    if (match) {
      const parsed = Number.parseFloat(match[0]);
      if (Number.isFinite(parsed)) {
        return parsed;
      }
    }
  }

  return fallback;
}

function resolveCanvasPadding(padding: React.CSSProperties["padding"]) {
  if (typeof padding === "number") {
    return {
      bottom: padding,
      left: padding,
      right: padding,
      top: padding
    };
  }

  if (typeof padding === "string") {
    const parts = padding
      .trim()
      .split(/\s+/)
      .map((part) => parseCanvasLength(part, 0));

    if (parts.length === 1) {
      const [value = 0] = parts;
      return {
        bottom: value,
        left: value,
        right: value,
        top: value
      };
    }

    if (parts.length === 2) {
      const [vertical = 0, horizontal = 0] = parts;
      return {
        bottom: vertical,
        left: horizontal,
        right: horizontal,
        top: vertical
      };
    }

    if (parts.length === 3) {
      const [top = 0, horizontal = 0, bottom = 0] = parts;
      return {
        bottom,
        left: horizontal,
        right: horizontal,
        top
      };
    }

    const [top = 0, right = 0, bottom = 0, left = 0] = parts;
    return {
      bottom,
      left,
      right,
      top
    };
  }

  return {
    bottom: 0,
    left: 0,
    right: 0,
    top: 0
  };
}

function resolveSpreadsheetTextRotation(textRotation: unknown) {
  if (typeof textRotation !== "number" || !Number.isFinite(textRotation)) {
    return null;
  }

  if (textRotation < 0 || textRotation > 180) {
    return null;
  }

  const excelDegrees = textRotation > 90 ? 90 - textRotation : textRotation;
  return -excelDegrees;
}

function resolveCanvasFont(style: React.CSSProperties, fallbackSize = 12) {
  if (typeof style.font === "string" && style.font.trim().length > 0) {
    return style.font;
  }

  const fontStyle = typeof style.fontStyle === "string" && style.fontStyle.trim().length > 0
    ? style.fontStyle
    : "normal";
  const fontSize = typeof style.fontSize === "string" && style.fontSize.trim().length > 0
    ? style.fontSize
    : typeof style.fontSize === "number" && Number.isFinite(style.fontSize)
      ? `${style.fontSize}px`
      : `${fallbackSize}px`;
  const fontWeight = typeof style.fontWeight === "number" || typeof style.fontWeight === "string"
    ? String(style.fontWeight)
    : "400";
  const fontFamily = typeof style.fontFamily === "string" && style.fontFamily.trim().length > 0
    ? style.fontFamily
    : "ui-sans-serif, system-ui, sans-serif";

  return `${fontStyle} ${fontWeight} ${fontSize} ${fontFamily}`;
}

function resolveCanvasFontWithPxSize(style: React.CSSProperties, fontSizePx: number) {
  const fontStyle = typeof style.fontStyle === "string" && style.fontStyle.trim().length > 0
    ? style.fontStyle
    : "normal";
  const fontWeight = typeof style.fontWeight === "number" || typeof style.fontWeight === "string"
    ? String(style.fontWeight)
    : "400";
  const fontFamily = typeof style.fontFamily === "string" && style.fontFamily.trim().length > 0
    ? style.fontFamily
    : "ui-sans-serif, system-ui, sans-serif";

  return `${fontStyle} ${fontWeight} ${Math.max(1, fontSizePx)}px ${fontFamily}`;
}

type CanvasBorderDeclaration = {
  color: string;
  style: string;
  width: number;
};

function parseCanvasBorderDeclaration(borderValue: React.CSSProperties["borderTop"]) {
  if (typeof borderValue !== "string") {
    return null;
  }

  const trimmed = borderValue.trim();
  if (!trimmed || trimmed === "none") {
    return null;
  }

  const parts = trimmed.split(/\s+/);
  if (parts.length < 3) {
    return null;
  }

  const width = parseCanvasLength(parts[0], 1);
  const style = parts[1] ?? "solid";
  const color = parts.slice(2).join(" ");
  if (!Number.isFinite(width) || width <= 0 || !color) {
    return null;
  }

  return { color, style, width } satisfies CanvasBorderDeclaration;
}

function getCanvasBorderPriority(border: CanvasBorderDeclaration | null) {
  if (!border) {
    return -1;
  }

  const stylePriority: Record<string, number> = {
    dashed: 1,
    dotted: 0,
    double: 3,
    solid: 2
  };
  return (border.width * 10) + (stylePriority[border.style] ?? 0);
}

function resolveCanvasBoundaryBorder(
  primary: CanvasBorderDeclaration | null,
  secondary: CanvasBorderDeclaration | null
) {
  const primaryPriority = getCanvasBorderPriority(primary);
  const secondaryPriority = getCanvasBorderPriority(secondary);
  return secondaryPriority > primaryPriority ? secondary : primary;
}

function applyCanvasBorderDash(
  context: CanvasRenderingContext2D,
  style: string,
  width: number
) {
  context.lineCap = "butt";
  if (style === "dashed") {
    context.setLineDash([Math.max(3, width * 3), Math.max(2, width * 2)]);
    return;
  }
  if (style === "dotted") {
    context.lineCap = "round";
    context.setLineDash([Math.max(0.01, width * 0.01), Math.max(2, width * 2.2)]);
    return;
  }
  context.setLineDash([]);
}

function strokeCanvasBorderSide(
  context: CanvasRenderingContext2D,
  side: "top" | "right" | "bottom" | "left",
  rect: { left: number; top: number; width: number; height: number },
  border: CanvasBorderDeclaration
) {
  const halfWidth = border.width / 2;
  const left = rect.left;
  const right = rect.left + rect.width;
  const top = rect.top;
  const bottom = rect.top + rect.height;

  context.strokeStyle = border.color;
  context.lineWidth = border.width;
  applyCanvasBorderDash(context, border.style, border.width);

  const strokeLine = (offset = 0) => {
    context.beginPath();
    if (side === "top") {
      context.moveTo(left, top + halfWidth + offset);
      context.lineTo(right, top + halfWidth + offset);
    } else if (side === "bottom") {
      context.moveTo(left, bottom - halfWidth - offset);
      context.lineTo(right, bottom - halfWidth - offset);
    } else if (side === "left") {
      context.moveTo(left + halfWidth + offset, top);
      context.lineTo(left + halfWidth + offset, bottom);
    } else {
      context.moveTo(right - halfWidth - offset, top);
      context.lineTo(right - halfWidth - offset, bottom);
    }
    context.stroke();
  };

  if (border.style === "double") {
    const inset = Math.max(1, border.width);
    context.lineWidth = Math.max(1, border.width / 3);
    context.setLineDash([]);
    strokeLine(0);
    strokeLine(inset);
  } else {
    strokeLine(0);
  }

  context.setLineDash([]);
  context.lineCap = "butt";
  context.lineWidth = 1;
}

function truncateCanvasText(
  context: CanvasRenderingContext2D,
  text: string,
  maxWidth: number
) {
  if (maxWidth <= 0 || text.length === 0) {
    return "";
  }

  const roundedMaxWidth = roundCanvasCacheWidth(maxWidth);
  const cacheKey = `${context.font}\u0000${roundedMaxWidth}\u0000${text}`;
  const cached = canvasTextTruncateCache.get(cacheKey);
  if (cached !== undefined) {
    return cached;
  }

  if (measureCanvasTextWidth(context, text) <= maxWidth) {
    rememberBoundedCacheValue(canvasTextTruncateCache, cacheKey, text, CANVAS_TEXT_TRUNCATE_CACHE_LIMIT);
    return text;
  }

  const ellipsis = "…";
  let low = 0;
  let high = text.length;
  while (low < high) {
    const mid = Math.ceil((low + high) / 2);
    const candidate = `${text.slice(0, mid)}${ellipsis}`;
    if (measureCanvasTextWidth(context, candidate) <= maxWidth) {
      low = mid;
    } else {
      high = mid - 1;
    }
  }

  return rememberBoundedCacheValue(
    canvasTextTruncateCache,
    cacheKey,
    low <= 0 ? ellipsis : `${text.slice(0, low)}${ellipsis}`,
    CANVAS_TEXT_TRUNCATE_CACHE_LIMIT
  );
}

function resolveCanvasFontSizePx(style: React.CSSProperties, fallbackSize = 12) {
  const rawFontSize = style.fontSize;
  if (typeof rawFontSize === "number" && Number.isFinite(rawFontSize)) {
    return rawFontSize;
  }

  if (typeof rawFontSize === "string") {
    const trimmed = rawFontSize.trim();
    const parsed = Number.parseFloat(trimmed);
    if (Number.isFinite(parsed)) {
      if (trimmed.endsWith("pt")) {
        return parsed * (96 / 72);
      }
      if (trimmed.endsWith("em") || trimmed.endsWith("rem")) {
        return parsed * 16;
      }
      return parsed;
    }
  }

  return fallbackSize;
}

function resolveCanvasLineHeight(style: React.CSSProperties, fallbackFontSize = 12) {
  const fontSizePx = resolveCanvasFontSizePx(style, fallbackFontSize);
  const rawLineHeight = style.lineHeight;

  if (typeof rawLineHeight === "number" && Number.isFinite(rawLineHeight)) {
    return rawLineHeight > 4 ? rawLineHeight : rawLineHeight * fontSizePx;
  }

  if (typeof rawLineHeight === "string") {
    const trimmed = rawLineHeight.trim();
    if (!trimmed || trimmed === "normal") {
      return fontSizePx * 1.2;
    }

    const parsed = Number.parseFloat(trimmed);
    if (Number.isFinite(parsed)) {
      if (/^-?\d*\.?\d+$/.test(trimmed)) {
        return parsed > 4 ? parsed : parsed * fontSizePx;
      }
      if (trimmed.endsWith("pt")) {
        return parsed * (96 / 72);
      }
      if (trimmed.endsWith("em") || trimmed.endsWith("rem")) {
        return parsed * 16;
      }
      return parsed;
    }
  }

  return fontSizePx * 1.2;
}

function resolveCanvasTextMiddleY(
  verticalAlign: React.CSSProperties["verticalAlign"],
  contentTop: number,
  contentHeight: number,
  lineHeight: number
) {
  if (verticalAlign === "top") {
    return contentTop + (lineHeight / 2);
  }
  if (verticalAlign === "middle") {
    return contentTop + (contentHeight / 2);
  }
  return contentTop + contentHeight - (lineHeight / 2);
}

function resolveCanvasWrapIndex(
  context: CanvasRenderingContext2D,
  text: string,
  maxWidth: number
) {
  if (text.length <= 1) {
    return text.length;
  }

  let low = 1;
  let high = text.length;
  let best = 1;
  while (low <= high) {
    const mid = Math.floor((low + high) / 2);
    const candidate = text.slice(0, mid);
    if (measureCanvasTextWidth(context, candidate) <= maxWidth) {
      best = mid;
      low = mid + 1;
    } else {
      high = mid - 1;
    }
  }

  return best;
}

function wrapCanvasText(
  context: CanvasRenderingContext2D,
  text: string,
  maxWidth: number
) {
  if (text.length === 0) {
    return [""];
  }

  if (maxWidth <= 0) {
    return text.replace(/\r\n?/g, "\n").split("\n");
  }

  const normalized = text.replace(/\r\n?/g, "\n");
  const roundedMaxWidth = roundCanvasCacheWidth(maxWidth);
  const cacheKey = `${context.font}\u0000${roundedMaxWidth}\u0000${normalized}`;
  const cached = canvasTextWrapCache.get(cacheKey);
  if (cached) {
    return cached;
  }

  const paragraphs = normalized.split("\n");
  const lines: string[] = [];

  for (const paragraph of paragraphs) {
    if (paragraph.length === 0) {
      lines.push("");
      continue;
    }

    let remaining = paragraph;
    while (remaining.length > 0) {
      if (measureCanvasTextWidth(context, remaining) <= maxWidth) {
        lines.push(remaining);
        break;
      }

      const fit = Math.max(1, resolveCanvasWrapIndex(context, remaining, maxWidth));
      const candidate = remaining.slice(0, fit);
      const whitespaceIndex = fit < remaining.length
        ? Math.max(candidate.lastIndexOf(" "), candidate.lastIndexOf("\t"))
        : -1;
      const breakIndex = whitespaceIndex > 0 ? whitespaceIndex : fit;
      const nextLine = remaining.slice(0, breakIndex);

      if (nextLine.length === 0) {
        lines.push(remaining.slice(0, fit));
        remaining = remaining.slice(fit);
        continue;
      }

      lines.push(nextLine.replace(/\s+$/g, ""));
      remaining = remaining.slice(breakIndex).replace(/^\s+/g, "");
    }
  }

  return rememberBoundedCacheValue(canvasTextWrapCache, cacheKey, lines, CANVAS_TEXT_WRAP_CACHE_LIMIT);
}

type CanvasCellStyleCache = {
  baseFont: string;
  leftBorder: CanvasBorderDeclaration | null;
  padding: ReturnType<typeof resolveCanvasPadding>;
  rightBorder: CanvasBorderDeclaration | null;
  textAlign: CanvasTextAlign;
  textColor: string;
  textDecoration?: string;
  textOverflowEllipsis: boolean;
  topBorder: CanvasBorderDeclaration | null;
  usesWrappedText: boolean;
  bottomBorder: CanvasBorderDeclaration | null;
};

function buildCanvasCellStyleCache(style: React.CSSProperties): CanvasCellStyleCache {
  const cacheKey = [
    style.font,
    style.fontStyle,
    style.fontWeight,
    style.fontSize,
    style.fontFamily,
    style.borderTop,
    style.borderRight,
    style.borderBottom,
    style.borderLeft,
    style.padding,
    style.textAlign,
    style.color,
    style.textDecoration,
    style.textOverflow,
    style.whiteSpace
  ].map((value) => value == null ? "" : String(value)).join("\u0001");
  const cached = canvasCellStyleCache.get(cacheKey);
  if (cached) {
    return cached;
  }

  return rememberBoundedCacheValue(canvasCellStyleCache, cacheKey, {
    baseFont: resolveCanvasFont(style, 12),
    bottomBorder: parseCanvasBorderDeclaration(style.borderBottom),
    leftBorder: parseCanvasBorderDeclaration(style.borderLeft),
    padding: resolveCanvasPadding(style.padding),
    rightBorder: parseCanvasBorderDeclaration(style.borderRight),
    textAlign: style.textAlign === "right" || style.textAlign === "center" ? style.textAlign : "left",
    textColor: typeof style.color === "string" ? style.color : "#000000",
    textDecoration: typeof style.textDecoration === "string" ? style.textDecoration : undefined,
    textOverflowEllipsis: style.textOverflow === "ellipsis",
    topBorder: parseCanvasBorderDeclaration(style.borderTop),
    usesWrappedText: style.whiteSpace === "pre-wrap"
  }, CANVAS_CELL_STYLE_CACHE_LIMIT);
}

function splitCssGradientArgs(value: string) {
  const parts: string[] = [];
  let current = "";
  let depth = 0;
  for (const character of value) {
    if (character === "(") {
      depth += 1;
      current += character;
      continue;
    }
    if (character === ")") {
      depth = Math.max(0, depth - 1);
      current += character;
      continue;
    }
    if (character === "," && depth === 0) {
      parts.push(current.trim());
      current = "";
      continue;
    }
    current += character;
  }
  if (current.trim().length > 0) {
    parts.push(current.trim());
  }
  return parts;
}

function parseCanvasGradientStops(value: string) {
  return splitCssGradientArgs(value)
    .map((part) => {
      const match = part.match(/^(.*?)(?:\s+(-?\d*\.?\d+)%)?$/);
      if (!match) {
        return null;
      }
      const color = match[1]?.trim();
      const percent = match[2] ? Number.parseFloat(match[2]) : Number.NaN;
      if (!color) {
        return null;
      }
      return {
        color,
        offset: Number.isFinite(percent) ? Math.max(0, Math.min(1, percent / 100)) : null
      };
    })
    .filter((stop): stop is { color: string; offset: number | null } => Boolean(stop));
}

function resolveCanvasGradientFill(
  context: CanvasRenderingContext2D,
  rect: { height: number; left: number; top: number; width: number },
  backgroundImage: string
) {
  const linearMatch = backgroundImage.match(/^linear-gradient\((.*)\)$/i);
  if (linearMatch) {
    const parts = splitCssGradientArgs(linearMatch[1] ?? "");
    if (parts.length >= 2) {
      const rawAngle = parts[0]?.trim() ?? "180deg";
      const angleMatch = rawAngle.match(/(-?\d*\.?\d+)deg/i);
      const cssAngle = angleMatch ? Number.parseFloat(angleMatch[1] ?? "180") : 180;
      const radians = ((90 - cssAngle) * Math.PI) / 180;
      const halfWidth = rect.width / 2;
      const halfHeight = rect.height / 2;
      const centerX = rect.left + halfWidth;
      const centerY = rect.top + halfHeight;
      const projectedHalfLength = Math.abs(Math.cos(radians) * halfWidth) + Math.abs(Math.sin(radians) * halfHeight);
      const startX = centerX - (Math.cos(radians) * projectedHalfLength);
      const startY = centerY + (Math.sin(radians) * projectedHalfLength);
      const endX = centerX + (Math.cos(radians) * projectedHalfLength);
      const endY = centerY - (Math.sin(radians) * projectedHalfLength);
      const gradient = context.createLinearGradient(startX, startY, endX, endY);
      const stops = parseCanvasGradientStops(parts.slice(1).join(","));
      if (stops.length > 0) {
        const normalizedStops = stops.map((stop, index) => ({
          color: stop.color,
          offset: stop.offset ?? (stops.length === 1 ? 0 : index / (stops.length - 1))
        }));
        normalizedStops.forEach((stop) => {
          gradient.addColorStop(stop.offset, stop.color);
        });
        return gradient;
      }
    }
  }

  const radialMatch = backgroundImage.match(/^radial-gradient\((.*)\)$/i);
  if (radialMatch) {
    const parts = splitCssGradientArgs(radialMatch[1] ?? "");
    const stopParts = parts[0]?.startsWith("circle") ? parts.slice(1) : parts;
    const gradient = context.createRadialGradient(
      rect.left + (rect.width / 2),
      rect.top + (rect.height / 2),
      0,
      rect.left + (rect.width / 2),
      rect.top + (rect.height / 2),
      Math.max(rect.width, rect.height) / 2
    );
    const stops = parseCanvasGradientStops(stopParts.join(","));
    if (stops.length > 0) {
      const normalizedStops = stops.map((stop, index) => ({
        color: stop.color,
        offset: stop.offset ?? (stops.length === 1 ? 0 : index / (stops.length - 1))
      }));
      normalizedStops.forEach((stop) => {
        gradient.addColorStop(stop.offset, stop.color);
      });
      return gradient;
    }
  }

  return null;
}

function resolveCanvasDataBarFill(
  context: CanvasRenderingContext2D,
  left: number,
  top: number,
  width: number,
  _height: number,
  dataBar: NonNullable<CellRenderData["conditionalDataBar"]>
) {
  if (dataBar.gradient === false) {
    return dataBar.color;
  }

  const gradient = context.createLinearGradient(left, top, left + width, top);
  gradient.addColorStop(0, lightenColor(dataBar.color, 0.28));
  gradient.addColorStop(1, dataBar.color);
  return gradient;
}

function drawCanvasConditionalIcon(
  context: CanvasRenderingContext2D,
  icon: NonNullable<CellRenderData["conditionalIcon"]>,
  centerX: number,
  centerY: number,
  size: number
) {
  context.save();
  if (icon.glyph) {
    context.textAlign = "center";
    context.textBaseline = "middle";
    context.fillStyle = icon.color ?? icon.backgroundColor ?? "#6b7280";
    context.fillText(icon.glyph, centerX, centerY);
    context.restore();
    return;
  }

  if (icon.shape === "arrow") {
    const fill = icon.color ?? "#111827";
    const stroke = icon.borderColor ?? darkenColor(fill, 0.32);
    const scale = size / 16;
    context.translate(centerX, centerY);
    context.rotate(((icon.rotationDeg ?? 0) * Math.PI) / 180);
    context.translate(-8 * scale, -8 * scale);
    context.beginPath();
    context.moveTo(2.5 * scale, 8 * scale);
    context.lineTo(8.4 * scale, 2.4 * scale);
    context.lineTo(8.4 * scale, 5.2 * scale);
    context.lineTo(13.5 * scale, 5.2 * scale);
    context.lineTo(13.5 * scale, 10.8 * scale);
    context.lineTo(8.4 * scale, 10.8 * scale);
    context.lineTo(8.4 * scale, 13.6 * scale);
    context.closePath();
    context.fillStyle = fill;
    context.fill();
    context.strokeStyle = stroke;
    context.lineWidth = Math.max(1, 1.25 * scale);
    context.lineJoin = "round";
    context.stroke();
    context.restore();
    return;
  }

  context.beginPath();
  context.arc(centerX, centerY, size / 2, 0, Math.PI * 2);
  context.fillStyle = icon.color ?? icon.backgroundColor ?? "#6b7280";
  context.fill();
  context.restore();
}

function formatZoomScale(zoomScale: number) {
  return `${Math.round(zoomScale)}%`;
}

function clampValue(value: number, min: number, max: number) {
  return Math.min(max, Math.max(min, value));
}

function normalizeWheelDelta(event: WheelEvent) {
  const dominantDelta = Math.abs(event.deltaY) >= Math.abs(event.deltaX) ? event.deltaY : event.deltaX;
  switch (event.deltaMode) {
    case WheelEvent.DOM_DELTA_LINE:
      return dominantDelta * WHEEL_LINE_DELTA_PX;
    case WheelEvent.DOM_DELTA_PAGE:
      return dominantDelta * window.innerHeight;
    default:
      return dominantDelta;
  }
}

function resolveEventAnchor(clientX: number, clientY: number, rect: DOMRect): ZoomAnchor {
  return {
    x: clampValue(clientX - rect.left, 0, rect.width),
    y: clampValue(clientY - rect.top, 0, rect.height)
  };
}

function resolveTouchDistance(firstTouch: Touch, secondTouch: Touch) {
  return Math.hypot(firstTouch.clientX - secondTouch.clientX, firstTouch.clientY - secondTouch.clientY);
}

type ViewerPalette = {
  border: string;
  buttonSurface: string;
  buttonText: string;
  canvas: string;
  danger: string;
  headerSurface: string;
  mutedSurface: string;
  mutedText: string;
  rowHeaderSurface: string;
  shadow: string;
  sheetActiveSurface: string;
  sheetActiveText: string;
  sheetInactiveSurface: string;
  sheetInactiveText: string;
  strongBorder: string;
  subtleSurface: string;
  surface: string;
  text: string;
  toolbarSurface: string;
};

const LIGHT_PALETTE: ViewerPalette = {
  border: "#e4e4e7",
  buttonSurface: "#ffffff",
  buttonText: "#18181b",
  canvas: "#fafafa",
  danger: "#dc2626",
  headerSurface: "#f4f4f5",
  mutedSurface: "#f5f5f5",
  mutedText: "#71717a",
  rowHeaderSurface: "#f4f4f5",
  shadow: "0 1px 2px rgba(0, 0, 0, 0.04)",
  sheetActiveSurface: "#ffffff",
  sheetActiveText: "#18181b",
  sheetInactiveSurface: "#e4e4e7",
  sheetInactiveText: "#52525b",
  strongBorder: "#d4d4d8",
  subtleSurface: "#fafafa",
  surface: "#ffffff",
  text: "#18181b",
  toolbarSurface: "#f5f5f5"
};

const DARK_PALETTE: ViewerPalette = {
  border: "rgba(255, 255, 255, 0.10)",
  buttonSurface: "rgba(255, 255, 255, 0.06)",
  buttonText: "#f4f4f5",
  canvas: "#09090b",
  danger: "#f87171",
  headerSurface: "#18181b",
  mutedSurface: "#111113",
  mutedText: "#a1a1aa",
  rowHeaderSurface: "#18181b",
  shadow: "0 1px 2px rgba(0, 0, 0, 0.28)",
  sheetActiveSurface: "#27272a",
  sheetActiveText: "#fafafa",
  sheetInactiveSurface: "#18181b",
  sheetInactiveText: "#a1a1aa",
  strongBorder: "rgba(255, 255, 255, 0.16)",
  subtleSurface: "#101012",
  surface: "#111113",
  text: "#f4f4f5",
  toolbarSurface: "#101012"
};

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
  const clamped = Math.max(0, Math.min(1, ratio));
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

const ViewerContext = React.createContext<XlsxViewerController | null>(null);
const ViewerAppearanceContext = React.createContext<{ isDark: boolean }>({ isDark: false });

type ZoomAnchor = {
  x: number;
  y: number;
};

type WebKitGestureEvent = Event & {
  clientX: number;
  clientY: number;
  scale: number;
};

type LiveGestureZoomState = {
  anchor: ZoomAnchor;
  baseZoomScale: number;
  targetZoomScale: number;
};

type ChartRangeHighlight = {
  fillColor: string;
  range: XlsxCellRange;
  strokeColor: string;
  workbookSheetIndex: number;
};

type CellChartHighlight = {
  borderBottom: boolean;
  borderLeft: boolean;
  borderRight: boolean;
  borderTop: boolean;
  fillColor: string;
  strokeColor: string;
};

type WorksheetWithRowsBatch = Worksheet & {
  getRowsBatch?: (startRow: number, maxRows: number, options?: unknown) => unknown;
};

type WorksheetBatchCellEntry = {
  col?: unknown;
  formula?: unknown;
  hyperlink?: unknown;
  isMergedSecondary?: unknown;
  mergeSpan?: unknown;
  style?: unknown;
  value?: unknown;
};

type WorksheetBatchRowEntry = {
  cells?: unknown;
  index?: unknown;
};

type BatchedCellData = {
  formula: string | null;
  hyperlink: {
    location?: string;
    target?: string;
    tooltip?: string;
  } | null;
  isMergedSecondary: boolean;
  mergeSpan: { colSpan?: number; rowSpan?: number } | null;
  style: Record<string, unknown> | null;
  value: string;
};

type WorksheetBatchWindow = {
  cells: Map<string, BatchedCellData>;
  endRow: number;
  startRow: number;
};

function classNames(...values: Array<string | false | null | undefined>) {
  return values.filter(Boolean).join(" ");
}

function resolveDarkModeSurface(themePalette: XlsxSheetData["themePalette"] | undefined, palette: ViewerPalette) {
  return "hsl(225 4% 6%)";
}

function resolveSheetSurface(sheet: XlsxSheetData | null, palette: ViewerPalette) {
  return paletteIsDark(palette)
    ? resolveDarkModeSurface(sheet?.themePalette, palette)
    : sheet?.themePalette.colorsByIndex[0] ?? SHEET_SURFACE;
}

function normalizeRange(range: XlsxCellRange): XlsxCellRange {
  return {
    start: {
      col: Math.min(range.start.col, range.end.col),
      row: Math.min(range.start.row, range.end.row)
    },
    end: {
      col: Math.max(range.start.col, range.end.col),
      row: Math.max(range.start.row, range.end.row)
    }
  };
}

function parseCellAddressAttribute(value: string | null): XlsxCellAddress | null {
  if (!value) {
    return null;
  }

  const [rowValue, colValue] = value.split(":");
  const row = Number(rowValue);
  const col = Number(colValue);
  if (!Number.isInteger(row) || !Number.isInteger(col) || row < 0 || col < 0) {
    return null;
  }

  return { row, col };
}

function isSameCell(left: XlsxCellAddress | null, right: XlsxCellAddress | null) {
  return Boolean(left && right && left.row === right.row && left.col === right.col);
}

function isCellInRange(cell: XlsxCellAddress, range: XlsxCellRange | null) {
  if (!range) {
    return false;
  }

  const normalized = normalizeRange(range);
  return (
    cell.row >= normalized.start.row &&
    cell.row <= normalized.end.row &&
    cell.col >= normalized.start.col &&
    cell.col <= normalized.end.col
  );
}

function resolveCellChartHighlight(
  cell: XlsxCellAddress,
  highlights: ChartRangeHighlight[]
): CellChartHighlight | null {
  let match: CellChartHighlight | null = null;

  for (const highlight of highlights) {
    if (!isCellInRange(cell, highlight.range)) {
      continue;
    }

    const normalized = normalizeRange(highlight.range);
    match = {
      borderBottom: cell.row === normalized.end.row,
      borderLeft: cell.col === normalized.start.col,
      borderRight: cell.col === normalized.end.col,
      borderTop: cell.row === normalized.start.row,
      fillColor: highlight.fillColor,
      strokeColor: highlight.strokeColor
    };
  }

  return match;
}

function rangesEqual(left: XlsxCellRange | null, right: XlsxCellRange | null) {
  if (!left || !right) {
    return left === right;
  }

  const normalizedLeft = normalizeRange(left);
  const normalizedRight = normalizeRange(right);
  return (
    normalizedLeft.start.row === normalizedRight.start.row &&
    normalizedLeft.start.col === normalizedRight.start.col &&
    normalizedLeft.end.row === normalizedRight.end.row &&
    normalizedLeft.end.col === normalizedRight.end.col
  );
}

function isPrintableKey(event: Pick<KeyboardEvent, "altKey" | "ctrlKey" | "key" | "metaKey">) {
  return event.key.length === 1 && !event.altKey && !event.ctrlKey && !event.metaKey;
}

function getInteractiveFocusTarget(target: EventTarget | null) {
  if (!(target instanceof Element)) {
    return null;
  }

  return target.closest(
    "a[href], button, input, select, textarea, [contenteditable=''], [contenteditable='true'], [role='button'], [role='menu'], [role='menuitem'], [role='textbox'], [tabindex]:not([tabindex='-1'])"
  );
}

function isInteractiveFocusTarget(target: EventTarget | null, container: HTMLElement) {
  const interactiveElement = getInteractiveFocusTarget(target);
  return Boolean(interactiveElement && interactiveElement !== container && container.contains(interactiveElement));
}

function sumSegment(values: number[], startIndex: number, endIndex: number) {
  let total = 0;
  for (let index = startIndex; index <= endIndex; index += 1) {
    total += values[index] ?? 0;
  }
  return total;
}

function findIndexForOffset(values: number[], offset: number) {
  if (offset <= 0) {
    return 0;
  }

  let currentOffset = 0;
  for (let index = 0; index < values.length; index += 1) {
    currentOffset += values[index] ?? 0;
    if (offset < currentOffset) {
      return index;
    }
  }

  return values.length > 0 ? values.length - 1 : -1;
}

function buildPrefixSums(values: number[]) {
  const prefix = new Array(values.length + 1).fill(0);
  for (let index = 0; index < values.length; index += 1) {
    prefix[index + 1] = prefix[index] + (values[index] ?? 0);
  }
  return prefix;
}

type IndexedMergedRegion = {
  endColIndex: number;
  endRowIndex: number;
  startColIndex: number;
  startRowIndex: number;
};

function setIntersectsIndexRange(indices: Set<number>, startIndex: number, endIndex: number) {
  for (let index = startIndex; index <= endIndex; index += 1) {
    if (indices.has(index)) {
      return true;
    }
  }
  return false;
}

function expandMergeAwareWindow(
  rowIndices: number[],
  colIndices: number[],
  mergedRegions: IndexedMergedRegion[]
) {
  const nextRowIndices = new Set(rowIndices);
  const nextColIndices = new Set(colIndices);

  if (nextRowIndices.size === 0 || nextColIndices.size === 0 || mergedRegions.length === 0) {
    return {
      colIndices: Array.from(nextColIndices).sort((left, right) => left - right),
      rowIndices: Array.from(nextRowIndices).sort((left, right) => left - right)
    };
  }

  let changed = true;
  while (changed) {
    changed = false;
    for (const region of mergedRegions) {
      const intersectsRows = setIntersectsIndexRange(nextRowIndices, region.startRowIndex, region.endRowIndex);
      const intersectsCols = setIntersectsIndexRange(nextColIndices, region.startColIndex, region.endColIndex);
      if (!intersectsRows || !intersectsCols) {
        continue;
      }

      for (let rowIndex = region.startRowIndex; rowIndex <= region.endRowIndex; rowIndex += 1) {
        if (!nextRowIndices.has(rowIndex)) {
          nextRowIndices.add(rowIndex);
          changed = true;
        }
      }
      for (let colIndex = region.startColIndex; colIndex <= region.endColIndex; colIndex += 1) {
        if (!nextColIndices.has(colIndex)) {
          nextColIndices.add(colIndex);
          changed = true;
        }
      }
    }
  }

  return {
    colIndices: Array.from(nextColIndices).sort((left, right) => left - right),
    rowIndices: Array.from(nextRowIndices).sort((left, right) => left - right)
  };
}

function buildRenderedAxisSignature(
  items: Array<{ index: number; size: number; start: number }>,
  resolveActualKey: (item: { index: number; size: number; start: number }) => number
) {
  return items
    .map((item) => `${resolveActualKey(item)}:${item.start}:${item.size}`)
    .join("|");
}

function buildRangeSignature(range: XlsxCellRange | null) {
  if (!range) {
    return "";
  }

  const normalized = normalizeRange(range);
  return [
    normalized.start.row,
    normalized.start.col,
    normalized.end.row,
    normalized.end.col
  ].join(":");
}

function sumPrefixRange(prefixSums: number[], startIndex: number, endIndex: number) {
  if (startIndex > endIndex) {
    return 0;
  }

  return (prefixSums[endIndex + 1] ?? 0) - (prefixSums[startIndex] ?? 0);
}

function buildStickyOffsets(actualIndices: number[], sizesByActualIndex: number[], leadingOffset: number) {
  const offsets = new Map<number, number>();
  let nextOffset = leadingOffset;

  for (const actualIndex of actualIndices) {
    offsets.set(actualIndex, nextOffset);
    nextOffset += sizesByActualIndex[actualIndex] ?? 0;
  }

  return offsets;
}

function findIndexForOffsetPrefix(prefixSums: number[], offset: number) {
  const count = prefixSums.length - 1;
  if (count <= 0) {
    return -1;
  }
  if (offset <= 0) {
    return 0;
  }
  if (offset >= prefixSums[count]) {
    return count - 1;
  }

  let low = 0;
  let high = count - 1;
  while (low < high) {
    const mid = Math.floor((low + high) / 2);
    if ((prefixSums[mid + 1] ?? 0) <= offset) {
      low = mid + 1;
    } else {
      high = mid;
    }
  }

  return low;
}

function clampContentOffset(offset: number, maxOffset: number) {
  if (!Number.isFinite(offset) || maxOffset <= 0) {
    return 0;
  }

  return Math.max(0, Math.min(offset, Math.max(0, maxOffset - 0.001)));
}

function resolveOpenGridExtent(maxUsedIndex: number, minimum: number, padding: number) {
  return Math.max(maxUsedIndex + 1 + padding, minimum);
}

function buildVisibleAxisIndices(
  precomputed: number[],
  displayLimit: number,
  maxUsedIndex: number,
  hiddenIndices?: Set<number>
) {
  if (displayLimit <= 0) {
    return [];
  }

  if (precomputed.length === 0) {
    const visible: number[] = [];
    const usedLimit = Math.min(displayLimit, Math.max(0, maxUsedIndex + 1));
    for (let index = 0; index < usedLimit; index += 1) {
      if (!hiddenIndices?.has(index)) {
        visible.push(index);
      }
    }
    for (let index = usedLimit; index < displayLimit; index += 1) {
      visible.push(index);
    }
    return visible;
  }

  const visible = precomputed.filter((value) => value < displayLimit);
  const appendStart = Math.max(0, maxUsedIndex + 1);
  for (let index = appendStart; index < displayLimit; index += 1) {
    visible.push(index);
  }
  return visible;
}

function resolveInitialDisplayExtent(
  maxUsedIndex: number,
  minimum: number,
  padding: number,
  isWorkerBacked: boolean,
  initialCap: number
) {
  const fullExtent = resolveOpenGridExtent(maxUsedIndex, minimum, padding);
  return isWorkerBacked ? Math.min(fullExtent, Math.max(minimum, initialCap)) : fullExtent;
}

function sumBeforeActualIndex(actualIndices: number[], sizes: number[], actualIndex: number) {
  let total = 0;
  for (let index = 0; index < actualIndices.length; index += 1) {
    if ((actualIndices[index] ?? 0) >= actualIndex) {
      break;
    }
    total += sizes[index] ?? 0;
  }
  return total;
}

function resolveAxisStartOffset(
  actualIndex: number,
  actualIndices: number[],
  sizes: number[],
  indexByActual?: Map<number, number>,
  prefixSums?: number[],
  actualPrefixSums?: number[]
) {
  if (actualPrefixSums && actualIndex >= 0 && actualIndex < actualPrefixSums.length) {
    return actualPrefixSums[actualIndex] ?? 0;
  }

  const visibleIndex = indexByActual?.get(actualIndex);
  if (visibleIndex !== undefined && prefixSums) {
    return prefixSums[visibleIndex] ?? 0;
  }

  return sumBeforeActualIndex(actualIndices, sizes, actualIndex);
}

function resolveAnchoredRect(
  anchor: XlsxImage["anchor"] | XlsxShape["anchor"],
  visibleRows: number[],
  visibleCols: number[],
  rowHeights: number[],
  colWidths: number[],
  options?: {
    actualColPrefixSums?: number[];
    actualRowPrefixSums?: number[];
    colIndexByActual?: Map<number, number>;
    colPrefixSums?: number[];
    headerHeight?: number;
    rowIndexByActual?: Map<number, number>;
    rowHeaderWidth?: number;
    rowPrefixSums?: number[];
  }
): XlsxImageRect {
  const headerHeight = options?.headerHeight ?? HEADER_HEIGHT;
  const rowHeaderWidth = options?.rowHeaderWidth ?? ROW_HEADER_WIDTH;
  const resolveMarkerLeft = (col: number, colOffsetEmu: number) =>
    rowHeaderWidth +
    resolveAxisStartOffset(
      col,
      visibleCols,
      colWidths,
      options?.colIndexByActual,
      options?.colPrefixSums,
      options?.actualColPrefixSums
    ) +
    emuToPixels(colOffsetEmu);
  const resolveMarkerTop = (row: number, rowOffsetEmu: number) =>
    headerHeight +
    resolveAxisStartOffset(
      row,
      visibleRows,
      rowHeights,
      options?.rowIndexByActual,
      options?.rowPrefixSums,
      options?.actualRowPrefixSums
    ) +
    emuToPixels(rowOffsetEmu);

  if (anchor.kind === "absolute") {
    return {
      height: Math.max(1, emuToPixels(anchor.sizeEmu.cy)),
      left: rowHeaderWidth + emuToPixels(anchor.positionEmu.x),
      top: headerHeight + emuToPixels(anchor.positionEmu.y),
      width: Math.max(1, emuToPixels(anchor.sizeEmu.cx))
    };
  }

  if (anchor.kind === "one-cell") {
    return {
      height: Math.max(1, emuToPixels(anchor.sizeEmu.cy)),
      left: resolveMarkerLeft(anchor.from.col, anchor.from.colOffsetEmu),
      top: resolveMarkerTop(anchor.from.row, anchor.from.rowOffsetEmu),
      width: Math.max(1, emuToPixels(anchor.sizeEmu.cx))
    };
  }

  const left = resolveMarkerLeft(anchor.from.col, anchor.from.colOffsetEmu);
  const top = resolveMarkerTop(anchor.from.row, anchor.from.rowOffsetEmu);
  const right = resolveMarkerLeft(anchor.to.col, anchor.to.colOffsetEmu);
  const bottom = resolveMarkerTop(anchor.to.row, anchor.to.rowOffsetEmu);

  return {
    height: Math.max(1, bottom - top),
    left,
    top,
    width: Math.max(1, right - left)
  };
}

function resolveImageRect(
  image: XlsxImage,
  visibleRows: number[],
  visibleCols: number[],
  rowHeights: number[],
  colWidths: number[],
  options?: {
    actualColPrefixSums?: number[];
    actualRowPrefixSums?: number[];
    colIndexByActual?: Map<number, number>;
    colPrefixSums?: number[];
    headerHeight?: number;
    rowIndexByActual?: Map<number, number>;
    rowHeaderWidth?: number;
    rowPrefixSums?: number[];
  }
): XlsxImageRect {
  return resolveAnchoredRect(image.anchor, visibleRows, visibleCols, rowHeights, colWidths, options);
}

function resolveImageAnchorExtents(image: XlsxImage) {
  if (image.anchor.kind === "absolute") {
    return { maxCol: 0, maxRow: 0 };
  }

  if (image.anchor.kind === "one-cell") {
    return {
      maxCol: image.anchor.from.col,
      maxRow: image.anchor.from.row
    };
  }

  return {
    maxCol: Math.max(image.anchor.from.col, image.anchor.to.col),
    maxRow: Math.max(image.anchor.from.row, image.anchor.to.row)
  };
}

function resolveAnchoredBounds(anchor: XlsxImage["anchor"] | XlsxShape["anchor"]) {
  if (anchor.kind === "absolute") {
    return null;
  }

  if (anchor.kind === "one-cell") {
    return {
      maxCol: anchor.from.col,
      maxRow: anchor.from.row,
      minCol: anchor.from.col,
      minRow: anchor.from.row
    };
  }

  return {
    maxCol: Math.max(anchor.from.col, anchor.to.col),
    maxRow: Math.max(anchor.from.row, anchor.to.row),
    minCol: Math.min(anchor.from.col, anchor.to.col),
    minRow: Math.min(anchor.from.row, anchor.to.row)
  };
}

function resolveShapeAnchorExtents(shape: XlsxShape) {
  if (shape.anchor.kind === "absolute") {
    return { maxCol: 0, maxRow: 0 };
  }

  if (shape.anchor.kind === "one-cell") {
    return {
      maxCol: shape.anchor.from.col,
      maxRow: shape.anchor.from.row
    };
  }

  return {
    maxCol: Math.max(shape.anchor.from.col, shape.anchor.to.col),
    maxRow: Math.max(shape.anchor.from.row, shape.anchor.to.row)
  };
}

function resolveChartAnchorExtents(chart: XlsxChart) {
  if (chart.anchor.kind === "absolute") {
    return { maxCol: 0, maxRow: 0 };
  }

  if (chart.anchor.kind === "one-cell") {
    return {
      maxCol: chart.anchor.from.col,
      maxRow: chart.anchor.from.row
    };
  }

  return {
    maxCol: Math.max(chart.anchor.from.col, chart.anchor.to.col),
    maxRow: Math.max(chart.anchor.from.row, chart.anchor.to.row)
  };
}

function rectIntersectsViewport(
  rect: XlsxImageRect,
  viewport: { height: number; left: number; top: number; width: number },
  overscan = 240
) {
  const viewportLeft = viewport.left - overscan;
  const viewportTop = viewport.top - overscan;
  const viewportRight = viewport.left + viewport.width + overscan;
  const viewportBottom = viewport.top + viewport.height + overscan;
  const rectRight = rect.left + rect.width;
  const rectBottom = rect.top + rect.height;
  return rectRight >= viewportLeft
    && rect.left <= viewportRight
    && rectBottom >= viewportTop
    && rect.top <= viewportBottom;
}

type FrozenDrawingPane = "corner" | "left" | "scroll" | "top";
type DrawingViewport = {
  height: number;
  left: number;
  top: number;
  width: number;
};

function resolveFrozenDrawingPane(
  rect: XlsxImageRect,
  frozenRows: number[],
  frozenCols: number[],
  actualRowHeights: number[],
  actualColWidths: number[],
  freezePanes: XlsxSheetData["freezePanes"] | null,
  stickyTopByRow: Map<number, number>,
  stickyLeftByCol: Map<number, number>,
  options?: {
    defaultColWidth?: number;
    defaultRowHeight?: number;
    headerHeight?: number;
    rowHeaderWidth?: number;
  }
): FrozenDrawingPane {
  const headerHeight = options?.headerHeight ?? HEADER_HEIGHT;
  const rowHeaderWidth = options?.rowHeaderWidth ?? ROW_HEADER_WIDTH;
  const defaultRowHeight = options?.defaultRowHeight ?? DEFAULT_ROW_HEIGHT;
  const defaultColWidth = options?.defaultColWidth ?? DEFAULT_COL_WIDTH;
  const frozenPaneBottom =
    freezePanes?.row && freezePanes.row > 0 && frozenRows.length > 0
      ? frozenRows.reduce(
          (max, row) => Math.max(max, (stickyTopByRow.get(row) ?? headerHeight) + (actualRowHeights[row] ?? defaultRowHeight)),
          headerHeight
        )
      : null;
  const frozenPaneRight =
    freezePanes?.col && freezePanes.col > 0 && frozenCols.length > 0
      ? frozenCols.reduce(
          (max, col) => Math.max(max, (stickyLeftByCol.get(col) ?? rowHeaderWidth) + (actualColWidths[col] ?? defaultColWidth)),
          rowHeaderWidth
        )
      : null;

  const freezeTop = frozenPaneBottom !== null && rect.top + rect.height <= frozenPaneBottom + 0.5;
  const freezeLeft = frozenPaneRight !== null && rect.left + rect.width <= frozenPaneRight + 0.5;

  if (freezeTop && freezeLeft) {
    return "corner";
  }
  if (freezeTop) {
    return "top";
  }
  if (freezeLeft) {
    return "left";
  }
  return "scroll";
}

function buildShapeContainerStyle(shape: XlsxShape, rect: XlsxImageRect, viewerScale = 1): React.CSSProperties {
  const borderWidth = shape.stroke?.none ? 0 : Math.max(0, shape.stroke?.widthPx ?? 0) * viewerScale;
  const strokeColor = shape.stroke?.color ?? "transparent";
  const fillColor = shape.fill?.none ? "transparent" : (shape.fill?.color ?? "transparent");
  const hasVisibleText = shape.paragraphs.some((paragraph) => paragraph.runs.some((run) => run.text.trim().length > 0));
  const transformParts = [
    shape.rotationDeg ? `rotate(${shape.rotationDeg}deg)` : "",
    shape.flipH ? "scaleX(-1)" : "",
    shape.flipV ? "scaleY(-1)" : ""
  ].filter(Boolean);

  let borderRadius: React.CSSProperties["borderRadius"] = 0;
  if (shape.geometry === "ellipse") {
    borderRadius = "9999px";
  } else if (shape.geometry === "roundRect") {
    borderRadius = 12 * viewerScale;
  }

  return {
    alignItems:
      shape.textBox?.verticalAlign === "middle"
        ? "center"
        : shape.textBox?.verticalAlign === "bottom"
          ? "flex-end"
          : "flex-start",
    backgroundColor: fillColor,
    border: borderWidth > 0 ? `${borderWidth}px solid ${strokeColor}` : "none",
    borderRadius,
    boxSizing: "border-box",
    display: "flex",
    flexDirection: "column",
    height: rect.height,
    justifyContent: "flex-start",
    left: rect.left,
    opacity: Math.min(shape.fill?.opacity ?? 1, shape.stroke?.opacity ?? 1),
    overflow: hasVisibleText ? "hidden" : "visible",
    position: "absolute",
    top: rect.top,
    transform: transformParts.join(" ") || undefined,
    transformOrigin: "center center",
    width: rect.width,
    zIndex: shape.zIndex
  };
}

function clampPercent(value: number, fallback: number) {
  if (!Number.isFinite(value)) {
    return fallback;
  }
  return Math.max(0, Math.min(100, value));
}

function resolveShapeAdjustment(shape: XlsxShape, name: string, fallback: number) {
  const rawValue = shape.geometryAdjustments?.[name];
  if (typeof rawValue !== "number") {
    return fallback;
  }
  return clampPercent(rawValue / 1000, fallback);
}

function buildClosedPath(points: Array<[number, number]>) {
  if (points.length === 0) {
    return "";
  }

  return points
    .map(([x, y], index) => `${index === 0 ? "M" : "L"} ${x} ${y}`)
    .join(" ") + " Z";
}

function pointOnCircle(radius: number, angleDeg: number, center = 50): [number, number] {
  const radians = angleDeg * (Math.PI / 180);
  return [
    center + Math.cos(radians) * radius,
    center + Math.sin(radians) * radius
  ];
}

function buildRegularPolygonPath(sides: number, rotationDeg = -90, radius = 50) {
  return buildClosedPath(
    Array.from({ length: sides }, (_, index) => pointOnCircle(radius, rotationDeg + (360 / sides) * index))
  );
}

function buildStarPath(points: number, innerRadius: number, rotationDeg = -90, outerRadius = 50) {
  return buildClosedPath(
    Array.from({ length: points * 2 }, (_, index) => {
      const radius = index % 2 === 0 ? outerRadius : innerRadius;
      return pointOnCircle(radius, rotationDeg + (360 / (points * 2)) * index);
    })
  );
}

function buildRoundRectPath(radius: number) {
  const cornerRadius = Math.max(0, Math.min(50, radius));
  const edge = 100 - cornerRadius;
  return [
    `M ${cornerRadius} 0`,
    `L ${edge} 0`,
    `Q 100 0 100 ${cornerRadius}`,
    `L 100 ${edge}`,
    `Q 100 100 ${edge} 100`,
    `L ${cornerRadius} 100`,
    `Q 0 100 0 ${edge}`,
    `L 0 ${cornerRadius}`,
    `Q 0 0 ${cornerRadius} 0`,
    "Z"
  ].join(" ");
}

function buildEllipsePath() {
  return "M 50 0 A 50 50 0 1 1 49.999 0 Z";
}

function buildPiePath(startAngleDeg: number, endAngleDeg: number) {
  const [startX, startY] = pointOnCircle(50, startAngleDeg);
  const [endX, endY] = pointOnCircle(50, endAngleDeg);
  const sweep = endAngleDeg >= startAngleDeg ? endAngleDeg - startAngleDeg : 360 - (startAngleDeg - endAngleDeg);
  const largeArcFlag = sweep > 180 ? 1 : 0;
  return [
    "M 50 50",
    `L ${startX} ${startY}`,
    `A 50 50 0 ${largeArcFlag} 1 ${endX} ${endY}`,
    "Z"
  ].join(" ");
}

function buildChordPath(startAngleDeg: number, endAngleDeg: number) {
  const [startX, startY] = pointOnCircle(50, startAngleDeg);
  const [endX, endY] = pointOnCircle(50, endAngleDeg);
  const sweep = endAngleDeg >= startAngleDeg ? endAngleDeg - startAngleDeg : 360 - (startAngleDeg - endAngleDeg);
  const largeArcFlag = sweep > 180 ? 1 : 0;
  return [
    `M ${startX} ${startY}`,
    `A 50 50 0 ${largeArcFlag} 1 ${endX} ${endY}`,
    "Z"
  ].join(" ");
}

function buildPlusPath(thickness: number) {
  const arm = Math.max(10, Math.min(40, thickness));
  const inner = 50 - arm / 2;
  const outer = 50 + arm / 2;
  return buildClosedPath([
    [inner, 0],
    [outer, 0],
    [outer, inner],
    [100, inner],
    [100, outer],
    [outer, outer],
    [outer, 100],
    [inner, 100],
    [inner, outer],
    [0, outer],
    [0, inner],
    [inner, inner]
  ]);
}

function buildPentagonArrowPath() {
  return buildClosedPath([
    [0, 20],
    [70, 20],
    [70, 0],
    [100, 50],
    [70, 100],
    [70, 80],
    [0, 80]
  ]);
}

function buildChevronArrowPath() {
  return buildClosedPath([
    [0, 20],
    [54, 20],
    [54, 0],
    [100, 50],
    [54, 100],
    [54, 80],
    [0, 80],
    [32, 50]
  ]);
}

function buildNotchedRightArrowPath(shape: XlsxShape) {
  const headLength = resolveShapeAdjustment(shape, "adj2", 36);
  const shaftHalf = resolveShapeAdjustment(shape, "adj1", 20);
  const notchDepth = Math.min(24, headLength * 0.7);
  const bodyTop = 50 - shaftHalf;
  const bodyBottom = 50 + shaftHalf;
  const headStart = 100 - headLength;
  return buildClosedPath([
    [0, bodyTop],
    [headStart, bodyTop],
    [headStart, 0],
    [100, 50],
    [headStart, 100],
    [headStart, bodyBottom],
    [notchDepth, bodyBottom],
    [0, 50],
    [notchDepth, bodyTop]
  ]);
}

function buildStripedRightArrowPath(shape: XlsxShape) {
  const headLength = resolveShapeAdjustment(shape, "adj2", 34);
  const shaftHalf = resolveShapeAdjustment(shape, "adj1", 20);
  const stripeWidth = 12;
  const bodyTop = 50 - shaftHalf;
  const bodyBottom = 50 + shaftHalf;
  const headStart = 100 - headLength;
  return [
    `M 0 ${bodyTop} L ${stripeWidth} ${bodyTop} L ${stripeWidth} ${bodyBottom} L 0 ${bodyBottom} Z`,
    buildClosedPath([
      [stripeWidth * 1.6, bodyTop],
      [headStart, bodyTop],
      [headStart, 0],
      [100, 50],
      [headStart, 100],
      [headStart, bodyBottom],
      [stripeWidth * 1.6, bodyBottom]
    ])
  ].join(" ");
}

function buildLeftUpArrowPath(shape: XlsxShape) {
  const stemWidth = resolveShapeAdjustment(shape, "adj1", 28);
  const headLength = resolveShapeAdjustment(shape, "adj2", 34);
  return buildClosedPath([
    [100, 100 - stemWidth],
    [stemWidth, 100 - stemWidth],
    [stemWidth, 100],
    [0, 50],
    [stemWidth, 0],
    [stemWidth, headLength],
    [100, headLength]
  ]);
}

function buildBentUpArrowPath(shape: XlsxShape) {
  const stemWidth = resolveShapeAdjustment(shape, "adj1", 26);
  const headLength = resolveShapeAdjustment(shape, "adj2", 28);
  const shaftHalf = resolveShapeAdjustment(shape, "adj3", 12);
  const verticalCenter = 50;
  const bodyTop = verticalCenter - shaftHalf;
  const bodyBottom = verticalCenter + shaftHalf;
  return buildClosedPath([
    [0, bodyTop],
    [100 - stemWidth, bodyTop],
    [100 - stemWidth, headLength],
    [100 - stemWidth * 0.5, headLength],
    [100, 0],
    [100 - stemWidth * 1.5, headLength],
    [100 - stemWidth, headLength],
    [100 - stemWidth, 100],
    [100 - stemWidth * 2, 100],
    [100 - stemWidth * 2, bodyBottom],
    [0, bodyBottom]
  ]);
}

function buildUturnArrowPath(shape: XlsxShape) {
  const stemWidth = resolveShapeAdjustment(shape, "adj1", 22);
  const headLength = resolveShapeAdjustment(shape, "adj2", 28);
  return buildClosedPath([
    [100, 100],
    [100, 55],
    [45, 55],
    [45, 100],
    [0, 50],
    [45, 0],
    [45, 45],
    [100 - stemWidth, 45],
    [100 - stemWidth, 100]
  ]);
}

function buildLeftRightUpArrowPath(shape: XlsxShape) {
  const headLength = resolveShapeAdjustment(shape, "adj2", 22);
  const shaftHalf = resolveShapeAdjustment(shape, "adj1", 16);
  const bodyTop = 50 - shaftHalf;
  const bodyBottom = 50 + shaftHalf;
  const bodyLeft = 50 - shaftHalf;
  const bodyRight = 50 + shaftHalf;
  return buildClosedPath([
    [headLength, bodyTop],
    [bodyLeft, bodyTop],
    [bodyLeft, headLength],
    [0, headLength],
    [50, 0],
    [100, headLength],
    [bodyRight, headLength],
    [bodyRight, bodyTop],
    [100 - headLength, bodyTop],
    [100 - headLength, 0],
    [100, 50],
    [100 - headLength, 100],
    [100 - headLength, bodyBottom],
    [headLength, bodyBottom],
    [headLength, 100],
    [0, 50],
    [headLength, 0]
  ]);
}

function buildQuadArrowPath(shape: XlsxShape) {
  const headLength = resolveShapeAdjustment(shape, "adj2", 22);
  const shaftHalf = resolveShapeAdjustment(shape, "adj1", 14);
  const bodyLeft = 50 - shaftHalf;
  const bodyRight = 50 + shaftHalf;
  const bodyTop = 50 - shaftHalf;
  const bodyBottom = 50 + shaftHalf;
  return buildClosedPath([
    [bodyLeft, headLength],
    [headLength, headLength],
    [headLength, bodyTop],
    [0, 50],
    [headLength, 100 - bodyTop],
    [headLength, bodyBottom],
    [bodyLeft, bodyBottom],
    [bodyLeft, 100 - headLength],
    [bodyTop, 100 - headLength],
    [50, 100],
    [100 - bodyTop, 100 - headLength],
    [bodyRight, 100 - headLength],
    [bodyRight, bodyBottom],
    [100 - headLength, bodyBottom],
    [100 - headLength, 100 - bodyTop],
    [100, 50],
    [100 - headLength, bodyTop],
    [100 - headLength, headLength],
    [bodyRight, headLength],
    [bodyRight, bodyTop],
    [100 - bodyTop, headLength],
    [50, 0],
    [bodyTop, headLength],
    [bodyLeft, headLength]
  ]);
}

function buildHeartPath() {
  return [
    "M 50 92",
    "C 18 72 0 52 0 28",
    "C 0 10 14 0 28 0",
    "C 38 0 46 6 50 16",
    "C 54 6 62 0 72 0",
    "C 86 0 100 10 100 28",
    "C 100 52 82 72 50 92",
    "Z"
  ].join(" ");
}

function buildLightningBoltPath() {
  return buildClosedPath([
    [42, 0],
    [10, 58],
    [38, 58],
    [24, 100],
    [90, 36],
    [58, 36],
    [72, 0]
  ]);
}

function buildTeardropPath() {
  return [
    "M 50 0",
    "C 20 24 0 44 0 68",
    "C 0 86 14 100 32 100",
    "C 60 100 82 78 82 50",
    "C 82 26 68 10 50 0",
    "Z"
  ].join(" ");
}

function buildCloudPath() {
  return [
    "M 22 70",
    "C 8 70 0 62 0 50",
    "C 0 40 6 32 18 30",
    "C 20 16 30 8 44 8",
    "C 54 8 62 12 68 20",
    "C 72 14 80 12 88 14",
    "C 98 18 100 26 100 34",
    "C 100 44 94 50 88 52",
    "C 92 66 82 76 68 76",
    "C 62 84 52 88 42 88",
    "C 32 88 24 82 22 70",
    "Z"
  ].join(" ");
}

function buildRightArrowPath(shape: XlsxShape) {
  const headLength = resolveShapeAdjustment(shape, "adj2", 38);
  const shaftHalf = resolveShapeAdjustment(shape, "adj1", 20);
  const bodyTop = 50 - shaftHalf;
  const bodyBottom = 50 + shaftHalf;
  const headStart = 100 - headLength;
  return buildClosedPath([
    [0, bodyTop],
    [headStart, bodyTop],
    [headStart, 0],
    [100, 50],
    [headStart, 100],
    [headStart, bodyBottom],
    [0, bodyBottom]
  ]);
}

function buildLeftArrowPath(shape: XlsxShape) {
  const headLength = resolveShapeAdjustment(shape, "adj2", 38);
  const shaftHalf = resolveShapeAdjustment(shape, "adj1", 20);
  const bodyTop = 50 - shaftHalf;
  const bodyBottom = 50 + shaftHalf;
  const headEnd = headLength;
  return buildClosedPath([
    [100, bodyTop],
    [headEnd, bodyTop],
    [headEnd, 0],
    [0, 50],
    [headEnd, 100],
    [headEnd, bodyBottom],
    [100, bodyBottom]
  ]);
}

function buildUpArrowPath(shape: XlsxShape) {
  const headLength = resolveShapeAdjustment(shape, "adj2", 38);
  const shaftHalf = resolveShapeAdjustment(shape, "adj1", 20);
  const bodyLeft = 50 - shaftHalf;
  const bodyRight = 50 + shaftHalf;
  const headEnd = headLength;
  return buildClosedPath([
    [bodyLeft, 100],
    [bodyLeft, headEnd],
    [0, headEnd],
    [50, 0],
    [100, headEnd],
    [bodyRight, headEnd],
    [bodyRight, 100]
  ]);
}

function buildDownArrowPath(shape: XlsxShape) {
  const headLength = resolveShapeAdjustment(shape, "adj2", 38);
  const shaftHalf = resolveShapeAdjustment(shape, "adj1", 20);
  const bodyLeft = 50 - shaftHalf;
  const bodyRight = 50 + shaftHalf;
  const headStart = 100 - headLength;
  return buildClosedPath([
    [bodyLeft, 0],
    [bodyLeft, headStart],
    [0, headStart],
    [50, 100],
    [100, headStart],
    [bodyRight, headStart],
    [bodyRight, 0]
  ]);
}

function buildLeftRightArrowPath(shape: XlsxShape) {
  const headLength = resolveShapeAdjustment(shape, "adj2", 24);
  const shaftHalf = resolveShapeAdjustment(shape, "adj1", 18);
  const bodyTop = 50 - shaftHalf;
  const bodyBottom = 50 + shaftHalf;
  return buildClosedPath([
    [headLength, bodyTop],
    [100 - headLength, bodyTop],
    [100 - headLength, 0],
    [100, 50],
    [100 - headLength, 100],
    [100 - headLength, bodyBottom],
    [headLength, bodyBottom],
    [headLength, 100],
    [0, 50],
    [headLength, 0]
  ]);
}

function buildUpDownArrowPath(shape: XlsxShape) {
  const headLength = resolveShapeAdjustment(shape, "adj2", 24);
  const shaftHalf = resolveShapeAdjustment(shape, "adj1", 18);
  const bodyLeft = 50 - shaftHalf;
  const bodyRight = 50 + shaftHalf;
  return buildClosedPath([
    [bodyLeft, headLength],
    [bodyLeft, 100 - headLength],
    [0, 100 - headLength],
    [50, 100],
    [100, 100 - headLength],
    [bodyRight, 100 - headLength],
    [bodyRight, headLength],
    [100, headLength],
    [50, 0],
    [0, headLength]
  ]);
}

function buildBentArrowPath(shape: XlsxShape) {
  const stemWidth = resolveShapeAdjustment(shape, "adj1", 25);
  const headLength = resolveShapeAdjustment(shape, "adj2", 27);
  const shaftHalf = resolveShapeAdjustment(shape, "adj3", 12.5);
  const bendY = resolveShapeAdjustment(shape, "adj4", 43.75);
  const bodyTop = bendY - shaftHalf;
  const bodyBottom = bendY + shaftHalf;
  const headStart = 100 - headLength;
  return buildClosedPath([
    [0, bodyTop],
    [headStart, bodyTop],
    [headStart, 0],
    [100, bendY],
    [headStart, 100],
    [headStart, bodyBottom],
    [stemWidth, bodyBottom],
    [stemWidth, 100],
    [0, 100]
  ]);
}

function buildPresetShapePath(shape: XlsxShape) {
  switch (shape.geometry) {
    case "arc":
      return {
        path: "M 8 74 C 18 24 82 24 92 74",
        viewBox: { width: 100, height: 100 }
      };
    case "bentArrow":
      return {
        path: buildBentArrowPath(shape),
        viewBox: { width: 100, height: 100 }
      };
    case "bentUpArrow":
      return {
        path: buildBentUpArrowPath(shape),
        viewBox: { width: 100, height: 100 }
      };
    case "blockArc":
      return {
        path: "M 12 78 A 38 38 0 1 1 78 12 L 62 28 A 16 16 0 1 0 28 62 Z",
        viewBox: { width: 100, height: 100 }
      };
    case "can":
      return {
        path: "M 18 14 C 18 6 82 6 82 14 L 82 86 C 82 94 18 94 18 86 Z",
        viewBox: { width: 100, height: 100 }
      };
    case "chevron":
      return {
        path: "M 0 0 L 62 0 L 100 50 L 62 100 L 0 100 L 38 50 Z",
        viewBox: { width: 100, height: 100 }
      };
    case "chevronArrow":
      return {
        path: buildChevronArrowPath(),
        viewBox: { width: 100, height: 100 }
      };
    case "chord":
      return {
        path: buildChordPath(-30, 210),
        viewBox: { width: 100, height: 100 }
      };
    case "cloud":
    case "cloudCallout":
      return {
        path: buildCloudPath(),
        viewBox: { width: 100, height: 100 }
      };
    case "diamond":
    case "flowChartDecision":
      return {
        path: "M 50 0 L 100 50 L 50 100 L 0 50 Z",
        viewBox: { width: 100, height: 100 }
      };
    case "downArrow":
      return {
        path: buildDownArrowPath(shape),
        viewBox: { width: 100, height: 100 }
      };
    case "line":
      return {
        path: "M 0 0 L 100 100",
        viewBox: { width: 100, height: 100 }
      };
    case "ellipse":
    case "flowChartTerminator":
      return {
        path: buildEllipsePath(),
        viewBox: { width: 100, height: 100 }
      };
    case "flowChartAlternateProcess":
    case "flowChartDelay":
    case "roundRect":
      return {
        path: buildRoundRectPath(resolveShapeAdjustment(shape, "adj", 18)),
        viewBox: { width: 100, height: 100 }
      };
    case "flowChartCollate":
      return {
        path: "M 0 0 L 100 0 L 0 100 Z",
        viewBox: { width: 100, height: 100 }
      };
    case "flowChartData":
    case "flowChartInputOutput":
      return {
        path: "M 24 0 L 100 0 L 76 100 L 0 100 Z",
        viewBox: { width: 100, height: 100 }
      };
    case "flowChartDocument":
      return {
        path: "M 0 0 L 100 0 L 100 82 C 78 70 60 94 40 82 C 24 72 12 88 0 82 Z",
        viewBox: { width: 100, height: 100 }
      };
    case "flowChartMagneticDisk":
      return {
        path: "M 18 14 C 18 6 82 6 82 14 L 82 86 C 82 94 18 94 18 86 Z",
        viewBox: { width: 100, height: 100 }
      };
    case "flowChartManualInput":
      return {
        path: "M 16 0 L 100 0 L 84 100 L 0 100 Z",
        viewBox: { width: 100, height: 100 }
      };
    case "flowChartPreparation":
    case "hexagon":
      return {
        path: "M 25 0 L 75 0 L 100 50 L 75 100 L 25 100 L 0 50 Z",
        viewBox: { width: 100, height: 100 }
      };
    case "flowChartProcess":
    case "rect":
      return {
        path: "M 0 0 L 100 0 L 100 100 L 0 100 Z",
        viewBox: { width: 100, height: 100 }
      };
    case "heart":
      return {
        path: buildHeartPath(),
        viewBox: { width: 100, height: 100 }
      };
    case "heptagon":
      return {
        path: buildRegularPolygonPath(7),
        viewBox: { width: 100, height: 100 }
      };
    case "homePlate":
    case "flowChartOffpageConnector":
      return {
        path: "M 0 0 L 70 0 L 100 50 L 70 100 L 0 100 Z",
        viewBox: { width: 100, height: 100 }
      };
    case "leftArrow":
      return {
        path: buildLeftArrowPath(shape),
        viewBox: { width: 100, height: 100 }
      };
    case "leftBracket":
      return {
        path: "M 72 0 L 28 0 L 28 100 L 72 100",
        viewBox: { width: 100, height: 100 }
      };
    case "leftBrace":
      return {
        path: "M 82 0 C 46 0 52 24 52 38 C 52 46 46 50 24 50 C 46 50 52 54 52 62 C 52 76 46 100 82 100",
        viewBox: { width: 100, height: 100 }
      };
    case "leftRightArrowCallout":
      return {
        path: buildLeftRightArrowPath(shape),
        viewBox: { width: 100, height: 100 }
      };
    case "leftRightArrow":
      return {
        path: buildLeftRightArrowPath(shape),
        viewBox: { width: 100, height: 100 }
      };
    case "leftRightUpArrow":
      return {
        path: buildLeftRightUpArrowPath(shape),
        viewBox: { width: 100, height: 100 }
      };
    case "leftUpArrow":
      return {
        path: buildLeftUpArrowPath(shape),
        viewBox: { width: 100, height: 100 }
      };
    case "lightningBolt":
      return {
        path: buildLightningBoltPath(),
        viewBox: { width: 100, height: 100 }
      };
    case "mathPlus":
    case "plus":
      return {
        path: buildPlusPath(resolveShapeAdjustment(shape, "adj", 26)),
        viewBox: { width: 100, height: 100 }
      };
    case "cross":
      return {
        path: buildPlusPath(resolveShapeAdjustment(shape, "adj", 22)),
        viewBox: { width: 100, height: 100 }
      };
    case "mathMinus":
    case "minus":
      return {
        path: "M 0 40 L 100 40 L 100 60 L 0 60 Z",
        viewBox: { width: 100, height: 100 }
      };
    case "mathMultiply":
    case "multiply":
      return {
        path: "M 18 0 L 50 32 L 82 0 L 100 18 L 68 50 L 100 82 L 82 100 L 50 68 L 18 100 L 0 82 L 32 50 L 0 18 Z",
        viewBox: { width: 100, height: 100 }
      };
    case "mathDivide":
      return {
        path: "M 0 42 L 100 42 L 100 58 L 0 58 Z M 50 12 A 8 8 0 1 1 49.999 12 Z M 50 88 A 8 8 0 1 1 49.999 88 Z",
        viewBox: { width: 100, height: 100 }
      };
    case "moon":
      return {
        path: "M 70 6 C 42 10 22 34 22 62 C 22 80 32 94 48 100 C 20 98 0 76 0 48 C 0 20 20 0 48 0 C 56 0 64 2 70 6 Z",
        viewBox: { width: 100, height: 100 }
      };
    case "notchedRightArrow":
      return {
        path: buildNotchedRightArrowPath(shape),
        viewBox: { width: 100, height: 100 }
      };
    case "octagon":
      return {
        path: "M 30 0 L 70 0 L 100 30 L 100 70 L 70 100 L 30 100 L 0 70 L 0 30 Z",
        viewBox: { width: 100, height: 100 }
      };
    case "parallelogram":
      return {
        path: "M 24 0 L 100 0 L 76 100 L 0 100 Z",
        viewBox: { width: 100, height: 100 }
      };
    case "pentagon":
      return {
        path: "M 50 0 L 100 38 L 81 100 L 19 100 L 0 38 Z",
        viewBox: { width: 100, height: 100 }
      };
    case "pie":
    case "pieWedge":
      return {
        path: buildPiePath(-35, 225),
        viewBox: { width: 100, height: 100 }
      };
    case "pentagonArrow":
      return {
        path: buildPentagonArrowPath(),
        viewBox: { width: 100, height: 100 }
      };
    case "plaque":
      return {
        path: "M 12 0 L 88 0 L 100 12 L 100 88 L 88 100 L 12 100 L 0 88 L 0 12 Z",
        viewBox: { width: 100, height: 100 }
      };
    case "quadArrow":
      return {
        path: buildQuadArrowPath(shape),
        viewBox: { width: 100, height: 100 }
      };
    case "rightArrow":
      return {
        path: buildRightArrowPath(shape),
        viewBox: { width: 100, height: 100 }
      };
    case "rightArrowCallout":
      return {
        path: "M 0 18 L 62 18 L 62 0 L 100 50 L 62 100 L 62 82 L 0 82 L 0 18 Z",
        viewBox: { width: 100, height: 100 }
      };
    case "rightBracket":
      return {
        path: "M 28 0 L 72 0 L 72 100 L 28 100",
        viewBox: { width: 100, height: 100 }
      };
    case "downArrowCallout":
      return {
        path: "M 18 0 L 82 0 L 82 58 L 100 58 L 50 100 L 0 58 L 18 58 L 18 0 Z",
        viewBox: { width: 100, height: 100 }
      };
    case "rightBrace":
      return {
        path: "M 18 0 C 54 0 48 24 48 38 C 48 46 54 50 76 50 C 54 50 48 54 48 62 C 48 76 54 100 18 100",
        viewBox: { width: 100, height: 100 }
      };
    case "round1Rect":
      return {
        path: "M 18 0 L 100 0 L 100 100 L 18 100 Q 0 100 0 82 L 0 18 Q 0 0 18 0 Z",
        viewBox: { width: 100, height: 100 }
      };
    case "round2DiagRect":
      return {
        path: "M 18 0 L 100 0 L 100 82 Q 100 100 82 100 L 0 100 L 0 18 Q 0 0 18 0 Z",
        viewBox: { width: 100, height: 100 }
      };
    case "round2SameRect":
      return {
        path: "M 18 0 L 82 0 Q 100 0 100 18 L 100 82 Q 100 100 82 100 L 0 100 L 0 0 Z",
        viewBox: { width: 100, height: 100 }
      };
    case "decagon":
      return {
        path: buildRegularPolygonPath(10),
        viewBox: { width: 100, height: 100 }
      };
    case "dodecagon":
      return {
        path: buildRegularPolygonPath(12),
        viewBox: { width: 100, height: 100 }
      };
    case "rtTriangle":
      return {
        path: "M 0 0 L 100 100 L 0 100 Z",
        viewBox: { width: 100, height: 100 }
      };
    case "star4":
      return {
        path: buildStarPath(4, 20),
        viewBox: { width: 100, height: 100 }
      };
    case "star5":
      return {
        path: buildStarPath(5, 21),
        viewBox: { width: 100, height: 100 }
      };
    case "star6":
      return {
        path: buildStarPath(6, 23),
        viewBox: { width: 100, height: 100 }
      };
    case "star7":
      return {
        path: buildStarPath(7, 24),
        viewBox: { width: 100, height: 100 }
      };
    case "star8":
      return {
        path: buildStarPath(8, 25),
        viewBox: { width: 100, height: 100 }
      };
    case "star10":
      return {
        path: buildStarPath(10, 26),
        viewBox: { width: 100, height: 100 }
      };
    case "star12":
      return {
        path: buildStarPath(12, 26),
        viewBox: { width: 100, height: 100 }
      };
    case "star16":
      return {
        path: buildStarPath(16, 27),
        viewBox: { width: 100, height: 100 }
      };
    case "star24":
      return {
        path: buildStarPath(24, 28),
        viewBox: { width: 100, height: 100 }
      };
    case "star32":
      return {
        path: buildStarPath(32, 29),
        viewBox: { width: 100, height: 100 }
      };
    case "stripedRightArrow":
      return {
        path: buildStripedRightArrowPath(shape),
        viewBox: { width: 100, height: 100 }
      };
    case "sun":
      return {
        path: buildStarPath(12, 36),
        viewBox: { width: 100, height: 100 }
      };
    case "teardrop":
      return {
        path: buildTeardropPath(),
        viewBox: { width: 100, height: 100 }
      };
    case "triangle":
      return {
        path: "M 50 0 L 100 100 L 0 100 Z",
        viewBox: { width: 100, height: 100 }
      };
    case "trapezoid":
      return {
        path: "M 20 0 L 80 0 L 100 100 L 0 100 Z",
        viewBox: { width: 100, height: 100 }
      };
    case "upArrow":
      return {
        path: buildUpArrowPath(shape),
        viewBox: { width: 100, height: 100 }
      };
    case "upArrowCallout":
      return {
        path: "M 50 0 L 100 42 L 82 42 L 82 100 L 18 100 L 18 42 L 0 42 L 50 0 Z",
        viewBox: { width: 100, height: 100 }
      };
    case "upDownArrow":
      return {
        path: buildUpDownArrowPath(shape),
        viewBox: { width: 100, height: 100 }
      };
    case "upDownArrowCallout":
      return {
        path: buildUpDownArrowPath(shape),
        viewBox: { width: 100, height: 100 }
      };
    case "uturnArrow":
      return {
        path: buildUturnArrowPath(shape),
        viewBox: { width: 100, height: 100 }
      };
    default:
      return null;
  }
}

function resolveShapeVector(shape: XlsxShape) {
  if (shape.svgPath && shape.svgViewBox) {
    return {
      path: shape.svgPath,
      viewBox: shape.svgViewBox
    };
  }

  return buildPresetShapePath(shape);
}

function drawCanvasRoundedRect(
  context: CanvasRenderingContext2D,
  left: number,
  top: number,
  width: number,
  height: number,
  radius: number
) {
  const safeRadius = Math.max(0, Math.min(radius, width / 2, height / 2));
  context.beginPath();
  context.moveTo(left + safeRadius, top);
  context.lineTo(left + width - safeRadius, top);
  context.quadraticCurveTo(left + width, top, left + width, top + safeRadius);
  context.lineTo(left + width, top + height - safeRadius);
  context.quadraticCurveTo(left + width, top + height, left + width - safeRadius, top + height);
  context.lineTo(left + safeRadius, top + height);
  context.quadraticCurveTo(left, top + height, left, top + height - safeRadius);
  context.lineTo(left, top + safeRadius);
  context.quadraticCurveTo(left, top, left + safeRadius, top);
  context.closePath();
}

function applyCanvasShapeDash(context: CanvasRenderingContext2D, dash: string | undefined, lineWidth: number) {
  if (!dash) {
    context.setLineDash([]);
    return;
  }

  const unit = Math.max(1, lineWidth);
  switch (dash) {
    case "dash":
      context.setLineDash([4 * unit, 3 * unit]);
      break;
    case "dashDot":
      context.setLineDash([4 * unit, 2 * unit, unit, 2 * unit]);
      break;
    case "dot":
      context.setLineDash([unit, 2 * unit]);
      break;
    case "lgDash":
      context.setLineDash([8 * unit, 3 * unit]);
      break;
    default:
      context.setLineDash([]);
      break;
  }
}

function resolveShapeLineEndMarker(
  type: string | undefined,
  markerId: string,
  color: string,
  strokeWidth: number,
  rect: XlsxImageRect,
  viewBox: { height: number; width: number }
) {
  if (type !== "triangle") {
    return null;
  }

  const pxToUserSpace = ((viewBox.width / Math.max(1, rect.width)) + (viewBox.height / Math.max(1, rect.height))) / 2;
  const markerSize = Math.max(pxToUserSpace * 8, pxToUserSpace * strokeWidth * 4);
  return (
    <marker
      id={markerId}
      key={markerId}
      markerHeight={markerSize}
      markerUnits="userSpaceOnUse"
      markerWidth={markerSize}
      orient="auto-start-reverse"
      overflow="visible"
      refX={markerSize}
      refY={markerSize / 2}
      viewBox={`0 0 ${markerSize} ${markerSize}`}
    >
      <path d={`M 0 0 L ${markerSize} ${markerSize / 2} L 0 ${markerSize} z`} fill={color} />
    </marker>
  );
}

function renderShapeParagraph(
  paragraph: XlsxShape["paragraphs"][number],
  index: number,
  fallbackAlign: React.CSSProperties["textAlign"] = "left",
  textScale = 1
) {
  return (
    <p
      key={index}
      style={{
        margin: 0,
        textAlign: paragraph.align ?? fallbackAlign,
        whiteSpace: "pre-wrap"
      }}
    >
      {paragraph.runs.map((run, runIndex) => (
        <span
          key={runIndex}
          style={{
            color: run.color,
            fontFamily: run.fontFamily,
            fontSize: run.fontSizePt ? `${run.fontSizePt * textScale}pt` : undefined,
            fontStyle: run.italic ? "italic" : undefined,
            fontWeight: run.bold ? 700 : undefined,
            textDecoration: run.underline ? "underline" : undefined
          }}
        >
          {run.text}
        </span>
      ))}
    </p>
  );
}

function clampImageRect(
  rect: XlsxImageRect,
  options?: {
    contentOffsetLeft?: number;
    contentOffsetTop?: number;
    minSizePx?: number;
  }
): XlsxImageRect {
  const contentOffsetLeft = options?.contentOffsetLeft ?? ROW_HEADER_WIDTH;
  const contentOffsetTop = options?.contentOffsetTop ?? HEADER_HEIGHT;
  const minSizePx = options?.minSizePx ?? IMAGE_MIN_SIZE_PX;
  return {
    height: Math.max(minSizePx, rect.height),
    left: Math.max(contentOffsetLeft, rect.left),
    top: Math.max(contentOffsetTop, rect.top),
    width: Math.max(minSizePx, rect.width)
  };
}

function resolveImageHandleStyle(
  position: XlsxImageResizeHandlePosition,
  stroke: string,
  surface: string,
  scale = 1
): React.CSSProperties {
  const handleSize = IMAGE_HANDLE_SIZE_PX * scale;
  const offset = handleSize / 2;
  const style: React.CSSProperties = {
    backgroundColor: surface,
    border: `${Math.max(1, scale)}px solid ${stroke}`,
    borderRadius: 3 * scale,
    boxShadow: `0 1px ${3 * scale}px ${DRAWING_SELECTION_HANDLE_SHADOW}`,
    boxSizing: "border-box",
    cursor: IMAGE_HANDLE_CURSOR[position],
    height: handleSize,
    pointerEvents: "auto",
    position: "absolute",
    width: handleSize
  };

  if (position.includes("n")) {
    style.top = -offset;
  }
  if (position.includes("s")) {
    style.bottom = -offset;
  }
  if (position.includes("w")) {
    style.left = -offset;
  }
  if (position.includes("e")) {
    style.right = -offset;
  }
  if (position === "n" || position === "s") {
    style.left = `calc(50% - ${offset}px)`;
  }
  if (position === "e" || position === "w") {
    style.top = `calc(50% - ${offset}px)`;
  }

  return style;
}

function useViewerPalette(isDark = false) {
  return isDark ? DARK_PALETTE : LIGHT_PALETTE;
}

type SheetThumbnailAxisItem = {
  actualIndex: number;
  size: number;
  start: number;
};

type SheetThumbnailRenderPlan = {
  aspectRatio: number;
  contentHeight: number;
  contentWidth: number;
  height: number;
  paint: (canvas: HTMLCanvasElement | null) => boolean;
  sourceRange: XlsxCellRange;
  width: number;
};

function normalizeThumbnailResolution(
  resolution?: UseXlsxViewerThumbnailsOptions["resolution"]
) {
  if (typeof resolution === "number" && Number.isFinite(resolution) && resolution > 0) {
    return {
      maxHeight: resolution,
      maxWidth: resolution
    };
  }

  const resolvedObject = typeof resolution === "object" && resolution ? resolution : undefined;
  return {
    maxHeight: resolvedObject?.maxHeight && Number.isFinite(resolvedObject.maxHeight) && resolvedObject.maxHeight > 0
      ? resolvedObject.maxHeight
      : THUMBNAIL_DEFAULT_MAX_DIMENSION,
    maxWidth: resolvedObject?.maxWidth && Number.isFinite(resolvedObject.maxWidth) && resolvedObject.maxWidth > 0
      ? resolvedObject.maxWidth
      : THUMBNAIL_DEFAULT_MAX_DIMENSION
  };
}

function resolveThumbnailOutputSize(
  sourceWidth: number,
  sourceHeight: number,
  resolution?: UseXlsxViewerThumbnailsOptions["resolution"]
) {
  const normalized = normalizeThumbnailResolution(resolution);
  const safeSourceWidth = Math.max(1, sourceWidth);
  const safeSourceHeight = Math.max(1, sourceHeight);
  const scale = Math.min(
    normalized.maxWidth / safeSourceWidth,
    normalized.maxHeight / safeSourceHeight,
    1
  );

  return {
    height: Math.max(1, Math.round(safeSourceHeight * scale)),
    maxHeight: normalized.maxHeight,
    maxWidth: normalized.maxWidth,
    scale,
    width: Math.max(1, Math.round(safeSourceWidth * scale))
  };
}

function resolveThumbnailAxisIndices({
  fallbackCount,
  getSizePx,
  maxCount,
  maxLogicalPixels,
  maxUsedIndex,
  minUsedIndex,
  precomputed
}: {
  fallbackCount: number;
  getSizePx: (actualIndex: number) => number;
  maxCount: number;
  maxLogicalPixels: number;
  maxUsedIndex: number;
  minUsedIndex: number;
  precomputed: number[];
}) {
  const sourceIndices = precomputed.length > 0
    ? precomputed
    : Array.from({
        length: Math.max(fallbackCount, Math.max(0, maxUsedIndex + 1))
      }, (_, index) => index);
  const hasUsedRange = maxUsedIndex >= minUsedIndex && maxUsedIndex >= 0;
  const preferredStart = hasUsedRange ? Math.max(0, minUsedIndex) : 0;
  const preferredEnd = hasUsedRange ? maxUsedIndex : Math.max(0, preferredStart + fallbackCount - 1);
  const picked: number[] = [];
  let totalSize = 0;

  for (const actualIndex of sourceIndices) {
    if (actualIndex < preferredStart) {
      continue;
    }
    if (hasUsedRange && actualIndex > preferredEnd) {
      break;
    }

    const size = Math.max(1, getSizePx(actualIndex));
    picked.push(actualIndex);
    totalSize += size;
    if (picked.length >= maxCount || totalSize >= maxLogicalPixels) {
      break;
    }
  }

  if (picked.length > 0) {
    return picked;
  }

  for (const actualIndex of sourceIndices) {
    const size = Math.max(1, getSizePx(actualIndex));
    picked.push(actualIndex);
    totalSize += size;
    if (picked.length >= fallbackCount || picked.length >= maxCount || totalSize >= maxLogicalPixels) {
      break;
    }
  }

  return picked.length > 0 ? picked : [0];
}

function buildThumbnailAxisItems(actualIndices: number[], getSizePx: (actualIndex: number) => number) {
  const items: SheetThumbnailAxisItem[] = [];
  let offset = 0;

  for (const actualIndex of actualIndices) {
    const size = Math.max(1, getSizePx(actualIndex));
    items.push({
      actualIndex,
      size,
      start: offset
    });
    offset += size;
  }

  return {
    items,
    totalSize: offset
  };
}

function resolveThumbnailAxisAbsoluteOffset(
  actualIndices: number[],
  getSizePx: (actualIndex: number) => number,
  actualIndex: number
) {
  let total = 0;

  for (const candidate of actualIndices) {
    if (candidate >= actualIndex) {
      break;
    }
    total += Math.max(1, getSizePx(candidate));
  }

  return total;
}

function resolveThumbnailAxisMarkerOffset({
  actualIndex,
  actualIndices,
  getSizePx,
  offsetEmu,
  previewStartActualIndex
}: {
  actualIndex: number;
  actualIndices: number[];
  getSizePx: (actualIndex: number) => number;
  offsetEmu: number;
  previewStartActualIndex: number;
}) {
  return (
    resolveThumbnailAxisAbsoluteOffset(actualIndices, getSizePx, actualIndex)
    - resolveThumbnailAxisAbsoluteOffset(actualIndices, getSizePx, previewStartActualIndex)
    + emuToPixels(offsetEmu)
  );
}

function resolveThumbnailAnchoredRect(
  anchor: XlsxImage["anchor"],
  visibleRows: number[],
  visibleCols: number[],
  getRowHeightPx: (actualIndex: number) => number,
  getColWidthPx: (actualIndex: number) => number,
  previewRows: number[],
  previewCols: number[],
  options: {
    headerHeight: number;
    rowHeaderWidth: number;
  }
): XlsxImageRect {
  const previewStartRow = previewRows[0] ?? 0;
  const previewStartCol = previewCols[0] ?? 0;
  const previewOriginX = resolveThumbnailAxisAbsoluteOffset(visibleCols, getColWidthPx, previewStartCol);
  const previewOriginY = resolveThumbnailAxisAbsoluteOffset(visibleRows, getRowHeightPx, previewStartRow);
  const resolveMarkerLeft = (col: number, colOffsetEmu: number) =>
    options.rowHeaderWidth + resolveThumbnailAxisMarkerOffset({
      actualIndex: col,
      actualIndices: visibleCols,
      getSizePx: getColWidthPx,
      offsetEmu: colOffsetEmu,
      previewStartActualIndex: previewStartCol
    });
  const resolveMarkerTop = (row: number, rowOffsetEmu: number) =>
    options.headerHeight + resolveThumbnailAxisMarkerOffset({
      actualIndex: row,
      actualIndices: visibleRows,
      getSizePx: getRowHeightPx,
      offsetEmu: rowOffsetEmu,
      previewStartActualIndex: previewStartRow
    });

  if (anchor.kind === "absolute") {
    return {
      height: Math.max(1, emuToPixels(anchor.sizeEmu.cy)),
      left: options.rowHeaderWidth + emuToPixels(anchor.positionEmu.x) - previewOriginX,
      top: options.headerHeight + emuToPixels(anchor.positionEmu.y) - previewOriginY,
      width: Math.max(1, emuToPixels(anchor.sizeEmu.cx))
    };
  }

  if (anchor.kind === "one-cell") {
    return {
      height: Math.max(1, emuToPixels(anchor.sizeEmu.cy)),
      left: resolveMarkerLeft(anchor.from.col, anchor.from.colOffsetEmu),
      top: resolveMarkerTop(anchor.from.row, anchor.from.rowOffsetEmu),
      width: Math.max(1, emuToPixels(anchor.sizeEmu.cx))
    };
  }

  const left = resolveMarkerLeft(anchor.from.col, anchor.from.colOffsetEmu);
  const top = resolveMarkerTop(anchor.from.row, anchor.from.rowOffsetEmu);
  const right = resolveMarkerLeft(anchor.to.col, anchor.to.colOffsetEmu);
  const bottom = resolveMarkerTop(anchor.to.row, anchor.to.rowOffsetEmu);

  return {
    height: Math.max(1, bottom - top),
    left,
    top,
    width: Math.max(1, right - left)
  };
}

function columnLabel(col: number): string {
  let label = "";
  let nextValue = col;

  while (nextValue >= 0) {
    label = String.fromCharCode(65 + (nextValue % 26)) + label;
    nextValue = Math.floor(nextValue / 26) - 1;
  }

  return label;
}

function cellAddressToA1(cell: XlsxCellAddress): string {
  return `${columnLabel(cell.col)}${cell.row + 1}`;
}

function parseA1CellReference(reference: string): XlsxCellAddress | null {
  const match = /^([A-Z]+)(\d+)$/i.exec(reference.trim());
  if (!match) {
    return null;
  }

  const [, columnPart, rowPart] = match;
  let col = 0;
  for (const char of columnPart.toUpperCase()) {
    col = col * 26 + (char.charCodeAt(0) - 64);
  }

  return {
    col: col - 1,
    row: Number(rowPart) - 1
  };
}

function resolveInheritedCellStyle(sheet: XlsxSheetData | null | undefined, row: number, col: number) {
  if (!sheet) {
    return null;
  }

  const colStyleId = sheet.colStyleIds[col];
  const rowStyleId = sheet.rowStyleIds[row];
  const colStyle = colStyleId !== undefined ? sheet.styleById[colStyleId] ?? null : null;
  const rowStyle = rowStyleId !== undefined ? sheet.styleById[rowStyleId] ?? null : null;
  return mergeResolvedCellStyle(colStyle, rowStyle, { replaceXfSubtrees: true });
}

function parseInternalSheetLink(target?: string | null) {
  if (!target) {
    return null;
  }

  const normalized = target.startsWith("#") ? target.slice(1) : target;
  const separatorIndex = normalized.lastIndexOf("!");
  if (separatorIndex < 0) {
    return null;
  }

  const rawSheetName = normalized.slice(0, separatorIndex).trim();
  const rawCellRef = normalized.slice(separatorIndex + 1).trim();
  const cell = parseA1CellReference(rawCellRef);
  if (!cell) {
    return null;
  }

  const sheetName = rawSheetName.startsWith("'") && rawSheetName.endsWith("'")
    ? rawSheetName.slice(1, -1).replace(/''/g, "'")
    : rawSheetName;

  return {
    cell,
    sheetName
  };
}

function decodeHtmlEntities(value: string): string {
  return value
    .replace(/&quot;/g, "\"")
    .replace(/&#34;/g, "\"")
    .replace(/&apos;/g, "'")
    .replace(/&#39;/g, "'")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&amp;/g, "&");
}

function mapBorder(edge: { style: string; color?: Record<string, unknown> }, themePalette?: XlsxSheetData["themePalette"]): string {
  const color = resolveWorkbookColor(edge.color as Record<string, unknown> | undefined, themePalette) ?? "#000";
  const widthMap: Record<string, string> = {
    dashed: "1px",
    dotted: "1px",
    double: "3px",
    hair: "0.5px",
    medium: "2px",
    thick: "3px",
    thin: "1px"
  };
  const styleMap: Record<string, string> = {
    dashDot: "dashed",
    dashDotDot: "dotted",
    dashed: "dashed",
    dotted: "dotted",
    double: "double",
    hair: "solid",
    medium: "solid",
    mediumDashDot: "dashed",
    mediumDashDotDot: "dotted",
    mediumDashed: "dashed",
    slantDashDot: "dashed",
    thick: "solid",
    thin: "solid"
  };

  return `${widthMap[edge.style] ?? "1px"} ${styleMap[edge.style] ?? "solid"} ${color}`;
}

function paletteIsDark(palette: ViewerPalette) {
  return palette.surface === DARK_PALETTE.surface;
}

function parseHexColor(color: string): [number, number, number] | null {
  const normalized = color.trim().toLowerCase();
  const hexMatch = /^#([0-9a-f]{6})$/.exec(normalized);
  if (hexMatch) {
    const hex = hexMatch[1];
    return [
      Number.parseInt(hex.slice(0, 2), 16),
      Number.parseInt(hex.slice(2, 4), 16),
      Number.parseInt(hex.slice(4, 6), 16)
    ];
  }

  const rgbMatch = /^rgb\(\s*(\d{1,3})\s*[,\s]\s*(\d{1,3})\s*[,\s]\s*(\d{1,3})\s*\)$/.exec(normalized);
  if (rgbMatch) {
    return [
      Math.max(0, Math.min(255, Number.parseInt(rgbMatch[1], 10))),
      Math.max(0, Math.min(255, Number.parseInt(rgbMatch[2], 10))),
      Math.max(0, Math.min(255, Number.parseInt(rgbMatch[3], 10)))
    ];
  }

  const hslMatch = /^hsl\(\s*(-?\d+(?:\.\d+)?)\s*(?:deg)?(?:\s+|,\s*)(\d+(?:\.\d+)?)%(?:\s+|,\s*)(\d+(?:\.\d+)?)%\s*\)$/.exec(normalized);
  if (!hslMatch) {
    return null;
  }

  const hue = ((Number.parseFloat(hslMatch[1]) % 360) + 360) % 360 / 360;
  const saturation = Math.max(0, Math.min(1, Number.parseFloat(hslMatch[2]) / 100));
  const lightness = Math.max(0, Math.min(1, Number.parseFloat(hslMatch[3]) / 100));
  return hslToRgb(hue, saturation, lightness);
}

function applyAlphaToColor(color: string, alpha: number) {
  const rgb = parseHexColor(color);
  if (!rgb) {
    return color;
  }

  const [red, green, blue] = rgb;
  return `rgba(${red}, ${green}, ${blue}, ${Math.max(0, Math.min(1, alpha))})`;
}

function relativeLuminance(color: string) {
  const rgb = parseHexColor(color);
  if (!rgb) {
    return null;
  }

  const [red, green, blue] = rgb.map((channel) => {
    const normalized = channel / 255;
    return normalized <= 0.03928 ? normalized / 12.92 : ((normalized + 0.055) / 1.055) ** 2.4;
  });

  return 0.2126 * red + 0.7152 * green + 0.0722 * blue;
}

function contrastRatio(foreground: string, background: string) {
  const foregroundLuminance = relativeLuminance(foreground);
  const backgroundLuminance = relativeLuminance(background);
  if (foregroundLuminance === null || backgroundLuminance === null) {
    return null;
  }

  const lighter = Math.max(foregroundLuminance, backgroundLuminance);
  const darker = Math.min(foregroundLuminance, backgroundLuminance);
  return (lighter + 0.05) / (darker + 0.05);
}

function resolveReadableTextColor(
  requestedColor: string | null,
  backgroundColor: string,
  palette: ViewerPalette
) {
  if (!paletteIsDark(palette)) {
    return requestedColor ?? "#000000";
  }

  if (!requestedColor) {
    return palette.text;
  }

  const requestedContrast = contrastRatio(requestedColor, backgroundColor);
  const paletteContrast = contrastRatio(palette.text, backgroundColor);
  if (requestedContrast !== null && requestedContrast >= 4.5) {
    return requestedColor;
  }
  if (paletteContrast !== null && (requestedContrast === null || paletteContrast > requestedContrast)) {
    return palette.text;
  }

  return requestedColor;
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
    .map((channel) => channel.toString(16).padStart(2, "0"))
    .join("")}`;
}

function invertHexLightness(color: string): string | null {
  const rgb = parseHexColor(color);
  if (!rgb) {
    return null;
  }

  const [hue, saturation, lightness] = rgbToHsl(rgb[0], rgb[1], rgb[2]);
  const [nextRed, nextGreen, nextBlue] = hslToRgb(hue, saturation, 1 - lightness);
  return rgbToHex(nextRed, nextGreen, nextBlue);
}

function escapeCssFontFamily(name: string) {
  return /[\s"'(),]/.test(name) ? `"${name.replace(/\\/g, "\\\\").replace(/"/g, '\\"')}"` : name;
}

function resolveSpreadsheetThemeFont(
  font: Record<string, unknown> | undefined,
  themePalette?: XlsxThemePalette
) {
  const scheme = typeof font?.scheme === "string" ? font.scheme : null;
  if (scheme === "major") {
    return themePalette?.majorLatinFont;
  }
  if (scheme === "minor") {
    return themePalette?.minorLatinFont;
  }
  return undefined;
}

function buildSpreadsheetFontFamily(
  font: Record<string, unknown> | undefined,
  themePalette?: XlsxThemePalette
) {
  const candidates = [
    typeof font?.name === "string" ? font.name.trim() : "",
    resolveSpreadsheetThemeFont(font, themePalette)?.trim() ?? ""
  ].filter((value, index, array) => value.length > 0 && array.indexOf(value) === index);

  if (candidates.length === 0) {
    return undefined;
  }

  const primaryFont = candidates[0]?.toLowerCase();
  if (primaryFont === "calibri") {
    return [
      "Calibri",
      "Carlito",
      "\"Aptos\"",
      "\"Segoe UI\"",
      "Tahoma",
      "Arial",
      "sans-serif"
    ].join(", ");
  }

  return [
    ...candidates.map(escapeCssFontFamily),
    "\"Segoe UI\"",
    "Tahoma",
    "Arial",
    "sans-serif"
  ].join(", ");
}

function resolveSpreadsheetLineHeight(fontSizePt?: number) {
  const sizePt = typeof fontSizePt === "number" && Number.isFinite(fontSizePt) ? fontSizePt : 11;
  if (sizePt <= 11) {
    return 1.2;
  }
  if (sizePt <= 14) {
    return 1.25;
  }
  return 1.3;
}

function buildGridlineShadow(color: string, options?: { bottom?: boolean; right?: boolean }) {
  const parts: string[] = [];
  if (options?.right !== false) {
    parts.push(`inset -1px 0 0 ${color}`);
  }
  if (options?.bottom !== false) {
    parts.push(`inset 0 -1px 0 ${color}`);
  }
  return parts.join(", ");
}

function buildCellStyle(
  style: Record<string, unknown> | null | undefined,
  palette: ViewerPalette,
  themePalette?: XlsxSheetData["themePalette"],
  options?: { showGridLines?: boolean }
): React.CSSProperties {
  const showGridLines = options?.showGridLines ?? true;
  const baseSurface = paletteIsDark(palette)
    ? resolveDarkModeSurface(themePalette, palette)
    : themePalette?.colorsByIndex[0] ?? SHEET_SURFACE;
  const gridlineShadow = showGridLines ? buildGridlineShadow(palette.border) : undefined;
  const css: React.CSSProperties = {
    backgroundColor: baseSurface,
    borderBottom: "none",
    borderRight: "none",
    boxShadow: gridlineShadow,
    color: resolveReadableTextColor(null, baseSurface, palette),
    fontFamily: buildSpreadsheetFontFamily({ scheme: "minor" }, themePalette) ?? "\"Segoe UI\", Tahoma, Arial, sans-serif",
    fontSize: "12px",
    lineHeight: String(resolveSpreadsheetLineHeight(11)),
    overflow: "hidden",
    padding: DEFAULT_CELL_PADDING,
    textOverflow: "clip",
    verticalAlign: "bottom",
    whiteSpace: "nowrap"
  };

  if (!style) {
    return css;
  }

  const fill = style.fill as Record<string, unknown> | undefined;
  let resolvedFillColor: string | null = null;
  let hasExplicitFill = false;
  if (fill) {
    const fillStyle = resolveWorkbookFillStyle(fill, themePalette);
    if (fillStyle.backgroundColor || fillStyle.backgroundImage) {
      hasExplicitFill = true;
      resolvedFillColor = fillStyle.backgroundColor;
      if (fillStyle.backgroundColor) {
        css.backgroundColor = fillStyle.backgroundColor;
      }
      if (fillStyle.backgroundImage) {
        css.backgroundImage = fillStyle.backgroundImage;
      }
    }
  }

  const font = style.font as Record<string, unknown> | undefined;
  let resolvedFontColor: string | null = null;
  let resolvedFontSizePt = 11;
  if (font) {
    if (font.bold) {
      css.fontWeight = "bold";
    }
    if (font.italic) {
      css.fontStyle = "italic";
    }
    if (font.underline && font.underline !== "none") {
      css.textDecoration = "underline";
    }
    if (font.strikethrough) {
      css.textDecoration = `${css.textDecoration ?? ""} line-through`.trim();
    }
    const fontColor = resolveWorkbookColor(font.color as Record<string, unknown> | undefined, themePalette);
    if (fontColor) {
      resolvedFontColor = fontColor;
      css.color = fontColor;
    }
    if (typeof font.size === "number" && font.size !== 11) {
      resolvedFontSizePt = font.size;
      css.fontSize = `${font.size}pt`;
    } else if (typeof font.size === "number") {
      resolvedFontSizePt = font.size;
    }

    const fontFamily = buildSpreadsheetFontFamily(font, themePalette);
    if (fontFamily) {
      css.fontFamily = fontFamily;
    }
  }

  css.lineHeight = String(resolveSpreadsheetLineHeight(resolvedFontSizePt));

  if (paletteIsDark(palette) && !css.backgroundImage) {
    const effectiveBackgroundColor = typeof css.backgroundColor === "string" ? css.backgroundColor : baseSurface;
    css.color = resolveReadableTextColor(resolvedFontColor, effectiveBackgroundColor, palette);
  } else if (!hasExplicitFill) {
    css.color = resolveReadableTextColor(resolvedFontColor, baseSurface, palette);
  }

  const alignment = style.alignment as Record<string, unknown> | undefined;
  if (alignment) {
    if (alignment.horizontal && alignment.horizontal !== "general") {
      css.textAlign = alignment.horizontal as React.CSSProperties["textAlign"];
    }
    if (alignment.vertical) {
      const verticalMap: Record<string, string> = {
        bottom: "bottom",
        center: "middle",
        top: "top"
      };
      const verticalValue = verticalMap[String(alignment.vertical)];
      if (verticalValue) {
        css.verticalAlign = verticalValue as React.CSSProperties["verticalAlign"];
      }
    }
    if (alignment.wrapText) {
      css.whiteSpace = "pre-wrap";
      css.wordBreak = "break-word";
    }
  }

  const border = style.border as Record<string, Record<string, unknown>> | undefined;
  if (border) {
    const hasRightBorder = Boolean(border.right?.style && border.right.style !== "none");
    const hasBottomBorder = Boolean(border.bottom?.style && border.bottom.style !== "none");
    if (showGridLines && (hasRightBorder || hasBottomBorder)) {
      css.boxShadow = buildGridlineShadow(palette.border, {
        bottom: !hasBottomBorder,
        right: !hasRightBorder
      }) || undefined;
    }
    if (border.top?.style && border.top.style !== "none") {
      css.borderTop = mapBorder(border.top as { style: string; color?: Record<string, unknown> }, themePalette);
    }
    if (border.right?.style && border.right.style !== "none") {
      css.borderRight = mapBorder(border.right as { style: string; color?: Record<string, unknown> }, themePalette);
    }
    if (border.bottom?.style && border.bottom.style !== "none") {
      css.borderBottom = mapBorder(border.bottom as { style: string; color?: Record<string, unknown> }, themePalette);
    }
    if (border.left?.style && border.left.style !== "none") {
      css.borderLeft = mapBorder(border.left as { style: string; color?: Record<string, unknown> }, themePalette);
    }
  }

  return css;
}

let textMeasureCanvas: HTMLCanvasElement | null = null;

function getHorizontalPadding(padding: React.CSSProperties["padding"]) {
  if (typeof padding !== "string") {
    return 8;
  }

  const values = padding
    .trim()
    .split(/\s+/)
    .map((value) => Number.parseFloat(value))
    .filter((value) => Number.isFinite(value));

  if (values.length === 1) {
    return values[0] * 2;
  }
  if (values.length === 2) {
    return values[1] * 2;
  }
  if (values.length === 3) {
    return values[1] * 2;
  }
  if (values.length >= 4) {
    return values[1] + values[3];
  }

  return 8;
}

function buildCanvasFont(style: React.CSSProperties) {
  return resolveCanvasFont(style, 12);
}

function measureTextWidth(value: string, style: React.CSSProperties) {
  if (!value) {
    return 0;
  }

  if (typeof document === "undefined") {
    return value.length * 7;
  }

  textMeasureCanvas ??= document.createElement("canvas");
  const context = textMeasureCanvas.getContext("2d", { alpha: false });
  if (!context) {
    return value.length * 7;
  }

  context.font = buildCanvasFont(style);
  return measureCanvasTextWidth(context, value);
}

function measureWrappedTextHeight(value: string, style: React.CSSProperties, widthPx: number) {
  if (!value) {
    return 0;
  }

  const availableWidth = Math.max(1, widthPx - getHorizontalPadding(style.padding));
  const lineHeight = resolveCanvasLineHeight(style, 12);
  const padding = resolveCanvasPadding(style.padding);
  const fallbackLineCount = value.replace(/\r\n?/g, "\n").split("\n").length;

  if (typeof document === "undefined") {
    return (fallbackLineCount * lineHeight) + padding.top + padding.bottom;
  }

  textMeasureCanvas ??= document.createElement("canvas");
  const context = textMeasureCanvas.getContext("2d", { alpha: false });
  if (!context) {
    return (fallbackLineCount * lineHeight) + padding.top + padding.bottom;
  }

  context.font = buildCanvasFont(style);
  const wrappedLines = wrapCanvasText(context, value, availableWidth);
  return (wrappedLines.length * lineHeight) + padding.top + padding.bottom;
}

function canCellTextOverflow(data: CellRenderData) {
  if (!data.value || data.isMergedSecondary || data.shrinkToFit || data.style.whiteSpace === "pre-wrap") {
    return false;
  }

  const textAlign = data.style.textAlign;
  if (textAlign && textAlign !== "left" && textAlign !== "start" && textAlign !== "center") {
    return false;
  }

  return true;
}

function canReceiveOverflowText(data: CellRenderData) {
  return !data.isMergedSecondary && !data.colSpan && data.value.length === 0;
}

function getCellDisplayValue(worksheet: Worksheet, row: number, col: number, activeSheet?: XlsxSheetData | null): string {
  const formula = worksheet.getFormulaAt(row, col);
  const cachedFormulaValue = formula ? activeSheet?.cachedFormulaValues?.[cellAddressToA1({ row, col })] : undefined;
  const formatted = worksheet.getFormattedValueAt(row, col);
  if (formatted && !(formula && cachedFormulaValue !== undefined && formatted.startsWith("#"))) {
    return decodeHtmlEntities(formatted);
  }

  const cellValue = worksheet.getCalculatedValueAt(row, col);
  if (formula && cachedFormulaValue !== undefined && cellValue.is_error) {
    return cachedFormulaValue;
  }
  if (cellValue.is_error) {
    return cellValue.asError() ?? "";
  }
  if (cellValue.is_empty) {
    return "";
  }

  return cellValue.toString();
}

function getCellNumericValue(worksheet: Worksheet, row: number, col: number) {
  const cellValue = worksheet.getCalculatedValueAt(row, col);
  if (cellValue.is_number) {
    return cellValue.asNumber() ?? null;
  }
  if (cellValue.is_boolean) {
    return cellValue.asBoolean() ? 1 : 0;
  }
  if (cellValue.is_empty || cellValue.is_error) {
    return null;
  }

  const parsedValue = Number(cellValue.asText() ?? cellValue.toString());
  return Number.isFinite(parsedValue) ? parsedValue : null;
}

function getCellBooleanValue(worksheet: Worksheet, row: number, col: number) {
  const cellValue = worksheet.getCalculatedValueAt(row, col);
  if (cellValue.is_boolean) {
    return cellValue.asBoolean() ?? null;
  }
  if (cellValue.is_number) {
    const numeric = cellValue.asNumber();
    return numeric == null ? null : numeric !== 0;
  }
  const text = (cellValue.asText() ?? cellValue.toString()).trim().toLowerCase();
  if (text === "true") {
    return true;
  }
  if (text === "false") {
    return false;
  }
  return null;
}

function parseA1RangeReference(reference: string): XlsxCellRange | null {
  const [startRef, endRef = startRef] = reference.split(":");
  const start = parseA1CellReference(startRef ?? "");
  const end = parseA1CellReference(endRef ?? "");
  return start && end ? { end, start } : null;
}

function unescapeSheetNameReference(value: string) {
  const trimmed = value.trim();
  return trimmed.startsWith("'") && trimmed.endsWith("'")
    ? trimmed.slice(1, -1).replace(/''/g, "'")
    : trimmed;
}

function splitSheetFormulaReference(reference: string) {
  let inQuotes = false;
  for (let index = 0; index < reference.length; index += 1) {
    const character = reference[index];
    if (character === "'") {
      if (inQuotes && reference[index + 1] === "'") {
        index += 1;
        continue;
      }
      inQuotes = !inQuotes;
      continue;
    }
    if (character === "!" && !inQuotes) {
      return {
        range: reference.slice(index + 1),
        sheetName: unescapeSheetNameReference(reference.slice(0, index))
      };
    }
  }
  return null;
}

function resolveChartFormulaRange(
  formula: string,
  fallbackSheetIndex: number,
  sheets: XlsxSheetData[],
  workbook: Workbook | null
) {
  let reference = formula.trim().replace(/^=/, "");
  if (!reference) {
    return null;
  }

  if (!reference.includes("!") && workbook) {
    try {
      const namedRange = workbook.getNamedRange(reference);
      if (typeof namedRange === "string" && namedRange.length > 0 && namedRange !== reference) {
        reference = namedRange;
      }
    } catch {
      // Ignore missing named ranges and fall back to direct parsing.
    }
  }

  const split = splitSheetFormulaReference(reference);
  const range = parseA1RangeReference((split?.range ?? reference).replace(/\$/g, ""));
  if (!range) {
    return null;
  }

  let workbookSheetIndex = fallbackSheetIndex;
  if (split?.sheetName) {
    const workbookResolvedIndex = workbook?.sheetIndex(split.sheetName);
    if (typeof workbookResolvedIndex === "number" && Number.isInteger(workbookResolvedIndex) && workbookResolvedIndex >= 0) {
      workbookSheetIndex = workbookResolvedIndex;
    } else {
      const matchingSheet = sheets.find((sheet) => sheet.name === split.sheetName);
      if (matchingSheet) {
        workbookSheetIndex = matchingSheet.workbookSheetIndex;
      }
    }
  }

  return {
    range: normalizeRange(range),
    workbookSheetIndex
  };
}

function resolveRangePointCell(range: XlsxCellRange, pointIndex: number): XlsxCellRange | null {
  const normalized = normalizeRange(range);
  const rowCount = normalized.end.row - normalized.start.row + 1;
  const colCount = normalized.end.col - normalized.start.col + 1;
  const cellCount = rowCount * colCount;
  if (pointIndex < 0 || pointIndex >= cellCount) {
    return null;
  }

  const rowOffset = Math.floor(pointIndex / colCount);
  const colOffset = pointIndex % colCount;
  const cell = {
    col: normalized.start.col + colOffset,
    row: normalized.start.row + rowOffset
  };
  return { start: cell, end: cell };
}

function collectChartRangeHighlights(
  chart: XlsxChart | null,
  selectedElement: XlsxChartElementSelection | null,
  sheets: XlsxSheetData[],
  workbook: Workbook | null,
  isDark: boolean
) {
  if (!chart) {
    return [] as ChartRangeHighlight[];
  }

  const highlights: ChartRangeHighlight[] = [];
  const seenRanges = new Set<string>();
  const selectedSeriesIndex = selectedElement?.chartId === chart.id && selectedElement.kind !== "chart"
    ? selectedElement.seriesIndex
    : null;
  const seriesEntries = chart.series
    .map((series, seriesIndex) => ({ series, seriesIndex }))
    .filter(({ series, seriesIndex }) => (
      series.hidden !== true
      && (selectedSeriesIndex == null || seriesIndex === selectedSeriesIndex)
    ));
  const references = seriesEntries.flatMap(({ series, seriesIndex }) => [
    {
      formula: typeof series.raw?.name === "string" ? series.raw.name : undefined,
      role: "name" as const,
      seriesIndex
    },
    {
      formula: series.categoriesRef?.formula,
      role: "category" as const,
      seriesIndex
    },
    {
      formula: series.valuesRef?.formula,
      role: "value" as const,
      seriesIndex
    },
    {
      formula: series.bubbleSizeRef?.formula,
      role: "bubble" as const,
      seriesIndex
    }
  ]);

  references.forEach((reference) => {
    if (!reference.formula) {
      return;
    }

    const resolved = resolveChartFormulaRange(reference.formula, chart.workbookSheetIndex, sheets, workbook);
    if (!resolved) {
      return;
    }

    const key = [
      resolved.workbookSheetIndex,
      resolved.range.start.row,
      resolved.range.start.col,
      resolved.range.end.row,
      resolved.range.end.col
    ].join(":");
    if (seenRanges.has(key)) {
      return;
    }

    seenRanges.add(key);
    const strokeColor = CHART_SOURCE_HIGHLIGHT_COLOR_BY_ROLE[reference.role];
    highlights.push({
      fillColor: applyAlphaToColor(strokeColor, isDark ? 0.24 : 0.14),
      range: resolved.range,
      strokeColor,
      workbookSheetIndex: resolved.workbookSheetIndex
    });
  });

  if (selectedElement?.chartId === chart.id && selectedElement.kind === "point") {
    references.forEach((reference) => {
      if (!reference.formula || reference.role === "name") {
        return;
      }

      const resolved = resolveChartFormulaRange(reference.formula, chart.workbookSheetIndex, sheets, workbook);
      if (!resolved) {
        return;
      }

      const range = resolveRangePointCell(resolved.range, selectedElement.pointIndex);
      if (!range) {
        return;
      }

      const strokeColor = CHART_SOURCE_HIGHLIGHT_COLOR_BY_ROLE.point;
      highlights.push({
        fillColor: applyAlphaToColor(strokeColor, isDark ? 0.34 : 0.22),
        range,
        strokeColor,
        workbookSheetIndex: resolved.workbookSheetIndex
      });
    });
  }

  return highlights;
}

function clampSparklineValue(value: number, min: number, max: number) {
  if (max <= min) {
    return 0.5;
  }
  return Math.max(0, Math.min(1, (value - min) / (max - min)));
}

function renderSparkline(
  sparkline: XlsxSheetData["sparklines"][number],
  values: Array<number | null>,
  palette: ViewerPalette
) {
  const width = 100;
  const height = 18;
  const innerWidth = width - 2;
  const innerHeight = height - 4;
  const points = values.map((value, index) => ({ index, value })).filter((entry): entry is { index: number; value: number } => typeof entry.value === "number" && Number.isFinite(entry.value));
  if (points.length === 0) {
    return null;
  }

  const negativeColor = sparkline.negativeColor ?? "#c2410c";
  const seriesColor = sparkline.color ?? "#2563eb";
  const markerColor = sparkline.markerColor ?? seriesColor;

  if (sparkline.type === "winLoss") {
    const normalizedValues = points.map((entry) => ({ ...entry, value: entry.value >= 0 ? 1 : -1 }));
    const segmentWidth = Math.max(4, innerWidth / Math.max(normalizedValues.length * 1.9, 1));
    const gap = normalizedValues.length > 1 ? (innerWidth - segmentWidth * normalizedValues.length) / (normalizedValues.length - 1) : 0;
    const positiveY = 4.5;
    const negativeY = height - 4.5;

    return (
      <svg aria-hidden="true" height={height} style={{ display: "block", overflow: "visible", width: "100%" }} viewBox={`0 0 ${width} ${height}`} width="100%">
        {normalizedValues.map((entry, index) => {
          const left = 1 + index * (segmentWidth + Math.max(0, gap));
          const y = entry.value >= 0 ? positiveY : negativeY;
          return (
            <line
              key={`spark-winloss-${index}`}
              stroke={seriesColor}
              strokeLinecap="round"
              strokeWidth={1.8}
              x1={left}
              x2={left + segmentWidth}
              y1={y}
              y2={y}
            />
          );
        })}
      </svg>
    );
  }

  if (sparkline.type === "column") {
    const normalizedValues = points;
    const minValue = Math.min(0, ...normalizedValues.map((entry) => entry.value));
    const maxValue = Math.max(0, ...normalizedValues.map((entry) => entry.value));
    const zeroY = 2 + innerHeight - clampSparklineValue(0, minValue, maxValue) * innerHeight;
    const barWidth = Math.max(2, innerWidth / Math.max(normalizedValues.length * 1.8, 1));
    const gap = normalizedValues.length > 1 ? (innerWidth - barWidth * normalizedValues.length) / (normalizedValues.length - 1) : 0;

    return (
      <svg aria-hidden="true" height={height} style={{ display: "block", overflow: "visible", width: "100%" }} viewBox={`0 0 ${width} ${height}`} width="100%">
        <line stroke={palette.border} strokeWidth={1} x1={1} x2={width - 1} y1={zeroY} y2={zeroY} />
        {normalizedValues.map((entry, index) => {
          const left = 1 + index * (barWidth + Math.max(0, gap));
          const y = 2 + innerHeight - clampSparklineValue(entry.value, minValue, maxValue) * innerHeight;
          const top = Math.min(y, zeroY);
          const barHeight = Math.max(1, Math.abs(y - zeroY));
          const fill = entry.value < 0 ? negativeColor : seriesColor;
          return (
            <rect
              fill={fill}
              height={barHeight}
              key={`spark-bar-${index}`}
              rx={sparkline.type === "column" ? 0 : 0.8}
              ry={sparkline.type === "column" ? 0 : 0.8}
              width={barWidth}
              x={left}
              y={top}
            />
          );
        })}
      </svg>
    );
  }

  const minValue = Math.min(...points.map((entry) => entry.value));
  const maxValue = Math.max(...points.map((entry) => entry.value));
  const xStep = points.length > 1 ? innerWidth / (points.length - 1) : 0;
  const path = points.map((entry, index) => {
    const x = 1 + index * xStep;
    const y = 2 + innerHeight - clampSparklineValue(entry.value, minValue, maxValue) * innerHeight;
    return `${index === 0 ? "M" : "L"} ${x.toFixed(2)} ${y.toFixed(2)}`;
  }).join(" ");

  const highValue = Math.max(...points.map((entry) => entry.value));
  const lowValue = Math.min(...points.map((entry) => entry.value));

  return (
    <svg aria-hidden="true" height={height} style={{ display: "block", overflow: "visible", width: "100%" }} viewBox={`0 0 ${width} ${height}`} width="100%">
      <path d={path} fill="none" stroke={seriesColor} strokeLinecap="round" strokeLinejoin="round" strokeWidth={1.6} />
      {sparkline.markers ? points.map((entry, index) => {
        const x = 1 + index * xStep;
        const y = 2 + innerHeight - clampSparklineValue(entry.value, minValue, maxValue) * innerHeight;
        let fill = markerColor;
        if (entry.value === highValue && sparkline.highColor) {
          fill = sparkline.highColor;
        } else if (entry.value === lowValue && sparkline.lowColor) {
          fill = sparkline.lowColor;
        } else if (index === 0 && sparkline.firstColor) {
          fill = sparkline.firstColor;
        } else if (index === points.length - 1 && sparkline.lastColor) {
          fill = sparkline.lastColor;
        } else if (entry.value < 0 && sparkline.negative && sparkline.negativeColor) {
          fill = sparkline.negativeColor;
        }
        return <circle cx={x} cy={y} fill={fill} key={`spark-point-${index}`} r={1.75} />;
      }) : null}
    </svg>
  );
}

function asNonNegativeInteger(value: unknown) {
  return typeof value === "number" && Number.isInteger(value) && value >= 0 ? value : null;
}

function asRecord(value: unknown) {
  return value && typeof value === "object" && !Array.isArray(value) ? (value as Record<string, unknown>) : null;
}

function normalizeBatchedHyperlink(value: unknown): BatchedCellData["hyperlink"] {
  const hyperlink = asRecord(value);
  if (!hyperlink) {
    return null;
  }

  return {
    location: typeof hyperlink.location === "string" ? hyperlink.location : undefined,
    target: typeof hyperlink.target === "string" ? hyperlink.target : undefined,
    tooltip: typeof hyperlink.tooltip === "string" ? hyperlink.tooltip : undefined
  };
}

function normalizeBatchedMergeSpan(value: unknown): BatchedCellData["mergeSpan"] {
  const mergeSpan = asRecord(value);
  if (!mergeSpan) {
    return null;
  }

  const rowSpan = asNonNegativeInteger(mergeSpan.rowSpan);
  const colSpan = asNonNegativeInteger(mergeSpan.colSpan);
  if (rowSpan === null && colSpan === null) {
    return null;
  }

  return {
    colSpan: colSpan ?? undefined,
    rowSpan: rowSpan ?? undefined
  };
}

function normalizeBatchedCellValue(
  rawValue: unknown,
  formula: string | null,
  row: number,
  col: number,
  activeSheet?: XlsxSheetData | null
) {
  const value =
    typeof rawValue === "string"
      ? decodeHtmlEntities(rawValue)
      : rawValue === null || rawValue === undefined
        ? ""
        : String(rawValue);
  const cachedFormulaValue = formula ? activeSheet?.cachedFormulaValues?.[cellAddressToA1({ row, col })] : undefined;

  if (formula && cachedFormulaValue !== undefined && value.startsWith("#")) {
    return cachedFormulaValue;
  }

  return value;
}

function buildWorksheetBatchWindow(
  rows: unknown[] | null,
  activeSheet: XlsxSheetData | null,
  startRow: number,
  endRow: number
): WorksheetBatchWindow {
  const cells = new Map<string, BatchedCellData>();

  if (Array.isArray(rows)) {
    for (const rowEntry of rows as WorksheetBatchRowEntry[]) {
      const row = asNonNegativeInteger(rowEntry.index);
      if (row === null || !Array.isArray(rowEntry.cells)) {
        continue;
      }

      for (const cellEntry of rowEntry.cells as WorksheetBatchCellEntry[]) {
        const col = asNonNegativeInteger(cellEntry.col);
        if (col === null) {
          continue;
        }

        const formula = typeof cellEntry.formula === "string" ? cellEntry.formula : null;
        cells.set(`${row}:${col}`, {
          formula,
          hyperlink: normalizeBatchedHyperlink(cellEntry.hyperlink),
          isMergedSecondary: cellEntry.isMergedSecondary === true,
          mergeSpan: normalizeBatchedMergeSpan(cellEntry.mergeSpan),
          style: asRecord(cellEntry.style),
          value: normalizeBatchedCellValue(cellEntry.value, formula, row, col, activeSheet)
        });
      }
    }
  }

  return {
    cells,
    endRow,
    startRow
  };
}

function getTableAtCell(tables: XlsxTable[], row: number, col: number) {
  return tables.find(
    (table) =>
      row >= table.start.row &&
      row <= table.end.row &&
      col >= table.start.col &&
      col <= table.end.col
  ) ?? null;
}

function mergeResolvedCellStyle(
  base: Record<string, unknown> | null | undefined,
  overlay: Record<string, unknown> | null | undefined,
  options?: { replaceXfSubtrees?: boolean }
) {
  if (!base && !overlay) {
    return null;
  }

  const nextStyle: Record<string, unknown> = {
    ...(base ?? {}),
    ...(overlay ?? {})
  };

  const baseAlignment = base?.alignment as Record<string, unknown> | undefined;
  const overlayAlignment = overlay?.alignment as Record<string, unknown> | undefined;
  if (baseAlignment || overlayAlignment) {
    nextStyle.alignment = {
      ...(baseAlignment ?? {}),
      ...(overlayAlignment ?? {})
    };
  }

  const baseBorder = base?.border as Record<string, Record<string, unknown>> | undefined;
  const overlayBorder = overlay?.border as Record<string, Record<string, unknown>> | undefined;
  if (baseBorder || overlayBorder) {
    nextStyle.border = options?.replaceXfSubtrees && overlayBorder
      ? overlayBorder
      : {
          ...(baseBorder ?? {}),
          ...(overlayBorder ?? {})
        };
  }

  const baseFill = base?.fill as Record<string, unknown> | undefined;
  const overlayFill = overlay?.fill as Record<string, unknown> | undefined;
  if (baseFill || overlayFill) {
    nextStyle.fill = options?.replaceXfSubtrees && overlayFill
      ? overlayFill
      : (overlayFill ?? baseFill);
  }

  const baseFont = base?.font as Record<string, unknown> | undefined;
  const overlayFont = overlay?.font as Record<string, unknown> | undefined;
  if (baseFont || overlayFont) {
    nextStyle.font = options?.replaceXfSubtrees && overlayFont
      ? overlayFont
      : {
          ...(baseFont ?? {}),
          ...(overlayFont ?? {})
        };
  }

  return nextStyle;
}

function normalizeTableStyleEdges(
  style: Record<string, unknown> | null | undefined,
  table: XlsxTable,
  row: number,
  col: number
) {
  if (!style) {
    return null;
  }

  const border = style.border as Record<string, Record<string, unknown>> | undefined;
  if (!border) {
    return style;
  }

  const nextBorder: Record<string, Record<string, unknown>> = { ...border };
  if (nextBorder.horizontal) {
    if (row < table.end.row) {
      nextBorder.bottom = nextBorder.bottom ?? nextBorder.horizontal;
    }
    delete nextBorder.horizontal;
  }
  if (nextBorder.vertical) {
    if (col < table.end.col) {
      nextBorder.right = nextBorder.right ?? nextBorder.vertical;
    }
    delete nextBorder.vertical;
  }

  return {
    ...style,
    border: nextBorder
  };
}

function resolveTableCellStyle(
  table: XlsxTable | null,
  row: number,
  col: number,
  activeSheet: XlsxSheetData | null | undefined
) {
  if (!table || !activeSheet) {
    return null;
  }

  const styleName = table.styleInfo?.name;
  if (!styleName) {
    return null;
  }

  const tableStyle = activeSheet.tableStyleByName[styleName];
  if (!tableStyle) {
    return null;
  }

  let resolved: Record<string, unknown> | null = null;
  const applyElement = (elementType: string, enabled = true) => {
    if (!enabled) {
      return;
    }

    const nextStyle = normalizeTableStyleEdges(tableStyle[elementType] ?? null, table, row, col);
    if (!nextStyle) {
      return;
    }
    resolved = mergeResolvedCellStyle(resolved, nextStyle);
  };

  applyElement("wholeTable");
  applyElement("firstColumn", Boolean(table.styleInfo?.showFirstColumn) && col === table.start.col);
  applyElement("lastColumn", Boolean(table.styleInfo?.showLastColumn) && col === table.end.col);

  const headerRowCount = Math.max(table.headerRowCount, 1);
  const isHeaderRow = row >= table.start.row && row < table.start.row + headerRowCount;
  applyElement("headerRow", isHeaderRow);

  if (isHeaderRow && table.headerRowCellStyle) {
    const namedHeaderStyle = activeSheet.namedCellStyleByName[table.headerRowCellStyle];
    if (namedHeaderStyle) {
      resolved = mergeResolvedCellStyle(resolved, namedHeaderStyle);
    }
  }

  if (table.totalsRowShown) {
    const totalsRowCount = Math.max(table.totalsRowCount, 1);
    applyElement("totalRow", row > table.end.row - totalsRowCount);
  }

  return resolved;
}

function getTableHeaderColumn(table: XlsxTable | null, row: number, col: number): XlsxTableColumn | null {
  if (!table || row !== table.start.row) {
    return null;
  }

  const index = col - table.start.col;
  return table.columns[index] ?? null;
}

function DefaultTableHeaderMenu({
  direction,
  sortAscending,
  sortDescending
}: XlsxTableHeaderMenuRenderProps) {
  return (
    <div
      style={{
        backgroundColor: "var(--xlsx-menu-surface)",
        border: "1px solid var(--xlsx-menu-border)",
        borderRadius: 10,
        boxShadow: "0 10px 30px rgba(0, 0, 0, 0.16)",
        display: "grid",
        gap: 4,
        minWidth: 160,
        padding: 6
      }}
    >
      <button
        onClick={sortAscending}
        style={{
          background: direction === "ascending" ? "var(--xlsx-menu-active)" : "transparent",
          border: "none",
          borderRadius: 8,
          color: "inherit",
          cursor: "pointer",
          fontSize: 12,
          padding: "8px 10px",
          textAlign: "left"
        }}
        type="button"
      >
        Sort A to Z
      </button>
      <button
        onClick={sortDescending}
        style={{
          background: direction === "descending" ? "var(--xlsx-menu-active)" : "transparent",
          border: "none",
          borderRadius: 8,
          color: "inherit",
          cursor: "pointer",
          fontSize: 12,
          padding: "8px 10px",
          textAlign: "left"
        }}
        type="button"
      >
        Sort Z to A
      </button>
    </div>
  );
}

function resolveTableHeaderTriggerStyle(palette: ViewerPalette, zoomFactor: number): React.CSSProperties {
  return {
    alignItems: "center",
    background: "transparent",
    border: "none",
    color: palette.mutedText,
    cursor: "pointer",
    display: "inline-flex",
    fontSize: 10 * zoomFactor,
    height: 16 * zoomFactor,
    justifyContent: "center",
    padding: 0,
    position: "absolute",
    right: 4 * zoomFactor,
    top: 3 * zoomFactor,
    width: 16 * zoomFactor,
    zIndex: 6
  };
}

function SegmentedControl({
  items,
  onValueChange,
  palette,
  value
}: {
  items: Array<{ id: string; label: string; title?: string; visibility?: "hidden" | "veryHidden" | "visible" }>;
  onValueChange: (value: string) => void;
  palette: ViewerPalette;
  value: string;
}) {
  return (
    <div
      aria-label="Workbook sheets"
      role="tablist"
      style={{
        alignItems: "center",
        backgroundColor: palette.sheetInactiveSurface,
        border: `1px solid ${palette.strongBorder}`,
        borderRadius: 10,
        display: "inline-flex",
        gap: 2,
        minHeight: 36,
        padding: 2
      }}
    >
      {items.map((item) => {
        const selected = item.id === value;
        const isHidden = item.visibility === "hidden" || item.visibility === "veryHidden";
        return (
          <button
            aria-selected={selected}
            key={item.id}
            onClick={() => onValueChange(item.id)}
            role="tab"
            style={{
              backgroundColor: selected ? palette.sheetActiveSurface : "transparent",
              border: "none",
              borderRadius: 8,
              boxShadow: selected ? palette.shadow : "none",
              color: selected ? palette.sheetActiveText : palette.sheetInactiveText,
              cursor: "pointer",
              fontSize: 12,
              fontWeight: selected ? 600 : 500,
              opacity: isHidden && !selected ? 0.68 : 1,
              padding: "7px 12px",
              transition: "background-color 120ms ease, color 120ms ease, box-shadow 120ms ease",
              whiteSpace: "nowrap"
            }}
            title={item.title}
            type="button"
          >
            {item.label}
          </button>
        );
      })}
    </div>
  );
}

function DefaultToolbar({ controller, palette }: { controller: XlsxViewerController; palette: ViewerPalette }) {
  const {
    activeTabIndex,
    canDownload,
    canZoomIn,
    canZoomOut,
    defaultZoomScale,
    displayFileName,
    download,
    resetZoom,
    setActiveTabIndex,
    tabs,
    zoomIn,
    zoomOut,
    zoomScale
  } = controller;

  return (
    <>
      <div
        style={{
          alignItems: "center",
          backgroundColor: palette.toolbarSurface,
          borderBottom: `1px solid ${palette.border}`,
          color: palette.text,
          display: "flex",
          gap: 12,
          justifyContent: "space-between",
          minHeight: 48,
          padding: "0 16px"
        }}
      >
        <div style={{ minWidth: 0 }}>
          <div
            style={{
              color: palette.text,
              fontSize: 14,
              fontWeight: 600,
              overflow: "hidden",
              textOverflow: "ellipsis",
              whiteSpace: "nowrap"
            }}
          >
            {displayFileName}
          </div>
        </div>
        <div style={{ alignItems: "center", display: "flex", gap: 8 }}>
          <div
            style={{
              alignItems: "center",
              background: palette.buttonSurface,
              border: `1px solid ${palette.strongBorder}`,
              borderRadius: 8,
              display: "flex",
              overflow: "hidden"
            }}
          >
            <button
              disabled={!canZoomOut}
              onClick={zoomOut}
              style={{
                background: "transparent",
                border: "none",
                color: canZoomOut ? palette.buttonText : palette.mutedText,
                cursor: canZoomOut ? "pointer" : "default",
                fontSize: 14,
                fontWeight: 600,
                padding: "6px 10px"
              }}
              type="button"
            >
              -
            </button>
            <button
              onClick={resetZoom}
              style={{
                background: Math.round(zoomScale) === Math.round(defaultZoomScale) ? "transparent" : palette.subtleSurface,
                border: "none",
                borderLeft: `1px solid ${palette.strongBorder}`,
                borderRight: `1px solid ${palette.strongBorder}`,
                color: palette.buttonText,
                cursor: "pointer",
                fontSize: 12,
                fontWeight: 600,
                minWidth: 60,
                padding: "6px 10px"
              }}
              type="button"
            >
              {formatZoomScale(zoomScale)}
            </button>
            <button
              disabled={!canZoomIn}
              onClick={zoomIn}
              style={{
                background: "transparent",
                border: "none",
                color: canZoomIn ? palette.buttonText : palette.mutedText,
                cursor: canZoomIn ? "pointer" : "default",
                fontSize: 14,
                fontWeight: 600,
                padding: "6px 10px"
              }}
              type="button"
            >
              +
            </button>
          </div>
          {canDownload ? (
            <button
              aria-label="Download workbook"
              onClick={download}
              style={{
                alignItems: "center",
                background: palette.buttonSurface,
                border: `1px solid ${palette.strongBorder}`,
                borderRadius: 8,
                color: palette.buttonText,
                cursor: "pointer",
                display: "inline-flex",
                fontSize: 16,
                fontWeight: 500,
                height: 32,
                justifyContent: "center",
                padding: 0,
                width: 32
              }}
              title="Download workbook"
              type="button"
            >
              ↓
            </button>
          ) : null}
        </div>
      </div>
      {tabs.length > 1 ? (
        <div
          style={{
            backgroundColor: palette.subtleSurface,
            borderBottom: `1px solid ${palette.border}`,
            overflowX: "auto",
            padding: "8px 12px"
          }}
        >
          <SegmentedControl
            items={tabs.map((tab, index) => ({
              id: String(index),
              label: tab.name,
              title:
                tab.visibility === "hidden"
                  ? `${tab.name} (hidden)`
                  : tab.visibility === "veryHidden"
                    ? `${tab.name} (very hidden)`
                    : tab.name,
              visibility: tab.visibility
            }))}
            onValueChange={(value) => setActiveTabIndex(Number(value))}
            palette={palette}
            value={String(activeTabIndex)}
          />
        </div>
      ) : null}
    </>
  );
}

function resolveToolbar(
  toolbar: XlsxViewerProps["toolbar"],
  showDefaultToolbar: boolean,
  controller: XlsxViewerController,
  palette: ViewerPalette
) {
  if (typeof toolbar === "function") {
    return toolbar(controller);
  }

  if (toolbar !== undefined) {
    return toolbar;
  }

  if (!showDefaultToolbar) {
    return null;
  }

  return <DefaultToolbar controller={controller} palette={palette} />;
}

function renderError(errorState: XlsxViewerProps["errorState"], error: Error, palette: ViewerPalette) {
  if (typeof errorState === "function") {
    return errorState(error);
  }
  if (errorState !== undefined) {
    return errorState;
  }

  return (
    <div
      style={{
        alignItems: "center",
        color: palette.danger,
        display: "flex",
        fontSize: 14,
        height: "100%",
        justifyContent: "center",
        padding: 16,
        textAlign: "center"
      }}
    >
      {error.message}
    </div>
  );
}

function renderLoading(
  loadingComponent: XlsxViewerProps["loadingComponent"],
  loadingState: XlsxViewerProps["loadingState"],
  palette: ViewerPalette
) {
  if (loadingComponent !== undefined) {
    return loadingComponent;
  }

  if (loadingState !== undefined) {
    return loadingState;
  }

  return (
    <div
      style={{
        alignItems: "center",
        color: palette.mutedText,
        display: "flex",
        fontSize: 14,
        height: "100%",
        justifyContent: "center"
      }}
    >
      Loading workbook...
    </div>
  );
}

function renderEmpty(emptyState: XlsxViewerProps["emptyState"], palette: ViewerPalette) {
  if (emptyState !== undefined) {
    return emptyState;
  }

  return (
    <div
      style={{
        alignItems: "center",
        color: palette.mutedText,
        display: "flex",
        fontSize: 14,
        height: "100%",
        justifyContent: "center",
        padding: 16,
        textAlign: "center"
      }}
    >
      No workbook loaded.
    </div>
  );
}

function formatBinaryBytes(bytes: number) {
  if (!Number.isFinite(bytes) || bytes <= 0) {
    return "0 B";
  }

  const units = ["B", "KB", "MB", "GB"];
  let value = bytes;
  let unitIndex = 0;
  while (value >= 1024 && unitIndex < units.length - 1) {
    value /= 1024;
    unitIndex += 1;
  }

  return `${value >= 10 || unitIndex === 0 ? value.toFixed(0) : value.toFixed(1)} ${units[unitIndex]}`;
}

function renderFileTooLarge(
  fileTooLargeState: XlsxViewerProps["fileTooLargeState"],
  renderProps: {
    displayFileName: string;
    fileSizeBytes: number;
    maxFileSizeBytes: number;
  },
  palette: ViewerPalette
) {
  const defaultNode = (
    <div
      style={{
        alignItems: "center",
        color: palette.mutedText,
        display: "flex",
        flexDirection: "column",
        fontSize: 14,
        gap: 8,
        height: "100%",
        justifyContent: "center",
        padding: 24,
        textAlign: "center"
      }}
    >
      <div style={{ color: palette.text, fontWeight: 600 }}>
        {renderProps.displayFileName}
      </div>
      <div>
        File size {formatBinaryBytes(renderProps.fileSizeBytes)} exceeds the configured limit of{" "}
        {formatBinaryBytes(renderProps.maxFileSizeBytes)}.
      </div>
    </div>
  );

  if (typeof fileTooLargeState === "function") {
    return fileTooLargeState({
      defaultNode,
      displayFileName: renderProps.displayFileName,
      fileSizeBytes: renderProps.fileSizeBytes,
      maxFileSizeBytes: renderProps.maxFileSizeBytes
    });
  }

  if (fileTooLargeState !== undefined) {
    return fileTooLargeState;
  }

  return defaultNode;
}

function renderCustomFileTooLarge(
  fileTooLargeState: XlsxViewerProps["fileTooLargeState"],
  renderProps: {
    displayFileName: string;
    fileSizeBytes: number;
    maxFileSizeBytes: number;
  },
  palette: ViewerPalette
) {
  if (fileTooLargeState === undefined) {
    return undefined;
  }

  return renderFileTooLarge(fileTooLargeState, renderProps, palette);
}

function renderDefaultChartLoadingCard(rect: XlsxImageRect) {
  const bars = [18, 32, 24];
  const barWidth = Math.max(8, Math.min(12, Math.round(rect.width * 0.018)));
  const barGap = 8;

  return (
    <div
      style={{
        backgroundColor: "#ffffff",
        border: "1px solid #e5e7eb",
        borderRadius: 10,
        boxSizing: "border-box",
        display: "flex",
        height: "100%",
        justifyContent: "center",
        padding: 12,
        width: "100%"
      }}
    >
      <div
        style={{
          alignItems: "flex-end",
          display: "flex",
          gap: barGap,
          justifyContent: "center"
        }}
      >
        {bars.map((heightPx, index) => (
        <div
          key={index}
          style={{
            backgroundColor: "#e5e7eb",
            borderRadius: 999,
            height: heightPx,
            width: barWidth
          }}
        />
        ))}
      </div>
    </div>
  );
}

function renderChartLoadingNode(
  renderChartLoading: XlsxViewerProps["renderChartLoading"],
  chart: XlsxChart,
  rect: XlsxImageRect
) {
  const defaultNode = renderDefaultChartLoadingCard(rect);
  if (!renderChartLoading) {
    return defaultNode;
  }

  return renderChartLoading({
    chart,
    defaultNode,
    rect
  });
}

function formatBytes(value: number) {
  if (!Number.isFinite(value) || value <= 0) {
    return "0 B";
  }

  const units = ["B", "KB", "MB", "GB"];
  let size = value;
  let unitIndex = 0;
  while (size >= 1024 && unitIndex < units.length - 1) {
    size /= 1024;
    unitIndex += 1;
  }

  return `${size >= 10 || unitIndex === 0 ? size.toFixed(0) : size.toFixed(1)} ${units[unitIndex]}`;
}

function renderDeferredLoad(
  controller: XlsxViewerController,
  palette: ViewerPalette
) {
  return (
    <div
      style={{
        alignItems: "center",
        color: palette.text,
        display: "flex",
        height: "100%",
        justifyContent: "center",
        padding: 24
      }}
    >
      <div
        style={{
          backgroundColor: palette.surface,
          border: `1px solid ${palette.strongBorder}`,
          borderRadius: 12,
          boxShadow: palette.shadow,
          maxWidth: 420,
          padding: 20,
          width: "100%"
        }}
      >
        <div style={{ fontSize: 16, fontWeight: 600 }}>Large workbook detected</div>
        <div style={{ color: palette.mutedText, fontSize: 13, lineHeight: 1.5, marginTop: 8 }}>
          This workbook is {formatBytes(controller.deferredLoadFileSize ?? 0)}. Loading it immediately can block the main thread and freeze the page.
        </div>
        <div style={{ color: palette.mutedText, fontSize: 13, lineHeight: 1.5, marginTop: 8 }}>
          Best practice is to gate large files or move parsing into a worker. You can still load it manually below.
        </div>
        <div style={{ display: "flex", gap: 10, marginTop: 16 }}>
          <button
            onClick={controller.continueDeferredLoad}
            style={{
              background: palette.buttonSurface,
              border: `1px solid ${palette.strongBorder}`,
              borderRadius: 8,
              color: palette.buttonText,
              cursor: controller.canLoadDeferred ? "pointer" : "default",
              fontSize: 12,
              fontWeight: 600,
              opacity: controller.canLoadDeferred ? 1 : 0.6,
              padding: "8px 12px"
            }}
            type="button"
          >
            Load Anyway
          </button>
        </div>
      </div>
    </div>
  );
}

function resolveSelectionColors({
  palette,
  selectionColor,
  selectionFillColor,
  selectionHeaderColor
}: {
  palette: ViewerPalette;
  selectionColor?: string;
  selectionFillColor?: string;
  selectionHeaderColor?: string;
}) {
  const stroke = selectionColor ?? (paletteIsDark(palette) ? "#60a5fa" : "#2563eb");
  const fill =
    selectionFillColor ??
    `color-mix(in srgb, ${stroke} ${paletteIsDark(palette) ? "16%" : "10%"}, transparent)`;
  const header =
    selectionHeaderColor ??
    `color-mix(in srgb, ${stroke} ${paletteIsDark(palette) ? "24%" : "16%"}, ${palette.headerSurface})`;

  return {
    fill,
    header,
    stroke
  };
}

type CellRenderData = {
  canvas?: CanvasCellStyleCache;
  chartHighlight?: CellChartHighlight | null;
  checkboxState?: boolean | null;
  conditionalColorScale?: {
    color: string;
  } | null;
  colSpan?: number;
  conditionalDataBar?: {
    axisColor?: string;
    border?: boolean;
    borderColor?: string;
    color: string;
    gradient?: boolean;
    negativeBorderColor?: string;
    negativeFillColor?: string;
    widthPercent: number;
  } | null;
  conditionalIcon?: {
    backgroundColor?: string;
    borderColor?: string;
    color?: string;
    glyph?: string;
    rotationDeg?: number;
    shape?: "arrow";
  } | null;
  hyperlink?: {
    location?: string;
    target?: string;
    tooltip?: string;
  } | null;
  isMergedSecondary: boolean;
  shrinkToFit?: boolean;
  shrinkToFitFontSizePx?: number;
  isTableHeader?: boolean;
  rowSpan?: number;
  sparkline?: {
    config: XlsxSheetData["sparklines"][number];
    values: Array<number | null>;
  } | null;
  spillWidth?: number;
  style: React.CSSProperties;
  validation?: {
    message?: string;
    showDropdown: boolean;
    validationType: string;
  } | null;
  value: string;
  textRotationDeg?: number | null;
};

type PaintedBodyCanvasSignature = {
  activeSheet: XlsxSheetData | null;
  bakedDrawingSignature: string;
  bodyHeight: number;
  bodyWidth: number;
  colSignature: string;
  getCellData: (row: number, col: number) => CellRenderData;
  headerHeight: number;
  palette: ViewerPalette;
  rowHeaderWidth: number;
  rowSignature: string;
  zoomFactor: number;
};

type PaintedHeaderCanvasSignature = {
  activeSheet: XlsxSheetData | null;
  bodyHeight: number;
  bodyWidth: number;
  colSignature: string;
  headerHeight: number;
  palette: ViewerPalette;
  rangeSignature: string;
  rowHeaderWidth: number;
  rowSignature: string;
  viewportLeft: number;
  viewportTop: number;
  zoomFactor: number;
};

type CanvasDirtyRect = {
  height: number;
  left: number;
  top: number;
  width: number;
};

type CanvasStickyColumnHeaderCell = {
  actualCol: number;
  height: number;
  isFrozen: boolean;
  left: number;
  localLeft: number;
  width: number;
};

type CanvasStickyRowHeaderCell = {
  actualRow: number;
  height: number;
  isFrozen: boolean;
  localTop: number;
  top: number;
};

function buildFullCanvasDirtyRect(width: number, height: number): CanvasDirtyRect[] {
  if (width <= 0 || height <= 0) {
    return [];
  }

  return [{
    height,
    left: 0,
    top: 0,
    width
  }];
}

function intersectsCanvasDirtyRects(
  left: number,
  top: number,
  width: number,
  height: number,
  dirtyRects: CanvasDirtyRect[]
) {
  if (dirtyRects.length === 0 || width <= 0 || height <= 0) {
    return false;
  }

  const right = left + width;
  const bottom = top + height;
  return dirtyRects.some((dirtyRect) => (
    left < dirtyRect.left + dirtyRect.width
    && right > dirtyRect.left
    && top < dirtyRect.top + dirtyRect.height
    && bottom > dirtyRect.top
  ));
}

function blitCanvasWithScrollDelta(
  canvas: HTMLCanvasElement,
  context: CanvasRenderingContext2D,
  bufferCanvas: HTMLCanvasElement,
  dpr: number,
  width: number,
  height: number,
  deltaX: number,
  deltaY: number
): CanvasDirtyRect[] | null {
  if (width <= 0 || height <= 0) {
    return [];
  }

  const clampedDeltaX = Math.max(-width, Math.min(width, deltaX));
  const clampedDeltaY = Math.max(-height, Math.min(height, deltaY));
  if (Math.abs(clampedDeltaX) < 0.01 && Math.abs(clampedDeltaY) < 0.01) {
    return [];
  }

  if (Math.abs(clampedDeltaX) >= width || Math.abs(clampedDeltaY) >= height) {
    return null;
  }

  const deviceWidth = Math.max(1, Math.round(width * dpr));
  const deviceHeight = Math.max(1, Math.round(height * dpr));
  if (bufferCanvas.width !== deviceWidth) {
    bufferCanvas.width = deviceWidth;
  }
  if (bufferCanvas.height !== deviceHeight) {
    bufferCanvas.height = deviceHeight;
  }

  const bufferContext = bufferCanvas.getContext("2d", { alpha: false });
  if (!bufferContext) {
    return null;
  }

  bufferContext.setTransform(1, 0, 0, 1, 0, 0);
  bufferContext.clearRect(0, 0, deviceWidth, deviceHeight);
  bufferContext.drawImage(canvas, 0, 0);

  const deltaDeviceX = clampedDeltaX * dpr;
  const deltaDeviceY = clampedDeltaY * dpr;
  const overlapDeviceWidth = deviceWidth - Math.abs(deltaDeviceX);
  const overlapDeviceHeight = deviceHeight - Math.abs(deltaDeviceY);
  if (overlapDeviceWidth <= 0 || overlapDeviceHeight <= 0) {
    return null;
  }

  const sourceX = deltaDeviceX > 0 ? deltaDeviceX : 0;
  const sourceY = deltaDeviceY > 0 ? deltaDeviceY : 0;
  const destinationX = deltaDeviceX > 0 ? 0 : -deltaDeviceX;
  const destinationY = deltaDeviceY > 0 ? 0 : -deltaDeviceY;

  context.setTransform(1, 0, 0, 1, 0, 0);
  context.drawImage(
    bufferCanvas,
    sourceX,
    sourceY,
    overlapDeviceWidth,
    overlapDeviceHeight,
    destinationX,
    destinationY,
    overlapDeviceWidth,
    overlapDeviceHeight
  );
  context.setTransform(dpr, 0, 0, dpr, 0, 0);

  const dirtyRects: CanvasDirtyRect[] = [];
  if (clampedDeltaX > 0) {
    dirtyRects.push({
      height,
      left: Math.max(0, width - clampedDeltaX),
      top: 0,
      width: Math.min(width, clampedDeltaX)
    });
  } else if (clampedDeltaX < 0) {
    dirtyRects.push({
      height,
      left: 0,
      top: 0,
      width: Math.min(width, -clampedDeltaX)
    });
  }

  if (clampedDeltaY > 0) {
    dirtyRects.push({
      height: Math.min(height, clampedDeltaY),
      left: 0,
      top: Math.max(0, height - clampedDeltaY),
      width
    });
  } else if (clampedDeltaY < 0) {
    dirtyRects.push({
      height: Math.min(height, -clampedDeltaY),
      left: 0,
      top: 0,
      width
    });
  }

  return dirtyRects;
}

function buildConditionalFormatRuleKey(rule: XlsxSheetData["conditionalFormatRules"][number]) {
  return `${rule.kind}:${rule.priority}:${rule.ranges
    .map((range) => `${range.start.row}:${range.start.col}-${range.end.row}:${range.end.col}`)
    .join("|")}`;
}

function resolveConditionalRuleThreshold(
  threshold: XlsxSheetData["conditionalFormatRules"][number]["cfvos"][number] | undefined,
  numericValues: number[]
) {
  if (!threshold) {
    return null;
  }

  const fallbackValue = typeof threshold.value === "number" ? threshold.value : null;
  if (threshold.type === "num" || threshold.type === "formula") {
    return fallbackValue;
  }

  if (numericValues.length === 0) {
    return fallbackValue;
  }

  const minValue = Math.min(...numericValues);
  const maxValue = Math.max(...numericValues);
  if (threshold.type === "min") {
    return minValue;
  }
  if (threshold.type === "max") {
    return maxValue;
  }

  if ((threshold.type === "percent" || threshold.type === "percentile") && fallbackValue !== null) {
    return minValue + (maxValue - minValue) * (fallbackValue / 100);
  }

  return fallbackValue;
}

function buildConditionalIcon(iconSet: string, iconId: number): NonNullable<CellRenderData["conditionalIcon"]> | null {
  switch (iconSet) {
    case "NoIcons":
      return null;
    case "4TrafficLights": {
      const definitions = [
        { backgroundColor: "#111827", borderColor: "#6b7280" },
        { backgroundColor: "#22c55e", borderColor: "#166534" },
        { backgroundColor: "#ef4444", borderColor: "#991b1b" },
        { backgroundColor: "#facc15", borderColor: "#a16207" }
      ];
      return definitions[iconId] ?? definitions[definitions.length - 1] ?? null;
    }
    case "3TrafficLights1":
    case "3TrafficLights2": {
      const definitions = [
        { backgroundColor: "#22c55e", borderColor: iconSet === "3TrafficLights2" ? "#111827" : "#166534" },
        { backgroundColor: "#facc15", borderColor: iconSet === "3TrafficLights2" ? "#111827" : "#a16207" },
        { backgroundColor: "#ef4444", borderColor: iconSet === "3TrafficLights2" ? "#111827" : "#991b1b" }
      ];
      return definitions[iconId] ?? definitions[definitions.length - 1] ?? null;
    }
    case "4RedToBlack": {
      const definitions = [
        { backgroundColor: "#ef4444", borderColor: "#991b1b" },
        { backgroundColor: "#fca5a5", borderColor: "#b91c1c" },
        { backgroundColor: "#9ca3af", borderColor: "#4b5563" },
        { backgroundColor: "#111827", borderColor: "#6b7280" }
      ];
      return definitions[iconId] ?? definitions[definitions.length - 1] ?? null;
    }
    case "3Symbols":
      return [
        { backgroundColor: "#22c55e", color: "#ffffff", glyph: "✓" },
        { backgroundColor: "#facc15", color: "#111827", glyph: "!" },
        { backgroundColor: "#ef4444", color: "#ffffff", glyph: "×" }
      ][iconId] ?? null;
    case "3Symbols2":
      return [
        { color: "#16a34a", glyph: "✓" },
        { color: "#ca8a04", glyph: "!" },
        { color: "#dc2626", glyph: "×" }
      ][iconId] ?? null;
    case "3Signs":
      return [
        { backgroundColor: "#22c55e", borderColor: "#166534" },
        { backgroundColor: "#facc15", borderColor: "#a16207", glyph: "▲", color: "#111827" },
        { backgroundColor: "#ef4444", borderColor: "#991b1b", glyph: "◆", color: "#ffffff" }
      ][iconId] ?? null;
    default:
      if (iconSet.includes("Arrows")) {
        const arrowColors = iconSet.includes("Gray")
          ? ["#9ca3af", "#9ca3af", "#9ca3af", "#9ca3af", "#9ca3af"]
          : ["#16a34a", "#ca8a04", "#ca8a04", "#ca8a04", "#dc2626"];
        const arrowRotations = iconSet.startsWith("3")
          ? [-90, 0, 90]
          : iconSet.startsWith("4")
            ? [-90, -35, 35, 90]
            : [-90, -35, 0, 35, 90];
        const resolvedIndex = Math.min(iconId, arrowColors.length - 1, arrowRotations.length - 1);
        return {
          borderColor: iconSet.includes("Gray") ? "#6b7280" : darkenColor(arrowColors[resolvedIndex] ?? "#16a34a", 0.34),
          color: arrowColors[resolvedIndex],
          rotationDeg: arrowRotations[resolvedIndex] ?? 0,
          shape: "arrow"
        };
      }

      return {
        backgroundColor: "#111827",
        borderColor: "#6b7280"
      };
  }
}

function renderConditionalIcon(icon: NonNullable<CellRenderData["conditionalIcon"]>, scale = 1) {
  const iconSize = 14 * scale;
  if (icon.shape === "arrow") {
    const fill = icon.color ?? "#111827";
    const stroke = icon.borderColor ?? darkenColor(fill, 0.32);
    return (
      <svg
        aria-hidden="true"
        height={iconSize}
        style={{ display: "block" }}
        viewBox="0 0 16 16"
        width={iconSize}
      >
        <g transform={`rotate(${icon.rotationDeg ?? 0} 8 8)`}>
          <path
            d="M2.5 8 L8.4 2.4 L8.4 5.2 L13.5 5.2 L13.5 10.8 L8.4 10.8 L8.4 13.6 Z"
            fill={fill}
            stroke={stroke}
            strokeLinejoin="round"
            strokeWidth={1.25}
          />
        </g>
      </svg>
    );
  }
  if (icon.glyph) {
    return (
      <span
        style={{
          alignItems: "center",
          color: icon.color ?? "#111827",
          display: "inline-flex",
          fontSize: 13 * scale,
          fontWeight: 700,
          height: iconSize,
          justifyContent: "center",
          lineHeight: 1,
          width: iconSize
        }}
      >
        {icon.glyph}
      </span>
    );
  }

  return (
    <span
      style={{
        backgroundColor: icon.backgroundColor ?? "#111827",
        border: icon.borderColor ? `1px solid ${icon.borderColor}` : "none",
        borderRadius: "999px",
        display: "inline-block",
        height: 12 * scale,
        width: 12 * scale
      }}
    />
  );
}

function renderCheckboxControl(checked: boolean, palette: ViewerPalette, scale = 1) {
  const stroke = paletteIsDark(palette) ? "#cbd5e1" : "#475569";
  const fill = checked ? (paletteIsDark(palette) ? "#60a5fa" : "#2563eb") : "transparent";
  const check = paletteIsDark(palette) ? "#020617" : "#ffffff";
  return (
    <svg aria-hidden="true" height={14 * scale} style={{ display: "block" }} viewBox="0 0 16 16" width={14 * scale}>
      <rect fill={fill} height={11} rx={2} ry={2} stroke={stroke} strokeWidth={1.2} width={11} x={2.5} y={2.5} />
      {checked ? (
        <path
          d="M5 8.1 7.1 10.2 11.3 5.8"
          fill="none"
          stroke={check}
          strokeLinecap="round"
          strokeLinejoin="round"
          strokeWidth={1.8}
        />
      ) : null}
    </svg>
  );
}

function renderRadioControl(checked: boolean, palette: ViewerPalette, scale = 1) {
  const stroke = paletteIsDark(palette) ? "#cbd5e1" : "#475569";
  const dot = paletteIsDark(palette) ? "#60a5fa" : "#2563eb";
  return (
    <svg aria-hidden="true" height={14 * scale} style={{ display: "block" }} viewBox="0 0 16 16" width={14 * scale}>
      <circle cx={8} cy={8} fill="transparent" r={5.5} stroke={stroke} strokeWidth={1.2} />
      {checked ? <circle cx={8} cy={8} fill={dot} r={2.75} /> : null}
    </svg>
  );
}

function resolveFormControlLabel(control: XlsxFormControl) {
  const label = control.label ?? control.name;
  if (!label) {
    return "";
  }

  const normalized = label.replace(/\u00a0/g, " ").replace(/\s+/g, " ").trim().toLowerCase();
  if (/^(option button|group box|check box|drop down|dropdown|list box|edit box|scroll bar|spinner|spin button|button)\s+\d+$/.test(normalized)) {
    return "";
  }

  return label.replace(/\u00a0/g, " ").replace(/^\s+/, "");
}

function drawStaticShapeText(
  context: CanvasRenderingContext2D,
  shape: XlsxShape,
  left: number,
  top: number,
  width: number,
  height: number,
  zoomFactor: number
) {
  if (shape.paragraphs.length === 0) {
    return;
  }

  const inset = shape.textBox?.insetPx;
  const paddingLeft = (inset?.left ?? 6) * zoomFactor;
  const paddingRight = (inset?.right ?? 6) * zoomFactor;
  const paddingTop = (inset?.top ?? 4) * zoomFactor;
  const paddingBottom = (inset?.bottom ?? 4) * zoomFactor;
  const textLeft = left + paddingLeft;
  const textTop = top + paddingTop;
  const textWidth = Math.max(0, width - paddingLeft - paddingRight);
  const textHeight = Math.max(0, height - paddingTop - paddingBottom);
  if (textWidth <= 0 || textHeight <= 0) {
    return;
  }

  const lineMetrics = shape.paragraphs.map((paragraph) => {
    const lineHeight = Math.max(
      12 * zoomFactor,
      ...paragraph.runs.map((run) => ((run.fontSizePt ?? 11) * 96 / 72) * zoomFactor * 1.2)
    );
    const widthPx = paragraph.runs.reduce((total, run) => {
      const fontSize = ((run.fontSizePt ?? 11) * 96 / 72) * zoomFactor;
      context.font = `${run.italic ? "italic " : ""}${run.bold ? "700 " : "400 "}${fontSize}px ${run.fontFamily ?? "Calibri, sans-serif"}`;
      return total + measureCanvasTextWidth(context, run.text);
    }, 0);
    return { lineHeight, widthPx };
  });
  const totalTextHeight = lineMetrics.reduce((total, metric) => total + metric.lineHeight, 0);
  let y = textTop;
  if (shape.textBox?.verticalAlign === "middle") {
    y = textTop + Math.max(0, (textHeight - totalTextHeight) / 2);
  } else if (shape.textBox?.verticalAlign === "bottom") {
    y = textTop + Math.max(0, textHeight - totalTextHeight);
  }

  context.save();
  context.beginPath();
  context.rect(textLeft, textTop, textWidth, textHeight);
  context.clip();
  context.textBaseline = "middle";
  shape.paragraphs.forEach((paragraph, paragraphIndex) => {
    const metric = lineMetrics[paragraphIndex] ?? { lineHeight: 14 * zoomFactor, widthPx: 0 };
    let x = textLeft;
    const align = paragraph.align ?? shape.textBox?.horizontalAlign ?? "left";
    if (align === "center") {
      x = textLeft + Math.max(0, (textWidth - metric.widthPx) / 2);
    } else if (align === "right") {
      x = textLeft + Math.max(0, textWidth - metric.widthPx);
    }

    paragraph.runs.forEach((run) => {
      const fontSize = ((run.fontSizePt ?? 11) * 96 / 72) * zoomFactor;
      context.font = `${run.italic ? "italic " : ""}${run.bold ? "700 " : "400 "}${fontSize}px ${run.fontFamily ?? "Calibri, sans-serif"}`;
      context.fillStyle = run.color ?? "#000000";
      context.textAlign = "left";
      const textY = y + metric.lineHeight / 2;
      context.fillText(run.text, x, textY);
      if (run.underline && run.text.length > 0) {
        const textWidthPx = measureCanvasTextWidth(context, run.text);
        context.strokeStyle = run.color ?? "#000000";
        context.lineWidth = Math.max(1, zoomFactor * 0.75);
        context.beginPath();
        context.moveTo(x, textY + Math.max(2, fontSize * 0.28));
        context.lineTo(x + textWidthPx, textY + Math.max(2, fontSize * 0.28));
        context.stroke();
      }
      x += measureCanvasTextWidth(context, run.text);
    });
    y += metric.lineHeight;
  });
  context.restore();
}

function drawStaticShape(
  context: CanvasRenderingContext2D,
  shape: XlsxShape,
  rect: XlsxImageRect,
  zoomFactor: number
) {
  const fillColor = shape.fill?.none ? "transparent" : (shape.fill?.color ?? "transparent");
  const strokeColor = shape.stroke?.none ? "transparent" : (shape.stroke?.color ?? "transparent");
  const lineWidth = Math.max(0, (shape.stroke?.widthPx ?? (shape.geometry === "line" ? 2 : 1)) * zoomFactor);
  const vectorShape = resolveShapeVector(shape);
  const opacity = Math.min(shape.fill?.opacity ?? 1, shape.stroke?.opacity ?? 1);

  context.save();
  context.translate(rect.left + rect.width / 2, rect.top + rect.height / 2);
  if (shape.rotationDeg) {
    context.rotate((shape.rotationDeg * Math.PI) / 180);
  }
  context.scale(shape.flipH ? -1 : 1, shape.flipV ? -1 : 1);
  context.globalAlpha *= opacity;
  context.lineWidth = lineWidth;
  context.strokeStyle = strokeColor;
  context.fillStyle = fillColor;
  applyCanvasShapeDash(context, shape.stroke?.dash, lineWidth);

  const localLeft = -rect.width / 2;
  const localTop = -rect.height / 2;
  if (vectorShape && typeof Path2D !== "undefined") {
    context.save();
    context.translate(localLeft, localTop);
    context.scale(
      rect.width / Math.max(1, vectorShape.viewBox.width),
      rect.height / Math.max(1, vectorShape.viewBox.height)
    );
    const path = getCachedCanvasPath2D(vectorShape.path);
    if (!path) {
      context.restore();
      context.restore();
      return;
    }
    if (fillColor !== "transparent") {
      context.fill(path);
    }
    if (strokeColor !== "transparent" && lineWidth > 0) {
      context.stroke(path);
    }
    context.restore();
  } else if (shape.geometry === "ellipse") {
    context.beginPath();
    context.ellipse(0, 0, rect.width / 2, rect.height / 2, 0, 0, Math.PI * 2);
    if (fillColor !== "transparent") {
      context.fill();
    }
    if (strokeColor !== "transparent" && lineWidth > 0) {
      context.stroke();
    }
  } else {
    const radius = shape.geometry === "roundRect" ? 12 * zoomFactor : 0;
    drawCanvasRoundedRect(context, localLeft, localTop, rect.width, rect.height, radius);
    if (fillColor !== "transparent") {
      context.fill();
    }
    if (strokeColor !== "transparent" && lineWidth > 0) {
      context.stroke();
    }
  }

  drawStaticShapeText(context, shape, localLeft, localTop, rect.width, rect.height, zoomFactor);
  context.restore();
}

function drawStaticFormControl(
  context: CanvasRenderingContext2D,
  control: XlsxFormControl,
  rect: XlsxImageRect,
  palette: ViewerPalette,
  zoomFactor: number,
  sheetSurface = SHEET_SURFACE
) {
  const label = resolveFormControlLabel(control);
  const stroke = paletteIsDark(palette) ? "#cbd5e1" : "#475569";
  const textColor = control.textColor ?? "#000000";
  const fontSizePx = Math.max(9 * zoomFactor, ((control.fontSizePt ?? 9) * 96 / 72) * zoomFactor);
  const iconSize = Math.min(14 * zoomFactor, Math.max(0, rect.height - 4 * zoomFactor));
  context.save();
  context.font = `400 ${fontSizePx}px ${control.fontFamily ?? "Calibri, sans-serif"}`;
  context.textBaseline = "middle";
  context.fillStyle = textColor;
  context.strokeStyle = stroke;
  context.lineWidth = Math.max(1, zoomFactor);

  if (control.kind === "group-box") {
    const labelInset = label ? Math.max(7, fontSizePx * 0.5) : 0;
    drawCanvasRoundedRect(context, rect.left, rect.top + labelInset, rect.width, Math.max(0, rect.height - labelInset), 2 * zoomFactor);
    context.stroke();
    if (label) {
      context.fillStyle = sheetSurface;
      const labelWidth = Math.min(rect.width - 16 * zoomFactor, measureCanvasTextWidth(context, label) + 8 * zoomFactor);
      context.fillRect(rect.left + 8 * zoomFactor, rect.top, labelWidth, fontSizePx * 1.2);
      context.fillStyle = textColor;
      context.textAlign = "left";
      context.fillText(label, rect.left + 12 * zoomFactor, rect.top + fontSizePx * 0.55);
    }
    context.restore();
    return;
  }

  if (control.kind === "button" || control.kind === "dropdown" || control.kind === "editbox" || control.kind === "listbox" || control.kind === "scrollbar" || control.kind === "spinner" || control.kind === "unknown") {
    if (control.kind === "button") {
      const gradient = context.createLinearGradient(rect.left, rect.top, rect.left, rect.top + rect.height);
      gradient.addColorStop(0, "#f8fafc");
      gradient.addColorStop(1, "#e2e8f0");
      context.fillStyle = gradient;
    } else {
      context.fillStyle = "transparent";
    }
    drawCanvasRoundedRect(context, rect.left, rect.top, rect.width, rect.height, control.kind === "button" ? 4 * zoomFactor : 2 * zoomFactor);
    if (control.kind === "button") {
      context.fill();
    }
    context.stroke();
  }

  let textLeft = rect.left + 2 * zoomFactor;
  const textRightInset = control.kind === "dropdown" ? 18 * zoomFactor : 6 * zoomFactor;
  if (control.kind === "checkbox") {
    context.strokeRect(rect.left + 2 * zoomFactor, rect.top + (rect.height - iconSize) / 2, iconSize, iconSize);
    if (control.checked) {
      context.fillStyle = paletteIsDark(palette) ? "#60a5fa" : "#2563eb";
      context.fillRect(rect.left + 2 * zoomFactor + 1.5, rect.top + (rect.height - iconSize) / 2 + 1.5, Math.max(0, iconSize - 3), Math.max(0, iconSize - 3));
      context.strokeStyle = paletteIsDark(palette) ? "#020617" : "#ffffff";
      context.beginPath();
      context.moveTo(rect.left + 2 * zoomFactor + iconSize * 0.24, rect.top + rect.height / 2 + iconSize * 0.06);
      context.lineTo(rect.left + 2 * zoomFactor + iconSize * 0.45, rect.top + rect.height / 2 + iconSize * 0.26);
      context.lineTo(rect.left + 2 * zoomFactor + iconSize * 0.8, rect.top + rect.height / 2 - iconSize * 0.2);
      context.stroke();
    }
    textLeft += iconSize + 4 * zoomFactor;
  } else if (control.kind === "radio") {
    context.beginPath();
    context.arc(rect.left + 2 * zoomFactor + iconSize / 2, rect.top + rect.height / 2, iconSize / 2, 0, Math.PI * 2);
    context.stroke();
    if (control.checked) {
      context.fillStyle = paletteIsDark(palette) ? "#60a5fa" : "#2563eb";
      context.beginPath();
      context.arc(rect.left + 2 * zoomFactor + iconSize / 2, rect.top + rect.height / 2, iconSize * 0.25, 0, Math.PI * 2);
      context.fill();
    }
    textLeft += iconSize + 4 * zoomFactor;
  }

  if (label) {
    const maxTextWidth = Math.max(0, rect.left + rect.width - textLeft - textRightInset);
    const text = truncateCanvasText(context, label, maxTextWidth);
    context.fillStyle = textColor;
    context.textAlign = control.textAlign === "right" ? "right" : control.textAlign === "center" ? "center" : "left";
    const textX = context.textAlign === "right"
      ? rect.left + rect.width - textRightInset
      : context.textAlign === "center"
        ? textLeft + maxTextWidth / 2
        : textLeft;
    context.fillText(text, textX, rect.top + rect.height / 2);
  }
  if (control.kind === "dropdown") {
    context.fillStyle = textColor;
    context.textAlign = "center";
    context.fillText("▼", rect.left + rect.width - 10 * zoomFactor, rect.top + rect.height / 2);
  }
  context.restore();
}

function resolveConditionalDataBarForCell(
  row: number,
  col: number,
  worksheet: Worksheet,
  sheet: XlsxSheetData | null | undefined,
  metricsCache: Map<string, number[]>
): CellRenderData["conditionalDataBar"] {
  const rules = sheet?.conditionalFormatRules ?? [];
  const matchingRule = rules.find(
    (rule): rule is Extract<XlsxSheetData["conditionalFormatRules"][number], { kind: "dataBar" }> =>
      rule.kind === "dataBar" && rule.ranges.some((range) => isCellInRange({ row, col }, range))
  );
  if (!matchingRule) {
    return null;
  }

  const numericValue = getCellNumericValue(worksheet, row, col);
  if (numericValue === null) {
    return null;
  }

  const cacheKey = buildConditionalFormatRuleKey(matchingRule);
  let ruleValues = metricsCache.get(cacheKey);
  if (!ruleValues) {
    ruleValues = matchingRule.ranges.flatMap((range) => {
      const values: number[] = [];
      for (let targetRow = range.start.row; targetRow <= range.end.row; targetRow += 1) {
        for (let targetCol = range.start.col; targetCol <= range.end.col; targetCol += 1) {
          const value = getCellNumericValue(worksheet, targetRow, targetCol);
          if (value !== null) {
            values.push(value);
          }
        }
      }
      return values;
    });
    metricsCache.set(cacheKey, ruleValues);
  }

  const minValue = resolveConditionalRuleThreshold(matchingRule.cfvos[0], ruleValues);
  const maxValue = resolveConditionalRuleThreshold(matchingRule.cfvos[matchingRule.cfvos.length - 1], ruleValues);
  if (minValue === null || maxValue === null) {
    return null;
  }

  const minLength = Number.isFinite(matchingRule.minLength ?? Number.NaN) ? (matchingRule.minLength ?? 0) : 0;
  const maxLength = Number.isFinite(matchingRule.maxLength ?? Number.NaN) ? (matchingRule.maxLength ?? 100) : 100;
  const span = maxValue - minValue;
  const ratio = span === 0 ? (numericValue >= maxValue ? 1 : 0) : (numericValue - minValue) / span;
  const widthPercent = minLength + Math.max(0, Math.min(1, ratio)) * (maxLength - minLength);
  const color = resolveWorkbookColor(matchingRule.color, sheet?.themePalette);
  if (!color || widthPercent <= 0) {
    return null;
  }

  return {
    axisColor: resolveWorkbookColor(matchingRule.axisColor, sheet?.themePalette) ?? undefined,
    border: matchingRule.border ?? undefined,
    borderColor: resolveWorkbookColor(matchingRule.borderColor, sheet?.themePalette) ?? undefined,
    color,
    gradient: matchingRule.gradient ?? undefined,
    negativeBorderColor: resolveWorkbookColor(matchingRule.negativeBorderColor, sheet?.themePalette) ?? undefined,
    negativeFillColor: resolveWorkbookColor(matchingRule.negativeFillColor, sheet?.themePalette) ?? undefined,
    widthPercent
  };
}

function resolveConditionalColorScaleForCell(
  row: number,
  col: number,
  worksheet: Worksheet,
  sheet: XlsxSheetData | null | undefined,
  metricsCache: Map<string, number[]>
): CellRenderData["conditionalColorScale"] {
  const rules = sheet?.conditionalFormatRules ?? [];
  const matchingRule = rules.find(
    (rule): rule is Extract<XlsxSheetData["conditionalFormatRules"][number], { kind: "colorScale" }> =>
      rule.kind === "colorScale" && rule.ranges.some((range) => isCellInRange({ row, col }, range))
  );
  if (!matchingRule) {
    return null;
  }

  const numericValue = getCellNumericValue(worksheet, row, col);
  if (numericValue === null) {
    return null;
  }

  const cacheKey = buildConditionalFormatRuleKey(matchingRule);
  let ruleValues = metricsCache.get(cacheKey);
  if (!ruleValues) {
    ruleValues = matchingRule.ranges.flatMap((range) => {
      const values: number[] = [];
      for (let targetRow = range.start.row; targetRow <= range.end.row; targetRow += 1) {
        for (let targetCol = range.start.col; targetCol <= range.end.col; targetCol += 1) {
          const value = getCellNumericValue(worksheet, targetRow, targetCol);
          if (value !== null) {
            values.push(value);
          }
        }
      }
      return values;
    });
    metricsCache.set(cacheKey, ruleValues);
  }

  const colors = matchingRule.colors
    .map((color) => resolveWorkbookColor(color, sheet?.themePalette))
    .filter((color): color is string => typeof color === "string");
  const thresholds = matchingRule.cfvos
    .map((threshold) => resolveConditionalRuleThreshold(threshold, ruleValues))
    .filter((value): value is number => typeof value === "number" && Number.isFinite(value));
  if (colors.length < 2 || thresholds.length < 2) {
    return null;
  }

  const segments = Math.min(colors.length, thresholds.length) - 1;
  let resolvedColor = colors[colors.length - 1];
  for (let index = 0; index < segments; index += 1) {
    const startValue = thresholds[index] ?? thresholds[0];
    const endValue = thresholds[index + 1] ?? thresholds[thresholds.length - 1];
    if (numericValue <= endValue || index === segments - 1) {
      const ratio = endValue === startValue ? 1 : (numericValue - startValue) / (endValue - startValue);
      resolvedColor = mixRgbColor(colors[index] ?? colors[0], colors[index + 1] ?? colors[colors.length - 1], Math.max(0, Math.min(1, ratio)));
      break;
    }
  }

  return { color: resolvedColor };
}

function resolveConditionalIconForCell(
  row: number,
  col: number,
  worksheet: Worksheet,
  sheet: XlsxSheetData | null | undefined,
  metricsCache: Map<string, number[]>
): CellRenderData["conditionalIcon"] {
  const rules = sheet?.conditionalFormatRules ?? [];
  const matchingRule = rules.find(
    (rule): rule is Extract<XlsxSheetData["conditionalFormatRules"][number], { kind: "iconSet" }> =>
      rule.kind === "iconSet" && rule.ranges.some((range) => isCellInRange({ row, col }, range))
  );
  if (!matchingRule) {
    return null;
  }

  const numericValue = getCellNumericValue(worksheet, row, col);
  if (numericValue === null) {
    return null;
  }

  const cacheKey = buildConditionalFormatRuleKey(matchingRule);
  let ruleValues = metricsCache.get(cacheKey);
  if (!ruleValues) {
    ruleValues = matchingRule.ranges.flatMap((range) => {
      const values: number[] = [];
      for (let targetRow = range.start.row; targetRow <= range.end.row; targetRow += 1) {
        for (let targetCol = range.start.col; targetCol <= range.end.col; targetCol += 1) {
          const value = getCellNumericValue(worksheet, targetRow, targetCol);
          if (value !== null) {
            values.push(value);
          }
        }
      }
      return values;
    });
    metricsCache.set(cacheKey, ruleValues);
  }

  const thresholds = matchingRule.cfvos.map((threshold) => resolveConditionalRuleThreshold(threshold, ruleValues));
  let selectedIndex = 0;
  thresholds.forEach((threshold, index) => {
    if (threshold !== null && numericValue >= threshold) {
      selectedIndex = index;
    }
  });

  const iconIndex = matchingRule.reverse
    ? Math.max(0, matchingRule.icons.length - 1 - selectedIndex)
    : selectedIndex;
  const icon = matchingRule.icons[Math.min(iconIndex, matchingRule.icons.length - 1)];
  return icon ? buildConditionalIcon(icon.iconSet, icon.iconId) : null;
}

function resolveCellDataValidation(
  row: number,
  col: number,
  sheet: XlsxSheetData | null | undefined
): CellRenderData["validation"] {
  const validation = sheet?.dataValidations.find((entry) => entry.ranges.some((range) => isCellInRange({ row, col }, range)));
  if (!validation) {
    return null;
  }

  const messageParts = [validation.inputMessage, validation.errorMessage].filter(
    (value): value is string => typeof value === "string" && value.trim().length > 0
  );

  return {
    message: messageParts.join("\n"),
    showDropdown: validation.showDropdown !== false,
    validationType: validation.validationType
  };
}

function resolveCellContentJustify(verticalAlign: React.CSSProperties["verticalAlign"]) {
  if (verticalAlign === "middle") {
    return "center";
  }
  if (verticalAlign === "top") {
    return "flex-start";
  }
  return "flex-end";
}

function resolveCellContentAlignItems(textAlign: React.CSSProperties["textAlign"]) {
  if (textAlign === "center") {
    return "center";
  }
  if (textAlign === "right" || textAlign === "end") {
    return "flex-end";
  }
  return "flex-start";
}

function resolveFirstUsedVisibleIndex(visibleIndices: number[], minUsedIndex: number) {
  if (visibleIndices.length === 0) {
    return -1;
  }
  if (minUsedIndex <= 0) {
    return visibleIndices[0] ?? -1;
  }

  const firstVisibleUsed = visibleIndices.find((index) => index >= minUsedIndex);
  return firstVisibleUsed ?? visibleIndices[0] ?? -1;
}

function resolveLastUsedVisibleIndex(visibleIndices: number[], maxUsedIndex: number) {
  if (visibleIndices.length === 0) {
    return -1;
  }
  if (maxUsedIndex < 0) {
    return visibleIndices[visibleIndices.length - 1] ?? -1;
  }

  for (let index = visibleIndices.length - 1; index >= 0; index -= 1) {
    const visibleIndex = visibleIndices[index];
    if (visibleIndex !== undefined && visibleIndex <= maxUsedIndex) {
      return visibleIndex;
    }
  }

  return visibleIndices[visibleIndices.length - 1] ?? -1;
}

type RenderedAxisItem = {
  end: number;
  index: number;
  key: React.Key;
  size: number;
  start: number;
};

type RenderedColumn = {
  actualCol: number;
  key: React.Key;
  size: number;
  virtualIndex: number;
};

type GridRowProps = {
  actualRow: number;
  editingCell: XlsxCellAddress | null;
  editingValue: string;
  frozenRowHeaderZIndex: number;
  getCellData: (row: number, col: number) => CellRenderData;
  headerLabelLiveScale: number;
  leadingSpacerWidth: number;
  onCellDoubleClick: (cell: XlsxCellAddress) => void;
  onCellClick: (cell: XlsxCellAddress, cellData: CellRenderData) => void;
  onCellPointerDown: (event: React.PointerEvent<HTMLTableCellElement>, cell: XlsxCellAddress) => void;
  onEditingBlur: () => void;
  onEditingCancel: () => void;
  onEditingCommit: () => void;
  onEditingNavigate: (key: string, value: string) => void;
  onEditingValueChange: (value: string) => void;
  onRowHeaderRef: (actualRow: number, element: HTMLTableCellElement | null) => void;
  onRowPointerDown: (event: React.PointerEvent<HTMLTableCellElement>, actualRow: number) => void;
  onRowResizePointerDown: (event: React.PointerEvent<HTMLDivElement>, actualRow: number, rowHeight: number) => void;
  palette: ViewerPalette;
  readOnly: boolean;
  renderCellAdornment?: (cell: XlsxCellAddress) => React.ReactNode;
  rowHeight: number;
  rowHeaderZIndex: number;
  rowHeaderWidth: number;
  stickyLeftByCol: Map<number, number>;
  stickyTop?: number;
  trailingSpacerWidth: number;
  visibleCols: RenderedColumn[];
  zoomFactor: number;
};

function GridRow({
  actualRow,
  editingCell,
  editingValue,
  frozenRowHeaderZIndex,
  getCellData,
  headerLabelLiveScale,
  leadingSpacerWidth,
  onCellClick,
  onCellDoubleClick,
  onCellPointerDown,
  onEditingBlur,
  onEditingCancel,
  onEditingCommit,
  onEditingNavigate,
  onEditingValueChange,
  onRowHeaderRef,
  onRowPointerDown,
  onRowResizePointerDown,
  palette,
  readOnly,
  renderCellAdornment,
  rowHeight,
  rowHeaderZIndex,
  rowHeaderWidth,
  stickyLeftByCol,
  stickyTop,
  trailingSpacerWidth,
  visibleCols,
  zoomFactor
}: GridRowProps) {
  const gutterSeparatorShadow = `inset -1px 0 0 ${palette.border}, inset 0 -1px 0 ${palette.border}`;

  return (
    <tr data-xlsx-row={actualRow} style={{ height: rowHeight }}>
      <td
        ref={(element) => onRowHeaderRef(actualRow, element)}
        onPointerDown={(event) => onRowPointerDown(event, actualRow)}
        style={{
          backgroundColor: palette.rowHeaderSurface,
          borderBottom: "none",
          borderRight: "none",
          boxSizing: "border-box",
          boxShadow: gutterSeparatorShadow,
          color: palette.mutedText,
          fontSize: scaleCssLengthExpression("11px", zoomFactor),
          height: rowHeight,
          left: 0,
          maxHeight: rowHeight,
          minWidth: rowHeaderWidth,
          padding: scaleCssLengthExpression("2px 4px", zoomFactor),
          position: "sticky",
          top: stickyTop,
          textAlign: "center",
          userSelect: "none",
          width: rowHeaderWidth,
          zIndex: stickyTop !== undefined ? frozenRowHeaderZIndex : rowHeaderZIndex
        }}
      >
        <div style={{ position: "relative" }}>
          <span
            style={{
              display: "inline-block",
              transform: headerLabelLiveScale !== 1 ? `scale(${1 / headerLabelLiveScale})` : undefined,
              transformOrigin: "center center"
            }}
          >
            {actualRow + 1}
          </span>
          <div
            onPointerDown={(event) => onRowResizePointerDown(event, actualRow, rowHeight)}
            style={{
              backgroundColor: "transparent",
              bottom: -8 * zoomFactor,
              cursor: "row-resize",
              height: 16 * zoomFactor,
              left: 0,
              position: "absolute",
              width: "100%",
              zIndex: 5
            }}
          />
        </div>
      </td>
      {leadingSpacerWidth > 0 ? (
        <td
          aria-hidden="true"
          style={{
            backgroundColor: "transparent",
            border: "none",
            padding: 0,
            width: leadingSpacerWidth
          }}
        />
      ) : null}
      {visibleCols.map(({ actualCol, key }) => {
        const cellData = getCellData(actualRow, actualCol);
        if (cellData.isMergedSecondary) {
          return null;
        }

        const cell = { row: actualRow, col: actualCol };
        const isEditing = isSameCell(editingCell, cell);
        const isSpilling = Boolean(cellData.spillWidth && cellData.spillWidth > 0);
        const adornment = renderCellAdornment ? renderCellAdornment(cell) : null;
        const stickyLeft = stickyLeftByCol.get(actualCol);
        const validationRight = adornment ? 24 : 4;
        const conditionalIconRight = validationRight;
        const cellStyle: React.CSSProperties = {
          ...cellData.style,
          boxSizing: "border-box",
          cursor: isEditing ? "text" : cellData.hyperlink ? "pointer" : "cell"
        };

        if (!cellData.rowSpan) {
          cellStyle.height = rowHeight;
          cellStyle.maxHeight = rowHeight;
        }

        if (isSpilling) {
          cellStyle.position = "relative";
          cellStyle.zIndex = 2;
        }
        if (cellData.conditionalColorScale) {
          cellStyle.backgroundColor = cellData.conditionalColorScale.color;
        }
        if (cellData.conditionalColorScale || cellData.conditionalDataBar || cellData.conditionalIcon) {
          cellStyle.position = "relative";
        }
        if (cellData.chartHighlight) {
          cellStyle.position = cellStyle.position ?? "relative";
        }
        if (cellData.isTableHeader) {
          cellStyle.position = "relative";
          cellStyle.zIndex = Math.max(Number(cellStyle.zIndex ?? 0), 4);
        }
        if (stickyTop !== undefined) {
          cellStyle.position = "sticky";
          cellStyle.top = stickyTop;
          cellStyle.zIndex = Math.max(Number(cellStyle.zIndex ?? 0), 28);
        }
        if (stickyLeft !== undefined) {
          cellStyle.left = stickyLeft;
          cellStyle.position = "sticky";
          cellStyle.zIndex = Math.max(Number(cellStyle.zIndex ?? 0), stickyTop !== undefined ? 30 : 24);
        }
        if (isSpilling) {
          cellStyle.overflow = "visible";
          cellStyle.textOverflow = "clip";
          cellStyle.position = cellStyle.position ?? "relative";
          cellStyle.zIndex = Math.max(Number(cellStyle.zIndex ?? 0), stickyTop !== undefined ? 32 : 6);
        }
        if (isEditing) {
          cellStyle.padding = 0;
        }

        const cellContentStyle: React.CSSProperties = {
          alignItems: "stretch",
          display: "flex",
          flexDirection: "column",
          font: "inherit",
          height: "100%",
          justifyContent: resolveCellContentJustify(cellStyle.verticalAlign),
          maxHeight: cellData.rowSpan ? undefined : rowHeight,
          minHeight: 0,
          overflow: "hidden",
          pointerEvents: "none",
          textAlign: "inherit",
          textDecoration: "inherit",
          textOverflow: "inherit",
          whiteSpace: "inherit",
          width: "100%",
          wordBreak: "inherit"
        };
        const trailingInset = (adornment ? 20 : 0) + (cellData.conditionalIcon ? 18 : 0);
        if (cellData.conditionalDataBar) {
          cellContentStyle.position = "relative";
          cellContentStyle.zIndex = 1;
        }
        if (cellData.chartHighlight) {
          cellContentStyle.position = cellContentStyle.position ?? "relative";
          cellContentStyle.zIndex = Math.max(Number(cellContentStyle.zIndex ?? 0), 1);
        }
        if (trailingInset > 0) {
          cellContentStyle.paddingRight = (trailingInset + 4) * zoomFactor;
        }
        if (cellData.conditionalIcon) {
          cellContentStyle.position = "relative";
          cellContentStyle.zIndex = 1;
        }
        if (cellData.sparkline || cellData.checkboxState !== null) {
          cellContentStyle.alignItems = "center";
          cellContentStyle.justifyContent = "center";
        }
        if (cellData.textRotationDeg) {
          cellContentStyle.alignItems = resolveCellContentAlignItems(cellStyle.textAlign);
          cellContentStyle.overflow = "visible";
        }
        if (cellData.shrinkToFitFontSizePx) {
          cellContentStyle.fontSize = cellData.shrinkToFitFontSizePx;
          cellContentStyle.lineHeight = `${resolveCanvasLineHeight(cellData.style, cellData.shrinkToFitFontSizePx)}px`;
        }
        const rotatedTextStyle = cellData.textRotationDeg
          ? {
              display: "inline-block",
              maxWidth: "none",
              transform: `rotate(${cellData.textRotationDeg}deg)`,
              transformOrigin: "center center",
              whiteSpace: "nowrap",
              width: "max-content"
            } satisfies React.CSSProperties
          : null;
        const title = [cellData.hyperlink?.tooltip, cellData.validation?.message, cellData.value]
          .filter((value, index, values): value is string => typeof value === "string" && value.length > 0 && values.indexOf(value) === index)
          .join("\n");

        return (
          <td
            data-xlsx-cell={`${actualRow}:${actualCol}`}
            key={key}
            colSpan={cellData.colSpan}
            rowSpan={cellData.rowSpan}
            onDoubleClick={() => {
              if (readOnly) {
                return;
              }

              onCellDoubleClick(cell);
            }}
            onClick={() => onCellClick(cell, cellData)}
            onPointerDown={(event) => onCellPointerDown(event, cell)}
            style={cellStyle}
            title={title}
          >
            {cellData.chartHighlight ? (
              <div
                aria-hidden="true"
                style={{
                  backgroundColor: cellData.chartHighlight.fillColor,
                  borderBottom: cellData.chartHighlight.borderBottom ? `${Math.max(1, zoomFactor)}px solid ${cellData.chartHighlight.strokeColor}` : undefined,
                  borderLeft: cellData.chartHighlight.borderLeft ? `${Math.max(1, zoomFactor)}px solid ${cellData.chartHighlight.strokeColor}` : undefined,
                  borderRight: cellData.chartHighlight.borderRight ? `${Math.max(1, zoomFactor)}px solid ${cellData.chartHighlight.strokeColor}` : undefined,
                  borderTop: cellData.chartHighlight.borderTop ? `${Math.max(1, zoomFactor)}px solid ${cellData.chartHighlight.strokeColor}` : undefined,
                  boxSizing: "border-box",
                  inset: Math.max(1, zoomFactor * 0.5),
                  pointerEvents: "none",
                  position: "absolute"
                }}
              />
            ) : null}
            {cellData.conditionalDataBar ? (
              <div
                aria-hidden="true"
                style={{
                  alignItems: "center",
                  bottom: 4 * zoomFactor,
                  display: "flex",
                  left: 4 * zoomFactor,
                  pointerEvents: "none",
                  position: "absolute",
                  right: 4 * zoomFactor,
                  top: 4 * zoomFactor,
                  zIndex: 0
                }}
              >
                <div
                  style={{
                    background: cellData.conditionalDataBar.gradient === false
                      ? cellData.conditionalDataBar.color
                      : `linear-gradient(90deg, ${lightenColor(cellData.conditionalDataBar.color, 0.28)} 0%, ${cellData.conditionalDataBar.color} 100%)`,
                    border: cellData.conditionalDataBar.border !== false && cellData.conditionalDataBar.borderColor
                      ? `1px solid ${cellData.conditionalDataBar.borderColor}`
                      : "none",
                    borderRadius: 2,
                    height: "100%",
                    opacity: 1,
                    width: `${cellData.conditionalDataBar.widthPercent}%`
                  }}
                />
              </div>
            ) : null}
            {adornment}
            {cellData.conditionalIcon ? (
              <div
                aria-hidden="true"
                style={{
                  alignItems: "center",
                  display: "inline-flex",
                  justifyContent: "center",
                  pointerEvents: "none",
                  position: "absolute",
                  right: conditionalIconRight * zoomFactor,
                  top: "50%",
                  transform: "translateY(-50%)",
                  zIndex: 2
                }}
              >
                {renderConditionalIcon(cellData.conditionalIcon, zoomFactor)}
              </div>
            ) : null}
            {isEditing ? (
              <input
                autoFocus
                onBlur={onEditingBlur}
                onChange={(event) => onEditingValueChange(event.target.value)}
                onKeyDown={(event) => {
                  event.stopPropagation();
                  if (event.key === "Enter") {
                    event.preventDefault();
                    onEditingCommit();
                    return;
                  }

                  if (event.key === "Escape") {
                    event.preventDefault();
                    onEditingCancel();
                    return;
                  }

                  if (
                    event.key === "ArrowDown" ||
                    event.key === "ArrowLeft" ||
                    event.key === "ArrowRight" ||
                    event.key === "ArrowUp"
                  ) {
                    event.preventDefault();
                    onEditingNavigate(event.key, event.currentTarget.value);
                  }
                }}
                style={{
                  backgroundColor: "transparent",
                  border: 0,
                  color: "inherit",
                  font: "inherit",
                  height: "100%",
                  margin: 0,
                  outline: "none",
                  padding: scaleCssLengthExpression(DEFAULT_CELL_PADDING, zoomFactor),
                  width: "100%"
                }}
                value={editingValue}
              />
            ) : isSpilling ? (
              <div
                style={{
                  maxWidth: cellData.spillWidth,
                  overflow: "visible",
                  pointerEvents: "none",
                  position: "relative",
                  whiteSpace: "inherit",
                  width: cellData.spillWidth
                }}
              >
                {cellData.value}
              </div>
            ) : cellData.sparkline ? (
              <div
                style={{
                  alignItems: "center",
                  display: "flex",
                  height: "100%",
                  justifyContent: "center",
                  pointerEvents: "none",
                  width: "100%"
                }}
              >
                {renderSparkline(cellData.sparkline.config, cellData.sparkline.values, palette)}
              </div>
            ) : cellData.checkboxState != null ? (
              <div
                style={{
                  alignItems: "center",
                  display: "flex",
                  height: "100%",
                  justifyContent: "center",
                  pointerEvents: "none",
                  width: "100%"
                }}
              >
                {renderCheckboxControl(cellData.checkboxState, palette, zoomFactor)}
              </div>
            ) : (
              <div style={cellContentStyle}>
                {rotatedTextStyle ? <span style={rotatedTextStyle}>{cellData.value}</span> : cellData.value}
              </div>
            )}
          </td>
        );
      })}
      {trailingSpacerWidth > 0 ? (
        <td
          aria-hidden="true"
          style={{
            backgroundColor: "transparent",
            border: "none",
            padding: 0,
            width: trailingSpacerWidth
          }}
        />
      ) : null}
    </tr>
  );
}

const MemoGridRow = React.memo(GridRow, (prev, next) => {
  if (
    prev.actualRow !== next.actualRow ||
    prev.rowHeight !== next.rowHeight ||
    prev.headerLabelLiveScale !== next.headerLabelLiveScale ||
    prev.palette !== next.palette ||
    prev.readOnly !== next.readOnly ||
    prev.visibleCols !== next.visibleCols ||
    prev.leadingSpacerWidth !== next.leadingSpacerWidth ||
    prev.rowHeaderWidth !== next.rowHeaderWidth ||
    prev.stickyLeftByCol !== next.stickyLeftByCol ||
    prev.stickyTop !== next.stickyTop ||
    prev.trailingSpacerWidth !== next.trailingSpacerWidth ||
    prev.zoomFactor !== next.zoomFactor ||
    prev.getCellData !== next.getCellData ||
    prev.onCellClick !== next.onCellClick ||
    prev.onCellDoubleClick !== next.onCellDoubleClick ||
    prev.onCellPointerDown !== next.onCellPointerDown ||
    prev.onEditingBlur !== next.onEditingBlur ||
    prev.onEditingCancel !== next.onEditingCancel ||
    prev.onEditingCommit !== next.onEditingCommit ||
    prev.onEditingNavigate !== next.onEditingNavigate ||
    prev.onEditingValueChange !== next.onEditingValueChange ||
    prev.onRowHeaderRef !== next.onRowHeaderRef ||
    prev.onRowPointerDown !== next.onRowPointerDown ||
    prev.onRowResizePointerDown !== next.onRowResizePointerDown ||
    prev.renderCellAdornment !== next.renderCellAdornment
  ) {
    return false;
  }

  const prevEditingCol = prev.editingCell?.row === prev.actualRow ? prev.editingCell.col : -1;
  const nextEditingCol = next.editingCell?.row === next.actualRow ? next.editingCell.col : -1;
  if (prevEditingCol !== nextEditingCol) {
    return false;
  }

  if (prevEditingCol !== -1 && prev.editingValue !== next.editingValue) {
    return false;
  }

  return true;
});

function XlsxGrid({
  allowResizeInReadOnly = false,
  controller,
  emptyState,
  enableCanvasSelectionAnimation = true,
  errorState,
  fileTooLargeState,
  getCellStyle,
  loadingComponent,
  loadingState,
  renderChartLoading,
  palette,
  renderImage,
  renderImageSelection,
  renderTableHeaderMenu,
  renderScroller,
  enableGestureZoom = true,
  experimentalCanvas = true,
  selectionColor,
  selectionFillColor,
  selectionHeaderColor,
  showImages = true
}: Pick<
  XlsxViewerProps,
  "allowResizeInReadOnly" | "emptyState" | "enableCanvasSelectionAnimation" | "enableGestureZoom" | "errorState" | "experimentalCanvas" | "fileTooLargeState" | "getCellStyle" | "loadingComponent" | "loadingState" | "renderChartLoading" | "renderImage" | "renderImageSelection" | "renderScroller" | "renderTableHeaderMenu" | "selectionColor" | "selectionFillColor" | "selectionHeaderColor" | "showImages"
> & {
  controller: XlsxViewerController;
  palette: ViewerPalette;
}) {
  const xlsxCanvasProfileTarget = typeof window !== "undefined"
    ? (window as unknown as {
        __xlsxCanvasProfile?: Array<Record<string, number | boolean>>;
      }).__xlsxCanvasProfile
    : undefined;
  const xlsxGridRenderStart = xlsxCanvasProfileTarget ? performance.now() : 0;
  const {
    activeCell,
    activeSheet,
    activeSheetIndex,
    activeTab,
    activeTabIndex,
    canLoadDeferred,
    charts,
    clearSelectedChart,
    clearSelectedImage,
    clearSelectedCells,
    continueDeferredLoad,
    deferredLoadFileSize,
    displayFileName,
    error,
    fillSelection,
    formControls,
    getActiveWorksheet,
    getChartById,
    getRowsBatchAsync,
    getClipboardData,
    getCellDisplayValue: getControllerCellDisplayValue,
    images,
    shapes,
    isLoadDeferred,
    isLoading,
    isChartsLoading,
    isWorkerBacked,
    maxZoomScale,
    minZoomScale,
    copySelectionToClipboard,
    pasteFromClipboard,
    pasteStructuredClipboardData,
    pasteText,
    readOnly,
    redo,
    revision,
    selectedChart,
    selectedChartElement,
    selectedChartId,
    selectedImage,
    selectedImageId,
    selectCell,
    selectChart,
    selectChartElement,
    selectImage,
    selectRange,
    selection,
    setActiveSheetIndex,
    setChartRect,
    setImageRect,
    setCellValue,
    setZoomScale,
    sheets,
    tabs,
    sortState,
    sortTable,
    tables,
    undo,
    workbook,
    zoomScale
  } = controller;
  const canResizeHeaders = !readOnly || allowResizeInReadOnly;
  const scrollRef = React.useRef<HTMLDivElement>(null);
  const wrapperRef = React.useRef<HTMLDivElement>(null);
  const tableRef = React.useRef<HTMLTableElement>(null);
  const gridKeyboardActiveRef = React.useRef(false);
  const gridKeyboardHandlerRef = React.useRef<((event: KeyboardEvent) => void) | null>(null);
  const axisSelectionRef = React.useRef<
    | { axis: "column"; endCol: number; startCol: number }
    | { axis: "row"; endRow: number; startRow: number }
    | null
  >(null);
  const scrollBodyCanvasRef = React.useRef<HTMLCanvasElement>(null);
  const topBodyCanvasRef = React.useRef<HTMLCanvasElement>(null);
  const leftBodyCanvasRef = React.useRef<HTMLCanvasElement>(null);
  const cornerBodyCanvasRef = React.useRef<HTMLCanvasElement>(null);
  const bodyBlitBufferCanvasRef = React.useRef<Record<FrozenDrawingPane, HTMLCanvasElement | null>>({
    corner: null,
    left: null,
    scroll: null,
    top: null
  });
  const headerBlitBufferCanvasRef = React.useRef<{ left: HTMLCanvasElement | null; top: HTMLCanvasElement | null }>({
    left: null,
    top: null
  });
  const topFrozenHeaderCanvasRef = React.useRef<HTMLCanvasElement>(null);
  const topScrollHeaderCanvasRef = React.useRef<HTMLCanvasElement>(null);
  const leftFrozenHeaderCanvasRef = React.useRef<HTMLCanvasElement>(null);
  const leftScrollHeaderCanvasRef = React.useRef<HTMLCanvasElement>(null);
  const cornerHeaderCanvasRef = React.useRef<HTMLCanvasElement>(null);
  const canvasScrollOverlayContentRef = React.useRef<HTMLDivElement>(null);
  const canvasTopOverlayContentRef = React.useRef<HTMLDivElement>(null);
  const canvasLeftOverlayContentRef = React.useRef<HTMLDivElement>(null);
  const canvasCornerOverlayContentRef = React.useRef<HTMLDivElement>(null);
  const canvasImageCacheRef = React.useRef(new Map<string, {
    failed: boolean;
    image: HTMLImageElement;
    loaded: boolean;
    src: string;
  }>());
  const selectionOverlayRef = React.useRef<HTMLDivElement>(null);
  const activeValidationOverlayRef = React.useRef<HTMLDivElement>(null);
  const fillHandleRef = React.useRef<HTMLDivElement>(null);
  const tableMenuRef = React.useRef<HTMLDivElement>(null);
  const colElementRefs = React.useRef(new Map<number, HTMLTableColElement>());
  const colHeaderCellRefs = React.useRef(new Map<number, HTMLTableCellElement>());
  const rowElementRefs = React.useRef(new Map<number, HTMLTableRowElement>());
  const rowHeaderCellRefs = React.useRef(new Map<number, HTMLTableCellElement>());
  const columnPreviewRef = React.useRef<{ actualIndex: number; size: number } | null>(null);
  const rowPreviewRef = React.useRef<{ actualIndex: number; size: number } | null>(null);
  const activeCellRef = React.useRef<XlsxCellAddress | null>(activeCell);
  const selectedChartIdRef = React.useRef<string | null>(selectedChartId);
  const selectedImageIdRef = React.useRef<string | null>(selectedImageId);
  const pendingSelectionCommitRef = React.useRef<XlsxCellRange | null>(null);
  const selectionCommitFrameRef = React.useRef<number | null>(null);
  const selectionRef = React.useRef<XlsxCellRange | null>(null);
  const editingCellRef = React.useRef<XlsxCellAddress | null>(null);
  const commitEditingRef = React.useRef<() => void>(() => {});
  const editingInputRef = React.useRef<HTMLInputElement>(null);
  const readOnlyRef = React.useRef(readOnly);
  const committedZoomScaleRef = React.useRef(zoomScale);
  const gestureZoomScaleRef = React.useRef(zoomScale);
  const liveGestureZoomRef = React.useRef<LiveGestureZoomState | null>(null);
  const pendingLiveGestureZoomStateRef = React.useRef<LiveGestureZoomState | null>(null);
  const liveGestureZoomFrameRef = React.useRef<number | null>(null);
  const pendingZoomAnchorRef = React.useRef<ZoomAnchor | null>(null);
  const liveZoomCommitTimerRef = React.useRef<number | null>(null);
  const pendingLiveZoomCommitRef = React.useRef<number | null>(null);
  const touchPinchStateRef = React.useRef<{ startDistance: number; startZoomScale: number } | null>(null);
  const safariPinchStartZoomRef = React.useRef<number | null>(null);
  const canvasTransformStyleCacheRef = React.useRef(new WeakMap<HTMLElement, {
    transform: string;
    transformOrigin: string | null;
    willChange: string;
  }>());
  const displayedSelectionRef = React.useRef<XlsxCellRange | null>(null);
  const firstVisibleColRef = React.useRef<number | undefined>(undefined);
  const lastVisibleColRef = React.useRef<number | undefined>(undefined);
  const firstVisibleRowRef = React.useRef<number | undefined>(undefined);
  const lastVisibleRowRef = React.useRef<number | undefined>(undefined);
  const cellRenderCacheRef = React.useRef(new Map<string, CellRenderData>());
  const conditionalFormatMetricsCacheRef = React.useRef(new Map<string, number[]>());
  const resizeStateRef = React.useRef<
    | {
        actualIndex: number;
        initialPx: number;
        pointerId: number;
        startPosition: number;
        type: "column" | "row";
      }
    | null
  >(null);
  const resizeFrameRef = React.useRef<number | null>(null);
  const pendingResizePreviewRef = React.useRef<
    | {
        actualIndex: number;
        size: number;
        type: "column" | "row";
      }
    | null
  >(null);
  const selectionDragRef = React.useRef<
    SelectionDragState | null
  >(null);
  const pendingSelectionDragRef = React.useRef<SelectionDragState | null>(null);
  const fillDragRef = React.useRef<
    | {
        pointerId: number;
        previewRange: XlsxCellRange;
        sourceRange: XlsxCellRange;
      }
    | null
  >(null);
  const selectionDragCleanupRef = React.useRef<(() => void) | null>(null);
  const pendingSelectionDragCleanupRef = React.useRef<(() => void) | null>(null);
  const fillDragCleanupRef = React.useRef<(() => void) | null>(null);
  const cachedScrollerRectRef = React.useRef<DOMRect | null>(null);
  const imageInteractionCleanupRef = React.useRef<(() => void) | null>(null);
  const imageInteractionRef = React.useRef<
    | {
        baseRect: XlsxImageRect;
        didMove: boolean;
        imageId: string;
        pointerId: number;
        startClientX: number;
        startClientY: number;
        type: "move";
      }
    | {
        baseRect: XlsxImageRect;
        didMove: boolean;
        handle: XlsxImageResizeHandlePosition;
        imageId: string;
        pointerId: number;
        startClientX: number;
        startClientY: number;
        type: "resize";
      }
    | null
  >(null);
  const chartInteractionCleanupRef = React.useRef<(() => void) | null>(null);
  const chartInteractionRef = React.useRef<
    | {
        baseRect: XlsxImageRect;
        didMove: boolean;
        chartId: string;
        pointerId: number;
        startClientX: number;
        startClientY: number;
        type: "move";
      }
    | {
        baseRect: XlsxImageRect;
        didMove: boolean;
        chartId: string;
        handle: XlsxImageResizeHandlePosition;
        pointerId: number;
        startClientX: number;
        startClientY: number;
        type: "resize";
      }
    | null
  >(null);
  const [editingCell, setEditingCell] = React.useState<XlsxCellAddress | null>(null);
  const [editingValue, setEditingValue] = React.useState("");
  const [openTableMenu, setOpenTableMenu] = React.useState<{ col: number; row: number; tableName: string } | null>(null);
  const [fillPreviewRange, setFillPreviewRange] = React.useState<XlsxCellRange | null>(null);
  const [chartPreviewRect, setChartPreviewRect] = React.useState<{ id: string; rect: XlsxImageRect } | null>(null);
  const [liveGestureZoom, setLiveGestureZoom] = React.useState<LiveGestureZoomState | null>(null);
  const [canvasImageLoadVersion, setCanvasImageLoadVersion] = React.useState(0);
  const liveDrawingViewportRef = React.useRef<DrawingViewport>({
    height: 0,
    left: 0,
    top: 0,
    width: 0
  });
  const canvasLayoutMetricsRef = React.useRef({
    displayHeaderHeight: HEADER_HEIGHT,
    displayRowHeaderWidth: ROW_HEADER_WIDTH,
    frozenPaneBottom: HEADER_HEIGHT,
    frozenPaneRight: ROW_HEADER_WIDTH
  });
  const paintedDrawingViewportRef = React.useRef<DrawingViewport>({
    height: 0,
    left: 0,
    top: 0,
    width: 0
  });
  const paintedBodyCanvasSignatureRef = React.useRef<PaintedBodyCanvasSignature | null>(null);
  const paintedHeaderCanvasSignatureRef = React.useRef<PaintedHeaderCanvasSignature | null>(null);
  const pendingDrawingViewportRef = React.useRef<DrawingViewport | null>(null);
  const drawingViewportFrameRef = React.useRef<number | null>(null);
  const initialScrollKeyRef = React.useRef<string | null>(null);
  const chartPreviewRectRef = React.useRef<{ id: string; rect: XlsxImageRect } | null>(null);
  const skipNextChartClickRef = React.useRef<string | null>(null);
  const paneDrawingNodesCacheRef = React.useRef<{
    chartRects: Array<{ chart: XlsxChart; rect: XlsxImageRect }>;
    drawingViewportSignature: string;
    formControlRects: Array<{ control: XlsxFormControl; rect: XlsxImageRect }>;
    imageRects: Array<{ image: XlsxImage; rect: XlsxImageRect }>;
    palette: ViewerPalette;
    readOnly: boolean;
    renderChartLoading: XlsxViewerProps["renderChartLoading"];
    renderImage: XlsxViewerProps["renderImage"];
    renderImageSelection: XlsxViewerProps["renderImageSelection"];
      isChartsLoading: boolean;
      selectedChartElement: XlsxChartElementSelection | null;
      selectedChartId: string | null;
    selectedImageId: string | null;
    selectionStroke: string;
    shapeRects: Array<{ rect: XlsxImageRect; shape: XlsxShape }>;
    showImages: boolean;
    value: Record<FrozenDrawingPane, React.ReactNode>;
  } | null>(null);
  const [drawingViewport, setDrawingViewport] = React.useState<DrawingViewport>({
    height: 0,
    left: 0,
    top: 0,
    width: 0
  });
  const drawingViewportStateRef = React.useRef<DrawingViewport>({
    height: 0,
    left: 0,
    top: 0,
    width: 0
  });
  const [resizeGuide, setResizeGuide] = React.useState<{ position: number; type: "column" | "row" } | null>(null);
  const selectionPreviewRangeRef = React.useRef<XlsxCellRange | null>(null);
  const [imagePreviewRect, setImagePreviewRect] = React.useState<{ id: string; rect: XlsxImageRect } | null>(null);
  const imagePreviewRectRef = React.useRef<{ id: string; rect: XlsxImageRect } | null>(null);
  const imagePreviewFrameRef = React.useRef<number | null>(null);
  const pendingImagePreviewRef = React.useRef<{ id: string; rect: XlsxImageRect } | null>(null);
  const chartPreviewFrameRef = React.useRef<number | null>(null);
  const pendingChartPreviewRef = React.useRef<{ id: string; rect: XlsxImageRect } | null>(null);
  const skipNextImageClickRef = React.useRef<string | null>(null);
  const [pendingNavigation, setPendingNavigation] = React.useState<{ cell: XlsxCellAddress; sheetIndex: number } | null>(null);
  const interactionModeRef = React.useRef<"idle" | "fill" | "select">("idle");
  const setInteractionMode = React.useCallback((mode: "idle" | "fill" | "select") => {
    interactionModeRef.current = mode;
  }, []);
  const worksheet = getActiveWorksheet();
  const mergedRegions = React.useMemo<XlsxCellRange[]>(() => {
    const regions: XlsxCellRange[] = [];
    const rawMergedRegions = Array.isArray(worksheet?.mergedRegions) ? worksheet.mergedRegions : [];

    for (const entry of rawMergedRegions) {
      const region = asRecord(entry);
      const startRow = asNonNegativeInteger(region?.startRow);
      const startCol = asNonNegativeInteger(region?.startCol);
      const endRow = asNonNegativeInteger(region?.endRow);
      const endCol = asNonNegativeInteger(region?.endCol);
      if (startRow === null || startCol === null || endRow === null || endCol === null) {
        continue;
      }

      regions.push({
        end: { col: endCol, row: endRow },
        start: { col: startCol, row: startRow }
      });
    }

    return regions;
  }, [worksheet]);
  const mergedSecondaryAnchorMap = React.useMemo(() => {
    const map = new Map<string, XlsxCellAddress>();

    for (const region of mergedRegions) {
      for (let mergeRow: number = region.start.row; mergeRow <= region.end.row; mergeRow += 1) {
        for (let mergeCol: number = region.start.col; mergeCol <= region.end.col; mergeCol += 1) {
          if (mergeRow === region.start.row && mergeCol === region.start.col) {
            continue;
          }
          map.set(`${mergeRow}:${mergeCol}`, { row: region.start.row, col: region.start.col });
        }
      }
    }

    return map;
  }, [mergedRegions]);
  const resolveMergeAnchorCell = React.useCallback((cell: XlsxCellAddress): XlsxCellAddress => {
    return mergedSecondaryAnchorMap.get(`${cell.row}:${cell.col}`) ?? cell;
  }, [mergedSecondaryAnchorMap]);
  const normalizedSelection = React.useMemo(() => (selection ? normalizeRange(selection) : null), [selection]);
  const zoomFactor = React.useMemo(() => Math.max(0.1, zoomScale / 100), [zoomScale]);
  const previousZoomFactorRef = React.useRef(zoomFactor);
  const effectiveTables = tables;
  const [displayRowLimit, setDisplayRowLimit] = React.useState(() =>
    resolveInitialDisplayExtent(
      activeSheet?.maxUsedRow ?? -1,
      MIN_OPEN_GRID_ROWS,
      OPEN_GRID_ROW_PADDING,
      Boolean(isWorkerBacked),
      INITIAL_WORKER_GRID_ROWS
    )
  );
  const [displayColLimit, setDisplayColLimit] = React.useState(() =>
    resolveInitialDisplayExtent(
      activeSheet?.maxUsedCol ?? -1,
      MIN_OPEN_GRID_COLS,
      OPEN_GRID_COL_PADDING,
      Boolean(isWorkerBacked),
      INITIAL_WORKER_GRID_COLS
    )
  );
  const hiddenRowSet = React.useMemo(() => new Set(activeSheet?.hiddenRows ?? []), [activeSheet?.hiddenRows]);
  const hiddenColSet = React.useMemo(() => new Set(activeSheet?.hiddenCols ?? []), [activeSheet?.hiddenCols]);
  const displayDefaultRowHeight = DEFAULT_ROW_HEIGHT * zoomFactor;
  const displayDefaultColWidth = DEFAULT_COL_WIDTH * zoomFactor;
  const displayHeaderHeight = HEADER_HEIGHT * zoomFactor;
  const displayRowHeaderWidth = ROW_HEADER_WIDTH * zoomFactor;
  const displayImageMinSize = IMAGE_MIN_SIZE_PX * zoomFactor;
  const applyCanvasViewportCompensation = React.useCallback((liveViewport?: DrawingViewport) => {
    const nextLiveViewport = liveViewport ?? liveDrawingViewportRef.current;
    const paintedViewport = paintedDrawingViewportRef.current;
    const {
      displayHeaderHeight: liveDisplayHeaderHeight,
      displayRowHeaderWidth: liveDisplayRowHeaderWidth,
      frozenPaneBottom: liveFrozenPaneBottom,
      frozenPaneRight: liveFrozenPaneRight
    } = canvasLayoutMetricsRef.current;
    const currentLiveGestureZoom = liveGestureZoomRef.current;
    const isLiveZooming = currentLiveGestureZoom !== null && zoomScale === currentLiveGestureZoom.baseZoomScale;
    const liveZoomScale = isLiveZooming
      ? Math.max(0.1, currentLiveGestureZoom.targetZoomScale / currentLiveGestureZoom.baseZoomScale)
      : 1;
    const scrollDeltaX = paintedViewport.left - nextLiveViewport.left;
    const scrollDeltaY = paintedViewport.top - nextLiveViewport.top;
    const liveZoomTranslateX = isLiveZooming
      ? currentLiveGestureZoom.anchor.x * (1 - liveZoomScale)
      : 0;
    const liveZoomTranslateY = isLiveZooming
      ? currentLiveGestureZoom.anchor.y * (1 - liveZoomScale)
      : 0;

    const applyElementTransform = (
      element: HTMLElement | null,
      transform: string,
      willChange: string,
      transformOrigin: string | null = null
    ) => {
      if (!element) {
        return;
      }

      const cached = canvasTransformStyleCacheRef.current.get(element);
      if (cached?.transform !== transform || element.style.transform !== transform) {
        element.style.transform = transform;
      }
      if (cached?.willChange !== willChange || element.style.willChange !== willChange) {
        element.style.willChange = willChange;
      }
      if (
        transformOrigin !== null
        && (cached?.transformOrigin !== transformOrigin || element.style.transformOrigin !== transformOrigin)
      ) {
        element.style.transformOrigin = transformOrigin;
      }
      canvasTransformStyleCacheRef.current.set(element, {
        transform,
        transformOrigin,
        willChange
      });
    };
    const applyCanvasTransform = (
      canvas: HTMLCanvasElement | null,
      translateX: number,
      translateY: number
    ) => {
      const shouldTransform = translateX !== 0 || translateY !== 0 || liveZoomScale !== 1;
      applyElementTransform(
        canvas,
        shouldTransform ? `translate3d(${translateX}px, ${translateY}px, 0) scale(${liveZoomScale})` : "",
        shouldTransform ? "transform" : ""
      );
    };
    const applyOverlayTransform = (
      element: HTMLDivElement | null,
      translateX: number,
      translateY: number
    ) => {
      applyElementTransform(
        element,
        liveZoomScale !== 1
          ? `translate3d(${translateX + liveZoomTranslateX}px, ${translateY + liveZoomTranslateY}px, 0) scale(${liveZoomScale})`
          : `translate3d(${translateX}px, ${translateY}px, 0)`,
        "transform",
        "0 0"
      );
    };

    applyCanvasTransform(
      scrollBodyCanvasRef.current,
      scrollDeltaX + liveZoomTranslateX,
      scrollDeltaY + liveZoomTranslateY
    );
    applyCanvasTransform(
      topBodyCanvasRef.current,
      scrollDeltaX + liveZoomTranslateX,
      liveZoomTranslateY
    );
    applyCanvasTransform(
      leftBodyCanvasRef.current,
      liveZoomTranslateX,
      scrollDeltaY + liveZoomTranslateY
    );
    applyCanvasTransform(
      cornerBodyCanvasRef.current,
      liveZoomTranslateX,
      liveZoomTranslateY
    );

    applyCanvasTransform(
      topScrollHeaderCanvasRef.current,
      scrollDeltaX + liveZoomTranslateX,
      liveZoomTranslateY
    );
    applyCanvasTransform(
      topFrozenHeaderCanvasRef.current,
      liveZoomTranslateX,
      liveZoomTranslateY
    );
    applyCanvasTransform(
      leftScrollHeaderCanvasRef.current,
      liveZoomTranslateX,
      scrollDeltaY + liveZoomTranslateY
    );
    applyCanvasTransform(
      leftFrozenHeaderCanvasRef.current,
      liveZoomTranslateX,
      liveZoomTranslateY
    );
    applyCanvasTransform(
      cornerHeaderCanvasRef.current,
      liveZoomTranslateX,
      liveZoomTranslateY
    );
    applyOverlayTransform(
      canvasScrollOverlayContentRef.current,
      -nextLiveViewport.left - liveFrozenPaneRight,
      -nextLiveViewport.top - liveFrozenPaneBottom
    );
    applyOverlayTransform(
      canvasTopOverlayContentRef.current,
      -nextLiveViewport.left - liveFrozenPaneRight,
      -liveDisplayHeaderHeight
    );
    applyOverlayTransform(
      canvasLeftOverlayContentRef.current,
      -liveDisplayRowHeaderWidth,
      -nextLiveViewport.top - liveFrozenPaneBottom
    );
    applyOverlayTransform(
      canvasCornerOverlayContentRef.current,
      -liveDisplayRowHeaderWidth,
      -liveDisplayHeaderHeight
    );
  }, [zoomScale]);
  const updateLiveGestureZoomState = React.useCallback((
    nextState:
      | LiveGestureZoomState
      | null
      | ((current: LiveGestureZoomState | null) => LiveGestureZoomState | null)
  ) => {
    const resolvedState = typeof nextState === "function"
      ? nextState(liveGestureZoomRef.current)
      : nextState;
    liveGestureZoomRef.current = resolvedState;
    pendingLiveGestureZoomStateRef.current = resolvedState;
    applyCanvasViewportCompensation();
    if (liveGestureZoomFrameRef.current !== null) {
      return;
    }
    liveGestureZoomFrameRef.current = window.requestAnimationFrame(() => {
      liveGestureZoomFrameRef.current = null;
      const pendingState = pendingLiveGestureZoomStateRef.current;
      pendingLiveGestureZoomStateRef.current = null;
      setLiveGestureZoom(pendingState);
    });
  }, [applyCanvasViewportCompensation]);
  const clearLiveZoomCommitTimer = React.useCallback(() => {
    if (liveZoomCommitTimerRef.current !== null) {
      window.clearTimeout(liveZoomCommitTimerRef.current);
      liveZoomCommitTimerRef.current = null;
    }
  }, []);
  const clearLiveGestureZoom = React.useCallback(() => {
    clearLiveZoomCommitTimer();
    pendingLiveZoomCommitRef.current = null;
    pendingZoomAnchorRef.current = null;
    touchPinchStateRef.current = null;
    safariPinchStartZoomRef.current = null;
    gestureZoomScaleRef.current = committedZoomScaleRef.current;
    updateLiveGestureZoomState(null);
  }, [clearLiveZoomCommitTimer, updateLiveGestureZoomState]);
  const updateLiveGestureZoomTarget = React.useCallback((nextZoomScale: number, anchor: ZoomAnchor) => {
    const clampedZoomScale = clampValue(nextZoomScale, minZoomScale, maxZoomScale);
    gestureZoomScaleRef.current = clampedZoomScale;
    pendingLiveZoomCommitRef.current = null;
    updateLiveGestureZoomState((current) => ({
      anchor,
      baseZoomScale: current?.baseZoomScale ?? committedZoomScaleRef.current,
      targetZoomScale: clampedZoomScale
    }));
  }, [maxZoomScale, minZoomScale, updateLiveGestureZoomState]);
  const commitLiveGestureZoom = React.useCallback(() => {
    clearLiveZoomCommitTimer();
    const currentLiveGestureZoom = liveGestureZoomRef.current;
    if (!currentLiveGestureZoom) {
      return;
    }

    const nextCommittedZoomScale = clampValue(
      Math.round(currentLiveGestureZoom.targetZoomScale),
      minZoomScale,
      maxZoomScale
    );
    if (nextCommittedZoomScale === committedZoomScaleRef.current) {
      clearLiveGestureZoom();
      return;
    }

    if (pendingLiveZoomCommitRef.current === nextCommittedZoomScale) {
      return;
    }

    pendingLiveZoomCommitRef.current = nextCommittedZoomScale;
    pendingZoomAnchorRef.current = currentLiveGestureZoom.anchor;
    gestureZoomScaleRef.current = nextCommittedZoomScale;
    setZoomScale(nextCommittedZoomScale);
  }, [clearLiveGestureZoom, clearLiveZoomCommitTimer, maxZoomScale, minZoomScale, setZoomScale]);
  const scheduleLiveGestureZoomCommit = React.useCallback(() => {
    clearLiveZoomCommitTimer();
    liveZoomCommitTimerRef.current = window.setTimeout(() => {
      liveZoomCommitTimerRef.current = null;
      commitLiveGestureZoom();
    }, LIVE_ZOOM_COMMIT_IDLE_MS);
  }, [clearLiveZoomCommitTimer, commitLiveGestureZoom]);
  const syncDrawingViewport = React.useCallback((scroller: HTMLDivElement | null, options?: { immediate?: boolean }) => {
    if (!scroller) {
      return;
    }

    const nextViewport = {
      height: scroller.clientHeight,
      left: scroller.scrollLeft,
      top: scroller.scrollTop,
      width: scroller.clientWidth
    };
    const stateViewport = drawingViewportStateRef.current;
    const matchesStateViewport =
      stateViewport.left === nextViewport.left
      && stateViewport.top === nextViewport.top
      && stateViewport.width === nextViewport.width
      && stateViewport.height === nextViewport.height;
    const liveViewport = liveDrawingViewportRef.current;
    const matchesLiveViewport =
      liveViewport.left === nextViewport.left
      && liveViewport.top === nextViewport.top
      && liveViewport.width === nextViewport.width
      && liveViewport.height === nextViewport.height;
    liveDrawingViewportRef.current = nextViewport;
    applyCanvasViewportCompensation(nextViewport);

    if (options?.immediate) {
      pendingDrawingViewportRef.current = null;
      if (drawingViewportFrameRef.current !== null) {
        window.cancelAnimationFrame(drawingViewportFrameRef.current);
        drawingViewportFrameRef.current = null;
      }
      if (matchesStateViewport) {
        return;
      }
      setDrawingViewport((current) => (
        current.left === nextViewport.left
        && current.top === nextViewport.top
        && current.width === nextViewport.width
        && current.height === nextViewport.height
          ? current
          : nextViewport
      ));
      return;
    }

    if (matchesLiveViewport && matchesStateViewport) {
      return;
    }

    const shouldCommitDeferredViewport =
      !matchesStateViewport
      && (
        stateViewport.width !== nextViewport.width
        || stateViewport.height !== nextViewport.height
        || Math.abs(nextViewport.left - stateViewport.left) >= CANVAS_DEFERRED_VIEWPORT_SYNC_THRESHOLD_PX
        || Math.abs(nextViewport.top - stateViewport.top) >= CANVAS_DEFERRED_VIEWPORT_SYNC_THRESHOLD_PX
      );
    if (!shouldCommitDeferredViewport) {
      pendingDrawingViewportRef.current = null;
      return;
    }

    pendingDrawingViewportRef.current = nextViewport;
    if (drawingViewportFrameRef.current !== null) {
      return;
    }
    drawingViewportFrameRef.current = window.requestAnimationFrame(() => {
      drawingViewportFrameRef.current = null;
      const pendingViewport = pendingDrawingViewportRef.current;
      pendingDrawingViewportRef.current = null;
      if (!pendingViewport) {
        return;
      }
      setDrawingViewport((current) => (
        current.left === pendingViewport.left
        && current.top === pendingViewport.top
        && current.width === pendingViewport.width
        && current.height === pendingViewport.height
          ? current
          : pendingViewport
      ));
    });
  }, [applyCanvasViewportCompensation]);
  const setGlobalCursor = React.useCallback((cursor: string) => {
    document.body.style.cursor = cursor;
    document.documentElement.style.cursor = cursor;
    if (scrollRef.current) {
      scrollRef.current.style.cursor = cursor;
    }
  }, []);
  const clearGlobalCursor = React.useCallback(() => {
    document.body.style.cursor = "";
    document.documentElement.style.cursor = "";
    if (scrollRef.current) {
      scrollRef.current.style.cursor = "";
    }
  }, []);
  const getBodyBlitBufferCanvas = React.useCallback((pane: FrozenDrawingPane) => {
    let bufferCanvas = bodyBlitBufferCanvasRef.current[pane];
    if (!bufferCanvas && typeof document !== "undefined") {
      bufferCanvas = document.createElement("canvas");
      bodyBlitBufferCanvasRef.current[pane] = bufferCanvas;
    }
    return bufferCanvas;
  }, []);
  const getHeaderBlitBufferCanvas = React.useCallback((axis: "left" | "top") => {
    let bufferCanvas = headerBlitBufferCanvasRef.current[axis];
    if (!bufferCanvas && typeof document !== "undefined") {
      bufferCanvas = document.createElement("canvas");
      headerBlitBufferCanvasRef.current[axis] = bufferCanvas;
    }
    return bufferCanvas;
  }, []);
  const visibleRows = React.useMemo(() => {
    return buildVisibleAxisIndices(
      activeSheet?.visibleRows ?? [],
      displayRowLimit,
      activeSheet?.maxUsedRow ?? -1,
      hiddenRowSet
    );
  }, [activeSheet?.maxUsedRow, activeSheet?.visibleRows, displayRowLimit, hiddenRowSet]);
  const visibleCols = React.useMemo(() => {
    return buildVisibleAxisIndices(
      activeSheet?.visibleCols ?? [],
      displayColLimit,
      activeSheet?.maxUsedCol ?? -1,
      hiddenColSet
    );
  }, [activeSheet?.maxUsedCol, activeSheet?.visibleCols, displayColLimit, hiddenColSet]);
  const frozenRows = React.useMemo(() => {
    const freezeRow = activeSheet?.freezePanes?.row ?? 0;
    if (freezeRow <= 0) {
      return [];
    }

    return visibleRows.filter((row) => row < freezeRow);
  }, [activeSheet?.freezePanes?.row, visibleRows]);
  const frozenCols = React.useMemo(() => {
    const freezeCol = activeSheet?.freezePanes?.col ?? 0;
    if (freezeCol <= 0) {
      return [];
    }

    return visibleCols.filter((col) => col < freezeCol);
  }, [activeSheet?.freezePanes?.col, visibleCols]);
  const actualColWidths = React.useMemo(
    () => {
      const widths = new Array(displayColLimit).fill(0);
      const showGridLines = activeSheet?.showGridLines ?? true;
      const fallbackWidth = Math.max(
        resolveRenderedSheetAxisPixels(activeSheet?.defaultColWidthPx ?? DEFAULT_COL_WIDTH, showGridLines),
        DEFAULT_COL_WIDTH / 2
      );

      if (worksheet) {
        for (let col = 0; col < displayColLimit; col += 1) {
          if (worksheet.isColumnHidden(col)) {
            continue;
          }

          const width = worksheet.getColumnWidth(col);
          if (width !== undefined && width !== null) {
            widths[col] = Math.max(
              resolveRenderedSheetAxisPixels(
                resolveSheetColumnWidthPixels(width, activeSheet?.columnWidthCharacterWidthPx),
                showGridLines
              ),
              DEFAULT_COL_WIDTH / 2
            );
            continue;
          }

          widths[col] = Math.max(
            resolveRenderedSheetAxisPixels(
              activeSheet?.colWidthOverridesPx[col] ?? activeSheet?.defaultColWidthPx ?? DEFAULT_COL_WIDTH,
              showGridLines
            ),
            DEFAULT_COL_WIDTH / 2
          );
        }

        return widths;
      }

      const precomputedCols = activeSheet?.visibleCols ?? [];
      const precomputedWidths = activeSheet?.colWidths ?? [];
      for (let index = 0; index < precomputedCols.length; index += 1) {
        const col = precomputedCols[index];
        if (col === undefined || col >= displayColLimit) {
          continue;
        }

        widths[col] = Math.max(
          resolveRenderedSheetAxisPixels(precomputedWidths[index] ?? activeSheet?.defaultColWidthPx ?? DEFAULT_COL_WIDTH, showGridLines),
          DEFAULT_COL_WIDTH / 2
        );
      }

      for (let col = Math.max(0, (activeSheet?.maxUsedCol ?? -1) + 1); col < displayColLimit; col += 1) {
        widths[col] = fallbackWidth;
      }

      return widths;
    },
    [
      activeSheet?.colWidthOverridesPx,
      activeSheet?.colWidths,
      activeSheet?.defaultColWidthPx,
      activeSheet?.maxUsedCol,
      activeSheet?.showGridLines,
      activeSheet?.visibleCols,
      displayColLimit,
      worksheet,
      revision
    ]
  );
  const actualRowHeights = React.useMemo(
    () => {
      const heights = new Array(displayRowLimit).fill(0);
      const showGridLines = activeSheet?.showGridLines ?? true;
      const fallbackHeight = Math.max(
        resolveRenderedSheetAxisPixels(activeSheet?.defaultRowHeightPx ?? DEFAULT_ROW_HEIGHT, showGridLines),
        DEFAULT_ROW_HEIGHT / 1.5
      );

      if (worksheet) {
        for (let row = 0; row < displayRowLimit; row += 1) {
          if (worksheet.isRowHidden(row)) {
            continue;
          }

          const height = worksheet.getRowHeight(row);
          if (height !== undefined && height !== null) {
            heights[row] = Math.max(
              resolveRenderedSheetAxisPixels(resolveSheetRowHeightPixels(height), showGridLines),
              DEFAULT_ROW_HEIGHT / 1.5
            );
            continue;
          }

          const maxMeasuredCol = Math.min(displayColLimit - 1, activeSheet?.maxUsedCol ?? -1);
          let estimatedAutoHeight = 0;
          if (maxMeasuredCol >= 0) {
            for (let col = 0; col <= maxMeasuredCol; col += 1) {
              if (worksheet.isColumnHidden(col) || worksheet.isMergedSecondary(row, col)) {
                continue;
              }

              const cellValue = worksheet.getFormattedValueAt(row, col);
              if (!cellValue) {
                continue;
              }

              const rawStyle = (worksheet.getCellStyleAt(row, col) as Record<string, unknown> | null | undefined) ?? null;
              const cellStyle = buildCellStyle(rawStyle, palette, activeSheet?.themePalette, {
                showGridLines: activeSheet?.showGridLines
              });
              if (cellStyle.whiteSpace !== "pre-wrap" && !cellValue.includes("\n") && !cellValue.includes("\r")) {
                continue;
              }

              const merge = worksheet.getMergeSpan(row, col) as { colSpan?: number } | null | undefined;
              const colSpan = Math.max(1, merge?.colSpan ?? 1);
              let cellWidth = 0;
              for (let spanCol = col; spanCol < Math.min(displayColLimit, col + colSpan); spanCol += 1) {
                cellWidth += actualColWidths[spanCol] ?? (activeSheet?.defaultColWidthPx ?? DEFAULT_COL_WIDTH);
              }
              estimatedAutoHeight = Math.max(
                estimatedAutoHeight,
                measureWrappedTextHeight(cellValue, cellStyle, cellWidth)
              );
            }
          }
          if (estimatedAutoHeight > 0) {
            heights[row] = Math.max(
              resolveRenderedSheetAxisPixels(Math.ceil(estimatedAutoHeight), showGridLines),
              fallbackHeight
            );
            continue;
          }

          heights[row] = Math.max(
            resolveRenderedSheetAxisPixels(
              activeSheet?.rowHeightOverridesPx[row] ?? activeSheet?.defaultRowHeightPx ?? DEFAULT_ROW_HEIGHT,
              showGridLines
            ),
            DEFAULT_ROW_HEIGHT / 1.5
          );
        }

        return heights;
      }

      const precomputedRows = activeSheet?.visibleRows ?? [];
      const precomputedHeights = activeSheet?.rowHeights ?? [];
      for (let index = 0; index < precomputedRows.length; index += 1) {
        const row = precomputedRows[index];
        if (row === undefined || row >= displayRowLimit) {
          continue;
        }

        heights[row] = Math.max(
          resolveRenderedSheetAxisPixels(precomputedHeights[index] ?? activeSheet?.defaultRowHeightPx ?? DEFAULT_ROW_HEIGHT, showGridLines),
          DEFAULT_ROW_HEIGHT / 1.5
        );
      }

      for (let row = Math.max(0, (activeSheet?.maxUsedRow ?? -1) + 1); row < displayRowLimit; row += 1) {
        heights[row] = fallbackHeight;
      }

      return heights;
    },
    [
      activeSheet?.defaultRowHeightPx,
      activeSheet?.maxUsedRow,
      activeSheet?.maxUsedCol,
      activeSheet?.rowHeightOverridesPx,
      activeSheet?.rowHeights,
      activeSheet?.showGridLines,
      activeSheet?.themePalette,
      activeSheet?.visibleRows,
      actualColWidths,
      displayRowLimit,
      displayColLimit,
      palette,
      worksheet,
      revision
    ]
  );
  const displayActualColWidths = React.useMemo(
    () => actualColWidths.map((width) => width * zoomFactor),
    [actualColWidths, zoomFactor]
  );
  const displayActualRowHeights = React.useMemo(
    () => actualRowHeights.map((height) => height * zoomFactor),
    [actualRowHeights, zoomFactor]
  );
  const displayEffectiveColWidths = React.useMemo(
    () => visibleCols.map((col) => displayActualColWidths[col] ?? displayDefaultColWidth),
    [displayActualColWidths, displayDefaultColWidth, visibleCols]
  );
  const displayEffectiveRowHeights = React.useMemo(
    () => visibleRows.map((row) => displayActualRowHeights[row] ?? displayDefaultRowHeight),
    [displayActualRowHeights, displayDefaultRowHeight, visibleRows]
  );
  const rowIndexByActual = React.useMemo(() => new Map(visibleRows.map((row, index) => [row, index])), [visibleRows]);
  const colIndexByActual = React.useMemo(() => new Map(visibleCols.map((col, index) => [col, index])), [visibleCols]);
  const visibleRowsRef = React.useRef<number[]>(visibleRows);
  const visibleColsRef = React.useRef<number[]>(visibleCols);
  const effectiveRowHeightsRef = React.useRef<number[]>(displayEffectiveRowHeights);
  const effectiveColWidthsRef = React.useRef<number[]>(displayEffectiveColWidths);
  const rowPrefixSums = React.useMemo(() => buildPrefixSums(displayEffectiveRowHeights), [displayEffectiveRowHeights]);
  const colPrefixSums = React.useMemo(() => buildPrefixSums(displayEffectiveColWidths), [displayEffectiveColWidths]);
  const actualRowPrefixSums = React.useMemo(() => buildPrefixSums(displayActualRowHeights), [displayActualRowHeights]);
  const actualColPrefixSums = React.useMemo(() => buildPrefixSums(displayActualColWidths), [displayActualColWidths]);
  const stickyTopByRow = React.useMemo(
    () => buildStickyOffsets(frozenRows, displayActualRowHeights, displayHeaderHeight),
    [displayActualRowHeights, displayHeaderHeight, frozenRows]
  );
  const stickyLeftByCol = React.useMemo(
    () => buildStickyOffsets(frozenCols, displayActualColWidths, displayRowHeaderWidth),
    [displayActualColWidths, displayRowHeaderWidth, frozenCols]
  );
  const frozenPaneBottom = React.useMemo(
    () => (
      frozenRows.length > 0
        ? frozenRows.reduce(
            (max, row) => Math.max(max, (stickyTopByRow.get(row) ?? displayHeaderHeight) + (displayActualRowHeights[row] ?? displayDefaultRowHeight)),
            displayHeaderHeight
          )
        : displayHeaderHeight
    ),
    [displayActualRowHeights, displayDefaultRowHeight, displayHeaderHeight, frozenRows, stickyTopByRow]
  );
  const frozenPaneRight = React.useMemo(
    () => (
      frozenCols.length > 0
        ? frozenCols.reduce(
            (max, col) => Math.max(max, (stickyLeftByCol.get(col) ?? displayRowHeaderWidth) + (displayActualColWidths[col] ?? displayDefaultColWidth)),
            displayRowHeaderWidth
          )
        : displayRowHeaderWidth
    ),
    [displayActualColWidths, displayDefaultColWidth, displayRowHeaderWidth, frozenCols, stickyLeftByCol]
  );
  canvasLayoutMetricsRef.current = {
    displayHeaderHeight,
    displayRowHeaderWidth,
    frozenPaneBottom,
    frozenPaneRight
  };
  const rowPrefixSumsRef = React.useRef<number[]>(rowPrefixSums);
  const colPrefixSumsRef = React.useRef<number[]>(colPrefixSums);
  const firstVisibleRow = visibleRows[0];
  const lastVisibleRow = visibleRows[visibleRows.length - 1];
  const firstVisibleCol = visibleCols[0];
  const lastVisibleCol = visibleCols[visibleCols.length - 1];
  const displayedSelection = fillPreviewRange ?? normalizedSelection;
  const toLogicalRect = React.useCallback((rect: XlsxImageRect): XlsxImageRect => ({
    height: rect.height / zoomFactor,
    left: rect.left / zoomFactor,
    top: rect.top / zoomFactor,
    width: rect.width / zoomFactor
  }), [zoomFactor]);
  const drawingExtents = React.useMemo(() => {
    const imageExtents = images.reduce(
      (current, image) => {
        const extents = resolveImageAnchorExtents(image);
        return {
          maxCol: Math.max(current.maxCol, extents.maxCol),
          maxRow: Math.max(current.maxRow, extents.maxRow)
        };
      },
      { maxCol: -1, maxRow: -1 }
    );
    const shapeExtents = shapes.reduce(
      (current, shape) => {
        const extents = resolveShapeAnchorExtents(shape);
        return {
          maxCol: Math.max(current.maxCol, extents.maxCol),
          maxRow: Math.max(current.maxRow, extents.maxRow)
        };
      },
      { maxCol: -1, maxRow: -1 }
    );
    const chartExtents = charts.reduce(
      (current, chart) => {
        const extents = resolveChartAnchorExtents(chart);
        return {
          maxCol: Math.max(current.maxCol, extents.maxCol),
          maxRow: Math.max(current.maxRow, extents.maxRow)
        };
      },
      { maxCol: -1, maxRow: -1 }
    );

    return {
      maxCol: Math.max(imageExtents.maxCol, shapeExtents.maxCol, chartExtents.maxCol),
      maxRow: Math.max(imageExtents.maxRow, shapeExtents.maxRow, chartExtents.maxRow)
    };
  }, [charts, images, shapes]);
  const shouldVirtualizeRows = visibleRows.length > 0;
  const shouldVirtualizeCols = visibleCols.length > 0 && frozenCols.length === 0;
  const shouldUseDomVirtualizer = !experimentalCanvas;
  const getScrollElement = React.useCallback(() => scrollRef.current, []);
  const estimateRowSize = React.useCallback(
    (index: number) => displayEffectiveRowHeights[index] ?? displayDefaultRowHeight,
    [displayDefaultRowHeight, displayEffectiveRowHeights]
  );
  const getRowItemKey = React.useCallback((index: number) => visibleRows[index] ?? index, [visibleRows]);
  const estimateColSize = React.useCallback(
    (index: number) => displayEffectiveColWidths[index] ?? displayDefaultColWidth,
    [displayDefaultColWidth, displayEffectiveColWidths]
  );
  const getColItemKey = React.useCallback((index: number) => visibleCols[index] ?? index, [visibleCols]);

  const rowVirtualizer = useVirtualizer({
    count: visibleRows.length,
    enabled: shouldUseDomVirtualizer,
    estimateSize: estimateRowSize,
    getScrollElement,
    getItemKey: getRowItemKey,
    overscan: 10
  });
  const colVirtualizer = useVirtualizer({
    count: visibleCols.length,
    enabled: shouldUseDomVirtualizer,
    estimateSize: estimateColSize,
    getScrollElement,
    getItemKey: getColItemKey,
    horizontal: true,
    overscan: 6
  });
  const currentRowVirtualItems = shouldUseDomVirtualizer ? rowVirtualizer.getVirtualItems() : EMPTY_VIRTUAL_ITEMS;
  const currentColVirtualItems = shouldUseDomVirtualizer ? colVirtualizer.getVirtualItems() : EMPTY_VIRTUAL_ITEMS;
  const frozenRowVirtualIndices = React.useMemo(
    () => frozenRows
      .map((row) => rowIndexByActual.get(row))
      .filter((index): index is number => index !== undefined),
    [frozenRows, rowIndexByActual]
  );
  const frozenColVirtualIndices = React.useMemo(
    () => frozenCols
      .map((col) => colIndexByActual.get(col))
      .filter((index): index is number => index !== undefined),
    [colIndexByActual, frozenCols]
  );
  const indexedMergedRegions = React.useMemo(
    () =>
      mergedRegions.flatMap((region) => {
        const startRowIndex = rowIndexByActual.get(region.start.row);
        const endRowIndex = rowIndexByActual.get(region.end.row);
        const startColIndex = colIndexByActual.get(region.start.col);
        const endColIndex = colIndexByActual.get(region.end.col);
        if (
          startRowIndex === undefined
          || endRowIndex === undefined
          || startColIndex === undefined
          || endColIndex === undefined
        ) {
          return [];
        }

        return [{
          endColIndex,
          endRowIndex,
          startColIndex,
          startRowIndex
        } satisfies IndexedMergedRegion];
      }),
    [colIndexByActual, mergedRegions, rowIndexByActual]
  );
  const domBaseRowIndices = React.useMemo(
    () => Array.from(new Set([
      ...currentRowVirtualItems.map((virtualRow) => virtualRow.index),
      ...frozenRowVirtualIndices
    ])).sort((left, right) => left - right),
    [currentRowVirtualItems, frozenRowVirtualIndices]
  );
  const domBaseColIndices = React.useMemo(
    () => Array.from(new Set([
      ...(shouldVirtualizeCols
        ? currentColVirtualItems.map((virtualCol) => virtualCol.index)
        : visibleCols.map((_, index) => index)),
      ...frozenColVirtualIndices
    ])).sort((left, right) => left - right),
    [currentColVirtualItems, frozenColVirtualIndices, shouldVirtualizeCols, visibleCols]
  );
  const domExpandedWindow = React.useMemo(
    () => expandMergeAwareWindow(domBaseRowIndices, domBaseColIndices, indexedMergedRegions),
    [domBaseColIndices, domBaseRowIndices, indexedMergedRegions]
  );
  const rowVirtualItemByIndex = React.useMemo(
    () => new Map(currentRowVirtualItems.map((virtualRow) => [virtualRow.index, virtualRow])),
    [currentRowVirtualItems]
  );
  const colVirtualItemByIndex = React.useMemo(
    () => new Map(currentColVirtualItems.map((virtualCol) => [virtualCol.index, virtualCol])),
    [currentColVirtualItems]
  );
  const maxRowDisplayLimit = React.useMemo(
    () => resolveOpenGridExtent(activeSheet?.maxUsedRow ?? -1, MIN_OPEN_GRID_ROWS, OPEN_GRID_ROW_PADDING),
    [activeSheet?.maxUsedRow]
  );
  const maxColDisplayLimit = React.useMemo(
    () => resolveOpenGridExtent(activeSheet?.maxUsedCol ?? -1, MIN_OPEN_GRID_COLS, OPEN_GRID_COL_PADDING),
    [activeSheet?.maxUsedCol]
  );

  React.useEffect(() => {
    if (!isWorkerBacked) {
      return;
    }

    const lastVirtualRowIndex = currentRowVirtualItems[currentRowVirtualItems.length - 1]?.index ?? -1;
    const lastVirtualColIndex = currentColVirtualItems[currentColVirtualItems.length - 1]?.index ?? -1;

    if (
      lastVirtualRowIndex >= visibleRows.length - WORKER_GRID_GROW_THRESHOLD_ROWS
      && displayRowLimit < maxRowDisplayLimit
    ) {
      setDisplayRowLimit((current) => Math.min(maxRowDisplayLimit, current + WORKER_GRID_GROW_ROWS));
    }

    if (
      lastVirtualColIndex >= visibleCols.length - WORKER_GRID_GROW_THRESHOLD_COLS
      && displayColLimit < maxColDisplayLimit
    ) {
      setDisplayColLimit((current) => Math.min(maxColDisplayLimit, current + WORKER_GRID_GROW_COLS));
    }
  }, [
    currentColVirtualItems,
    currentRowVirtualItems,
    displayColLimit,
    displayRowLimit,
    isWorkerBacked,
    maxColDisplayLimit,
    maxRowDisplayLimit,
    visibleCols.length,
    visibleRows.length
  ]);

  React.useEffect(() => {
    activeCellRef.current = activeCell;
  }, [activeCell]);

  React.useEffect(() => {
    selectedChartIdRef.current = selectedChartId;
  }, [selectedChartId]);

  React.useEffect(() => {
    selectedImageIdRef.current = selectedImageId;
  }, [selectedImageId]);

  React.useEffect(() => {
    committedZoomScaleRef.current = zoomScale;
    if (pendingLiveZoomCommitRef.current !== null && zoomScale === pendingLiveZoomCommitRef.current) {
      pendingLiveZoomCommitRef.current = null;
      touchPinchStateRef.current = null;
      safariPinchStartZoomRef.current = null;
      clearLiveZoomCommitTimer();
      updateLiveGestureZoomState(null);
      gestureZoomScaleRef.current = zoomScale;
      return;
    }

    const currentLiveGestureZoom = liveGestureZoomRef.current;
    if (currentLiveGestureZoom && zoomScale !== currentLiveGestureZoom.baseZoomScale) {
      clearLiveGestureZoom();
      return;
    }

    if (currentLiveGestureZoom === null) {
      gestureZoomScaleRef.current = zoomScale;
    }
  }, [clearLiveGestureZoom, clearLiveZoomCommitTimer, updateLiveGestureZoomState, zoomScale]);

  React.useLayoutEffect(() => {
    drawingViewportStateRef.current = drawingViewport;
    applyCanvasViewportCompensation();
  }, [applyCanvasViewportCompensation, drawingViewport]);

  React.useEffect(() => {
    selectionRef.current = normalizedSelection;
  }, [normalizedSelection]);

  React.useEffect(() => {
    editingCellRef.current = editingCell;
  }, [editingCell]);

  React.useLayoutEffect(() => {
    if (!editingCell) {
      return;
    }

    const input = editingInputRef.current;
    if (!input) {
      return;
    }

    input.focus();
    const caret = input.value.length;
    input.setSelectionRange(caret, caret);
  }, [editingCell]);

  React.useEffect(() => {
    readOnlyRef.current = readOnly;
  }, [readOnly]);

  React.useEffect(() => {
    chartPreviewRectRef.current = chartPreviewRect;
  }, [chartPreviewRect]);

  React.useEffect(() => {
    imagePreviewRectRef.current = imagePreviewRect;
  }, [imagePreviewRect]);

  React.useEffect(() => () => {
    if (selectionCommitFrameRef.current !== null) {
      window.cancelAnimationFrame(selectionCommitFrameRef.current);
      selectionCommitFrameRef.current = null;
    }
    if (liveGestureZoomFrameRef.current !== null) {
      window.cancelAnimationFrame(liveGestureZoomFrameRef.current);
      liveGestureZoomFrameRef.current = null;
    }
    if (drawingViewportFrameRef.current !== null) {
      window.cancelAnimationFrame(drawingViewportFrameRef.current);
      drawingViewportFrameRef.current = null;
    }
    clearLiveZoomCommitTimer();
    if (imagePreviewFrameRef.current !== null) {
      window.cancelAnimationFrame(imagePreviewFrameRef.current);
      imagePreviewFrameRef.current = null;
    }
    if (chartPreviewFrameRef.current !== null) {
      window.cancelAnimationFrame(chartPreviewFrameRef.current);
      chartPreviewFrameRef.current = null;
    }
  }, [clearLiveZoomCommitTimer]);

  const scheduleImagePreviewRect = React.useCallback((preview: { id: string; rect: XlsxImageRect }) => {
    pendingImagePreviewRef.current = preview;
    if (imagePreviewFrameRef.current !== null) {
      return;
    }
    imagePreviewFrameRef.current = window.requestAnimationFrame(() => {
      imagePreviewFrameRef.current = null;
      const nextPreview = pendingImagePreviewRef.current;
      pendingImagePreviewRef.current = null;
      if (!nextPreview) {
        return;
      }
      imagePreviewRectRef.current = nextPreview;
      setImagePreviewRect(nextPreview);
    });
  }, []);

  const scheduleChartPreviewRect = React.useCallback((preview: { id: string; rect: XlsxImageRect }) => {
    pendingChartPreviewRef.current = preview;
    if (chartPreviewFrameRef.current !== null) {
      return;
    }
    chartPreviewFrameRef.current = window.requestAnimationFrame(() => {
      chartPreviewFrameRef.current = null;
      const nextPreview = pendingChartPreviewRef.current;
      pendingChartPreviewRef.current = null;
      if (!nextPreview) {
        return;
      }
      chartPreviewRectRef.current = nextPreview;
      setChartPreviewRect(nextPreview);
    });
  }, []);

  React.useEffect(() => {
    displayedSelectionRef.current = selectionPreviewRangeRef.current ?? displayedSelection;
  }, [displayedSelection]);

  React.useEffect(() => {
    const previewRange = selectionPreviewRangeRef.current;
    if (!previewRange || pendingSelectionDragRef.current || selectionDragRef.current || fillDragRef.current) {
      return;
    }

    if (normalizedSelection && rangesEqual(previewRange, normalizedSelection)) {
      selectionPreviewRangeRef.current = null;
      displayedSelectionRef.current = displayedSelection;
      return;
    }

    if (!normalizedSelection && !displayedSelection) {
      selectionPreviewRangeRef.current = null;
      displayedSelectionRef.current = null;
    }
  }, [displayedSelection, normalizedSelection]);

  React.useEffect(() => {
    firstVisibleColRef.current = firstVisibleCol;
    lastVisibleColRef.current = lastVisibleCol;
    firstVisibleRowRef.current = firstVisibleRow;
    lastVisibleRowRef.current = lastVisibleRow;
  }, [firstVisibleCol, firstVisibleRow, lastVisibleCol, lastVisibleRow]);

  React.useEffect(() => {
    visibleRowsRef.current = visibleRows;
    visibleColsRef.current = visibleCols;
    effectiveRowHeightsRef.current = displayEffectiveRowHeights;
    effectiveColWidthsRef.current = displayEffectiveColWidths;
    rowPrefixSumsRef.current = rowPrefixSums;
    colPrefixSumsRef.current = colPrefixSums;
  }, [colPrefixSums, displayEffectiveColWidths, displayEffectiveRowHeights, rowPrefixSums, visibleCols, visibleRows]);

  React.useLayoutEffect(() => {
    const scroller = scrollRef.current;
    const previousZoomFactor = previousZoomFactorRef.current;
    if (!scroller || previousZoomFactor === zoomFactor) {
      previousZoomFactorRef.current = zoomFactor;
      return;
    }

    const zoomAnchor = pendingZoomAnchorRef.current;
    if (zoomAnchor) {
      scroller.scrollLeft = ((scroller.scrollLeft + zoomAnchor.x) / previousZoomFactor) * zoomFactor - zoomAnchor.x;
      scroller.scrollTop = ((scroller.scrollTop + zoomAnchor.y) / previousZoomFactor) * zoomFactor - zoomAnchor.y;
      pendingZoomAnchorRef.current = null;
    } else {
      scroller.scrollLeft = (scroller.scrollLeft / previousZoomFactor) * zoomFactor;
      scroller.scrollTop = (scroller.scrollTop / previousZoomFactor) * zoomFactor;
    }
    previousZoomFactorRef.current = zoomFactor;
    if (shouldUseDomVirtualizer) {
      rowVirtualizer.measure();
      colVirtualizer.measure();
    }
    syncDrawingViewport(scroller, { immediate: true });
  }, [shouldUseDomVirtualizer, syncDrawingViewport, zoomFactor]);

  React.useLayoutEffect(() => {
    syncDrawingViewport(scrollRef.current, { immediate: true });
  }, [activeSheet, activeTabIndex, displayColLimit, displayRowLimit, syncDrawingViewport, zoomFactor]);

  React.useLayoutEffect(() => {
    const scroller = scrollRef.current;
    if (!scroller) {
      return;
    }
    syncDrawingViewport(scroller, { immediate: true });
    if (typeof ResizeObserver === "undefined") {
      return;
    }
    const observer = new ResizeObserver(() => {
      syncDrawingViewport(scroller, { immediate: true });
    });
    observer.observe(scroller);
    return () => {
      observer.disconnect();
    };
  }, [activeSheet, syncDrawingViewport]);

  React.useEffect(() => {
    setDisplayRowLimit(
      resolveInitialDisplayExtent(
        activeSheet?.maxUsedRow ?? -1,
        MIN_OPEN_GRID_ROWS,
        OPEN_GRID_ROW_PADDING,
        Boolean(isWorkerBacked),
        INITIAL_WORKER_GRID_ROWS
      )
    );
    setDisplayColLimit(
      resolveInitialDisplayExtent(
        activeSheet?.maxUsedCol ?? -1,
        MIN_OPEN_GRID_COLS,
        OPEN_GRID_COL_PADDING,
        Boolean(isWorkerBacked),
        INITIAL_WORKER_GRID_COLS
      )
    );
  }, [activeSheet?.maxUsedCol, activeSheet?.maxUsedRow, activeSheetIndex, isWorkerBacked]);

  React.useEffect(() => {
    const selectionEndRow = normalizedSelection?.end.row ?? -1;
    const selectionEndCol = normalizedSelection?.end.col ?? -1;
    const nextRowLimit = Math.max(
      resolveInitialDisplayExtent(
        activeSheet?.maxUsedRow ?? -1,
        MIN_OPEN_GRID_ROWS,
        OPEN_GRID_ROW_PADDING,
        false,
        INITIAL_WORKER_GRID_ROWS
      ),
      (activeCell?.row ?? -1) + OPEN_GRID_ROW_PADDING + 1,
      selectionEndRow + OPEN_GRID_ROW_PADDING + 1,
      drawingExtents.maxRow + OPEN_GRID_ROW_PADDING + 1
    );
    const nextColLimit = Math.max(
      resolveInitialDisplayExtent(
        activeSheet?.maxUsedCol ?? -1,
        MIN_OPEN_GRID_COLS,
        OPEN_GRID_COL_PADDING,
        false,
        INITIAL_WORKER_GRID_COLS
      ),
      (activeCell?.col ?? -1) + OPEN_GRID_COL_PADDING + 1,
      selectionEndCol + OPEN_GRID_COL_PADDING + 1,
      drawingExtents.maxCol + OPEN_GRID_COL_PADDING + 1
    );

    setDisplayRowLimit((current) => (current < nextRowLimit ? nextRowLimit : current));
    setDisplayColLimit((current) => (current < nextColLimit ? nextColLimit : current));
  }, [
    activeCell,
    activeSheet?.maxUsedCol,
    activeSheet?.maxUsedRow,
    drawingExtents.maxCol,
    drawingExtents.maxRow,
    normalizedSelection?.end.col,
    normalizedSelection?.end.row
  ]);

  React.useEffect(() => {
    const initialScrollKey = [
      displayFileName,
      activeSheetIndex,
      activeSheet?.workbookSheetIndex ?? -1,
      activeSheet?.name ?? "",
      activeSheet?.minUsedRow ?? -1,
      activeSheet?.minUsedCol ?? -1,
      isWorkerBacked ? "worker" : "main"
    ].join("|");
    if (initialScrollKeyRef.current === initialScrollKey) {
      return;
    }

    const initialUsedRow = resolveFirstUsedVisibleIndex(visibleRows, activeSheet?.minUsedRow ?? -1);
    const initialUsedCol = resolveFirstUsedVisibleIndex(visibleCols, activeSheet?.minUsedCol ?? -1);
    const initialScrollTop = initialUsedRow > 0
      ? sumPrefixRange(actualRowPrefixSums, 0, initialUsedRow - 1)
      : 0;
    const initialScrollLeft = initialUsedCol > 0
      ? sumPrefixRange(actualColPrefixSums, 0, initialUsedCol - 1)
      : 0;

    if (shouldUseDomVirtualizer && shouldVirtualizeRows) {
      rowVirtualizer.scrollToOffset(initialScrollTop);
    } else if (scrollRef.current) {
      scrollRef.current.scrollTop = initialScrollTop;
    }
    if (scrollRef.current) {
      scrollRef.current.scrollLeft = initialScrollLeft;
    }
    if (shouldUseDomVirtualizer && shouldVirtualizeCols && frozenCols.length === 0) {
      colVirtualizer.scrollToOffset(initialScrollLeft);
    }
    if (scrollRef.current) {
      syncDrawingViewport(scrollRef.current, { immediate: true });
    }
    initialScrollKeyRef.current = initialScrollKey;
  }, [
    activeSheet?.minUsedCol,
    activeSheet?.minUsedRow,
    activeSheet?.name,
    activeSheet?.workbookSheetIndex,
    activeSheetIndex,
    actualColPrefixSums,
    actualRowPrefixSums,
    colVirtualizer,
    displayFileName,
    frozenCols.length,
    isWorkerBacked,
    shouldVirtualizeCols,
    shouldVirtualizeRows,
    shouldUseDomVirtualizer,
    syncDrawingViewport,
    rowVirtualizer,
    visibleCols,
    visibleRows
  ]);

  React.useEffect(() => {
    clearLiveGestureZoom();
  }, [activeTabIndex, clearLiveGestureZoom]);

  React.useEffect(() => {
    setOpenTableMenu(null);
  }, [activeSheet, activeSheetIndex]);

  React.useEffect(() => {
    if (!pendingNavigation || pendingNavigation.sheetIndex !== activeSheetIndex) {
      return;
    }

    selectCell(pendingNavigation.cell);
    setPendingNavigation(null);
  }, [activeSheetIndex, pendingNavigation, selectCell]);

  React.useEffect(() => {
    if (shouldUseDomVirtualizer && shouldVirtualizeRows) {
      rowVirtualizer.measure();
    }
    if (shouldUseDomVirtualizer && shouldVirtualizeCols) {
      colVirtualizer.measure();
    }
  }, [
    activeSheetIndex,
    revision,
    shouldUseDomVirtualizer,
    shouldVirtualizeCols,
    shouldVirtualizeRows,
    visibleCols.length,
    visibleRows.length
  ]);

  const handleScrollerScroll = React.useCallback((event: React.UIEvent<HTMLDivElement>) => {
    const currentScroller = event.currentTarget;
    cachedScrollerRectRef.current = null;
    const paintedViewport = paintedDrawingViewportRef.current;
    const shouldSyncDrawingViewportImmediately =
      Math.abs(currentScroller.scrollLeft - paintedViewport.left) > CANVAS_IMMEDIATE_VIEWPORT_SYNC_THRESHOLD_PX
      || Math.abs(currentScroller.scrollTop - paintedViewport.top) > CANVAS_IMMEDIATE_VIEWPORT_SYNC_THRESHOLD_PX;
    syncDrawingViewport(currentScroller, { immediate: shouldSyncDrawingViewportImmediately });
    if (
      currentScroller.scrollHeight - (currentScroller.scrollTop + currentScroller.clientHeight) <
      OPEN_GRID_VERTICAL_EDGE_PX
    ) {
      setDisplayRowLimit((current) => {
        const nextLimit = current + OPEN_GRID_ROW_GROWTH;
        return readOnly && current < maxRowDisplayLimit ? Math.min(maxRowDisplayLimit, nextLimit) : nextLimit;
      });
    }

    if (
      currentScroller.scrollWidth - (currentScroller.scrollLeft + currentScroller.clientWidth) <
      OPEN_GRID_HORIZONTAL_EDGE_PX
    ) {
      setDisplayColLimit((current) => {
        const nextLimit = current + OPEN_GRID_COL_GROWTH;
        return readOnly && current < maxColDisplayLimit ? Math.min(maxColDisplayLimit, nextLimit) : nextLimit;
      });
    }
  }, [maxColDisplayLimit, maxRowDisplayLimit, readOnly, syncDrawingViewport]);

  React.useEffect(() => {
    const scroller = scrollRef.current;
    if (!scroller || !enableGestureZoom) {
      clearLiveGestureZoom();
      return;
    }

    const handleWheel = (event: WheelEvent) => {
      if (!experimentalCanvas) {
        return;
      }
      if (!event.ctrlKey && !event.metaKey) {
        return;
      }

      const delta = normalizeWheelDelta(event);
      if (delta === 0) {
        return;
      }

      event.preventDefault();
      const anchor = resolveEventAnchor(event.clientX, event.clientY, scroller.getBoundingClientRect());
      updateLiveGestureZoomTarget(
        gestureZoomScaleRef.current * Math.exp(-delta * WHEEL_ZOOM_SENSITIVITY),
        anchor
      );
      scheduleLiveGestureZoomCommit();
    };

    const handleTouchStart = (event: TouchEvent) => {
      if (event.touches.length !== 2) {
        return;
      }

      const [firstTouch, secondTouch] = [event.touches[0], event.touches[1]];
      if (!firstTouch || !secondTouch) {
        return;
      }

      const distance = resolveTouchDistance(firstTouch, secondTouch);
      if (distance <= 0) {
        return;
      }

      event.preventDefault();
      const anchor = resolveEventAnchor(
        (firstTouch.clientX + secondTouch.clientX) / 2,
        (firstTouch.clientY + secondTouch.clientY) / 2,
        scroller.getBoundingClientRect()
      );
      clearLiveZoomCommitTimer();
      updateLiveGestureZoomState({
        anchor,
        baseZoomScale: committedZoomScaleRef.current,
        targetZoomScale: committedZoomScaleRef.current
      });
      gestureZoomScaleRef.current = committedZoomScaleRef.current;
      touchPinchStateRef.current = {
        startDistance: distance,
        startZoomScale: committedZoomScaleRef.current
      };
    };

    const handleTouchMove = (event: TouchEvent) => {
      const pinchState = touchPinchStateRef.current;
      if (!pinchState || event.touches.length !== 2) {
        return;
      }

      const [firstTouch, secondTouch] = [event.touches[0], event.touches[1]];
      if (!firstTouch || !secondTouch) {
        return;
      }

      const distance = resolveTouchDistance(firstTouch, secondTouch);
      if (distance <= 0) {
        return;
      }

      event.preventDefault();
      const anchor = resolveEventAnchor(
        (firstTouch.clientX + secondTouch.clientX) / 2,
        (firstTouch.clientY + secondTouch.clientY) / 2,
        scroller.getBoundingClientRect()
      );
      updateLiveGestureZoomTarget(
        (pinchState.startZoomScale * distance) / pinchState.startDistance,
        anchor
      );
    };

    const handleTouchEnd = () => {
      if (touchPinchStateRef.current) {
        touchPinchStateRef.current = null;
        commitLiveGestureZoom();
      }
    };

    const handleTouchCancel = () => {
      if (touchPinchStateRef.current) {
        clearLiveGestureZoom();
      }
    };

    const handleGestureStart = (event: Event) => {
      const gestureEvent = event as WebKitGestureEvent;
      event.preventDefault();
      const anchor = resolveEventAnchor(
        gestureEvent.clientX,
        gestureEvent.clientY,
        scroller.getBoundingClientRect()
      );
      clearLiveZoomCommitTimer();
      safariPinchStartZoomRef.current = committedZoomScaleRef.current;
      updateLiveGestureZoomState({
        anchor,
        baseZoomScale: committedZoomScaleRef.current,
        targetZoomScale: committedZoomScaleRef.current
      });
      gestureZoomScaleRef.current = committedZoomScaleRef.current;
    };

    const handleGestureChange = (event: Event) => {
      const gestureEvent = event as WebKitGestureEvent;
      const startZoomScale = safariPinchStartZoomRef.current;
      if (startZoomScale === null || !Number.isFinite(gestureEvent.scale) || gestureEvent.scale <= 0) {
        return;
      }

      event.preventDefault();
      const anchor = resolveEventAnchor(
        gestureEvent.clientX,
        gestureEvent.clientY,
        scroller.getBoundingClientRect()
      );
      updateLiveGestureZoomTarget(startZoomScale * gestureEvent.scale, anchor);
    };

    const handleGestureEnd = () => {
      if (safariPinchStartZoomRef.current !== null) {
        safariPinchStartZoomRef.current = null;
        commitLiveGestureZoom();
      }
    };

    const handleGestureCancel = () => {
      if (safariPinchStartZoomRef.current !== null) {
        clearLiveGestureZoom();
      }
    };

    scroller.addEventListener("wheel", handleWheel, { passive: false });
    scroller.addEventListener("touchstart", handleTouchStart, { passive: false });
    scroller.addEventListener("touchmove", handleTouchMove, { passive: false });
    scroller.addEventListener("touchend", handleTouchEnd);
    scroller.addEventListener("touchcancel", handleTouchCancel);
    scroller.addEventListener("gesturestart", handleGestureStart as EventListener, { passive: false });
    scroller.addEventListener("gesturechange", handleGestureChange as EventListener, { passive: false });
    scroller.addEventListener("gestureend", handleGestureEnd as EventListener);

    return () => {
      scroller.removeEventListener("wheel", handleWheel);
      scroller.removeEventListener("touchstart", handleTouchStart);
      scroller.removeEventListener("touchmove", handleTouchMove);
      scroller.removeEventListener("touchend", handleTouchEnd);
      scroller.removeEventListener("touchcancel", handleTouchCancel);
      scroller.removeEventListener("gesturestart", handleGestureStart as EventListener);
      scroller.removeEventListener("gesturechange", handleGestureChange as EventListener);
      scroller.removeEventListener("gestureend", handleGestureEnd as EventListener);
      handleGestureCancel();
    };
  }, [
    clearLiveGestureZoom,
    clearLiveZoomCommitTimer,
    commitLiveGestureZoom,
    enableGestureZoom,
    experimentalCanvas,
    scheduleLiveGestureZoomCommit,
    updateLiveGestureZoomState,
    updateLiveGestureZoomTarget
  ]);

  React.useEffect(() => {
    if (!openTableMenu) {
      return;
    }

    function handlePointerDown(event: PointerEvent) {
      if (tableMenuRef.current?.contains(event.target as Node)) {
        return;
      }
      setOpenTableMenu(null);
    }

    function handleKeyDown(event: KeyboardEvent) {
      if (event.key === "Escape") {
        setOpenTableMenu(null);
      }
    }

    window.addEventListener("pointerdown", handlePointerDown);
    window.addEventListener("keydown", handleKeyDown);
    return () => {
      window.removeEventListener("pointerdown", handlePointerDown);
      window.removeEventListener("keydown", handleKeyDown);
    };
  }, [openTableMenu]);

  const resolvePointerCellFromGeometry = React.useCallback((clientX: number, clientY: number): XlsxCellAddress | null => {
    const wrapper = wrapperRef.current;
    const scroller = scrollRef.current;
    const visibleRowsCurrent = visibleRowsRef.current;
    const visibleColsCurrent = visibleColsRef.current;
    const rowPrefixSumsCurrent = rowPrefixSumsRef.current;
    const colPrefixSumsCurrent = colPrefixSumsRef.current;

    if (!wrapper || !scroller || visibleRowsCurrent.length === 0 || visibleColsCurrent.length === 0) {
      return null;
    }

    const scrollerRect = cachedScrollerRectRef.current ?? scroller.getBoundingClientRect();
    if (
      clientX < scrollerRect.left ||
      clientX > scrollerRect.right ||
      clientY < scrollerRect.top ||
      clientY > scrollerRect.bottom
    ) {
      return null;
    }

    const pointerOffsetX = clientX - scrollerRect.left;
    const pointerOffsetY = clientY - scrollerRect.top;
    const localX = pointerOffsetX + (pointerOffsetX >= frozenPaneRight ? scroller.scrollLeft : 0);
    const localY = pointerOffsetY + (pointerOffsetY >= frozenPaneBottom ? scroller.scrollTop : 0);
    const rowContentOffset = localY - displayHeaderHeight;
    const colContentOffset = localX - displayRowHeaderWidth;
    let geometryCell: XlsxCellAddress | null = null;

    if (localY >= displayHeaderHeight && localX < displayRowHeaderWidth) {
      const rowIndex = findIndexForOffsetPrefix(rowPrefixSumsCurrent, rowContentOffset);
      const actualRow = visibleRowsCurrent[rowIndex];
      if (actualRow !== undefined && firstVisibleColRef.current !== undefined) {
        geometryCell = { row: actualRow, col: firstVisibleColRef.current };
      }
    } else if (localY < displayHeaderHeight && localX >= displayRowHeaderWidth) {
      const colIndex = findIndexForOffsetPrefix(colPrefixSumsCurrent, colContentOffset);
      const actualCol = visibleColsCurrent[colIndex];
      if (actualCol !== undefined && firstVisibleRowRef.current !== undefined) {
        geometryCell = { row: firstVisibleRowRef.current, col: actualCol };
      }
    } else if (localY >= displayHeaderHeight && localX >= displayRowHeaderWidth) {
      const rowIndex = findIndexForOffsetPrefix(rowPrefixSumsCurrent, rowContentOffset);
      const colIndex = findIndexForOffsetPrefix(colPrefixSumsCurrent, colContentOffset);
      const actualRow = visibleRowsCurrent[rowIndex];
      const actualCol = visibleColsCurrent[colIndex];
      if (actualRow !== undefined && actualCol !== undefined) {
        geometryCell = { row: actualRow, col: actualCol };
      }
    }

    return geometryCell;
  }, [displayHeaderHeight, displayRowHeaderWidth, frozenPaneBottom, frozenPaneRight]);

  const resolvePointerCellFromHitTest = React.useCallback((clientX: number, clientY: number): XlsxCellAddress | null => {
    const elementsAtPoint = typeof document.elementsFromPoint === "function"
      ? document.elementsFromPoint(clientX, clientY)
      : [document.elementFromPoint(clientX, clientY)].filter((element): element is Element => Boolean(element));
    for (const element of elementsAtPoint) {
      if (!(element instanceof HTMLElement)) {
        continue;
      }

      const cell = parseCellAddressAttribute(element.closest<HTMLElement>("[data-xlsx-cell]")?.getAttribute("data-xlsx-cell") ?? null);
      if (cell) {
        return cell;
      }

      const colHeader = element.closest<HTMLElement>("[data-xlsx-col-header]");
      if (colHeader && firstVisibleRowRef.current !== undefined) {
        const actualCol = Number(colHeader.getAttribute("data-xlsx-col-header"));
        if (Number.isInteger(actualCol) && actualCol >= 0) {
          return { row: firstVisibleRowRef.current, col: actualCol };
        }
      }

      const rowElement = element.closest<HTMLTableRowElement>("tr[data-xlsx-row]");
      if (rowElement && firstVisibleColRef.current !== undefined) {
        const actualRow = Number(rowElement.getAttribute("data-xlsx-row"));
        if (Number.isInteger(actualRow) && actualRow >= 0) {
          return { row: actualRow, col: firstVisibleColRef.current };
        }
      }
    }

    return null;
  }, []);

  const resolvePointerCellFromClient = React.useCallback((clientX: number, clientY: number): XlsxCellAddress | null => {
    const geometryCell = resolvePointerCellFromGeometry(clientX, clientY);
    if (geometryCell) {
      return geometryCell;
    }

    const resolvedCell = resolvePointerCellFromHitTest(clientX, clientY);
    return resolvedCell ? resolveMergeAnchorCell(resolvedCell) : null;
  }, [resolveMergeAnchorCell, resolvePointerCellFromGeometry, resolvePointerCellFromHitTest]);

  const resolveDraggedSelectionCell = React.useCallback((
    dragState: {
      axis: "cell" | "column" | "row";
      contentScaleX: number;
      contentScaleY: number;
      originCell: XlsxCellAddress;
      originContentX: number;
      originContentY: number;
      startClientX: number;
      startClientY: number;
    },
    clientX: number,
    clientY: number
  ): XlsxCellAddress | null => {
    const geometryCell = resolvePointerCellFromGeometry(clientX, clientY);
    const hitCell = resolvePointerCellFromHitTest(clientX, clientY);
    const actualRow =
      geometryCell?.row ??
      (hitCell && rowIndexByActual.has(hitCell.row)
        ? hitCell.row
        : undefined);
    const actualCol = geometryCell?.col ?? hitCell?.col;
    if (actualRow === undefined || actualCol === undefined) {
      return null;
    }

    if (dragState.axis === "row") {
      return { row: actualRow, col: dragState.originCell.col };
    }

    if (dragState.axis === "column") {
      return { row: dragState.originCell.row, col: actualCol };
    }

    return { row: actualRow, col: actualCol };
  }, [resolvePointerCellFromGeometry, resolvePointerCellFromHitTest, rowIndexByActual]);

  const resolveCellPointerOrigin = React.useCallback((
    cell: XlsxCellAddress,
    rect: DOMRect,
    clientX: number,
    clientY: number
  ) => {
    const rowIndex = rowIndexByActual.get(cell.row);
    const colIndex = colIndexByActual.get(cell.col);
    if (rowIndex === undefined || colIndex === undefined) {
      return null;
    }

    const displayWidth = displayEffectiveColWidths[colIndex] ?? displayDefaultColWidth;
    const displayHeight = displayEffectiveRowHeights[rowIndex] ?? displayDefaultRowHeight;
    const contentScaleX = Math.max(0.0001, rect.width > 0 ? rect.width / displayWidth : 1);
    const contentScaleY = Math.max(0.0001, rect.height > 0 ? rect.height / displayHeight : 1);

    return {
      contentScaleX,
      contentScaleY,
      originContentX:
        (colPrefixSums[colIndex] ?? 0)
        + clampContentOffset((clientX - rect.left) / contentScaleX, displayWidth),
      originContentY:
        (rowPrefixSums[rowIndex] ?? 0)
        + clampContentOffset((clientY - rect.top) / contentScaleY, displayHeight)
    };
  }, [colIndexByActual, colPrefixSums, displayDefaultColWidth, displayDefaultRowHeight, displayEffectiveColWidths, displayEffectiveRowHeights, rowIndexByActual, rowPrefixSums]);

  const resolveCellPointerOriginFromClient = React.useCallback((
    cell: XlsxCellAddress,
    clientX: number,
    clientY: number
  ) => {
    const scroller = scrollRef.current;
    const rowIndex = rowIndexByActual.get(cell.row);
    const colIndex = colIndexByActual.get(cell.col);
    if (!scroller || rowIndex === undefined || colIndex === undefined) {
      return null;
    }

    const scrollerRect = cachedScrollerRectRef.current ?? scroller.getBoundingClientRect();
    const pointerOffsetX = clientX - scrollerRect.left;
    const pointerOffsetY = clientY - scrollerRect.top;
    const localX = pointerOffsetX + (pointerOffsetX >= frozenPaneRight ? scroller.scrollLeft : 0);
    const localY = pointerOffsetY + (pointerOffsetY >= frozenPaneBottom ? scroller.scrollTop : 0);
    const displayWidth = displayEffectiveColWidths[colIndex] ?? displayDefaultColWidth;
    const displayHeight = displayEffectiveRowHeights[rowIndex] ?? displayDefaultRowHeight;

    return {
      contentScaleX: 1,
      contentScaleY: 1,
      originContentX:
        (colPrefixSums[colIndex] ?? 0)
        + clampContentOffset(localX - displayRowHeaderWidth - (colPrefixSums[colIndex] ?? 0), displayWidth),
      originContentY:
        (rowPrefixSums[rowIndex] ?? 0)
        + clampContentOffset(localY - displayHeaderHeight - (rowPrefixSums[rowIndex] ?? 0), displayHeight)
    };
  }, [
    colIndexByActual,
    colPrefixSums,
    displayDefaultColWidth,
    displayDefaultRowHeight,
    displayEffectiveColWidths,
    displayEffectiveRowHeights,
    displayHeaderHeight,
    displayRowHeaderWidth,
    frozenPaneBottom,
    frozenPaneRight,
    rowIndexByActual,
    rowPrefixSums
  ]);

  const resolveRowPointerOrigin = React.useCallback((
    actualRow: number,
    rect: DOMRect,
    clientY: number
  ) => {
    const rowIndex = rowIndexByActual.get(actualRow);
    if (rowIndex === undefined) {
      return null;
    }

    const displayHeight = displayEffectiveRowHeights[rowIndex] ?? displayDefaultRowHeight;
    const contentScaleY = Math.max(0.0001, rect.height > 0 ? rect.height / displayHeight : 1);
    return {
      contentScaleX: 1,
      contentScaleY,
      originContentX: colPrefixSums[0] ?? 0,
      originContentY:
        (rowPrefixSums[rowIndex] ?? 0)
        + clampContentOffset((clientY - rect.top) / contentScaleY, displayHeight)
    };
  }, [colPrefixSums, displayDefaultRowHeight, displayEffectiveRowHeights, rowIndexByActual, rowPrefixSums]);

  const resolveColumnPointerOrigin = React.useCallback((
    actualCol: number,
    rect: DOMRect,
    clientX: number
  ) => {
    const colIndex = colIndexByActual.get(actualCol);
    if (colIndex === undefined) {
      return null;
    }

    const displayWidth = displayEffectiveColWidths[colIndex] ?? displayDefaultColWidth;
    const contentScaleX = Math.max(0.0001, rect.width > 0 ? rect.width / displayWidth : 1);
    return {
      contentScaleX,
      contentScaleY: 1,
      originContentX:
        (colPrefixSums[colIndex] ?? 0)
        + clampContentOffset((clientX - rect.left) / contentScaleX, displayWidth),
      originContentY: rowPrefixSums[0] ?? 0
    };
  }, [colIndexByActual, colPrefixSums, displayDefaultColWidth, displayEffectiveColWidths, rowPrefixSums]);

  const applyColumnPreview = React.useCallback((actualCol: number, widthPx: number | null) => {
    const colElement = colElementRefs.current.get(actualCol);
    if (colElement) {
      colElement.style.width = widthPx === null ? "" : `${widthPx}px`;
    }

    const baseIndex = visibleCols.indexOf(actualCol);
    const baseWidth = baseIndex >= 0 ? (displayEffectiveColWidths[baseIndex] ?? displayDefaultColWidth) : displayDefaultColWidth;
    const previewWidth = widthPx ?? baseWidth;
    const baseTotalWidth = displayEffectiveColWidths.reduce((sum, width) => sum + width, 0) + displayRowHeaderWidth;
    const widthDelta = previewWidth - baseWidth;
    if (tableRef.current) {
      tableRef.current.style.width = `${baseTotalWidth + widthDelta}px`;
    }
  }, [displayDefaultColWidth, displayEffectiveColWidths, displayRowHeaderWidth, visibleCols]);

  const applyRowPreview = React.useCallback((actualRow: number, heightPx: number | null) => {
    const rowElement =
      rowElementRefs.current.get(actualRow) ??
      wrapperRef.current?.querySelector<HTMLTableRowElement>(`tr[data-xlsx-row="${actualRow}"]`) ??
      null;
    if (rowElement) {
      rowElementRefs.current.set(actualRow, rowElement);
      rowElement.style.height = heightPx === null ? "" : `${heightPx}px`;
    }
  }, []);

  React.useEffect(() => {
    pendingSelectionDragCleanupRef.current?.();
    selectionDragCleanupRef.current?.();
    fillDragCleanupRef.current?.();
    chartInteractionCleanupRef.current?.();
    imageInteractionCleanupRef.current?.();
    pendingSelectionDragCleanupRef.current = null;
    selectionDragCleanupRef.current = null;
    fillDragCleanupRef.current = null;
    chartInteractionCleanupRef.current = null;
    imageInteractionCleanupRef.current = null;
    pendingSelectionDragRef.current = null;
    selectionDragRef.current = null;
    fillDragRef.current = null;
    chartInteractionRef.current = null;
    imageInteractionRef.current = null;
    pendingResizePreviewRef.current = null;
    columnPreviewRef.current = null;
    rowPreviewRef.current = null;
    setResizeGuide(null);
    clearGlobalCursor();
    setEditingCell(null);
    setEditingValue("");
    setFillPreviewRange(null);
    chartPreviewRectRef.current = null;
    setChartPreviewRect(null);
    imagePreviewRectRef.current = null;
    setImagePreviewRect(null);
    selectionPreviewRangeRef.current = null;
    setInteractionMode("idle");
  }, [activeSheetIndex, clearGlobalCursor, revision]);

  const focusGridElement = React.useCallback((element: HTMLDivElement, options?: FocusOptions) => {
    if (document.activeElement !== element) {
      element.focus(options ?? { preventScroll: true });
    }
  }, []);

  const focusGrid = React.useCallback((options?: FocusOptions) => {
    const scroller = scrollRef.current;
    if (scroller) {
      gridKeyboardActiveRef.current = true;
      focusGridElement(scroller, options);
    }
  }, [focusGridElement]);

  React.useEffect(() => {
    const handleDocumentPointerDown = (event: PointerEvent) => {
      const scroller = scrollRef.current;
      if (!scroller || !(event.target instanceof Node) || !scroller.contains(event.target)) {
        gridKeyboardActiveRef.current = false;
        return;
      }

      if (isInteractiveFocusTarget(event.target, scroller)) {
        gridKeyboardActiveRef.current = false;
        return;
      }

      gridKeyboardActiveRef.current = true;
      focusGridElement(scroller);
    };

    const handleDocumentKeyDown = (event: KeyboardEvent) => {
      if (!gridKeyboardActiveRef.current || event.defaultPrevented) {
        return;
      }

      const scroller = scrollRef.current;
      const interactiveTarget = getInteractiveFocusTarget(event.target);
      if (!scroller || (interactiveTarget && interactiveTarget !== scroller)) {
        return;
      }

      gridKeyboardHandlerRef.current?.(event);
    };

    document.addEventListener("pointerdown", handleDocumentPointerDown, true);
    document.addEventListener("keydown", handleDocumentKeyDown);
    return () => {
      document.removeEventListener("pointerdown", handleDocumentPointerDown, true);
      document.removeEventListener("keydown", handleDocumentKeyDown);
    };
  }, [focusGridElement]);

  const openHyperlink = React.useCallback((target?: string | null, location?: string | null) => {
    const internalTarget = parseInternalSheetLink(location ?? target);
    if (internalTarget) {
      const targetSheetIndex = sheets.findIndex((sheet) => sheet.name === internalTarget.sheetName);
      if (targetSheetIndex >= 0) {
        if (targetSheetIndex === activeSheetIndex) {
          selectCell(internalTarget.cell);
        } else {
          setPendingNavigation({
            cell: internalTarget.cell,
            sheetIndex: targetSheetIndex
          });
          setActiveSheetIndex(targetSheetIndex);
        }
        return;
      }
    }

    const externalTarget = target && !target.startsWith("#") ? target : null;
    if (externalTarget && typeof window !== "undefined") {
      window.open(externalTarget, "_blank", "noopener,noreferrer");
    }
  }, [activeSheetIndex, selectCell, setActiveSheetIndex, sheets]);

  const startEditing = React.useCallback(
    (cell: XlsxCellAddress, initialValue?: string) => {
      if (readOnly) {
        return;
      }

      selectCell(cell);
      setEditingCell(cell);
      setEditingValue(initialValue ?? getControllerCellDisplayValue(cell));
    },
    [getControllerCellDisplayValue, readOnly, selectCell]
  );

  const commitEditing = React.useCallback(() => {
    if (!editingCell) {
      return;
    }

    if (readOnly) {
      editingCellRef.current = null;
      setEditingCell(null);
      setEditingValue("");
      focusGrid();
      return;
    }

    setCellValue(editingCell, editingValue);
    editingCellRef.current = null;
    setEditingCell(null);
    setEditingValue("");
    focusGrid();
  }, [editingCell, editingValue, focusGrid, readOnly, setCellValue]);

  const cancelEditing = React.useCallback(() => {
    editingCellRef.current = null;
    setEditingCell(null);
    setEditingValue("");
    focusGrid();
  }, [focusGrid]);

  const handleEditingBlur = React.useCallback(() => {
    if (!editingCellRef.current) {
      return;
    }

    commitEditing();
  }, [commitEditing]);

  const handleEditingNavigate = React.useCallback((key: string, value: string) => {
    const cell = editingCellRef.current;
    if (!cell) {
      return;
    }

    const currentRowIndex = rowIndexByActual.get(cell.row);
    const currentColIndex = colIndexByActual.get(cell.col);
    if (currentRowIndex === undefined || currentColIndex === undefined) {
      if (!readOnlyRef.current) {
        setCellValue(cell, value);
      }
      editingCellRef.current = null;
      setEditingCell(null);
      setEditingValue("");
      focusGrid();
      return;
    }

    let nextRowIndex = currentRowIndex;
    let nextColIndex = currentColIndex;
    switch (key) {
      case "ArrowDown":
        nextRowIndex += 1;
        break;
      case "ArrowLeft":
        nextColIndex -= 1;
        break;
      case "ArrowRight":
        nextColIndex += 1;
        break;
      case "ArrowUp":
        nextRowIndex -= 1;
        break;
      default:
        return;
    }

    if (!readOnlyRef.current) {
      setCellValue(cell, value);
    }
    editingCellRef.current = null;
    setEditingCell(null);
    setEditingValue("");
    moveSelection(nextRowIndex, nextColIndex, false);
    focusGrid();
  }, [colIndexByActual, focusGrid, rowIndexByActual, setCellValue, visibleCols, visibleRows]);

  React.useEffect(() => {
    commitEditingRef.current = commitEditing;
  }, [commitEditing]);

  const [, startBatchTransition] = React.useTransition();
  const [, startSelectionTransition] = React.useTransition();
  const [asyncViewportRowBatch, setAsyncViewportRowBatch] = React.useState<WorksheetBatchWindow | null>(null);
  const viewportRequest = React.useMemo(() => {
    const firstVirtualRowIndex = domExpandedWindow.rowIndices[0];
    const lastVirtualRowIndex = domExpandedWindow.rowIndices[domExpandedWindow.rowIndices.length - 1];
    const firstVisibleRow = firstVirtualRowIndex === undefined ? undefined : visibleRows[firstVirtualRowIndex];
    const lastVisibleRow = lastVirtualRowIndex === undefined ? undefined : visibleRows[lastVirtualRowIndex];
    if (firstVisibleRow === undefined || lastVisibleRow === undefined || lastVisibleRow < firstVisibleRow) {
      return null;
    }

    const overscan = 48;
    const startRow = Math.max(0, firstVisibleRow - overscan);
    const endRow = lastVisibleRow + overscan;
    return {
      endRow,
      startRow
    };
  }, [domExpandedWindow.rowIndices, visibleRows]);

  const syncViewportRowBatch = React.useMemo<WorksheetBatchWindow | null>(() => {
    if (!shouldVirtualizeRows || !worksheet || getRowsBatchAsync || !viewportRequest) {
      return null;
    }

    const worksheetWithRowsBatch = worksheet as WorksheetWithRowsBatch;
    if (typeof worksheetWithRowsBatch.getRowsBatch !== "function") {
      return null;
    }

    try {
      const rows = worksheetWithRowsBatch.getRowsBatch(viewportRequest.startRow, viewportRequest.endRow - viewportRequest.startRow + 1, {
        includeFormulas: true,
        includeHyperlinks: true,
        includeMergeInfo: true,
        includeStyles: true,
        useFormattedValues: true
      });
      return buildWorksheetBatchWindow(rows as unknown[] | null, activeSheet ?? null, viewportRequest.startRow, viewportRequest.endRow);
    } catch {
      return null;
    }
  }, [activeSheet, getRowsBatchAsync, shouldVirtualizeRows, viewportRequest, worksheet]);

  React.useEffect(() => {
    if (!shouldVirtualizeRows || !getRowsBatchAsync || !activeSheet || !viewportRequest) {
      setAsyncViewportRowBatch(null);
      return;
    }

    if (
      asyncViewportRowBatch
      && viewportRequest.startRow >= asyncViewportRowBatch.startRow
      && viewportRequest.endRow <= asyncViewportRowBatch.endRow
    ) {
      return;
    }

    let isCurrent = true;
    void getRowsBatchAsync(
      activeSheet.workbookSheetIndex,
      viewportRequest.startRow,
      viewportRequest.endRow - viewportRequest.startRow + 1
    )
      .then((rows) => {
        if (!isCurrent) {
          return;
        }

        const nextBatch = buildWorksheetBatchWindow(
          rows,
          activeSheet,
          viewportRequest.startRow,
          viewportRequest.endRow
        );
        startBatchTransition(() => {
          setAsyncViewportRowBatch(nextBatch);
        });
      })
      .catch(() => {
        if (!isCurrent) {
          return;
        }

        startBatchTransition(() => {
          setAsyncViewportRowBatch(null);
        });
      });

    return () => {
      isCurrent = false;
    };
  }, [activeSheet, asyncViewportRowBatch, getRowsBatchAsync, shouldVirtualizeRows, startBatchTransition, viewportRequest]);

  const viewportRowBatch = getRowsBatchAsync ? asyncViewportRowBatch : syncViewportRowBatch;

  React.useEffect(() => {
    setAsyncViewportRowBatch(null);
  }, [activeSheetIndex, revision]);

  const sparklinesByCell = React.useMemo(() => {
    const map = new Map<string, XlsxSheetData["sparklines"][number]>();
    for (const sparkline of activeSheet?.sparklines ?? []) {
      map.set(`${sparkline.target.row}:${sparkline.target.col}`, sparkline);
    }
    return map;
  }, [activeSheet?.sparklines]);
  const chartRangeHighlights = React.useMemo(
    () => collectChartRangeHighlights(selectedChart, selectedChartElement, sheets, workbook, paletteIsDark(palette)),
    [palette, selectedChart, selectedChartElement, sheets, workbook]
  );
  const activeSheetChartHighlights = React.useMemo(
    () => activeSheet
      ? chartRangeHighlights.filter((highlight) => highlight.workbookSheetIndex === activeSheet.workbookSheetIndex)
      : [],
    [activeSheet, chartRangeHighlights]
  );

  const cellRenderCacheInvalidationKey = React.useMemo(
    () =>
      [
        activeSheetChartHighlights,
        activeSheetIndex,
        displayColLimit,
        displayRowLimit,
        getCellStyle,
        palette,
        revision,
        selectedChartElement,
        selectedChartId,
        viewportRowBatch,
        worksheet,
        zoomFactor
      ] as const,
    [
      activeSheetChartHighlights,
      activeSheetIndex,
      displayColLimit,
      displayRowLimit,
      getCellStyle,
      palette,
      revision,
      selectedChartElement,
      selectedChartId,
      viewportRowBatch,
      worksheet,
      zoomFactor
    ]
  );
  const cellRenderCacheInvalidationRef = React.useRef<typeof cellRenderCacheInvalidationKey | null>(null);
  if (cellRenderCacheInvalidationRef.current !== cellRenderCacheInvalidationKey) {
    cellRenderCacheRef.current.clear();
    cellRenderCacheInvalidationRef.current = cellRenderCacheInvalidationKey;
  }

  const getCellData = React.useCallback((row: number, col: number): CellRenderData => {
    const cacheKey = `${row}:${col}`;
    const cached = cellRenderCacheRef.current.get(cacheKey);
    if (cached) {
      return cached;
    }

    const batchCoversRow = viewportRowBatch ? row >= viewportRowBatch.startRow && row <= viewportRowBatch.endRow : false;
    const batchedCell = batchCoversRow ? viewportRowBatch?.cells.get(cacheKey) : undefined;
    const worksheetMergedSecondary = worksheet?.isMergedSecondary(row, col) ?? false;

    if (!worksheet && !batchedCell) {
      const emptyData: CellRenderData = {
        isMergedSecondary: false,
        style: scaleCssProperties({
          backgroundColor: resolveSheetSurface(activeSheet, palette),
          borderBottom: "none",
          borderRight: "none",
          boxShadow: activeSheet?.showGridLines ? buildGridlineShadow(palette.border) : undefined,
          padding: DEFAULT_CELL_PADDING,
          verticalAlign: "bottom",
          whiteSpace: "nowrap"
        }, zoomFactor),
        value: ""
      };
      cellRenderCacheRef.current.set(cacheKey, emptyData);
      return emptyData;
    }

    if (batchedCell?.isMergedSecondary || worksheetMergedSecondary) {
      const mergedSecondaryData: CellRenderData = {
        isMergedSecondary: true,
        style: {},
        value: ""
      };
      cellRenderCacheRef.current.set(cacheKey, mergedSecondaryData);
      return mergedSecondaryData;
    }

    const merge = worksheet
      ? (worksheet.getMergeSpan(row, col) as { colSpan?: number; rowSpan?: number } | null | undefined)
      : batchCoversRow
        ? batchedCell?.mergeSpan
        : null;
    const inheritedStyle = resolveInheritedCellStyle(activeSheet, row, col);
    const worksheetStyle = batchCoversRow
      ? batchedCell?.style ?? null
      : (worksheet?.getCellStyleAt(row, col) as Record<string, unknown> | null | undefined) ?? null;
    const rawStyle = mergeResolvedCellStyle(inheritedStyle, worksheetStyle, { replaceXfSubtrees: true });
    const table = getTableAtCell(effectiveTables, row, col);
    const tableStyle = resolveTableCellStyle(table, row, col, activeSheet);
    const mergedStyle = mergeResolvedCellStyle(rawStyle, tableStyle);
    const alignment = mergedStyle?.alignment as Record<string, unknown> | undefined;
    const textRotationDeg = resolveSpreadsheetTextRotation(alignment?.textRotation);
    const headerRowCount = table ? Math.max(table.headerRowCount, 1) : 0;
    const rawHyperlink = batchCoversRow
      ? batchedCell?.hyperlink
      : (worksheet?.getHyperlinkAt(row, col) as
          | { location?: string; target?: string; tooltip?: string }
          | null
          | undefined);
    const sparkline = sparklinesByCell.get(cacheKey) ?? null;
    const sparklineValues = sparkline && worksheet
      ? (
          sparkline.range.start.row === sparkline.range.end.row
            ? Array.from(
                { length: Math.abs(sparkline.range.end.col - sparkline.range.start.col) + 1 },
                (_, index) => getCellNumericValue(
                  worksheet,
                  sparkline.range.start.row,
                  Math.min(sparkline.range.start.col, sparkline.range.end.col) + index
                )
              )
            : Array.from(
                { length: Math.abs(sparkline.range.end.row - sparkline.range.start.row) + 1 },
                (_, index) => getCellNumericValue(
                  worksheet,
                  Math.min(sparkline.range.start.row, sparkline.range.end.row) + index,
                  sparkline.range.start.col
                )
              )
        )
      : null;
    const checkboxState = mergedStyle?.cellControl && worksheet
      ? getCellBooleanValue(worksheet, row, col)
      : null;
    const chartHighlight = resolveCellChartHighlight({ row, col }, activeSheetChartHighlights);
    const nextData: CellRenderData = {
      canvas: undefined,
      chartHighlight,
      checkboxState,
      colSpan: merge?.colSpan,
      conditionalDataBar: worksheet
        ? resolveConditionalDataBarForCell(
            row,
            col,
            worksheet,
            activeSheet,
            conditionalFormatMetricsCacheRef.current
          )
        : null,
      conditionalColorScale: worksheet
        ? resolveConditionalColorScaleForCell(
            row,
            col,
            worksheet,
            activeSheet,
            conditionalFormatMetricsCacheRef.current
          )
        : null,
      conditionalIcon: worksheet
        ? resolveConditionalIconForCell(
            row,
            col,
            worksheet,
            activeSheet,
            conditionalFormatMetricsCacheRef.current
          )
        : null,
      hyperlink: rawHyperlink ?? null,
      isMergedSecondary: false,
      shrinkToFit: alignment?.shrinkToFit === true,
      isTableHeader: Boolean(table && row >= table.start.row && row < table.start.row + headerRowCount),
      rowSpan: merge?.rowSpan,
      sparkline: sparkline && sparklineValues ? { config: sparkline, values: sparklineValues } : null,
      style: scaleCssProperties(buildCellStyle(mergedStyle, palette, activeSheet?.themePalette, {
        showGridLines: activeSheet?.showGridLines
      }), zoomFactor),
      textRotationDeg,
      validation: resolveCellDataValidation(row, col, activeSheet),
      value: sparkline
        ? ""
        : checkboxState !== null
          ? ""
          : batchCoversRow || !worksheet
            ? batchedCell?.value ?? ""
            : getCellDisplayValue(worksheet, row, col, activeSheet)
    };

    if (getCellStyle) {
      const styleOverrides = getCellStyle({
        cell: { row, col },
        hasChartHighlight: Boolean(nextData.chartHighlight),
        hasConditionalFormat: Boolean(
          nextData.conditionalColorScale || nextData.conditionalDataBar || nextData.conditionalIcon
        ),
        hasHyperlink: Boolean(nextData.hyperlink),
        hasValidation: Boolean(nextData.validation),
        isMerged: Boolean(nextData.colSpan || nextData.rowSpan),
        isTableHeader: Boolean(nextData.isTableHeader),
        resolvedStyle: { ...nextData.style },
        sheetName: activeSheet?.name ?? "",
        value: nextData.value,
        workbookSheetIndex: activeSheet?.workbookSheetIndex ?? -1
      });
      if (styleOverrides) {
        nextData.style = { ...nextData.style, ...styleOverrides };
        // getCellStyle is the final styling layer, so a host-provided
        // background replaces the cell's background fill — including a
        // conditional color-scale fill, since both target the full-cell
        // background. Data bars and icon sets are overlays drawn on top and
        // are intentionally left intact.
        if (
          nextData.conditionalColorScale &&
          (styleOverrides.backgroundColor !== undefined || styleOverrides.background !== undefined)
        ) {
          nextData.conditionalColorScale = null;
        }
      }
    }

    nextData.canvas = buildCanvasCellStyleCache(nextData.style);

    if (canCellTextOverflow(nextData)) {
      const startColIndex = colIndexByActual.get(col);
      if (startColIndex !== undefined) {
        const cellColSpan = Math.max(1, nextData.colSpan ?? 1);
        const endColIndex = Math.min(visibleCols.length - 1, startColIndex + cellColSpan - 1);
        const horizontalPadding = getHorizontalPadding(nextData.style.padding);
        const textWidth = measureTextWidth(nextData.value, nextData.style);
        const requiredWidth = textWidth + horizontalPadding + 2;
        const baseWidth = Math.max(
          displayEffectiveColWidths[startColIndex] ?? displayDefaultColWidth,
          sumPrefixRange(colPrefixSums, startColIndex, endColIndex)
        );
        let availableWidth = baseWidth;

        if (requiredWidth > availableWidth) {
          for (let nextColIndex = endColIndex + 1; nextColIndex < visibleCols.length; nextColIndex += 1) {
            const nextActualCol = visibleCols[nextColIndex];
            if (nextActualCol === undefined) {
              break;
            }

            const neighborData = getCellData(row, nextActualCol);
            if (!canReceiveOverflowText(neighborData)) {
              break;
            }

            availableWidth += displayEffectiveColWidths[nextColIndex] ?? displayDefaultColWidth;
            if (requiredWidth <= availableWidth) {
              break;
            }
          }

          if (availableWidth > baseWidth) {
            nextData.spillWidth = Math.max(0, availableWidth - horizontalPadding);
          }
        }
      }
    }

    if (nextData.shrinkToFit && nextData.value.length > 0 && nextData.style.whiteSpace !== "pre-wrap") {
      const startColIndex = colIndexByActual.get(col);
      if (startColIndex !== undefined) {
        const cellColSpan = Math.max(1, nextData.colSpan ?? 1);
        const endColIndex = Math.min(visibleCols.length - 1, startColIndex + cellColSpan - 1);
        const horizontalPadding = getHorizontalPadding(nextData.style.padding);
        const trailingInset = (nextData.conditionalIcon ? 18 * zoomFactor : 0) + (nextData.isTableHeader ? 16 * zoomFactor : 0);
        const availableWidth = Math.max(
          displayEffectiveColWidths[startColIndex] ?? displayDefaultColWidth,
          sumPrefixRange(colPrefixSums, startColIndex, endColIndex)
        );
        const availableTextWidth = Math.max(0, availableWidth - horizontalPadding - trailingInset);
        if (availableTextWidth > 0) {
          const textWidth = measureTextWidth(nextData.value, nextData.style);
          if (textWidth > availableTextWidth) {
            const baseFontSizePx = resolveCanvasFontSizePx(nextData.style, 12 * zoomFactor);
            const minimumFontSizePx = Math.max(6 * zoomFactor, baseFontSizePx * 0.4);
            let lowerBound = minimumFontSizePx;
            let upperBound = baseFontSizePx;
            let bestFontSizePx = minimumFontSizePx;

            while (upperBound - lowerBound > 0.25) {
              const candidateFontSizePx = (lowerBound + upperBound) / 2;
              const candidateStyle = {
                ...nextData.style,
                font: undefined,
                fontSize: candidateFontSizePx
              } as React.CSSProperties;
              const candidateWidth = measureTextWidth(nextData.value, candidateStyle);
              if (candidateWidth <= availableTextWidth) {
                bestFontSizePx = candidateFontSizePx;
                lowerBound = candidateFontSizePx;
              } else {
                upperBound = candidateFontSizePx;
              }
            }

            nextData.shrinkToFitFontSizePx = Math.min(baseFontSizePx, bestFontSizePx);
          }
        }
      }
    }

    cellRenderCacheRef.current.set(cacheKey, nextData);
    return nextData;
  }, [
    activeSheet,
    activeSheetChartHighlights,
    cellRenderCacheInvalidationKey,
    colIndexByActual,
    colPrefixSums,
    displayDefaultColWidth,
    displayEffectiveColWidths,
    effectiveTables,
    getCellStyle,
    palette,
    sparklinesByCell,
    viewportRowBatch,
    visibleCols,
    worksheet,
    zoomFactor
  ]);

  React.useEffect(() => {
    conditionalFormatMetricsCacheRef.current.clear();
  }, [activeSheet?.conditionalFormatRules, activeSheet?.workbookSheetIndex, revision]);

  const selectionOverlay = React.useMemo(() => {
    if (!displayedSelection) {
      return null;
    }

    const normalized = normalizeRange(displayedSelection);
    const startRowIndex = rowIndexByActual.get(normalized.start.row);
    const endRowIndex = rowIndexByActual.get(normalized.end.row);
    const startColIndex = colIndexByActual.get(normalized.start.col);
    const endColIndex = colIndexByActual.get(normalized.end.col);

    if (
      startRowIndex === undefined ||
      endRowIndex === undefined ||
      startColIndex === undefined ||
      endColIndex === undefined
    ) {
      return null;
    }

    return {
      height: sumPrefixRange(rowPrefixSums, startRowIndex, endRowIndex),
      left: displayRowHeaderWidth + sumPrefixRange(colPrefixSums, 0, startColIndex - 1),
      top: displayHeaderHeight + sumPrefixRange(rowPrefixSums, 0, startRowIndex - 1),
      width: sumPrefixRange(colPrefixSums, startColIndex, endColIndex)
    };
  }, [colIndexByActual, colPrefixSums, displayHeaderHeight, displayRowHeaderWidth, displayedSelection, rowIndexByActual, rowPrefixSums]);
  const resolvedSelectionOverlay = selectionOverlay;
  const { fill: selectionFill, header: selectionHeaderSurface, stroke: selectionStroke } = React.useMemo(() => resolveSelectionColors({
    palette,
    selectionColor,
    selectionFillColor,
    selectionHeaderColor
  }), [palette, selectionColor, selectionFillColor, selectionHeaderColor]);
  const virtualCols = React.useMemo<RenderedAxisItem[]>(
    () =>
      domExpandedWindow.colIndices.map((index) => {
        const virtualCol = colVirtualItemByIndex.get(index);
        return {
          end: virtualCol?.end ?? (colPrefixSums[index + 1] ?? 0),
          index,
          key: visibleCols[index] ?? index,
          size: virtualCol?.size ?? (displayEffectiveColWidths[index] ?? displayDefaultColWidth),
          start: virtualCol?.start ?? (colPrefixSums[index] ?? 0)
        };
      }),
    [colPrefixSums, colVirtualItemByIndex, displayDefaultColWidth, displayEffectiveColWidths, domExpandedWindow.colIndices, visibleCols]
  );
  const renderedCols = React.useMemo<RenderedColumn[]>(
    () => {
      const columns: RenderedColumn[] = [];
      virtualCols.forEach((virtualCol) => {
        const actualCol = visibleCols[virtualCol.index];
        if (actualCol === undefined) {
          return;
        }

        columns.push({
          actualCol,
          key: actualCol,
          size: virtualCol.size,
          virtualIndex: virtualCol.index
        });
      });
      return columns;
    },
    [virtualCols, visibleCols]
  );
  const canvasBaseRowIndices = React.useMemo(() => {
    if (visibleRows.length === 0 || drawingViewport.height <= 0) {
      return [] as number[];
    }

    const indices = new Set<number>(frozenRowVirtualIndices);
    const viewportStart = Math.max(0, drawingViewport.top - displayHeaderHeight - CANVAS_VIEWPORT_OVERSCAN_PX);
    const viewportEnd = Math.max(
      viewportStart,
      drawingViewport.top + drawingViewport.height - displayHeaderHeight + CANVAS_VIEWPORT_OVERSCAN_PX
    );
    const startIndex = findIndexForOffsetPrefix(rowPrefixSums, viewportStart);
    const endIndex = findIndexForOffsetPrefix(rowPrefixSums, viewportEnd);
    for (let index = Math.max(0, startIndex); index <= Math.max(startIndex, endIndex); index += 1) {
      indices.add(index);
    }

    return Array.from(indices).sort((left, right) => left - right);
  }, [
    displayHeaderHeight,
    drawingViewport.height,
    drawingViewport.top,
    frozenRowVirtualIndices,
    rowPrefixSums,
    visibleRows.length
  ]);
  const canvasBaseColIndices = React.useMemo(() => {
    if (visibleCols.length === 0 || drawingViewport.width <= 0) {
      return [] as number[];
    }

    const indices = new Set<number>(frozenColVirtualIndices);
    const viewportStart = Math.max(0, drawingViewport.left - displayRowHeaderWidth - CANVAS_VIEWPORT_OVERSCAN_PX);
    const viewportEnd = Math.max(
      viewportStart,
      drawingViewport.left + drawingViewport.width - displayRowHeaderWidth + CANVAS_VIEWPORT_OVERSCAN_PX
    );
    const startIndex = findIndexForOffsetPrefix(colPrefixSums, viewportStart);
    const endIndex = findIndexForOffsetPrefix(colPrefixSums, viewportEnd);
    for (let index = Math.max(0, startIndex); index <= Math.max(startIndex, endIndex); index += 1) {
      indices.add(index);
    }

    return Array.from(indices).sort((left, right) => left - right);
  }, [
    colPrefixSums,
    displayRowHeaderWidth,
    drawingViewport.left,
    drawingViewport.width,
    frozenColVirtualIndices,
    visibleCols.length
  ]);
  const canvasExpandedWindow = React.useMemo(
    () => expandMergeAwareWindow(canvasBaseRowIndices, canvasBaseColIndices, indexedMergedRegions),
    [canvasBaseColIndices, canvasBaseRowIndices, indexedMergedRegions]
  );
  const canvasVisibleRowItems = React.useMemo(
    () =>
      canvasExpandedWindow.rowIndices.map((index) => ({
        actualRow: visibleRows[index] ?? index,
        index,
        size: displayEffectiveRowHeights[index] ?? displayDefaultRowHeight,
        start: rowPrefixSums[index] ?? 0
      })),
    [canvasExpandedWindow.rowIndices, displayDefaultRowHeight, displayEffectiveRowHeights, rowPrefixSums, visibleRows]
  );
  const canvasVisibleColItems = React.useMemo(
    () =>
      canvasExpandedWindow.colIndices.map((index) => ({
        actualCol: visibleCols[index] ?? index,
        index,
        size: displayEffectiveColWidths[index] ?? displayDefaultColWidth,
        start: colPrefixSums[index] ?? 0
      })),
    [canvasExpandedWindow.colIndices, colPrefixSums, displayDefaultColWidth, displayEffectiveColWidths, visibleCols]
  );
  const canvasPaneAxisItems = React.useMemo(() => {
    const frozenRowIndexSet = new Set(frozenRows
      .map((actualRow) => rowIndexByActual.get(actualRow))
      .filter((index): index is number => index !== undefined));
    const frozenColIndexSet = new Set(frozenCols
      .map((actualCol) => colIndexByActual.get(actualCol))
      .filter((index): index is number => index !== undefined));

    const frozenRowItems = canvasVisibleRowItems.filter((item) => frozenRowIndexSet.has(item.index));
    const scrollRowItems = canvasVisibleRowItems.filter((item) => !frozenRowIndexSet.has(item.index));
    const frozenColItems = canvasVisibleColItems.filter((item) => frozenColIndexSet.has(item.index));
    const scrollColItems = canvasVisibleColItems.filter((item) => !frozenColIndexSet.has(item.index));

    return {
      corner: {
        cols: frozenColItems,
        rows: frozenRowItems
      },
      left: {
        cols: frozenColItems,
        rows: scrollRowItems
      },
      scroll: {
        cols: scrollColItems,
        rows: scrollRowItems
      },
      top: {
        cols: scrollColItems,
        rows: frozenRowItems
      }
    } satisfies Record<FrozenDrawingPane, {
      cols: Array<{ actualCol: number; index: number; size: number; start: number }>;
      rows: Array<{ actualRow: number; index: number; size: number; start: number }>;
    }>;
  }, [canvasVisibleColItems, canvasVisibleRowItems, colIndexByActual, frozenCols, frozenRows, rowIndexByActual]);
  const totalContentWidth = colPrefixSums[colPrefixSums.length - 1] ?? 0;
  const leadingColumnSpacerWidth = shouldVirtualizeCols ? (virtualCols[0]?.start ?? 0) : 0;
  const trailingColumnSpacerWidth = shouldVirtualizeCols
    ? totalContentWidth - (virtualCols[virtualCols.length - 1]?.end ?? 0)
    : 0;
  const imageRects = React.useMemo(
    () =>
      showImages
        ? images.map((image) => ({
            image,
            rect:
              imagePreviewRect && imagePreviewRect.id === image.id
                ? imagePreviewRect.rect
                : resolveImageRect(image, visibleRows, visibleCols, displayEffectiveRowHeights, displayEffectiveColWidths, {
                    actualColPrefixSums,
                    actualRowPrefixSums,
                    colIndexByActual,
                    colPrefixSums,
                    headerHeight: displayHeaderHeight,
                    rowIndexByActual,
                    rowHeaderWidth: displayRowHeaderWidth,
                    rowPrefixSums
                  })
          }))
        : [],
    [
      colIndexByActual,
      colPrefixSums,
      actualColPrefixSums,
      actualRowPrefixSums,
      displayHeaderHeight,
      displayEffectiveColWidths,
      displayEffectiveRowHeights,
      displayRowHeaderWidth,
      imagePreviewRect,
      images,
      rowIndexByActual,
      rowPrefixSums,
      showImages,
      visibleCols,
      visibleRows
    ]
  );
  const shapeRects = React.useMemo(
    () =>
      showImages
        ? shapes.map((shape) => ({
            rect: resolveAnchoredRect(shape.anchor, visibleRows, visibleCols, displayEffectiveRowHeights, displayEffectiveColWidths, {
              actualColPrefixSums,
              actualRowPrefixSums,
              colIndexByActual,
              colPrefixSums,
              headerHeight: displayHeaderHeight,
              rowIndexByActual,
              rowHeaderWidth: displayRowHeaderWidth,
              rowPrefixSums
            }),
            shape
          }))
        : [],
    [
      colIndexByActual,
      colPrefixSums,
      actualColPrefixSums,
      actualRowPrefixSums,
      displayHeaderHeight,
      displayEffectiveColWidths,
      displayEffectiveRowHeights,
      displayRowHeaderWidth,
      rowIndexByActual,
      rowPrefixSums,
      shapes,
      showImages,
      visibleCols,
      visibleRows
    ]
  );
  const formControlRects = React.useMemo(
    () =>
      showImages
        ? formControls
          .filter((control) => !control.hidden)
          .map((control) => ({
            control,
            rect: resolveAnchoredRect(control.anchor, visibleRows, visibleCols, displayEffectiveRowHeights, displayEffectiveColWidths, {
              actualColPrefixSums,
              actualRowPrefixSums,
              colIndexByActual,
              colPrefixSums,
              headerHeight: displayHeaderHeight,
              rowIndexByActual,
              rowHeaderWidth: displayRowHeaderWidth,
              rowPrefixSums
            })
          }))
        : [],
    [
      actualColPrefixSums,
      actualRowPrefixSums,
      colIndexByActual,
      colPrefixSums,
      displayHeaderHeight,
      displayEffectiveColWidths,
      displayEffectiveRowHeights,
      displayRowHeaderWidth,
      formControls,
      rowIndexByActual,
      rowPrefixSums,
      showImages,
      visibleCols,
      visibleRows
    ]
  );
  const chartRects = React.useMemo(
    () =>
      showImages
        ? charts.map((chart) => ({
            chart,
            rect:
              chartPreviewRect && chartPreviewRect.id === chart.id
                ? chartPreviewRect.rect
                : resolveAnchoredRect(chart.anchor, visibleRows, visibleCols, displayEffectiveRowHeights, displayEffectiveColWidths, {
                    actualColPrefixSums,
                    actualRowPrefixSums,
                    colIndexByActual,
                    colPrefixSums,
                    headerHeight: displayHeaderHeight,
                    rowIndexByActual,
                    rowHeaderWidth: displayRowHeaderWidth,
                    rowPrefixSums
                  })
          }))
        : [],
    [
      actualColPrefixSums,
      actualRowPrefixSums,
      chartPreviewRect,
      charts,
      colIndexByActual,
      colPrefixSums,
      displayHeaderHeight,
      displayEffectiveColWidths,
      displayEffectiveRowHeights,
      displayRowHeaderWidth,
      rowIndexByActual,
      rowPrefixSums,
      showImages,
      visibleCols,
      visibleRows
    ]
  );
  const shouldBakeCanvasStaticDrawings = Boolean(experimentalCanvas && readOnly && showImages);
  const shouldBakeCanvasImages = Boolean(shouldBakeCanvasStaticDrawings && !renderImage && !renderImageSelection);
  const bakedShapeRects = React.useMemo(
    () => (shouldBakeCanvasStaticDrawings ? shapeRects.filter(({ shape }) => !shape.hyperlink) : []),
    [shapeRects, shouldBakeCanvasStaticDrawings]
  );
  const domShapeRects = React.useMemo(
    () => (shouldBakeCanvasStaticDrawings ? shapeRects.filter(({ shape }) => shape.hyperlink) : shapeRects),
    [shapeRects, shouldBakeCanvasStaticDrawings]
  );
  const bakedFormControlRects = React.useMemo(
    () => (shouldBakeCanvasStaticDrawings ? formControlRects : []),
    [formControlRects, shouldBakeCanvasStaticDrawings]
  );
  const domFormControlRects = React.useMemo(
    () => (shouldBakeCanvasStaticDrawings ? [] : formControlRects),
    [formControlRects, shouldBakeCanvasStaticDrawings]
  );
  const bakedImageRects = React.useMemo(
    () => (
      shouldBakeCanvasImages
        ? imageRects.filter(({ image }) => !image.hyperlink && selectedImageId !== image.id)
        : []
    ),
    [imageRects, selectedImageId, shouldBakeCanvasImages]
  );
  const domImageRects = React.useMemo(
    () => (
      shouldBakeCanvasImages
        ? imageRects.filter(({ image }) => image.hyperlink || selectedImageId === image.id)
        : imageRects
    ),
    [imageRects, selectedImageId, shouldBakeCanvasImages]
  );
  const hasCanvasDomDrawingOverlays = Boolean(
    showImages
    && (
      chartRects.length > 0
      || domShapeRects.length > 0
      || domFormControlRects.length > 0
      || domImageRects.length > 0
    )
  );
  const bakedCanvasDrawingSignature = React.useMemo(() => {
    if (!shouldBakeCanvasStaticDrawings) {
      return "";
    }

    const shapeSignature = bakedShapeRects
      .map(({ rect, shape }) => `s:${shape.id}:${shape.zIndex}:${Math.round(rect.left)}:${Math.round(rect.top)}:${Math.round(rect.width)}:${Math.round(rect.height)}`)
      .join("|");
    const controlSignature = bakedFormControlRects
      .map(({ control, rect }) => `f:${control.id}:${control.zIndex}:${Boolean(control.checked)}:${Math.round(rect.left)}:${Math.round(rect.top)}:${Math.round(rect.width)}:${Math.round(rect.height)}`)
      .join("|");
    const imageSignature = bakedImageRects
      .map(({ image, rect }) => `i:${image.id}:${image.src}:${image.zIndex}:${Math.round(rect.left)}:${Math.round(rect.top)}:${Math.round(rect.width)}:${Math.round(rect.height)}`)
      .join("|");
    return `${revision}:${canvasImageLoadVersion}:${shapeSignature}:${controlSignature}:${imageSignature}`;
  }, [
    bakedFormControlRects,
    bakedImageRects,
    bakedShapeRects,
    canvasImageLoadVersion,
    revision,
    shouldBakeCanvasStaticDrawings
  ]);

  const resolveMountedCellOverlayRect = React.useCallback((element: HTMLElement) => {
    const wrapper = wrapperRef.current;
    if (!wrapper) {
      return null;
    }

    const wrapperRect = wrapper.getBoundingClientRect();
    const elementRect = element.getBoundingClientRect();
    const logicalWidth = wrapper.offsetWidth || 1;
    const logicalHeight = wrapper.offsetHeight || 1;
    const scaleX = Math.max(0.0001, wrapperRect.width > 0 ? wrapperRect.width / logicalWidth : 1);
    const scaleY = Math.max(0.0001, wrapperRect.height > 0 ? wrapperRect.height / logicalHeight : 1);

    return {
      height: elementRect.height / scaleY,
      left: (elementRect.left - wrapperRect.left) / scaleX,
      top: (elementRect.top - wrapperRect.top) / scaleY,
      width: elementRect.width / scaleX
    };
  }, []);

  const resolveMountedCellOverlayRectForAddress = React.useCallback((cell: XlsxCellAddress) => {
    const wrapper = wrapperRef.current;
    if (!wrapper) {
      return null;
    }

    const element = wrapper.querySelector<HTMLElement>(`[data-xlsx-cell="${cell.row}:${cell.col}"]`);
    if (!element) {
      return null;
    }

    return resolveMountedCellOverlayRect(element);
  }, [resolveMountedCellOverlayRect]);

  const resolveGeometryOverlayRect = React.useCallback((range: XlsxCellRange) => {
    const normalized = normalizeRange(range);
    const isSingleCellSelection = normalized.start.row === normalized.end.row && normalized.start.col === normalized.end.col;
    const isMergedSecondarySelection = isSingleCellSelection
      ? worksheet?.isMergedSecondary(normalized.start.row, normalized.start.col) === true
      : false;
    const startCell = isSingleCellSelection && !isMergedSecondarySelection
      ? resolveMergeAnchorCell(normalized.start)
      : normalized.start;
    const merge = isSingleCellSelection && !isMergedSecondarySelection
      ? (worksheet?.getMergeSpan(startCell.row, startCell.col) as { colSpan?: number; rowSpan?: number } | null | undefined)
      : null;
    const endCell = isSingleCellSelection && !isMergedSecondarySelection
      ? {
          row: startCell.row + Math.max(1, merge?.rowSpan ?? 1) - 1,
          col: startCell.col + Math.max(1, merge?.colSpan ?? 1) - 1
        }
      : normalized.end;
    const startRowIndex = rowIndexByActual.get(startCell.row);
    const endRowIndex = rowIndexByActual.get(endCell.row);
    const startColIndex = colIndexByActual.get(startCell.col);
    const endColIndex = colIndexByActual.get(endCell.col);

    if (
      startRowIndex === undefined ||
      endRowIndex === undefined ||
      startColIndex === undefined ||
      endColIndex === undefined
    ) {
      return null;
    }

    let left = displayRowHeaderWidth + sumPrefixRange(colPrefixSums, 0, startColIndex - 1);
    let top = displayHeaderHeight + sumPrefixRange(rowPrefixSums, 0, startRowIndex - 1);
    let width = sumPrefixRange(colPrefixSums, startColIndex, endColIndex);
    let height = sumPrefixRange(rowPrefixSums, startRowIndex, endRowIndex);

    const columnPreview = columnPreviewRef.current;
    if (columnPreview) {
      const previewIndex = colIndexByActual.get(columnPreview.actualIndex);
      if (previewIndex !== undefined) {
        const baseWidth = displayEffectiveColWidths[previewIndex] ?? displayDefaultColWidth;
        const widthDelta = columnPreview.size - baseWidth;
        if (previewIndex < startColIndex) {
          left += widthDelta;
        }
        if (previewIndex >= startColIndex && previewIndex <= endColIndex) {
          width += widthDelta;
        }
      }
    }

    const rowPreview = rowPreviewRef.current;
    if (rowPreview) {
      const previewIndex = rowIndexByActual.get(rowPreview.actualIndex);
      if (previewIndex !== undefined) {
        const baseHeight = displayEffectiveRowHeights[previewIndex] ?? displayDefaultRowHeight;
        const heightDelta = rowPreview.size - baseHeight;
        if (previewIndex < startRowIndex) {
          top += heightDelta;
        }
        if (previewIndex >= startRowIndex && previewIndex <= endRowIndex) {
          height += heightDelta;
        }
      }
    }

    return {
      height: Math.max(0, height),
      left: Math.max(displayRowHeaderWidth, left),
      top: Math.max(displayHeaderHeight, top),
      width: Math.max(0, width)
    };
  }, [colIndexByActual, colPrefixSums, displayDefaultColWidth, displayDefaultRowHeight, displayEffectiveColWidths, displayEffectiveRowHeights, displayHeaderHeight, displayRowHeaderWidth, resolveMergeAnchorCell, rowIndexByActual, rowPrefixSums, worksheet]);

  const resolveMountedRangeOverlayRect = React.useCallback((range: XlsxCellRange, geometryRect: {
    height: number;
    left: number;
    top: number;
    width: number;
  }) => {
    const normalized = normalizeRange(range);
    const startRect = resolveMountedCellOverlayRectForAddress(normalized.start);
    const topRightRect = resolveMountedCellOverlayRectForAddress({ row: normalized.start.row, col: normalized.end.col });
    const bottomLeftRect = resolveMountedCellOverlayRectForAddress({ row: normalized.end.row, col: normalized.start.col });
    const endRect = resolveMountedCellOverlayRectForAddress(normalized.end);

    const leftRect = startRect ?? bottomLeftRect;
    const topRect = startRect ?? topRightRect;
    const rightRect = topRightRect ?? endRect;
    const bottomRect = bottomLeftRect ?? endRect;

    const left = leftRect ? leftRect.left : geometryRect.left;
    const top = topRect ? topRect.top : geometryRect.top;
    const right = rightRect ? rightRect.left + rightRect.width : geometryRect.left + geometryRect.width;
    const bottom = bottomRect ? bottomRect.top + bottomRect.height : geometryRect.top + geometryRect.height;

    return {
      height: Math.max(0, bottom - top),
      left: Math.max(displayRowHeaderWidth, left),
      top: Math.max(displayHeaderHeight, top),
      width: Math.max(0, right - left)
    };
  }, [displayHeaderHeight, displayRowHeaderWidth, resolveMountedCellOverlayRectForAddress]);

  const resolveDragPreviewRect = React.useCallback((range: XlsxCellRange) => {
    const dragState = selectionDragRef.current;
    if (!dragState || !dragState.didDrag || dragState.axis !== "cell" || !dragState.originOverlayRect) {
      return null;
    }

    const rangeRect = resolveGeometryOverlayRect(range);
    if (!rangeRect) {
      return null;
    }

    const originRect = dragState.originOverlayRect;
    const normalized = normalizeRange(range);
    if (rangesEqual(normalized, { start: dragState.originCell, end: dragState.originCell })) {
      return originRect;
    }

    const originCell = dragState.originCell;
    const originOnLeft = originCell.col === normalized.start.col;
    const originOnTop = originCell.row === normalized.start.row;
    const adjustedRangeRect = resolveMountedRangeOverlayRect(range, rangeRect);
    const geometryRight = adjustedRangeRect.left + adjustedRangeRect.width;
    const geometryBottom = adjustedRangeRect.top + adjustedRangeRect.height;
    const originRight = originRect.left + originRect.width;
    const originBottom = originRect.top + originRect.height;
    const left = originOnLeft ? originRect.left : Math.min(originRect.left, adjustedRangeRect.left);
    const top = originOnTop ? originRect.top : Math.min(originRect.top, adjustedRangeRect.top);
    const right = originOnLeft ? Math.max(originRight, geometryRight) : originRight;
    const bottom = originOnTop ? Math.max(originBottom, geometryBottom) : originBottom;

    return {
      height: Math.max(originRect.height, bottom - top),
      left,
      top,
      width: Math.max(originRect.width, right - left)
    };
  }, [resolveGeometryOverlayRect, resolveMountedRangeOverlayRect]);

  const resolveOverlayRect = React.useCallback((range: XlsxCellRange) => {
    const normalized = normalizeRange(range);
    if (normalized.start.row === normalized.end.row && normalized.start.col === normalized.end.col) {
      const rect = resolveMountedCellOverlayRectForAddress(normalized.start);
      if (rect) {
        return rect;
      }
    }

    const geometryRect = resolveGeometryOverlayRect(range);
    if (geometryRect) {
      return resolveMountedRangeOverlayRect(range, geometryRect);
    }

    return null;
  }, [
    resolveGeometryOverlayRect,
    resolveMountedRangeOverlayRect,
  ]);

  const resolveCellDisplayRect = React.useCallback((cell: XlsxCellAddress) => {
    const rowIndex = rowIndexByActual.get(cell.row);
    const colIndex = colIndexByActual.get(cell.col);
    if (rowIndex === undefined || colIndex === undefined) {
      return null;
    }

    const cellData = getCellData(cell.row, cell.col);
    const colSpan = Math.max(1, cellData.colSpan ?? 1);
    const rowSpan = Math.max(1, cellData.rowSpan ?? 1);
    let endActualCol = Math.min(displayColLimit - 1, cell.col + colSpan - 1);
    let endActualRow = Math.min(displayRowLimit - 1, cell.row + rowSpan - 1);

    if (worksheet && (colSpan > 1 || rowSpan > 1)) {
      for (let nextCol = cell.col + 1; nextCol < displayColLimit; nextCol += 1) {
        const nextAnchor = resolveMergeAnchorCell({ row: cell.row, col: nextCol });
        if (nextAnchor.row !== cell.row || nextAnchor.col !== cell.col) {
          break;
        }
        endActualCol = nextCol;
      }
      for (let nextRow = cell.row + 1; nextRow < displayRowLimit; nextRow += 1) {
        const nextAnchor = resolveMergeAnchorCell({ row: nextRow, col: cell.col });
        if (nextAnchor.row !== cell.row || nextAnchor.col !== cell.col) {
          break;
        }
        endActualRow = nextRow;
      }
    }

    return {
      height: sumPrefixRange(actualRowPrefixSums, cell.row, endActualRow),
      left: displayRowHeaderWidth + sumPrefixRange(actualColPrefixSums, 0, cell.col - 1),
      top: displayHeaderHeight + sumPrefixRange(actualRowPrefixSums, 0, cell.row - 1),
      width: sumPrefixRange(actualColPrefixSums, cell.col, endActualCol)
    };
  }, [
    actualColPrefixSums,
    actualRowPrefixSums,
    colIndexByActual,
    displayColLimit,
    displayHeaderHeight,
    displayRowLimit,
    displayRowHeaderWidth,
    getCellData,
    resolveMergeAnchorCell,
    rowIndexByActual,
    worksheet
  ]);

  const openTableMenuState = React.useMemo(() => {
    if (!openTableMenu) {
      return null;
    }

    const table = effectiveTables.find((entry) => entry.name === openTableMenu.tableName) ?? null;
    const column = getTableHeaderColumn(table, openTableMenu.row, openTableMenu.col);
    const wrapper = wrapperRef.current;
    if (!table || !column || !wrapper) {
      return null;
    }

    const cell = wrapper.querySelector<HTMLElement>(`[data-xlsx-cell="${openTableMenu.row}:${openTableMenu.col}"]`);
    if (cell) {
      return {
        column,
        left: cell.offsetLeft + cell.offsetWidth - 170,
        table,
        top: cell.offsetTop + cell.offsetHeight - 2
      };
    }

    const rect = resolveCellDisplayRect({ row: openTableMenu.row, col: openTableMenu.col });
    if (!rect) {
      return null;
    }

    return {
      column,
      left: rect.left + rect.width - 170,
      table,
      top: rect.top + rect.height - 2
    };
  }, [effectiveTables, openTableMenu, resolveCellDisplayRect]);

  const applyHeaderSelection = React.useCallback((range: XlsxCellRange | null) => {
    if (experimentalCanvas) {
      return;
    }

    const normalized = range ? normalizeRange(range) : null;

    rowHeaderCellRefs.current.forEach((element, actualRow) => {
      element.style.backgroundColor = normalized && actualRow >= normalized.start.row && actualRow <= normalized.end.row
        ? "var(--xlsx-selection-header)"
        : palette.rowHeaderSurface;
    });

    colHeaderCellRefs.current.forEach((element, actualCol) => {
      element.style.backgroundColor = normalized && actualCol >= normalized.start.col && actualCol <= normalized.end.col
        ? "var(--xlsx-selection-header)"
        : palette.headerSurface;
    });
  }, [experimentalCanvas, palette.headerSurface, palette.rowHeaderSurface]);

  const setRowHeaderRef = React.useCallback((actualRow: number, element: HTMLTableCellElement | null) => {
    if (experimentalCanvas) {
      rowHeaderCellRefs.current.delete(actualRow);
      return;
    }

    if (element) {
      rowHeaderCellRefs.current.set(actualRow, element);
      const range = selectionPreviewRangeRef.current ?? displayedSelectionRef.current;
      element.style.backgroundColor = range && actualRow >= range.start.row && actualRow <= range.end.row
        ? "var(--xlsx-selection-header)"
        : palette.rowHeaderSurface;
      return;
    }

    rowHeaderCellRefs.current.delete(actualRow);
  }, [experimentalCanvas, palette.rowHeaderSurface]);

  const setColHeaderRef = React.useCallback((actualCol: number, element: HTMLTableCellElement | null) => {
    if (experimentalCanvas) {
      colHeaderCellRefs.current.delete(actualCol);
      return;
    }

    if (element) {
      colHeaderCellRefs.current.set(actualCol, element);
      const range = selectionPreviewRangeRef.current ?? displayedSelectionRef.current;
      element.style.backgroundColor = range && actualCol >= range.start.col && actualCol <= range.end.col
        ? "var(--xlsx-selection-header)"
        : palette.headerSurface;
      return;
    }

    colHeaderCellRefs.current.delete(actualCol);
  }, [experimentalCanvas, palette.headerSurface]);

  const applyPreviewOverlay = React.useCallback((range: XlsxCellRange | null) => {
    const overlay = selectionOverlayRef.current;
    if (!overlay || !range) {
      applyHeaderSelection(range);
      return;
    }

    const nextRect =
      selectionDragRef.current?.didDrag || fillDragRef.current
        ? resolveDragPreviewRect(range) ?? resolveGeometryOverlayRect(range) ?? resolveOverlayRect(range)
        : resolveOverlayRect(range);
    if (!nextRect) {
      applyHeaderSelection(range);
      return;
    }

    overlay.style.left = `${nextRect.left}px`;
    overlay.style.top = `${nextRect.top}px`;
    overlay.style.width = `${nextRect.width}px`;
    overlay.style.height = `${nextRect.height}px`;
    overlay.style.opacity = "1";
    overlay.style.visibility = "visible";
    const fillHandle = fillHandleRef.current;
    if (fillHandle) {
      fillHandle.style.left = `${nextRect.left + nextRect.width - (4 * zoomFactor)}px`;
      fillHandle.style.top = `${nextRect.top + nextRect.height - (4 * zoomFactor)}px`;
    }
    applyHeaderSelection(range);
  }, [applyHeaderSelection, resolveDragPreviewRect, resolveGeometryOverlayRect, resolveOverlayRect, zoomFactor]);

  const applyPreviewOverlayFromElement = React.useCallback((element: HTMLElement, range: XlsxCellRange) => {
    const overlay = selectionOverlayRef.current;
    if (!overlay) {
      return;
    }

    const nextRect = resolveMountedCellOverlayRect(element) ?? resolveOverlayRect(range);
    if (!nextRect) {
      applyHeaderSelection(range);
      return;
    }

    overlay.style.left = `${nextRect.left}px`;
    overlay.style.top = `${nextRect.top}px`;
    overlay.style.width = `${nextRect.width}px`;
    overlay.style.height = `${nextRect.height}px`;
    overlay.style.opacity = "1";
    overlay.style.visibility = "visible";
    const fillHandle = fillHandleRef.current;
    if (fillHandle) {
      fillHandle.style.left = `${nextRect.left + nextRect.width - (4 * zoomFactor)}px`;
      fillHandle.style.top = `${nextRect.top + nextRect.height - (4 * zoomFactor)}px`;
    }
    applyHeaderSelection(range);
  }, [applyHeaderSelection, resolveMountedCellOverlayRect, resolveOverlayRect, zoomFactor]);

  const syncActiveValidationOverlay = React.useCallback((cell: XlsxCellAddress | null) => {
    const overlay = activeValidationOverlayRef.current;
    if (!overlay || !cell || editingCellRef.current || selectionDragRef.current || fillDragRef.current) {
      if (overlay) {
        overlay.style.opacity = "0";
        overlay.style.visibility = "hidden";
      }
      return;
    }

    const cellData = getCellData(cell.row, cell.col);
    const shouldShow = cellData.validation?.validationType === "list" && cellData.validation.showDropdown;
    const rect = shouldShow ? resolveOverlayRect({ start: cell, end: cell }) : null;
    if (!rect) {
      overlay.style.opacity = "0";
      overlay.style.visibility = "hidden";
      return;
    }

    overlay.style.left = `${rect.left + rect.width - (16 * zoomFactor)}px`;
    overlay.style.top = `${rect.top + (rect.height / 2)}px`;
    overlay.style.opacity = "1";
    overlay.style.visibility = "visible";
  }, [getCellData, resolveOverlayRect, zoomFactor]);

  const commitSelectionRange = React.useCallback((range: XlsxCellRange) => {
    gridKeyboardActiveRef.current = true;
    const normalized = normalizeRange(range);
    if (
      selectionRef.current &&
      rangesEqual(selectionRef.current, normalized) &&
      isSameCell(activeCellRef.current, normalized.end) &&
      selectedChartIdRef.current === null &&
      selectedImageIdRef.current === null
    ) {
      return;
    }

    pendingSelectionCommitRef.current = normalized;
    if (selectionCommitFrameRef.current !== null) {
      return;
    }

    selectionCommitFrameRef.current = window.requestAnimationFrame(() => {
      selectionCommitFrameRef.current = null;
      const pendingRange = pendingSelectionCommitRef.current;
      pendingSelectionCommitRef.current = null;
      if (!pendingRange) {
        return;
      }

      if (
        selectionRef.current &&
        rangesEqual(selectionRef.current, pendingRange) &&
        isSameCell(activeCellRef.current, pendingRange.end) &&
        selectedChartIdRef.current === null &&
        selectedImageIdRef.current === null
      ) {
        return;
      }

      startSelectionTransition(() => {
        selectRange(pendingRange);
      });
    });
  }, [selectRange, startSelectionTransition]);

  const refreshOverlayFromCurrentSelection = React.useCallback(() => {
    if (displayedSelectionRef.current) {
      applyPreviewOverlay(displayedSelectionRef.current);
    }
  }, [applyPreviewOverlay]);
  const resolveResizeGuidePositionFromClient = React.useCallback((
    type: "column" | "row",
    clientPosition: number
  ) => {
    const scroller = scrollRef.current;
    const scrollerRect = scroller?.getBoundingClientRect();
    if (!scroller || !scrollerRect) {
      return null;
    }

    if (type === "column") {
      return scroller.scrollLeft + (clientPosition - scrollerRect.left);
    }

    return scroller.scrollTop + (clientPosition - scrollerRect.top);
  }, []);

  React.useEffect(() => {
    function handlePointerMove(event: PointerEvent) {
      const state = resizeStateRef.current;
      if (!state) {
        return;
      }

      const size =
        state.type === "column"
          ? Math.max(displayDefaultColWidth / 2, state.initialPx + (event.clientX - state.startPosition))
          : Math.max(displayDefaultRowHeight / 1.5, state.initialPx + (event.clientY - state.startPosition));
      pendingResizePreviewRef.current = {
        actualIndex: state.actualIndex,
        size,
        type: state.type
      };

      if (resizeFrameRef.current !== null) {
        return;
      }

      resizeFrameRef.current = window.requestAnimationFrame(() => {
        resizeFrameRef.current = null;
        const preview = pendingResizePreviewRef.current;
        if (!preview) {
          return;
        }

        if (preview.type === "column") {
          columnPreviewRef.current = { actualIndex: preview.actualIndex, size: preview.size };
          rowPreviewRef.current = null;
          const position = resolveResizeGuidePositionFromClient("column", event.clientX);
          setResizeGuide(position === null ? null : { position, type: "column" });
          setGlobalCursor("col-resize");
          return;
        }

        rowPreviewRef.current = { actualIndex: preview.actualIndex, size: preview.size };
        columnPreviewRef.current = null;
        const position = resolveResizeGuidePositionFromClient("row", event.clientY);
        setResizeGuide(position === null ? null : { position, type: "row" });
        setGlobalCursor("row-resize");
      });
    }

    function handlePointerUp(event: PointerEvent) {
      if (resizeStateRef.current?.pointerId === event.pointerId) {
        const resizeState = resizeStateRef.current;
        const preview = pendingResizePreviewRef.current;
        resizeStateRef.current = null;
        pendingResizePreviewRef.current = null;
        columnPreviewRef.current = null;
        rowPreviewRef.current = null;
        setInteractionMode("idle");
        setResizeGuide(null);
        clearGlobalCursor();
        document.body.style.userSelect = "";

        if (resizeFrameRef.current !== null) {
          window.cancelAnimationFrame(resizeFrameRef.current);
          resizeFrameRef.current = null;
        }

        if (preview && preview.actualIndex === resizeState.actualIndex && preview.type === resizeState.type) {
          if (preview.type === "column") {
            controller.resizeColumn(preview.actualIndex, preview.size / zoomFactor);
          } else {
            controller.resizeRow(preview.actualIndex, preview.size / zoomFactor);
          }
        }
      }
    }

    window.addEventListener("pointermove", handlePointerMove);
    window.addEventListener("pointerup", handlePointerUp);
    window.addEventListener("pointercancel", handlePointerUp);

    return () => {
      window.removeEventListener("pointermove", handlePointerMove);
      window.removeEventListener("pointerup", handlePointerUp);
      window.removeEventListener("pointercancel", handlePointerUp);
      clearGlobalCursor();
      if (resizeFrameRef.current !== null) {
        window.cancelAnimationFrame(resizeFrameRef.current);
        resizeFrameRef.current = null;
      }
    };
  }, [
    clearGlobalCursor,
    controller,
    displayDefaultColWidth,
    displayDefaultRowHeight,
    resolveResizeGuidePositionFromClient,
    setGlobalCursor,
    zoomFactor
  ]);

  function buildDraggedSelectionRange(
    dragState: { anchor: XlsxCellAddress; axis: "cell" | "column" | "row" },
    cell: XlsxCellAddress
  ): XlsxCellRange | null {
    if (dragState.axis === "row") {
      if (firstVisibleCol === undefined || lastVisibleCol === undefined) {
        return null;
      }

      return normalizeRange({
        start: { row: dragState.anchor.row, col: firstVisibleCol },
        end: { row: cell.row, col: lastVisibleCol }
      });
    }

    if (dragState.axis === "column") {
      if (firstVisibleRow === undefined || lastVisibleRow === undefined) {
        return null;
      }

      return normalizeRange({
        start: { row: firstVisibleRow, col: dragState.anchor.col },
        end: { row: lastVisibleRow, col: cell.col }
      });
    }

    return normalizeRange({ start: dragState.anchor, end: cell });
  }

  function updateSelectionDragPreview(clientX: number, clientY: number) {
    const dragState = selectionDragRef.current;
    if (!dragState) {
      return;
    }

    const nextCell = resolveDraggedSelectionCell(dragState, clientX, clientY);
    if (!nextCell) {
      return;
    }

    const nextRange = buildDraggedSelectionRange(dragState, nextCell);
    if (!nextRange || rangesEqual(nextRange, dragState.previewRange)) {
      return;
    }

    dragState.previewRange = nextRange;
    selectionPreviewRangeRef.current = nextRange;
    displayedSelectionRef.current = nextRange;
    applyPreviewOverlay(nextRange);
  }

  function clearPendingSelectionDrag() {
    const cleanup = pendingSelectionDragCleanupRef.current;
    pendingSelectionDragCleanupRef.current = null;
    pendingSelectionDragRef.current = null;
    cleanup?.();
  }

  function finishPendingSelectionDrag() {
    const pendingState = pendingSelectionDragRef.current;
    clearPendingSelectionDrag();
    if (pendingState && !pendingState.committedOnPointerDown) {
      commitSelectionRange(pendingState.previewRange);
    }
  }

  function promotePendingSelectionDrag(clientX: number, clientY: number) {
    const pendingState = pendingSelectionDragRef.current;
    if (!pendingState) {
      return;
    }

    clearPendingSelectionDrag();
    cachedScrollerRectRef.current = scrollRef.current?.getBoundingClientRect() ?? null;
    pendingState.didDrag = true;
    selectionDragRef.current = pendingState;
    selectionPreviewRangeRef.current = pendingState.previewRange;
    displayedSelectionRef.current = pendingState.previewRange;
    setInteractionMode("select");
    document.body.style.userSelect = "none";
    installSelectionDragListeners(pendingState.pointerId);
    updateSelectionDragPreview(clientX, clientY);
  }

  function installPendingSelectionDragListeners(pointerId: number, target: Element) {
    const existingCleanup = pendingSelectionDragCleanupRef.current;
    pendingSelectionDragCleanupRef.current = null;
    existingCleanup?.();

    const pointerTarget = target as Element & {
      hasPointerCapture?: (pointerId: number) => boolean;
      releasePointerCapture?: (pointerId: number) => void;
      setPointerCapture?: (pointerId: number) => void;
    };

    try {
      pointerTarget.setPointerCapture?.(pointerId);
    } catch {
      // Pointer capture can fail if the initiating element unmounts before the browser applies it.
    }

    const handlePointerMove = (event: Event) => {
      const pointerEvent = event as PointerEvent;
      if (pointerEvent.pointerId !== pointerId) {
        return;
      }

      const pendingState = pendingSelectionDragRef.current;
      if (!pendingState) {
        return;
      }

      const deltaX = Math.abs(pointerEvent.clientX - pendingState.startClientX);
      const deltaY = Math.abs(pointerEvent.clientY - pendingState.startClientY);
      if (deltaX < SELECTION_DRAG_THRESHOLD_PX && deltaY < SELECTION_DRAG_THRESHOLD_PX) {
        return;
      }

      event.preventDefault();
      promotePendingSelectionDrag(pointerEvent.clientX, pointerEvent.clientY);
    };

    const handlePointerEnd = (event: Event) => {
      const pointerEvent = event as PointerEvent;
      if (pointerEvent.pointerId !== pointerId) {
        return;
      }

      finishPendingSelectionDrag();
    };

    target.addEventListener("pointermove", handlePointerMove);
    target.addEventListener("pointerup", handlePointerEnd);
    target.addEventListener("pointercancel", handlePointerEnd);
    pendingSelectionDragCleanupRef.current = () => {
      target.removeEventListener("pointermove", handlePointerMove);
      target.removeEventListener("pointerup", handlePointerEnd);
      target.removeEventListener("pointercancel", handlePointerEnd);
      try {
        if (pointerTarget.hasPointerCapture?.(pointerId)) {
          pointerTarget.releasePointerCapture?.(pointerId);
        }
      } catch {
        // Pointer capture may already be released by the browser on pointerup/cancel.
      }
    };
  }

  function installSelectionDragListeners(pointerId: number) {
    selectionDragCleanupRef.current?.();
    let pendingClientPoint: { x: number; y: number } | null = null;
    let pointerMoveFrame: number | null = null;

    const flushPointerMove = () => {
      pointerMoveFrame = null;
      const pendingPoint = pendingClientPoint;
      pendingClientPoint = null;
      if (!pendingPoint) {
        return;
      }

      const dragState = selectionDragRef.current;
      if (!dragState) {
        return;
      }

      if (!dragState.didDrag) {
        const deltaX = Math.abs(pendingPoint.x - dragState.startClientX);
        const deltaY = Math.abs(pendingPoint.y - dragState.startClientY);
        if (deltaX < SELECTION_DRAG_THRESHOLD_PX && deltaY < SELECTION_DRAG_THRESHOLD_PX) {
          return;
        }

        dragState.didDrag = true;
      }

      updateSelectionDragPreview(pendingPoint.x, pendingPoint.y);
    };

    const handlePointerMove = (event: PointerEvent) => {
      if (event.pointerId !== pointerId) {
        return;
      }
      pendingClientPoint = { x: event.clientX, y: event.clientY };
      if (pointerMoveFrame !== null) {
        return;
      }
      pointerMoveFrame = window.requestAnimationFrame(flushPointerMove);
    };

    const finishSelectionDrag = (event: PointerEvent) => {
      if (event.pointerId !== pointerId) {
        return;
      }
      if (pointerMoveFrame !== null) {
        window.cancelAnimationFrame(pointerMoveFrame);
        pointerMoveFrame = null;
      }
      pendingClientPoint = null;

      const dragState = selectionDragRef.current;
      const nextCell = dragState ? resolveDraggedSelectionCell(dragState, event.clientX, event.clientY) : null;
      let nextRange = dragState?.previewRange ?? null;
      if (dragState?.didDrag && nextCell && dragState) {
        nextRange = buildDraggedSelectionRange(dragState, nextCell);
      }

      selectionDragRef.current = null;
      cachedScrollerRectRef.current = null;
      if (nextRange && (dragState?.didDrag || !dragState?.committedOnPointerDown)) {
        selectionPreviewRangeRef.current = nextRange;
        displayedSelectionRef.current = nextRange;
        if (dragState?.axis === "column") {
          const normalized = normalizeRange(nextRange);
          axisSelectionRef.current = {
            axis: "column",
            startCol: normalized.start.col,
            endCol: normalized.end.col
          };
        } else if (dragState?.axis === "row") {
          const normalized = normalizeRange(nextRange);
          axisSelectionRef.current = {
            axis: "row",
            startRow: normalized.start.row,
            endRow: normalized.end.row
          };
        } else {
          axisSelectionRef.current = null;
        }
        commitSelectionRange(nextRange);
      } else if (!nextRange) {
        selectionPreviewRangeRef.current = null;
      }
      setInteractionMode("idle");
      document.body.style.cursor = "";
      document.body.style.userSelect = "";
      selectionDragCleanupRef.current?.();
      selectionDragCleanupRef.current = null;
    };

    window.addEventListener("pointermove", handlePointerMove);
    window.addEventListener("pointerup", finishSelectionDrag);
    window.addEventListener("pointercancel", finishSelectionDrag);
    selectionDragCleanupRef.current = () => {
      if (pointerMoveFrame !== null) {
        window.cancelAnimationFrame(pointerMoveFrame);
        pointerMoveFrame = null;
      }
      pendingClientPoint = null;
      window.removeEventListener("pointermove", handlePointerMove);
      window.removeEventListener("pointerup", finishSelectionDrag);
      window.removeEventListener("pointercancel", finishSelectionDrag);
    };
  }

  function installFillDragListeners(pointerId: number, sourceRange: XlsxCellRange) {
    fillDragCleanupRef.current?.();
    let pendingClientPoint: { x: number; y: number } | null = null;
    let pointerMoveFrame: number | null = null;

    const flushPointerMove = () => {
      pointerMoveFrame = null;
      const pendingPoint = pendingClientPoint;
      pendingClientPoint = null;
      if (!pendingPoint) {
        return;
      }
      const nextCell = resolvePointerCellFromClient(pendingPoint.x, pendingPoint.y);
      if (nextCell) {
        updateFillPreview(nextCell);
      }
    };

    const handlePointerMove = (event: PointerEvent) => {
      if (event.pointerId !== pointerId) {
        return;
      }
      pendingClientPoint = { x: event.clientX, y: event.clientY };
      if (pointerMoveFrame !== null) {
        return;
      }
      pointerMoveFrame = window.requestAnimationFrame(flushPointerMove);
    };

    const finishFillDrag = (event: PointerEvent) => {
      if (event.pointerId !== pointerId) {
        return;
      }
      if (pointerMoveFrame !== null) {
        window.cancelAnimationFrame(pointerMoveFrame);
        pointerMoveFrame = null;
      }
      pendingClientPoint = null;

      const nextCell = resolvePointerCellFromClient(event.clientX, event.clientY);
      if (nextCell) {
        updateFillPreview(nextCell);
      }

      const nextRange = fillDragRef.current?.previewRange ?? sourceRange;
      fillDragRef.current = null;
      cachedScrollerRectRef.current = null;
      setInteractionMode("idle");
      setFillPreviewRange(null);
      document.body.style.cursor = "";
      document.body.style.userSelect = "";
      fillSelection(nextRange);
      fillDragCleanupRef.current?.();
      fillDragCleanupRef.current = null;
    };

    window.addEventListener("pointermove", handlePointerMove);
    window.addEventListener("pointerup", finishFillDrag);
    window.addEventListener("pointercancel", finishFillDrag);
    fillDragCleanupRef.current = () => {
      if (pointerMoveFrame !== null) {
        window.cancelAnimationFrame(pointerMoveFrame);
        pointerMoveFrame = null;
      }
      pendingClientPoint = null;
      window.removeEventListener("pointermove", handlePointerMove);
      window.removeEventListener("pointerup", finishFillDrag);
      window.removeEventListener("pointercancel", finishFillDrag);
    };
  }

  React.useLayoutEffect(() => {
    const overlayRange = selectionPreviewRangeRef.current ?? displayedSelection;
    if (!overlayRange || !wrapperRef.current) {
      applyHeaderSelection(null);
      if (selectionOverlayRef.current) {
        selectionOverlayRef.current.style.opacity = "0";
        selectionOverlayRef.current.style.visibility = "hidden";
      }
      return;
    }

    applyPreviewOverlay(overlayRange);
  }, [applyHeaderSelection, applyPreviewOverlay, displayedSelection, revision]);

  React.useLayoutEffect(() => {
    syncActiveValidationOverlay(activeCell);
  }, [activeCell, editingCell, revision, syncActiveValidationOverlay]);

  const handleCellDoubleClick = React.useCallback((cell: XlsxCellAddress) => {
    startEditing(cell);
  }, [startEditing]);

  const handleCellClick = React.useCallback((cell: XlsxCellAddress, cellData: CellRenderData) => {
    if (!cellData.hyperlink) {
      return;
    }

    openHyperlink(cellData.hyperlink.target, cellData.hyperlink.location);
  }, [openHyperlink]);

  const handleCellPointerDown = React.useCallback((
    event: React.PointerEvent<HTMLTableCellElement>,
    cell: XlsxCellAddress
  ) => {
    if (event.button !== 0) {
      return;
    }

    event.preventDefault();
    focusGrid();
    axisSelectionRef.current = null;
    const targetCell =
      event.currentTarget.colSpan > 1 || event.currentTarget.rowSpan > 1
        ? resolvePointerCellFromGeometry(event.clientX, event.clientY) ?? cell
        : cell;
    const currentSelection = selectionRef.current;
    const anchor = event.shiftKey && currentSelection ? currentSelection.start : targetCell;
    const initialRange = normalizeRange({ start: anchor, end: targetCell });
    const isActive = isSameCell(activeCellRef.current, targetCell);
    const committedOnPointerDown = !isActive || !editingCellRef.current;
    if (editingCellRef.current && !isActive) {
      commitEditingRef.current();
    }
    const pointerOrigin =
      targetCell.row === cell.row && targetCell.col === cell.col
        ? resolveCellPointerOrigin(cell, event.currentTarget.getBoundingClientRect(), event.clientX, event.clientY)
        : resolveCellPointerOriginFromClient(targetCell, event.clientX, event.clientY);
    const originOverlayRect =
      targetCell.row === cell.row && targetCell.col === cell.col
        ? resolveMountedCellOverlayRect(event.currentTarget)
        : resolveOverlayRect(initialRange);
    if (!pointerOrigin) {
      return;
    }

    startCellSelection(
      event.currentTarget,
      event.pointerId,
      anchor,
      "cell",
      targetCell,
      pointerOrigin,
      originOverlayRect,
      committedOnPointerDown,
      initialRange,
      event.clientX,
      event.clientY
    );
    if (targetCell.row === cell.row && targetCell.col === cell.col) {
      applyPreviewOverlayFromElement(event.currentTarget, initialRange);
    } else {
      applyPreviewOverlay(initialRange);
    }
    if (committedOnPointerDown) {
      commitSelectionRange(initialRange);
    }
  }, [
    applyPreviewOverlay,
    applyPreviewOverlayFromElement,
    commitSelectionRange,
    focusGrid,
    resolveCellPointerOrigin,
    resolveCellPointerOriginFromClient,
    resolveMountedCellOverlayRect,
    resolveOverlayRect,
    resolvePointerCellFromGeometry
  ]);

  const handleRowPointerDown = React.useCallback((
    event: React.PointerEvent<HTMLTableCellElement>,
    actualRow: number
  ) => {
    if (event.button !== 0 || firstVisibleCol === undefined || lastVisibleCol === undefined) {
      return;
    }

    event.preventDefault();
    focusGrid();
    const currentSelection = selectionRef.current;
    const anchorRow = event.shiftKey && currentSelection ? currentSelection.start.row : actualRow;
    const initialRange = normalizeRange({
      start: { row: anchorRow, col: firstVisibleCol },
      end: { row: actualRow, col: lastVisibleCol }
    });
    const pointerOrigin = resolveRowPointerOrigin(actualRow, event.currentTarget.getBoundingClientRect(), event.clientY);
    if (!pointerOrigin) {
      return;
    }

    startCellSelection(
      event.currentTarget,
      event.pointerId,
      { row: anchorRow, col: firstVisibleCol },
      "row",
      { row: actualRow, col: firstVisibleCol },
      pointerOrigin,
      null,
      true,
      initialRange,
      event.clientX,
      event.clientY
    );
    axisSelectionRef.current = { axis: "row", startRow: anchorRow, endRow: actualRow };
    commitSelectionRange(initialRange);
  }, [commitSelectionRange, firstVisibleCol, focusGrid, lastVisibleCol, resolveRowPointerOrigin]);

  const handleColumnPointerDown = React.useCallback((
    event: React.PointerEvent<HTMLTableCellElement>,
    actualCol: number
  ) => {
    if (event.button !== 0 || firstVisibleRow === undefined || lastVisibleRow === undefined) {
      return;
    }

    event.preventDefault();
    focusGrid();
    const currentSelection = selectionRef.current;
    const anchorCol =
      event.shiftKey && currentSelection ? currentSelection.start.col : actualCol;
    const initialRange = normalizeRange({
      start: { row: firstVisibleRow, col: anchorCol },
      end: { row: lastVisibleRow, col: actualCol }
    });
    const pointerOrigin = resolveColumnPointerOrigin(actualCol, event.currentTarget.getBoundingClientRect(), event.clientX);
    if (!pointerOrigin) {
      return;
    }

    startCellSelection(
      event.currentTarget,
      event.pointerId,
      { row: firstVisibleRow, col: anchorCol },
      "column",
      { row: firstVisibleRow, col: actualCol },
      pointerOrigin,
      null,
      true,
      initialRange,
      event.clientX,
      event.clientY
    );
    axisSelectionRef.current = { axis: "column", startCol: anchorCol, endCol: actualCol };
    commitSelectionRange(initialRange);
  }, [commitSelectionRange, firstVisibleRow, focusGrid, lastVisibleRow, resolveColumnPointerOrigin]);

  const handleRowResizePointerDown = React.useCallback((
    event: React.PointerEvent<HTMLDivElement>,
    actualRow: number,
    rowHeight: number
  ) => {
    if (!canResizeHeaders) {
      return;
    }

    event.preventDefault();
    event.stopPropagation();
    startRowResize(event.pointerId, actualRow, rowHeight, event.clientY);
  }, [canResizeHeaders]);

  const renderCellAdornment = React.useCallback((cell: XlsxCellAddress) => {
    const table = getTableAtCell(effectiveTables, cell.row, cell.col);
    const tableColumn = getTableHeaderColumn(table, cell.row, cell.col);
    if (!table || !tableColumn) {
      return null;
    }

    const direction =
      sortState && sortState.tableName === table.name && sortState.columnIndex === tableColumn.index
        ? sortState.direction
        : null;
    const triggerIcon = direction === "ascending" ? "▲" : direction === "descending" ? "▼" : "▾";

    if (renderTableHeaderMenu) {
      return renderTableHeaderMenu({
        cell,
        column: tableColumn,
        direction,
        sortAscending: () => sortTable(table.name, tableColumn.index, "ascending"),
        sortDescending: () => sortTable(table.name, tableColumn.index, "descending"),
        table,
        triggerIcon,
        triggerProps: {
          "aria-haspopup": "menu",
          "aria-label": `Open menu for ${tableColumn.name}`,
          onClick: (event) => {
            event.stopPropagation();
          },
          onPointerDown: (event) => {
            event.stopPropagation();
          },
          style: resolveTableHeaderTriggerStyle(palette, zoomFactor),
          type: "button"
        }
      });
    }

    return (
      <button
        onPointerDown={(event) => {
          event.preventDefault();
          event.stopPropagation();
        }}
        onClick={(event) => {
          event.preventDefault();
          event.stopPropagation();
          setOpenTableMenu((current) =>
            current &&
            current.tableName === table.name &&
            current.row === cell.row &&
            current.col === cell.col
              ? null
              : { col: cell.col, row: cell.row, tableName: table.name }
          );
        }}
        style={{
          ...resolveTableHeaderTriggerStyle(palette, zoomFactor)
        }}
        type="button"
      >
        {triggerIcon}
      </button>
    );
  }, [effectiveTables, palette, renderTableHeaderMenu, sortState, sortTable, zoomFactor]);
  const resolveCanvasColumnHeaderRect = React.useCallback((actualCol: number) => {
    const colIndex = colIndexByActual.get(actualCol);
    if (colIndex === undefined) {
      return null;
    }

    const width = displayEffectiveColWidths[colIndex] ?? displayDefaultColWidth;
    return {
      left: stickyLeftByCol.get(actualCol) ?? (displayRowHeaderWidth + (colPrefixSums[colIndex] ?? 0) - drawingViewport.left),
      width
    };
  }, [
    colIndexByActual,
    colPrefixSums,
    displayDefaultColWidth,
    displayEffectiveColWidths,
    displayRowHeaderWidth,
    drawingViewport.left,
    stickyLeftByCol
  ]);

  const resolveCanvasRowHeaderRect = React.useCallback((actualRow: number) => {
    const rowIndex = rowIndexByActual.get(actualRow);
    if (rowIndex === undefined) {
      return null;
    }

    const height = displayEffectiveRowHeights[rowIndex] ?? displayDefaultRowHeight;
    return {
      height,
      top: stickyTopByRow.get(actualRow) ?? (displayHeaderHeight + (rowPrefixSums[rowIndex] ?? 0) - drawingViewport.top)
    };
  }, [
    displayDefaultRowHeight,
    displayEffectiveRowHeights,
    displayHeaderHeight,
    drawingViewport.top,
    rowIndexByActual,
    rowPrefixSums,
    stickyTopByRow
  ]);
  const scrollBodyViewportWidth = Math.max(0, drawingViewport.width - frozenPaneRight);
  const scrollBodyViewportHeight = Math.max(0, drawingViewport.height - frozenPaneBottom);
  const canvasScrollBufferX = Math.min(CANVAS_SCROLL_BUFFER_PX, scrollBodyViewportWidth);
  const canvasScrollBufferY = Math.min(CANVAS_SCROLL_BUFFER_PX, scrollBodyViewportHeight);
  const scrollBodyCanvasLeft = frozenPaneRight - canvasScrollBufferX;
  const scrollBodyCanvasTop = frozenPaneBottom - canvasScrollBufferY;
  const scrollBodyCanvasWidth = scrollBodyViewportWidth + canvasScrollBufferX * 2;
  const scrollBodyCanvasHeight = scrollBodyViewportHeight + canvasScrollBufferY * 2;
  const topBodyCanvasWidth = scrollBodyCanvasWidth;
  const topBodyCanvasHeight = Math.max(0, frozenPaneBottom - displayHeaderHeight);
  const leftBodyCanvasWidth = Math.max(0, frozenPaneRight - displayRowHeaderWidth);
  const leftBodyCanvasHeight = scrollBodyCanvasHeight;
  const cornerBodyCanvasWidth = leftBodyCanvasWidth;
  const cornerBodyCanvasHeight = topBodyCanvasHeight;
  const topFrozenHeaderCanvasWidth = leftBodyCanvasWidth;
  const topScrollHeaderCanvasWidth = scrollBodyCanvasWidth;
  const leftFrozenHeaderCanvasHeight = topBodyCanvasHeight;
  const leftScrollHeaderCanvasHeight = scrollBodyCanvasHeight;
  const canvasColumnHeaderCells = React.useMemo(
    () =>
      canvasVisibleColItems.flatMap<CanvasStickyColumnHeaderCell>((column) => {
        const rect = resolveCanvasColumnHeaderRect(column.actualCol);
        if (!rect) {
          return [];
        }

        const isFrozen = stickyLeftByCol.has(column.actualCol);
        return [{
          actualCol: column.actualCol,
          height: displayHeaderHeight,
          isFrozen,
          left: rect.left,
          localLeft: rect.left - (isFrozen ? displayRowHeaderWidth : scrollBodyCanvasLeft),
          width: rect.width
        }];
      }),
    [
      canvasVisibleColItems,
      displayHeaderHeight,
      displayRowHeaderWidth,
      resolveCanvasColumnHeaderRect,
      scrollBodyCanvasLeft,
      stickyLeftByCol
    ]
  );
  const canvasRowHeaderCells = React.useMemo(
    () =>
      canvasVisibleRowItems.flatMap<CanvasStickyRowHeaderCell>((row) => {
        const rect = resolveCanvasRowHeaderRect(row.actualRow);
        if (!rect) {
          return [];
        }

        const isFrozen = stickyTopByRow.has(row.actualRow);
        return [{
          actualRow: row.actualRow,
          height: rect.height,
          isFrozen,
          localTop: rect.top - (isFrozen ? displayHeaderHeight : scrollBodyCanvasTop),
          top: rect.top
        }];
      }),
    [
      canvasVisibleRowItems,
      displayHeaderHeight,
      resolveCanvasRowHeaderRect,
      scrollBodyCanvasTop,
      stickyTopByRow
    ]
  );
  const resolveCanvasColumnResizeTarget = React.useCallback((clientX: number) => {
    if (!canResizeHeaders) {
      return null;
    }

    const scrollerRect = scrollRef.current?.getBoundingClientRect();
    if (!scrollerRect) {
      return null;
    }

    const localX = clientX - scrollerRect.left;
    for (const column of canvasColumnHeaderCells) {
      if (Math.abs(localX - (column.left + column.width)) <= CANVAS_RESIZE_HIT_SLOP_PX) {
        return { actualCol: column.actualCol, width: column.width };
      }
    }

    return null;
  }, [canResizeHeaders, canvasColumnHeaderCells]);
  const resolveCanvasColumnHeaderTarget = React.useCallback((clientX: number) => {
    const scrollerRect = scrollRef.current?.getBoundingClientRect();
    if (!scrollerRect) {
      return null;
    }

    const localX = clientX - scrollerRect.left;
    for (const column of canvasColumnHeaderCells) {
      if (localX >= column.left && localX <= column.left + column.width) {
        return column.actualCol;
      }
    }

    return null;
  }, [canvasColumnHeaderCells]);

  const resolveCanvasRowResizeTarget = React.useCallback((clientY: number) => {
    if (!canResizeHeaders) {
      return null;
    }

    const scrollerRect = scrollRef.current?.getBoundingClientRect();
    if (!scrollerRect) {
      return null;
    }

    const localY = clientY - scrollerRect.top;
    for (const row of canvasRowHeaderCells) {
      if (Math.abs(localY - (row.top + row.height)) <= CANVAS_RESIZE_HIT_SLOP_PX) {
        return { actualRow: row.actualRow, height: row.height };
      }
    }

    return null;
  }, [canResizeHeaders, canvasRowHeaderCells]);
  const resolveCanvasRowHeaderTarget = React.useCallback((clientY: number) => {
    const scrollerRect = scrollRef.current?.getBoundingClientRect();
    if (!scrollerRect) {
      return null;
    }

    const localY = clientY - scrollerRect.top;
    for (const row of canvasRowHeaderCells) {
      if (localY >= row.top && localY <= row.top + row.height) {
        return row.actualRow;
      }
    }

    return null;
  }, [canvasRowHeaderCells]);

  const handleCanvasColumnHeaderPointerMove = React.useCallback((event: React.PointerEvent<HTMLCanvasElement>) => {
    if (resizeStateRef.current?.type === "column") {
      event.currentTarget.style.cursor = "col-resize";
      return;
    }
    event.currentTarget.style.cursor = resolveCanvasColumnResizeTarget(event.clientX) ? "col-resize" : "default";
  }, [resolveCanvasColumnResizeTarget]);

  const handleCanvasRowHeaderPointerMove = React.useCallback((event: React.PointerEvent<HTMLCanvasElement>) => {
    if (resizeStateRef.current?.type === "row") {
      event.currentTarget.style.cursor = "row-resize";
      return;
    }
    event.currentTarget.style.cursor = resolveCanvasRowResizeTarget(event.clientY) ? "row-resize" : "default";
  }, [resolveCanvasRowResizeTarget]);

  const handleCanvasHeaderPointerLeave = React.useCallback((event: React.PointerEvent<HTMLCanvasElement>) => {
    if (resizeStateRef.current) {
      event.currentTarget.style.cursor = resizeStateRef.current.type === "column" ? "col-resize" : "row-resize";
      return;
    }
    event.currentTarget.style.cursor = "default";
  }, []);

  const handleCanvasBodyPointerDown = React.useCallback((event: React.PointerEvent<HTMLCanvasElement>) => {
    if (event.button !== 0) {
      return;
    }

    const cell = resolvePointerCellFromClient(event.clientX, event.clientY);
    if (!cell) {
      return;
    }

    event.preventDefault();
    focusGrid();
    axisSelectionRef.current = null;
    const currentSelection = selectionRef.current;
    const anchor = event.shiftKey && currentSelection ? currentSelection.start : cell;
    const initialRange = normalizeRange({ start: anchor, end: cell });
    const isActive = isSameCell(activeCellRef.current, cell);
    const committedOnPointerDown = !isActive || !editingCellRef.current;
    if (editingCellRef.current && !isActive) {
      commitEditingRef.current();
    }
    const rowIndex = rowIndexByActual.get(cell.row);
    const colIndex = colIndexByActual.get(cell.col);
    if (rowIndex === undefined || colIndex === undefined) {
      return;
    }

    startCellSelection(
      event.currentTarget,
      event.pointerId,
      anchor,
      "cell",
      cell,
      {
        contentScaleX: 1,
        contentScaleY: 1,
        originContentX: colPrefixSums[colIndex] ?? 0,
        originContentY: rowPrefixSums[rowIndex] ?? 0
      },
      resolveOverlayRect(initialRange),
      committedOnPointerDown,
      initialRange,
      event.clientX,
      event.clientY
    );
    applyPreviewOverlay(initialRange);
    if (committedOnPointerDown) {
      commitSelectionRange(initialRange);
    }
  }, [
    applyPreviewOverlay,
    colIndexByActual,
    colPrefixSums,
    commitSelectionRange,
    focusGrid,
    resolveOverlayRect,
    resolvePointerCellFromClient,
    rowIndexByActual,
    rowPrefixSums,
    startCellSelection
  ]);

  const handleCanvasBodyClick = React.useCallback((event: React.MouseEvent<HTMLCanvasElement>) => {
    const cell = resolvePointerCellFromClient(event.clientX, event.clientY);
    if (!cell) {
      return;
    }

    handleCellClick(cell, getCellData(cell.row, cell.col));
  }, [getCellData, handleCellClick, resolvePointerCellFromClient]);

  const handleCanvasBodyDoubleClick = React.useCallback((event: React.MouseEvent<HTMLCanvasElement>) => {
    if (readOnly) {
      return;
    }

    const cell = resolvePointerCellFromClient(event.clientX, event.clientY);
    if (!cell) {
      return;
    }

    startEditing(cell);
  }, [readOnly, resolvePointerCellFromClient, startEditing]);

  const handleCanvasColumnHeaderPointerDown = React.useCallback((event: React.PointerEvent<HTMLCanvasElement>) => {
    if (event.button !== 0 || firstVisibleRow === undefined || lastVisibleRow === undefined) {
      return;
    }

    const resizeTarget = resolveCanvasColumnResizeTarget(event.clientX);
    if (resizeTarget) {
      event.preventDefault();
      event.stopPropagation();
      startColumnResize(event.pointerId, resizeTarget.actualCol, resizeTarget.width, event.clientX);
      return;
    }

    const actualCol = resolveCanvasColumnHeaderTarget(event.clientX);
    if (actualCol === null) {
      return;
    }

    event.preventDefault();
    focusGrid();
    const currentSelection = selectionRef.current;
    const anchorCol = event.shiftKey && currentSelection ? currentSelection.start.col : actualCol;
    const initialRange = normalizeRange({
      start: { row: firstVisibleRow, col: anchorCol },
      end: { row: lastVisibleRow, col: actualCol }
    });
    const anchorColIndex = colIndexByActual.get(anchorCol);
    if (anchorColIndex === undefined) {
      return;
    }

    startCellSelection(
      event.currentTarget,
      event.pointerId,
      { row: firstVisibleRow, col: anchorCol },
      "column",
      { row: firstVisibleRow, col: actualCol },
      {
        contentScaleX: 1,
        contentScaleY: 1,
        originContentX: colPrefixSums[anchorColIndex] ?? 0,
        originContentY: rowPrefixSums[0] ?? 0
      },
      null,
      true,
      initialRange,
      event.clientX,
      event.clientY
    );
    axisSelectionRef.current = { axis: "column", startCol: anchorCol, endCol: actualCol };
    commitSelectionRange(initialRange);
  }, [
    colIndexByActual,
    colPrefixSums,
    commitSelectionRange,
    firstVisibleRow,
    focusGrid,
    lastVisibleRow,
    resolveCanvasColumnHeaderTarget,
    resolveCanvasColumnResizeTarget,
    rowPrefixSums,
    startCellSelection
  ]);

  const handleCanvasRowHeaderPointerDown = React.useCallback((event: React.PointerEvent<HTMLCanvasElement>) => {
    if (event.button !== 0 || firstVisibleCol === undefined || lastVisibleCol === undefined) {
      return;
    }

    const resizeTarget = resolveCanvasRowResizeTarget(event.clientY);
    if (resizeTarget) {
      event.preventDefault();
      event.stopPropagation();
      startRowResize(event.pointerId, resizeTarget.actualRow, resizeTarget.height, event.clientY);
      return;
    }

    const actualRow = resolveCanvasRowHeaderTarget(event.clientY);
    if (actualRow === null) {
      return;
    }

    event.preventDefault();
    focusGrid();
    const currentSelection = selectionRef.current;
    const anchorRow = event.shiftKey && currentSelection ? currentSelection.start.row : actualRow;
    const initialRange = normalizeRange({
      start: { row: anchorRow, col: firstVisibleCol },
      end: { row: actualRow, col: lastVisibleCol }
    });
    const anchorRowIndex = rowIndexByActual.get(anchorRow);
    if (anchorRowIndex === undefined) {
      return;
    }

    startCellSelection(
      event.currentTarget,
      event.pointerId,
      { row: anchorRow, col: firstVisibleCol },
      "row",
      { row: actualRow, col: firstVisibleCol },
      {
        contentScaleX: 1,
        contentScaleY: 1,
        originContentX: colPrefixSums[0] ?? 0,
        originContentY: rowPrefixSums[anchorRowIndex] ?? 0
      },
      null,
      true,
      initialRange,
      event.clientX,
      event.clientY
    );
    axisSelectionRef.current = { axis: "row", startRow: anchorRow, endRow: actualRow };
    commitSelectionRange(initialRange);
  }, [
    colPrefixSums,
    commitSelectionRange,
    firstVisibleCol,
    focusGrid,
    lastVisibleCol,
    resolveCanvasRowHeaderTarget,
    resolveCanvasRowResizeTarget,
    rowIndexByActual,
    rowPrefixSums,
    startCellSelection
  ]);

  const getCanvasImage = React.useCallback((image: XlsxImage) => {
    if (typeof Image === "undefined") {
      return null;
    }

    const cacheKey = image.id;
    const cached = canvasImageCacheRef.current.get(cacheKey);
    if (cached && cached.src === image.src) {
      return cached.loaded && !cached.failed ? cached.image : null;
    }

    const imageElement = new Image();
    const entry = {
      failed: false,
      image: imageElement,
      loaded: false,
      src: image.src
    };
    canvasImageCacheRef.current.set(cacheKey, entry);
    imageElement.onload = () => {
      entry.loaded = true;
      setCanvasImageLoadVersion((version) => version + 1);
    };
    imageElement.onerror = () => {
      entry.failed = true;
      setCanvasImageLoadVersion((version) => version + 1);
    };
    imageElement.src = image.src;
    return null;
  }, []);

  React.useLayoutEffect(() => {
    if (!experimentalCanvas) {
      return;
    }

    const canvasProfileTarget = typeof window !== "undefined"
      ? (window as unknown as {
          __xlsxCanvasProfile?: Array<{
            bodyMs: number;
            culledCells: number;
            dirtyRects: number;
            lookedUpCells: number;
            paintedCells: number;
            repaintBody: boolean;
            repaintHeaders: boolean;
            renderToLayoutMs: number;
            totalMs: number;
          }>;
        }).__xlsxCanvasProfile
      : undefined;
    const canvasProfileStart = canvasProfileTarget ? performance.now() : 0;
    let canvasProfileBodyStart = 0;
    let canvasProfileBodyMs = 0;
    let canvasProfileCulledCells = 0;
    let canvasProfileDirtyRects = 0;
    let canvasProfileLookedUpCells = 0;
    let canvasProfilePaintedCells = 0;

    const dpr = typeof window === "undefined" ? 1 : Math.max(1, window.devicePixelRatio || 1);

    function configureCanvas(
      canvas: HTMLCanvasElement | null,
      width: number,
      height: number,
      options?: { clear?: boolean }
    ) {
      if (!canvas) {
        return null;
      }

      const nextWidth = Math.max(1, Math.round(width * dpr));
      const nextHeight = Math.max(1, Math.round(height * dpr));
      if (canvas.width !== nextWidth) {
        canvas.width = nextWidth;
      }
      if (canvas.height !== nextHeight) {
        canvas.height = nextHeight;
      }
      if (canvas.style.width !== `${width}px`) {
        canvas.style.width = `${width}px`;
      }
      if (canvas.style.height !== `${height}px`) {
        canvas.style.height = `${height}px`;
      }
      const context = canvas.getContext("2d", { alpha: false });
      if (!context) {
        return null;
      }

      context.setTransform(dpr, 0, 0, dpr, 0, 0);
      if (options?.clear !== false) {
        context.clearRect(0, 0, width, height);
      }
      return context;
    }

    const bodyWidth = Math.max(0, drawingViewport.width);
    const bodyHeight = Math.max(0, drawingViewport.height);
    const headerHeight = Math.min(displayHeaderHeight, bodyHeight);
    const rowHeaderWidth = Math.min(displayRowHeaderWidth, bodyWidth);
    const visibleRange = selectionPreviewRangeRef.current ?? displayedSelection;
    const normalizedVisibleRange = visibleRange ? normalizeRange(visibleRange) : null;
    const rangeSignature = buildRangeSignature(normalizedVisibleRange);
    const nextBodyCanvasSignature: PaintedBodyCanvasSignature = {
      activeSheet: activeSheet ?? null,
      bakedDrawingSignature: bakedCanvasDrawingSignature,
      bodyHeight,
      bodyWidth,
      colSignature: buildRenderedAxisSignature(canvasVisibleColItems, (item) => item.index),
      getCellData,
      headerHeight,
      palette,
      rowHeaderWidth,
      rowSignature: buildRenderedAxisSignature(canvasVisibleRowItems, (item) => item.index),
      zoomFactor
    };
    const nextHeaderCanvasSignature: PaintedHeaderCanvasSignature = {
      activeSheet: activeSheet ?? null,
      bodyHeight,
      bodyWidth,
      colSignature: nextBodyCanvasSignature.colSignature,
      headerHeight,
      palette,
      rangeSignature,
      rowHeaderWidth,
      rowSignature: nextBodyCanvasSignature.rowSignature,
      viewportLeft: drawingViewport.left,
      viewportTop: drawingViewport.top,
      zoomFactor
    };
    const previousBodyCanvasSignature = paintedBodyCanvasSignatureRef.current;
    const previousHeaderCanvasSignature = paintedHeaderCanvasSignatureRef.current;
    const previousPaintedViewport = paintedDrawingViewportRef.current;
    const shouldRepaintBody = !(
      previousBodyCanvasSignature
      && previousBodyCanvasSignature.activeSheet === nextBodyCanvasSignature.activeSheet
      && previousBodyCanvasSignature.bakedDrawingSignature === nextBodyCanvasSignature.bakedDrawingSignature
      && previousBodyCanvasSignature.bodyHeight === nextBodyCanvasSignature.bodyHeight
      && previousBodyCanvasSignature.bodyWidth === nextBodyCanvasSignature.bodyWidth
      && previousBodyCanvasSignature.colSignature === nextBodyCanvasSignature.colSignature
      && previousBodyCanvasSignature.getCellData === nextBodyCanvasSignature.getCellData
      && previousBodyCanvasSignature.headerHeight === nextBodyCanvasSignature.headerHeight
      && previousBodyCanvasSignature.palette === nextBodyCanvasSignature.palette
      && previousBodyCanvasSignature.rowHeaderWidth === nextBodyCanvasSignature.rowHeaderWidth
      && previousBodyCanvasSignature.rowSignature === nextBodyCanvasSignature.rowSignature
      && previousBodyCanvasSignature.zoomFactor === nextBodyCanvasSignature.zoomFactor
    );
    const shouldRepaintHeaders = !(
      previousHeaderCanvasSignature
      && previousHeaderCanvasSignature.activeSheet === nextHeaderCanvasSignature.activeSheet
      && previousHeaderCanvasSignature.bodyHeight === nextHeaderCanvasSignature.bodyHeight
      && previousHeaderCanvasSignature.bodyWidth === nextHeaderCanvasSignature.bodyWidth
      && previousHeaderCanvasSignature.colSignature === nextHeaderCanvasSignature.colSignature
      && previousHeaderCanvasSignature.headerHeight === nextHeaderCanvasSignature.headerHeight
      && previousHeaderCanvasSignature.palette === nextHeaderCanvasSignature.palette
      && previousHeaderCanvasSignature.rangeSignature === nextHeaderCanvasSignature.rangeSignature
      && previousHeaderCanvasSignature.rowHeaderWidth === nextHeaderCanvasSignature.rowHeaderWidth
      && previousHeaderCanvasSignature.rowSignature === nextHeaderCanvasSignature.rowSignature
      && previousHeaderCanvasSignature.viewportLeft === nextHeaderCanvasSignature.viewportLeft
      && previousHeaderCanvasSignature.viewportTop === nextHeaderCanvasSignature.viewportTop
      && previousHeaderCanvasSignature.zoomFactor === nextHeaderCanvasSignature.zoomFactor
    );
    const canBlitBody = Boolean(
      shouldRepaintBody
      && previousBodyCanvasSignature
      && previousBodyCanvasSignature.activeSheet === nextBodyCanvasSignature.activeSheet
      && previousBodyCanvasSignature.bakedDrawingSignature === nextBodyCanvasSignature.bakedDrawingSignature
      && previousBodyCanvasSignature.bodyHeight === nextBodyCanvasSignature.bodyHeight
      && previousBodyCanvasSignature.bodyWidth === nextBodyCanvasSignature.bodyWidth
      && previousBodyCanvasSignature.getCellData === nextBodyCanvasSignature.getCellData
      && previousBodyCanvasSignature.headerHeight === nextBodyCanvasSignature.headerHeight
      && previousBodyCanvasSignature.palette === nextBodyCanvasSignature.palette
      && previousBodyCanvasSignature.rowHeaderWidth === nextBodyCanvasSignature.rowHeaderWidth
      && previousBodyCanvasSignature.zoomFactor === nextBodyCanvasSignature.zoomFactor
    );
    const canBlitTopHeader = Boolean(
      shouldRepaintHeaders
      && previousHeaderCanvasSignature
      && previousHeaderCanvasSignature.activeSheet === nextHeaderCanvasSignature.activeSheet
      && previousHeaderCanvasSignature.bodyHeight === nextHeaderCanvasSignature.bodyHeight
      && previousHeaderCanvasSignature.bodyWidth === nextHeaderCanvasSignature.bodyWidth
      && previousHeaderCanvasSignature.headerHeight === nextHeaderCanvasSignature.headerHeight
      && previousHeaderCanvasSignature.palette === nextHeaderCanvasSignature.palette
      && previousHeaderCanvasSignature.rangeSignature === nextHeaderCanvasSignature.rangeSignature
      && previousHeaderCanvasSignature.rowHeaderWidth === nextHeaderCanvasSignature.rowHeaderWidth
      && previousHeaderCanvasSignature.zoomFactor === nextHeaderCanvasSignature.zoomFactor
    );
    const canBlitLeftHeader = Boolean(
      shouldRepaintHeaders
      && previousHeaderCanvasSignature
      && previousHeaderCanvasSignature.activeSheet === nextHeaderCanvasSignature.activeSheet
      && previousHeaderCanvasSignature.bodyHeight === nextHeaderCanvasSignature.bodyHeight
      && previousHeaderCanvasSignature.bodyWidth === nextHeaderCanvasSignature.bodyWidth
      && previousHeaderCanvasSignature.headerHeight === nextHeaderCanvasSignature.headerHeight
      && previousHeaderCanvasSignature.palette === nextHeaderCanvasSignature.palette
      && previousHeaderCanvasSignature.rangeSignature === nextHeaderCanvasSignature.rangeSignature
      && previousHeaderCanvasSignature.rowHeaderWidth === nextHeaderCanvasSignature.rowHeaderWidth
      && previousHeaderCanvasSignature.zoomFactor === nextHeaderCanvasSignature.zoomFactor
    );
    if (!shouldRepaintBody && !shouldRepaintHeaders) {
      return;
    }
    const paneBounds: Record<FrozenDrawingPane, { height: number; left: number; top: number; width: number }> = {
      corner: {
        height: cornerBodyCanvasHeight,
        left: displayRowHeaderWidth,
        top: displayHeaderHeight,
        width: cornerBodyCanvasWidth
      },
      left: {
        height: leftBodyCanvasHeight,
        left: displayRowHeaderWidth,
        top: scrollBodyCanvasTop,
        width: leftBodyCanvasWidth
      },
      scroll: {
        height: scrollBodyCanvasHeight,
        left: scrollBodyCanvasLeft,
        top: scrollBodyCanvasTop,
        width: scrollBodyCanvasWidth
      },
      top: {
        height: topBodyCanvasHeight,
        left: scrollBodyCanvasLeft,
        top: displayHeaderHeight,
        width: topBodyCanvasWidth
      }
    };
    const bodyContexts: Record<FrozenDrawingPane, CanvasRenderingContext2D | null> = shouldRepaintBody
      ? {
          corner: configureCanvas(cornerBodyCanvasRef.current, cornerBodyCanvasWidth, cornerBodyCanvasHeight, { clear: !canBlitBody }),
          left: configureCanvas(leftBodyCanvasRef.current, leftBodyCanvasWidth, leftBodyCanvasHeight, { clear: !canBlitBody }),
          scroll: configureCanvas(scrollBodyCanvasRef.current, scrollBodyCanvasWidth, scrollBodyCanvasHeight, { clear: !canBlitBody }),
          top: configureCanvas(topBodyCanvasRef.current, topBodyCanvasWidth, topBodyCanvasHeight, { clear: !canBlitBody })
        }
      : {
          corner: null,
          left: null,
          scroll: null,
          top: null
        };
    const topHeaderContexts = shouldRepaintHeaders
      ? {
          frozen: configureCanvas(topFrozenHeaderCanvasRef.current, topFrozenHeaderCanvasWidth, headerHeight),
          scroll: configureCanvas(topScrollHeaderCanvasRef.current, topScrollHeaderCanvasWidth, headerHeight, { clear: !canBlitTopHeader })
        }
      : {
          frozen: null,
          scroll: null
        };
    const leftHeaderContexts = shouldRepaintHeaders
      ? {
          frozen: configureCanvas(leftFrozenHeaderCanvasRef.current, rowHeaderWidth, leftFrozenHeaderCanvasHeight),
          scroll: configureCanvas(leftScrollHeaderCanvasRef.current, rowHeaderWidth, leftScrollHeaderCanvasHeight, { clear: !canBlitLeftHeader })
        }
      : {
          frozen: null,
          scroll: null
        };
    const cornerContext = shouldRepaintHeaders
      ? configureCanvas(cornerHeaderCanvasRef.current, rowHeaderWidth, headerHeight)
      : null;
    if (
      (shouldRepaintBody && (!bodyContexts.scroll || !bodyContexts.top || !bodyContexts.left || !bodyContexts.corner))
      || (
        shouldRepaintHeaders
        && (
          !cornerContext
          || (topFrozenHeaderCanvasWidth > 0 && !topHeaderContexts.frozen)
          || (topScrollHeaderCanvasWidth > 0 && !topHeaderContexts.scroll)
          || (leftFrozenHeaderCanvasHeight > 0 && !leftHeaderContexts.frozen)
          || (leftScrollHeaderCanvasHeight > 0 && !leftHeaderContexts.scroll)
        )
      )
    ) {
      return;
    }
    const showGridLines = activeSheet?.showGridLines ?? true;
    const sheetSurface = resolveSheetSurface(activeSheet, palette);
    const deferredSpillTextsByPane: Record<FrozenDrawingPane, Array<{
      align: CanvasTextAlign;
      clipHeight: number;
      clipLeft: number;
      clipTop: number;
      clipWidth: number;
      color: string;
      ellipsize: boolean;
      font: string;
      maxWidth: number;
      text: string;
      textDecoration?: string;
      textX: number;
      textY: number;
      lineThroughY: number;
      underlineY: number;
    }>> = {
      corner: [],
      left: [],
      scroll: [],
      top: []
    };
    const bodyDirtyRectsByPane: Record<FrozenDrawingPane, CanvasDirtyRect[]> = {
      corner: [],
      left: [],
      scroll: [],
      top: []
    };
    let topScrollHeaderDirtyRects = buildFullCanvasDirtyRect(topScrollHeaderCanvasWidth, headerHeight);
    let leftScrollHeaderDirtyRects = buildFullCanvasDirtyRect(rowHeaderWidth, leftScrollHeaderCanvasHeight);
    type BakedCanvasDrawingEntry =
      | {
          kind: "formControl";
          control: XlsxFormControl;
          rect: XlsxImageRect;
          zIndex: number;
        }
      | {
          kind: "image";
          image: XlsxImage;
          rect: XlsxImageRect;
          zIndex: number;
        }
      | {
          kind: "shape";
          rect: XlsxImageRect;
          shape: XlsxShape;
          zIndex: number;
        };
    const bakedCanvasDrawingEntries: BakedCanvasDrawingEntry[] = shouldBakeCanvasStaticDrawings
      ? [
          ...bakedShapeRects.map(({ rect, shape }) => ({
            kind: "shape" as const,
            rect,
            shape,
            zIndex: shape.zIndex
          })),
          ...bakedFormControlRects.map(({ control, rect }) => ({
            control,
            kind: "formControl" as const,
            rect,
            zIndex: control.zIndex
          })),
          ...bakedImageRects.map(({ image, rect }) => ({
            image,
            kind: "image" as const,
            rect,
            zIndex: image.zIndex
          }))
        ].sort((left, right) => left.zIndex - right.zIndex)
      : [];
    const resolveBakedCanvasLocalRect = (rect: XlsxImageRect, pane: FrozenDrawingPane) => {
      const bounds = paneBounds[pane];
      const scrollX = pane === "scroll" || pane === "top" ? drawingViewport.left : 0;
      const scrollY = pane === "scroll" || pane === "left" ? drawingViewport.top : 0;
      return {
        height: rect.height,
        left: rect.left - scrollX - bounds.left,
        top: rect.top - scrollY - bounds.top,
        width: rect.width
      };
    };
    const drawBakedShapeText = (
      context: CanvasRenderingContext2D,
      shape: XlsxShape,
      left: number,
      top: number,
      width: number,
      height: number
    ) => {
      if (shape.paragraphs.length === 0) {
        return;
      }

      const inset = shape.textBox?.insetPx;
      const paddingLeft = (inset?.left ?? 6) * zoomFactor;
      const paddingRight = (inset?.right ?? 6) * zoomFactor;
      const paddingTop = (inset?.top ?? 4) * zoomFactor;
      const paddingBottom = (inset?.bottom ?? 4) * zoomFactor;
      const textLeft = left + paddingLeft;
      const textTop = top + paddingTop;
      const textWidth = Math.max(0, width - paddingLeft - paddingRight);
      const textHeight = Math.max(0, height - paddingTop - paddingBottom);
      if (textWidth <= 0 || textHeight <= 0) {
        return;
      }

      const lineMetrics = shape.paragraphs.map((paragraph) => {
        const lineHeight = Math.max(
          12 * zoomFactor,
          ...paragraph.runs.map((run) => ((run.fontSizePt ?? 11) * 96 / 72) * zoomFactor * 1.2)
        );
        const widthPx = paragraph.runs.reduce((total, run) => {
          const fontSize = ((run.fontSizePt ?? 11) * 96 / 72) * zoomFactor;
          context.font = `${run.italic ? "italic " : ""}${run.bold ? "700 " : "400 "}${fontSize}px ${run.fontFamily ?? "Calibri, sans-serif"}`;
          return total + measureCanvasTextWidth(context, run.text);
        }, 0);
        return { lineHeight, widthPx };
      });
      const totalTextHeight = lineMetrics.reduce((total, metric) => total + metric.lineHeight, 0);
      let y = textTop;
      if (shape.textBox?.verticalAlign === "middle") {
        y = textTop + Math.max(0, (textHeight - totalTextHeight) / 2);
      } else if (shape.textBox?.verticalAlign === "bottom") {
        y = textTop + Math.max(0, textHeight - totalTextHeight);
      }

      context.save();
      context.beginPath();
      context.rect(textLeft, textTop, textWidth, textHeight);
      context.clip();
      context.textBaseline = "middle";
      shape.paragraphs.forEach((paragraph, paragraphIndex) => {
        const metric = lineMetrics[paragraphIndex] ?? { lineHeight: 14 * zoomFactor, widthPx: 0 };
        let x = textLeft;
        const align = paragraph.align ?? shape.textBox?.horizontalAlign ?? "left";
        if (align === "center") {
          x = textLeft + Math.max(0, (textWidth - metric.widthPx) / 2);
        } else if (align === "right") {
          x = textLeft + Math.max(0, textWidth - metric.widthPx);
        }

        paragraph.runs.forEach((run) => {
          const fontSize = ((run.fontSizePt ?? 11) * 96 / 72) * zoomFactor;
          context.font = `${run.italic ? "italic " : ""}${run.bold ? "700 " : "400 "}${fontSize}px ${run.fontFamily ?? "Calibri, sans-serif"}`;
          context.fillStyle = run.color ?? "#000000";
          context.textAlign = "left";
          const textY = y + metric.lineHeight / 2;
          context.fillText(run.text, x, textY);
          if (run.underline && run.text.length > 0) {
            const textWidthPx = measureCanvasTextWidth(context, run.text);
            context.strokeStyle = run.color ?? "#000000";
            context.lineWidth = Math.max(1, zoomFactor * 0.75);
            context.beginPath();
            context.moveTo(x, textY + Math.max(2, fontSize * 0.28));
            context.lineTo(x + textWidthPx, textY + Math.max(2, fontSize * 0.28));
            context.stroke();
          }
          x += measureCanvasTextWidth(context, run.text);
        });
        y += metric.lineHeight;
      });
      context.restore();
    };
    const drawBakedShape = (context: CanvasRenderingContext2D, shape: XlsxShape, rect: XlsxImageRect) => {
      const fillColor = shape.fill?.none ? "transparent" : (shape.fill?.color ?? "transparent");
      const strokeColor = shape.stroke?.none ? "transparent" : (shape.stroke?.color ?? "transparent");
      const lineWidth = Math.max(0, (shape.stroke?.widthPx ?? (shape.geometry === "line" ? 2 : 1)) * zoomFactor);
      const vectorShape = resolveShapeVector(shape);
      const opacity = Math.min(shape.fill?.opacity ?? 1, shape.stroke?.opacity ?? 1);

      context.save();
      context.translate(rect.left + rect.width / 2, rect.top + rect.height / 2);
      if (shape.rotationDeg) {
        context.rotate((shape.rotationDeg * Math.PI) / 180);
      }
      context.scale(shape.flipH ? -1 : 1, shape.flipV ? -1 : 1);
      context.globalAlpha *= opacity;
      context.lineWidth = lineWidth;
      context.strokeStyle = strokeColor;
      context.fillStyle = fillColor;
      applyCanvasShapeDash(context, shape.stroke?.dash, lineWidth);

      const localLeft = -rect.width / 2;
      const localTop = -rect.height / 2;
      if (vectorShape && typeof Path2D !== "undefined") {
        context.save();
        context.translate(localLeft, localTop);
        context.scale(
          rect.width / Math.max(1, vectorShape.viewBox.width),
          rect.height / Math.max(1, vectorShape.viewBox.height)
        );
        const path = getCachedCanvasPath2D(vectorShape.path);
        if (!path) {
          context.restore();
          context.restore();
          return;
        }
        if (fillColor !== "transparent") {
          context.fill(path);
        }
        if (strokeColor !== "transparent" && lineWidth > 0) {
          context.stroke(path);
        }
        context.restore();
      } else if (shape.geometry === "ellipse") {
        context.beginPath();
        context.ellipse(0, 0, rect.width / 2, rect.height / 2, 0, 0, Math.PI * 2);
        if (fillColor !== "transparent") {
          context.fill();
        }
        if (strokeColor !== "transparent" && lineWidth > 0) {
          context.stroke();
        }
      } else {
        const radius = shape.geometry === "roundRect" ? 12 * zoomFactor : 0;
        drawCanvasRoundedRect(context, localLeft, localTop, rect.width, rect.height, radius);
        if (fillColor !== "transparent") {
          context.fill();
        }
        if (strokeColor !== "transparent" && lineWidth > 0) {
          context.stroke();
        }
      }

      drawBakedShapeText(context, shape, localLeft, localTop, rect.width, rect.height);
      context.restore();
    };
    const drawBakedFormControl = (context: CanvasRenderingContext2D, control: XlsxFormControl, rect: XlsxImageRect) => {
      const label = resolveFormControlLabel(control);
      const stroke = paletteIsDark(palette) ? "#cbd5e1" : "#475569";
      const textColor = control.textColor ?? "#000000";
      const fontSizePx = Math.max(9 * zoomFactor, ((control.fontSizePt ?? 9) * 96 / 72) * zoomFactor);
      const iconSize = Math.min(14 * zoomFactor, Math.max(0, rect.height - 4 * zoomFactor));
      context.save();
      context.font = `400 ${fontSizePx}px ${control.fontFamily ?? "Calibri, sans-serif"}`;
      context.textBaseline = "middle";
      context.fillStyle = textColor;
      context.strokeStyle = stroke;
      context.lineWidth = Math.max(1, zoomFactor);

      if (control.kind === "group-box") {
        const labelInset = label ? Math.max(7, fontSizePx * 0.5) : 0;
        drawCanvasRoundedRect(context, rect.left, rect.top + labelInset, rect.width, Math.max(0, rect.height - labelInset), 2 * zoomFactor);
        context.stroke();
        if (label) {
          context.fillStyle = SHEET_SURFACE;
          const labelWidth = Math.min(rect.width - 16 * zoomFactor, measureCanvasTextWidth(context, label) + 8 * zoomFactor);
          context.fillRect(rect.left + 8 * zoomFactor, rect.top, labelWidth, fontSizePx * 1.2);
          context.fillStyle = textColor;
          context.textAlign = "left";
          context.fillText(label, rect.left + 12 * zoomFactor, rect.top + fontSizePx * 0.55);
        }
        context.restore();
        return;
      }

      if (control.kind === "button" || control.kind === "dropdown" || control.kind === "editbox" || control.kind === "listbox" || control.kind === "scrollbar" || control.kind === "spinner" || control.kind === "unknown") {
        if (control.kind === "button") {
          const gradient = context.createLinearGradient(rect.left, rect.top, rect.left, rect.top + rect.height);
          gradient.addColorStop(0, "#f8fafc");
          gradient.addColorStop(1, "#e2e8f0");
          context.fillStyle = gradient;
        } else {
          context.fillStyle = "transparent";
        }
        drawCanvasRoundedRect(context, rect.left, rect.top, rect.width, rect.height, control.kind === "button" ? 4 * zoomFactor : 2 * zoomFactor);
        if (control.kind === "button") {
          context.fill();
        }
        context.stroke();
      }

      let textLeft = rect.left + 2 * zoomFactor;
      const textRightInset = control.kind === "dropdown" ? 18 * zoomFactor : 6 * zoomFactor;
      if (control.kind === "checkbox") {
        context.strokeRect(rect.left + 2 * zoomFactor, rect.top + (rect.height - iconSize) / 2, iconSize, iconSize);
        if (control.checked) {
          context.fillStyle = paletteIsDark(palette) ? "#60a5fa" : "#2563eb";
          context.fillRect(rect.left + 2 * zoomFactor + 1.5, rect.top + (rect.height - iconSize) / 2 + 1.5, Math.max(0, iconSize - 3), Math.max(0, iconSize - 3));
          context.strokeStyle = paletteIsDark(palette) ? "#020617" : "#ffffff";
          context.beginPath();
          context.moveTo(rect.left + 2 * zoomFactor + iconSize * 0.24, rect.top + rect.height / 2 + iconSize * 0.06);
          context.lineTo(rect.left + 2 * zoomFactor + iconSize * 0.45, rect.top + rect.height / 2 + iconSize * 0.26);
          context.lineTo(rect.left + 2 * zoomFactor + iconSize * 0.8, rect.top + rect.height / 2 - iconSize * 0.2);
          context.stroke();
        }
        textLeft += iconSize + 4 * zoomFactor;
      } else if (control.kind === "radio") {
        context.beginPath();
        context.arc(rect.left + 2 * zoomFactor + iconSize / 2, rect.top + rect.height / 2, iconSize / 2, 0, Math.PI * 2);
        context.stroke();
        if (control.checked) {
          context.fillStyle = paletteIsDark(palette) ? "#60a5fa" : "#2563eb";
          context.beginPath();
          context.arc(rect.left + 2 * zoomFactor + iconSize / 2, rect.top + rect.height / 2, iconSize * 0.25, 0, Math.PI * 2);
          context.fill();
        }
        textLeft += iconSize + 4 * zoomFactor;
      }

      if (label) {
        const maxTextWidth = Math.max(0, rect.left + rect.width - textLeft - textRightInset);
        const text = truncateCanvasText(context, label, maxTextWidth);
        context.fillStyle = textColor;
        context.textAlign = control.textAlign === "right" ? "right" : control.textAlign === "center" ? "center" : "left";
        const textX = context.textAlign === "right"
          ? rect.left + rect.width - textRightInset
          : context.textAlign === "center"
            ? textLeft + maxTextWidth / 2
            : textLeft;
        context.fillText(text, textX, rect.top + rect.height / 2);
      }
      if (control.kind === "dropdown") {
        context.fillStyle = textColor;
        context.textAlign = "center";
        context.fillText("▼", rect.left + rect.width - 10 * zoomFactor, rect.top + rect.height / 2);
      }
      context.restore();
    };
    const drawBakedCanvasDrawings = (
      pane: FrozenDrawingPane,
      context: CanvasRenderingContext2D,
      dirtyRects: CanvasDirtyRect[]
    ) => {
      if (bakedCanvasDrawingEntries.length === 0 || dirtyRects.length === 0) {
        return;
      }

      for (const entry of bakedCanvasDrawingEntries) {
        if (resolveDrawingPane(entry.rect) !== pane) {
          continue;
        }
        const localRect = resolveBakedCanvasLocalRect(entry.rect, pane);
        if (
          localRect.left + localRect.width < 0
          || localRect.top + localRect.height < 0
          || localRect.left > paneBounds[pane].width
          || localRect.top > paneBounds[pane].height
          || !intersectsCanvasDirtyRects(localRect.left, localRect.top, localRect.width, localRect.height, dirtyRects)
        ) {
          continue;
        }

        if (entry.kind === "shape") {
          drawBakedShape(context, entry.shape, localRect);
        } else if (entry.kind === "formControl") {
          drawBakedFormControl(context, entry.control, localRect);
        } else {
          const imageElement = getCanvasImage(entry.image);
          if (imageElement) {
            context.drawImage(imageElement, localRect.left, localRect.top, localRect.width, localRect.height);
          }
        }
      }
    };

    const cellPaneOrder: FrozenDrawingPane[] = ["scroll", "top", "left", "corner"];
    if (shouldRepaintBody) {
      canvasProfileBodyStart = canvasProfileTarget ? performance.now() : 0;
      for (const pane of Object.keys(bodyContexts) as FrozenDrawingPane[]) {
        const context = bodyContexts[pane];
        const bounds = paneBounds[pane];
        if (!context || bounds.width <= 0 || bounds.height <= 0) {
          continue;
        }

        let dirtyRects = buildFullCanvasDirtyRect(bounds.width, bounds.height);
        if (canBlitBody) {
          const bufferCanvas = getBodyBlitBufferCanvas(pane);
          const deltaX = pane === "scroll" || pane === "top"
            ? drawingViewport.left - previousPaintedViewport.left
            : 0;
          const deltaY = pane === "scroll" || pane === "left"
            ? drawingViewport.top - previousPaintedViewport.top
            : 0;
          if (bufferCanvas) {
            const blittedDirtyRects = blitCanvasWithScrollDelta(
              pane === "corner"
                ? cornerBodyCanvasRef.current!
                : pane === "left"
                  ? leftBodyCanvasRef.current!
                  : pane === "top"
                    ? topBodyCanvasRef.current!
                    : scrollBodyCanvasRef.current!,
              context,
              bufferCanvas,
              dpr,
              bounds.width,
              bounds.height,
              deltaX,
              deltaY
            );
            if (blittedDirtyRects) {
              dirtyRects = blittedDirtyRects;
            }
          }
        } else {
          context.clearRect(0, 0, bounds.width, bounds.height);
        }

        bodyDirtyRectsByPane[pane] = dirtyRects;
        canvasProfileDirtyRects += dirtyRects.length;
        if (dirtyRects.length === 0) {
          continue;
        }

        context.fillStyle = sheetSurface;
        for (const dirtyRect of dirtyRects) {
          context.fillRect(dirtyRect.left, dirtyRect.top, dirtyRect.width, dirtyRect.height);
        }
      }

      for (const pane of cellPaneOrder) {
        const paneContext = bodyContexts[pane];
        const paneBoundsForCell = paneBounds[pane];
        const paneDirtyRects = bodyDirtyRectsByPane[pane];
        const paneAxisItems = canvasPaneAxisItems[pane];
        if (
          !paneContext
          || paneDirtyRects.length === 0
          || paneBoundsForCell.width <= 0
          || paneBoundsForCell.height <= 0
          || paneAxisItems.rows.length === 0
          || paneAxisItems.cols.length === 0
        ) {
          continue;
        }
        let hasPendingGridlinePath = false;
        const enqueueGridlinePath = () => {
          if (!hasPendingGridlinePath) {
            paneContext.beginPath();
            hasPendingGridlinePath = true;
          }
        };
        const flushPendingGridlines = () => {
          if (!hasPendingGridlinePath) {
            return;
          }
          paneContext.strokeStyle = palette.border;
          paneContext.lineWidth = 1;
          paneContext.stroke();
          hasPendingGridlinePath = false;
        };
        const drawnMergedAnchorKeys = new Set<string>();
        for (const rowItem of paneAxisItems.rows) {
          for (const colItem of paneAxisItems.cols) {
          const cell = { row: rowItem.actualRow, col: colItem.actualCol };
          const anchorCell = resolveMergeAnchorCell(cell);
          const anchorKey = `${anchorCell.row}:${anchorCell.col}`;
          let drawCell = anchorCell;

          const drawRowIndex = rowIndexByActual.get(drawCell.row);
          const drawColIndex = colIndexByActual.get(drawCell.col);
          if (drawRowIndex === undefined || drawColIndex === undefined) {
            continue;
          }

          const baseCellLeft = displayRowHeaderWidth + (colPrefixSums[drawColIndex] ?? 0);
          const baseCellTop = displayHeaderHeight + (rowPrefixSums[drawRowIndex] ?? 0);
          const useFrozenHorizontalPosition = pane === "left" || pane === "corner";
          const useFrozenVerticalPosition = pane === "top" || pane === "corner";
          const roughLocalRect = {
            height: displayEffectiveRowHeights[drawRowIndex] ?? rowItem.size,
            left: (
              useFrozenHorizontalPosition
                ? (stickyLeftByCol.get(drawCell.col) ?? (baseCellLeft - drawingViewport.left))
                : (baseCellLeft - drawingViewport.left)
            ) - paneBoundsForCell.left,
            top: (
              useFrozenVerticalPosition
                ? (stickyTopByRow.get(drawCell.row) ?? (baseCellTop - drawingViewport.top))
                : (baseCellTop - drawingViewport.top)
            ) - paneBoundsForCell.top,
            width: displayEffectiveColWidths[drawColIndex] ?? colItem.size
          };
          const isMergedSecondaryProbe = anchorCell.row !== cell.row || anchorCell.col !== cell.col;
          if (
            !isMergedSecondaryProbe
            && !intersectsCanvasDirtyRects(
              roughLocalRect.left - CANVAS_DIRTY_CELL_CULL_MARGIN_PX,
              roughLocalRect.top - CANVAS_DIRTY_CELL_CULL_MARGIN_PX,
              roughLocalRect.width + CANVAS_DIRTY_CELL_CULL_MARGIN_PX * 2,
              roughLocalRect.height + CANVAS_DIRTY_CELL_CULL_MARGIN_PX * 2,
              paneDirtyRects
            )
          ) {
            canvasProfileCulledCells += 1;
            continue;
          }

          canvasProfileLookedUpCells += 1;
          let cellData = getCellData(drawCell.row, drawCell.col);
          if ((cellData.colSpan || cellData.rowSpan) && drawnMergedAnchorKeys.has(anchorKey)) {
            continue;
          }
          if (cellData.isMergedSecondary) {
            continue;
          }
          if (cellData.colSpan || cellData.rowSpan) {
            drawnMergedAnchorKeys.add(anchorKey);
          }

          const displayRect = cellData.colSpan || cellData.rowSpan
            ? resolveCellDisplayRect(drawCell)
            : null;
          const baseLeft = displayRect?.left ?? baseCellLeft;
          const baseTop = displayRect?.top ?? baseCellTop;
          const localRect = {
            height: displayRect?.height ?? (displayEffectiveRowHeights[drawRowIndex] ?? rowItem.size),
            left: (
              useFrozenHorizontalPosition
                ? (stickyLeftByCol.get(drawCell.col) ?? (baseLeft - drawingViewport.left))
                : (baseLeft - drawingViewport.left)
            ) - paneBoundsForCell.left,
            top: (
              useFrozenVerticalPosition
                ? (stickyTopByRow.get(drawCell.row) ?? (baseTop - drawingViewport.top))
                : (baseTop - drawingViewport.top)
            ) - paneBoundsForCell.top,
            width: displayRect?.width ?? (displayEffectiveColWidths[drawColIndex] ?? colItem.size)
          };
          const drawableWidth = Math.max(localRect.width, cellData.spillWidth ?? 0);

          if (
            localRect.left + drawableWidth < 0
            || localRect.top + localRect.height < 0
            || localRect.left > paneBoundsForCell.width
            || localRect.top > paneBoundsForCell.height
          ) {
            continue;
          }
          if (!intersectsCanvasDirtyRects(localRect.left, localRect.top, drawableWidth, localRect.height, paneDirtyRects)) {
            canvasProfileCulledCells += 1;
            continue;
          }

          canvasProfilePaintedCells += 1;
          const cellStyle = cellData.style;
          const canvasCellStyle = cellData.canvas ?? buildCanvasCellStyleCache(cellStyle);
          const fillColor = cellData.conditionalColorScale?.color
            ?? (typeof cellStyle.backgroundColor === "string" ? cellStyle.backgroundColor : sheetSurface);
          const gradientFill = !cellData.conditionalColorScale && typeof cellStyle.backgroundImage === "string"
            ? resolveCanvasGradientFill(paneContext, localRect, cellStyle.backgroundImage)
            : null;
          const hasExplicitCellFill =
            cellData.conditionalColorScale !== null
            || gradientFill !== null
            || (typeof cellStyle.backgroundColor === "string"
              && cellStyle.backgroundColor !== sheetSurface);
          if (hasExplicitCellFill || cellData.chartHighlight) {
            flushPendingGridlines();
          }
          paneContext.fillStyle = gradientFill ?? fillColor;
          paneContext.fillRect(localRect.left, localRect.top, localRect.width, localRect.height);
          if (cellData.chartHighlight) {
            paneContext.fillStyle = cellData.chartHighlight.fillColor;
            paneContext.fillRect(
              localRect.left + 1,
              localRect.top + 1,
              Math.max(0, localRect.width - 2),
              Math.max(0, localRect.height - 2)
            );
          }

          if (cellData.conditionalDataBar) {
            const barLeft = localRect.left + 4 * zoomFactor;
            const barTop = localRect.top + 4 * zoomFactor;
            const barWidth = Math.max(0, (localRect.width - (8 * zoomFactor)) * (cellData.conditionalDataBar.widthPercent / 100));
            const barHeight = Math.max(0, localRect.height - (8 * zoomFactor));
            if (barWidth > 0 && barHeight > 0) {
              paneContext.fillStyle = resolveCanvasDataBarFill(
                paneContext,
                barLeft,
                barTop,
                barWidth,
                barHeight,
                cellData.conditionalDataBar
              );
              paneContext.fillRect(barLeft, barTop, barWidth, barHeight);
              if (cellData.conditionalDataBar.border !== false && cellData.conditionalDataBar.borderColor) {
                paneContext.strokeStyle = cellData.conditionalDataBar.borderColor;
                paneContext.lineWidth = 1;
                paneContext.strokeRect(barLeft + 0.5, barTop + 0.5, Math.max(0, barWidth - 1), Math.max(0, barHeight - 1));
              }
            }
          }

          const topBorder = canvasCellStyle.topBorder;
          const rightBorder = canvasCellStyle.rightBorder;
          const bottomBorder = canvasCellStyle.bottomBorder;
          const leftBorder = canvasCellStyle.leftBorder;
          const rightNeighborCol = visibleCols[drawColIndex + Math.max(1, cellData.colSpan ?? 1)];
          const rightNeighborData = rightNeighborCol === undefined
            ? null
            : getCellData(drawCell.row, rightNeighborCol);
          const rightNeighborLeftBorder = rightNeighborData?.isMergedSecondary
            ? null
            : (rightNeighborData?.canvas?.leftBorder ?? parseCanvasBorderDeclaration(rightNeighborData?.style.borderLeft));
          const bottomNeighborRow = visibleRows[drawRowIndex + Math.max(1, cellData.rowSpan ?? 1)];
          const bottomNeighborData = bottomNeighborRow === undefined
            ? null
            : getCellData(bottomNeighborRow, drawCell.col);
          const bottomNeighborTopBorder = bottomNeighborData?.isMergedSecondary
            ? null
            : (bottomNeighborData?.canvas?.topBorder ?? parseCanvasBorderDeclaration(bottomNeighborData?.style.borderTop));
          const resolvedRightBorder = resolveCanvasBoundaryBorder(rightBorder, rightNeighborLeftBorder);
          const resolvedBottomBorder = resolveCanvasBoundaryBorder(bottomBorder, bottomNeighborTopBorder);

          if (showGridLines && !hasExplicitCellFill && (!resolvedRightBorder || !resolvedBottomBorder)) {
            enqueueGridlinePath();
            if (!resolvedRightBorder) {
              paneContext.moveTo(localRect.left + localRect.width - 0.5, localRect.top);
              paneContext.lineTo(localRect.left + localRect.width - 0.5, localRect.top + localRect.height);
            }
            if (!resolvedBottomBorder) {
              paneContext.moveTo(localRect.left, localRect.top + localRect.height - 0.5);
              paneContext.lineTo(localRect.left + localRect.width, localRect.top + localRect.height - 0.5);
            }
          }

          if (topBorder || resolvedRightBorder || resolvedBottomBorder || leftBorder || cellData.chartHighlight) {
            flushPendingGridlines();
          }

          if (topBorder && drawRowIndex === 0) {
            strokeCanvasBorderSide(paneContext, "top", localRect, topBorder);
          }
          if (resolvedRightBorder) {
            strokeCanvasBorderSide(paneContext, "right", localRect, resolvedRightBorder);
          }
          if (resolvedBottomBorder) {
            strokeCanvasBorderSide(paneContext, "bottom", localRect, resolvedBottomBorder);
          }
          if (leftBorder && drawColIndex === 0) {
            strokeCanvasBorderSide(paneContext, "left", localRect, leftBorder);
          }
          if (cellData.chartHighlight) {
            const highlightBorder = {
              color: cellData.chartHighlight.strokeColor,
              style: "solid",
              width: Math.max(1, zoomFactor)
            };
            if (cellData.chartHighlight.borderTop) {
              strokeCanvasBorderSide(paneContext, "top", localRect, highlightBorder);
            }
            if (cellData.chartHighlight.borderRight) {
              strokeCanvasBorderSide(paneContext, "right", localRect, highlightBorder);
            }
            if (cellData.chartHighlight.borderBottom) {
              strokeCanvasBorderSide(paneContext, "bottom", localRect, highlightBorder);
            }
            if (cellData.chartHighlight.borderLeft) {
              strokeCanvasBorderSide(paneContext, "left", localRect, highlightBorder);
            }
          }

          const rawText = cellData.value ?? "";
          const shouldDrawCanvasContent = (
            cellData.checkboxState != null
            || cellData.sparkline
            || rawText.length > 0
            || cellData.conditionalIcon
            || cellData.isTableHeader
          );
          if (!shouldDrawCanvasContent) {
            continue;
          }

          const padding = canvasCellStyle.padding;
          const contentLeft = localRect.left + padding.left;
          const contentTop = localRect.top + padding.top;
          const contentWidth = Math.max(0, localRect.width - padding.left - padding.right);
          const contentHeight = Math.max(0, localRect.height - padding.top - padding.bottom);
          if (contentWidth <= 0 || contentHeight <= 0) {
            continue;
          }

          const activeFontSizePx = cellData.shrinkToFitFontSizePx
            ?? resolveCanvasFontSizePx(cellStyle, 12 * zoomFactor);
          const textClipOverscan = Math.max(
            1,
            zoomFactor * 1.5,
            activeFontSizePx * (cellData.textRotationDeg ? 0.75 : 0.18)
          );

          flushPendingGridlines();
          paneContext.save();
          paneContext.beginPath();
          paneContext.rect(
            contentLeft - textClipOverscan,
            contentTop - textClipOverscan,
            contentWidth + (textClipOverscan * 2),
            contentHeight + (textClipOverscan * 2)
          );
          paneContext.clip();
          paneContext.font = cellData.shrinkToFitFontSizePx
            ? resolveCanvasFontWithPxSize(cellStyle, cellData.shrinkToFitFontSizePx)
            : canvasCellStyle.baseFont;
          paneContext.fillStyle = canvasCellStyle.textColor;
          paneContext.textBaseline = "middle";

          if (cellData.checkboxState != null) {
            const boxSize = Math.min(14 * zoomFactor, contentWidth, contentHeight);
            const boxLeft = localRect.left + (localRect.width - boxSize) / 2;
            const boxTop = localRect.top + (localRect.height - boxSize) / 2;
            paneContext.strokeStyle = paletteIsDark(palette) ? "#cbd5e1" : "#475569";
            paneContext.lineWidth = 1.25;
            paneContext.strokeRect(boxLeft, boxTop, boxSize, boxSize);
            if (cellData.checkboxState) {
              paneContext.fillStyle = paletteIsDark(palette) ? "#60a5fa" : "#2563eb";
              paneContext.fillRect(boxLeft + 1.5, boxTop + 1.5, Math.max(0, boxSize - 3), Math.max(0, boxSize - 3));
              paneContext.strokeStyle = paletteIsDark(palette) ? "#020617" : "#ffffff";
              paneContext.lineWidth = 1.5;
              paneContext.beginPath();
              paneContext.moveTo(boxLeft + (boxSize * 0.24), boxTop + (boxSize * 0.56));
              paneContext.lineTo(boxLeft + (boxSize * 0.45), boxTop + (boxSize * 0.76));
              paneContext.lineTo(boxLeft + (boxSize * 0.8), boxTop + (boxSize * 0.3));
              paneContext.stroke();
            }
          } else if (cellData.sparkline) {
            const sparkline = cellData.sparkline.config;
            const sparklineValues = cellData.sparkline.values;
            const points = sparklineValues
              .map((value, index) => ({ index, value }))
              .filter((entry): entry is { index: number; value: number } => typeof entry.value === "number" && Number.isFinite(entry.value));
            if (points.length > 0) {
              const negativeColor = sparkline.negativeColor ?? "#c2410c";
              const seriesColor = sparkline.color ?? "#2563eb";
              const markerColor = sparkline.markerColor ?? seriesColor;
              const sparkLeft = contentLeft + 1;
              const sparkTop = contentTop + 2;
              const sparkWidth = Math.max(1, contentWidth - 2);
              const sparkHeight = Math.max(1, contentHeight - 4);

              if (sparkline.type === "winLoss") {
                const normalizedValues = points.map((entry) => ({ ...entry, value: entry.value >= 0 ? 1 : -1 }));
                const segmentWidth = Math.max(4 * zoomFactor, sparkWidth / Math.max(normalizedValues.length * 1.9, 1));
                const gap = normalizedValues.length > 1
                  ? (sparkWidth - segmentWidth * normalizedValues.length) / (normalizedValues.length - 1)
                  : 0;
                const positiveY = sparkTop + Math.max(1.5 * zoomFactor, sparkHeight * 0.18);
                const negativeY = sparkTop + sparkHeight - Math.max(1.5 * zoomFactor, sparkHeight * 0.18);
                paneContext.strokeStyle = seriesColor;
                paneContext.lineCap = "round";
                paneContext.lineWidth = Math.max(1.5, 1.8 * zoomFactor);
                normalizedValues.forEach((entry, index) => {
                  const left = sparkLeft + index * (segmentWidth + Math.max(0, gap));
                  const y = entry.value >= 0 ? positiveY : negativeY;
                  paneContext.beginPath();
                  paneContext.moveTo(left, y);
                  paneContext.lineTo(left + segmentWidth, y);
                  paneContext.stroke();
                });
              } else if (sparkline.type === "column") {
                const minValue = Math.min(0, ...points.map((entry) => entry.value));
                const maxValue = Math.max(0, ...points.map((entry) => entry.value));
                const zeroY = sparkTop + sparkHeight - clampSparklineValue(0, minValue, maxValue) * sparkHeight;
                const barWidth = Math.max(2 * zoomFactor, sparkWidth / Math.max(points.length * 1.8, 1));
                const gap = points.length > 1 ? (sparkWidth - barWidth * points.length) / (points.length - 1) : 0;
                paneContext.strokeStyle = palette.border;
                paneContext.lineWidth = 1;
                paneContext.beginPath();
                paneContext.moveTo(sparkLeft, zeroY);
                paneContext.lineTo(sparkLeft + sparkWidth, zeroY);
                paneContext.stroke();
                points.forEach((entry, index) => {
                  const left = sparkLeft + index * (barWidth + Math.max(0, gap));
                  const y = sparkTop + sparkHeight - clampSparklineValue(entry.value, minValue, maxValue) * sparkHeight;
                  const top = Math.min(y, zeroY);
                  const barHeight = Math.max(1, Math.abs(y - zeroY));
                  paneContext.fillStyle = entry.value < 0 ? negativeColor : seriesColor;
                  paneContext.fillRect(left, top, barWidth, barHeight);
                });
              } else if (points.length > 1) {
                const minValue = Math.min(...points.map((entry) => entry.value));
                const maxValue = Math.max(...points.map((entry) => entry.value));
                const xStep = points.length > 1 ? sparkWidth / (points.length - 1) : 0;
                paneContext.strokeStyle = seriesColor;
                paneContext.lineCap = "round";
                paneContext.lineJoin = "round";
                paneContext.lineWidth = Math.max(1.2, 1.6 * zoomFactor);
                paneContext.beginPath();
                points.forEach((entry, index) => {
                  const x = sparkLeft + index * xStep;
                  const y = sparkTop + sparkHeight - clampSparklineValue(entry.value, minValue, maxValue) * sparkHeight;
                  if (index === 0) {
                    paneContext.moveTo(x, y);
                  } else {
                    paneContext.lineTo(x, y);
                  }
                });
                paneContext.stroke();

                if (sparkline.markers) {
                  const highValue = Math.max(...points.map((entry) => entry.value));
                  const lowValue = Math.min(...points.map((entry) => entry.value));
                  points.forEach((entry, index) => {
                    const x = sparkLeft + index * xStep;
                    const y = sparkTop + sparkHeight - clampSparklineValue(entry.value, minValue, maxValue) * sparkHeight;
                    let fill = markerColor;
                    if (entry.value === highValue && sparkline.highColor) {
                      fill = sparkline.highColor;
                    } else if (entry.value === lowValue && sparkline.lowColor) {
                      fill = sparkline.lowColor;
                    } else if (index === 0 && sparkline.firstColor) {
                      fill = sparkline.firstColor;
                    } else if (index === points.length - 1 && sparkline.lastColor) {
                      fill = sparkline.lastColor;
                    } else if (entry.value < 0 && sparkline.negative && sparkline.negativeColor) {
                      fill = sparkline.negativeColor;
                    }
                    paneContext.fillStyle = fill;
                    paneContext.beginPath();
                    paneContext.arc(x, y, Math.max(1.25, 1.75 * zoomFactor), 0, Math.PI * 2);
                    paneContext.fill();
                  });
                }
              }
            }
          } else {
            const align = canvasCellStyle.textAlign;
            paneContext.textAlign = align;
            const textX = align === "right"
              ? contentLeft + contentWidth
              : align === "center"
                ? contentLeft + (contentWidth / 2)
                : contentLeft;
            const trailingInset = (cellData.conditionalIcon ? 18 * zoomFactor : 0) + (cellData.isTableHeader ? 16 * zoomFactor : 0);
            const spillMaxWidth = cellData.spillWidth && cellData.spillWidth > 0
              ? Math.max(0, cellData.spillWidth - trailingInset)
              : null;
            const maxTextWidth = spillMaxWidth ?? Math.max(0, contentWidth - trailingInset);
            const textColor = canvasCellStyle.textColor;
            const shouldEllipsizeText = canvasCellStyle.textOverflowEllipsis;
            const shouldWrapText = canvasCellStyle.usesWrappedText || rawText.includes("\n");
            const singleLineHeight = cellData.shrinkToFitFontSizePx
              ? resolveCanvasLineHeight(cellStyle, cellData.shrinkToFitFontSizePx)
              : resolveCanvasLineHeight(cellStyle, 12 * zoomFactor);

            if (cellData.textRotationDeg && rawText.length > 0) {
              const textY = resolveCanvasTextMiddleY(cellStyle.verticalAlign, contentTop, contentHeight, singleLineHeight);
              const rotationOriginX = contentLeft + (contentWidth / 2);
              paneContext.save();
              paneContext.translate(rotationOriginX, textY);
              paneContext.rotate((cellData.textRotationDeg * Math.PI) / 180);
              paneContext.textAlign = align;
              paneContext.fillText(
                rawText,
                align === "right"
                  ? contentWidth / 2
                  : align === "center"
                    ? 0
                    : -(contentWidth / 2),
                0
              );
              paneContext.restore();
            } else if (shouldWrapText) {
              const wrappedLines = wrapCanvasText(paneContext, rawText, maxTextWidth);
              const lineHeight = cellData.shrinkToFitFontSizePx
                ? resolveCanvasLineHeight(cellStyle, cellData.shrinkToFitFontSizePx)
                : resolveCanvasLineHeight(cellStyle, 12 * zoomFactor);
              const textBlockHeight = wrappedLines.length * lineHeight;
              const verticalAlign = cellStyle.verticalAlign;
              let textBlockTop = contentTop;
              if (verticalAlign === "middle") {
                textBlockTop = contentTop + ((contentHeight - textBlockHeight) / 2);
              } else if (verticalAlign !== "top") {
                textBlockTop = contentTop + contentHeight - textBlockHeight;
              }

              wrappedLines.forEach((line, lineIndex) => {
                const textY = textBlockTop + (lineIndex * lineHeight) + (lineHeight / 2);
                paneContext.fillText(line, textX, textY);
                drawCanvasTextDecorations(paneContext, {
                  align,
                  color: textColor,
                  decoration: canvasCellStyle.textDecoration,
                  lineThroughY: textY,
                  maxWidth: maxTextWidth,
                  text: line,
                  textX,
                  underlineY: textY + Math.max(2, lineHeight * 0.24),
                  zoomFactor
                });
              });
            } else if (spillMaxWidth != null) {
              const text = shouldEllipsizeText ? truncateCanvasText(paneContext, rawText, maxTextWidth) : rawText;
              const textY = resolveCanvasTextMiddleY(cellStyle.verticalAlign, contentTop, contentHeight, singleLineHeight);
              deferredSpillTextsByPane[pane].push({
                align,
                clipHeight: contentHeight + (textClipOverscan * 2),
                clipLeft: contentLeft - textClipOverscan,
                clipTop: contentTop - textClipOverscan,
                clipWidth: spillMaxWidth + (textClipOverscan * 2),
                color: textColor,
                ellipsize: shouldEllipsizeText,
                font: paneContext.font,
                maxWidth: spillMaxWidth,
                text,
                textDecoration: canvasCellStyle.textDecoration,
                textX,
                textY,
                lineThroughY: textY,
                underlineY: textY + (6 * zoomFactor)
              });
            } else {
              const text = cellData.shrinkToFit
                ? rawText
                : shouldEllipsizeText
                  ? truncateCanvasText(paneContext, rawText, maxTextWidth)
                  : rawText;
              const textY = resolveCanvasTextMiddleY(cellStyle.verticalAlign, contentTop, contentHeight, singleLineHeight);
              paneContext.fillText(text, textX, textY);
              drawCanvasTextDecorations(paneContext, {
                align,
                color: textColor,
                decoration: canvasCellStyle.textDecoration,
                ellipsize: shouldEllipsizeText,
                lineThroughY: textY,
                maxWidth: maxTextWidth,
                text,
                textX,
                underlineY: textY + 6 * zoomFactor,
                zoomFactor
              });
            }
          }

          if (cellData.conditionalIcon) {
            const iconSize = 10 * zoomFactor;
            const iconX = localRect.left + localRect.width - (padding.right + iconSize + (cellData.isTableHeader ? 16 * zoomFactor : 4 * zoomFactor));
            const iconY = localRect.top + (localRect.height / 2);
            drawCanvasConditionalIcon(paneContext, cellData.conditionalIcon, iconX + (iconSize / 2), iconY, iconSize);
          }

          if (cellData.isTableHeader) {
            paneContext.fillStyle = palette.mutedText;
            paneContext.textAlign = "center";
            paneContext.fillText("▾", localRect.left + localRect.width - (10 * zoomFactor), localRect.top + (localRect.height / 2));
          }

          paneContext.restore();
        }
      }
        flushPendingGridlines();
      }

      for (const pane of cellPaneOrder) {
        const paneContext = bodyContexts[pane];
        if (!paneContext) {
          continue;
        }
        for (const spillText of deferredSpillTextsByPane[pane]) {
          paneContext.save();
          paneContext.beginPath();
          paneContext.rect(spillText.clipLeft, spillText.clipTop, spillText.clipWidth, spillText.clipHeight);
          paneContext.clip();
          paneContext.font = spillText.font;
          paneContext.fillStyle = spillText.color;
          paneContext.textAlign = spillText.align;
          paneContext.textBaseline = "middle";
          paneContext.fillText(spillText.text, spillText.textX, spillText.textY);
          drawCanvasTextDecorations(paneContext, {
            align: spillText.align,
            color: spillText.color,
            decoration: spillText.textDecoration,
            ellipsize: spillText.ellipsize,
            lineThroughY: spillText.lineThroughY,
            maxWidth: spillText.maxWidth,
            text: spillText.text,
            textX: spillText.textX,
            underlineY: spillText.underlineY,
            zoomFactor
          });
          paneContext.restore();
        }
      }

      for (const pane of cellPaneOrder) {
        const paneContext = bodyContexts[pane];
        if (!paneContext) {
          continue;
        }
        drawBakedCanvasDrawings(pane, paneContext, bodyDirtyRectsByPane[pane]);
      }
      canvasProfileBodyMs = canvasProfileTarget ? performance.now() - canvasProfileBodyStart : 0;
    }

    if (shouldRepaintHeaders && cornerContext) {
      const topFrozenHeaderContext = topHeaderContexts.frozen;
      const topScrollHeaderContext = topHeaderContexts.scroll;
      const leftFrozenHeaderContext = leftHeaderContexts.frozen;
      const leftScrollHeaderContext = leftHeaderContexts.scroll;

      if (canBlitTopHeader) {
        const bufferCanvas = getHeaderBlitBufferCanvas("top");
        if (bufferCanvas && topScrollHeaderContext && topScrollHeaderCanvasRef.current) {
          const blittedDirtyRects = blitCanvasWithScrollDelta(
            topScrollHeaderCanvasRef.current,
            topScrollHeaderContext,
            bufferCanvas,
            dpr,
            topScrollHeaderCanvasWidth,
            headerHeight,
            drawingViewport.left - previousPaintedViewport.left,
            0
          );
          if (blittedDirtyRects) {
            topScrollHeaderDirtyRects = blittedDirtyRects;
          }
        }
      }
      if (canBlitLeftHeader) {
        const bufferCanvas = getHeaderBlitBufferCanvas("left");
        if (bufferCanvas && leftScrollHeaderContext && leftScrollHeaderCanvasRef.current) {
          const blittedDirtyRects = blitCanvasWithScrollDelta(
            leftScrollHeaderCanvasRef.current,
            leftScrollHeaderContext,
            bufferCanvas,
            dpr,
            rowHeaderWidth,
            leftScrollHeaderCanvasHeight,
            0,
            drawingViewport.top - previousPaintedViewport.top
          );
          if (blittedDirtyRects) {
            leftScrollHeaderDirtyRects = blittedDirtyRects;
          }
        }
      }

      if (topFrozenHeaderContext && topFrozenHeaderCanvasWidth > 0) {
        topFrozenHeaderContext.fillStyle = palette.headerSurface;
        topFrozenHeaderContext.fillRect(0, 0, topFrozenHeaderCanvasWidth, headerHeight);
      }
      if (topScrollHeaderContext && topScrollHeaderCanvasWidth > 0) {
        topScrollHeaderContext.fillStyle = palette.headerSurface;
        for (const dirtyRect of topScrollHeaderDirtyRects) {
          topScrollHeaderContext.fillRect(dirtyRect.left, dirtyRect.top, dirtyRect.width, dirtyRect.height);
        }
      }
      if (topFrozenHeaderContext) {
        topFrozenHeaderContext.strokeStyle = palette.border;
        topFrozenHeaderContext.lineWidth = 1;
      }
      if (topScrollHeaderContext) {
        topScrollHeaderContext.strokeStyle = palette.border;
        topScrollHeaderContext.lineWidth = 1;
      }
      for (const column of canvasColumnHeaderCells) {
        const paneContext = column.isFrozen ? topFrozenHeaderContext : topScrollHeaderContext;
        if (!paneContext) {
          continue;
        }
        if (
          column.localLeft + column.width < 0
          || column.localLeft > (column.isFrozen ? topFrozenHeaderCanvasWidth : topScrollHeaderCanvasWidth)
        ) {
          continue;
        }
        if (
          !column.isFrozen
          && !intersectsCanvasDirtyRects(column.localLeft, 0, column.width, column.height, topScrollHeaderDirtyRects)
        ) {
          continue;
        }

        const selected =
          normalizedVisibleRange
          && column.actualCol >= normalizedVisibleRange.start.col
          && column.actualCol <= normalizedVisibleRange.end.col;
        paneContext.fillStyle = selected ? selectionHeaderSurface : palette.headerSurface;
        paneContext.fillRect(column.localLeft, 0, column.width, column.height);
        paneContext.strokeStyle = palette.border;
        paneContext.beginPath();
        paneContext.moveTo(column.localLeft + column.width - 0.5, 0);
        paneContext.lineTo(column.localLeft + column.width - 0.5, column.height);
        paneContext.moveTo(column.localLeft, column.height - 0.5);
        paneContext.lineTo(column.localLeft + column.width, column.height - 0.5);
        paneContext.stroke();
        paneContext.font = `600 ${11 * zoomFactor}px ui-sans-serif, system-ui, sans-serif`;
        paneContext.fillStyle = palette.mutedText;
        paneContext.textAlign = "center";
        paneContext.textBaseline = "middle";
        paneContext.fillText(
          columnLabel(column.actualCol),
          column.localLeft + (column.width / 2),
          column.height / 2
        );
      }

      if (leftFrozenHeaderContext && leftFrozenHeaderCanvasHeight > 0) {
        leftFrozenHeaderContext.fillStyle = palette.rowHeaderSurface;
        leftFrozenHeaderContext.fillRect(0, 0, rowHeaderWidth, leftFrozenHeaderCanvasHeight);
      }
      if (leftScrollHeaderContext && leftScrollHeaderCanvasHeight > 0) {
        leftScrollHeaderContext.fillStyle = palette.rowHeaderSurface;
        for (const dirtyRect of leftScrollHeaderDirtyRects) {
          leftScrollHeaderContext.fillRect(dirtyRect.left, dirtyRect.top, dirtyRect.width, dirtyRect.height);
        }
      }
      if (leftFrozenHeaderContext) {
        leftFrozenHeaderContext.strokeStyle = palette.border;
        leftFrozenHeaderContext.lineWidth = 1;
      }
      if (leftScrollHeaderContext) {
        leftScrollHeaderContext.strokeStyle = palette.border;
        leftScrollHeaderContext.lineWidth = 1;
      }
      for (const row of canvasRowHeaderCells) {
        const paneContext = row.isFrozen ? leftFrozenHeaderContext : leftScrollHeaderContext;
        if (!paneContext) {
          continue;
        }
        if (
          row.localTop + row.height < 0
          || row.localTop > (row.isFrozen ? leftFrozenHeaderCanvasHeight : leftScrollHeaderCanvasHeight)
        ) {
          continue;
        }
        if (
          !row.isFrozen
          && !intersectsCanvasDirtyRects(0, row.localTop, rowHeaderWidth, row.height, leftScrollHeaderDirtyRects)
        ) {
          continue;
        }

        const selected =
          normalizedVisibleRange
          && row.actualRow >= normalizedVisibleRange.start.row
          && row.actualRow <= normalizedVisibleRange.end.row;
        paneContext.fillStyle = selected ? selectionHeaderSurface : palette.rowHeaderSurface;
        paneContext.fillRect(0, row.localTop, rowHeaderWidth, row.height);
        paneContext.beginPath();
        paneContext.moveTo(0, row.localTop + row.height - 0.5);
        paneContext.lineTo(rowHeaderWidth, row.localTop + row.height - 0.5);
        paneContext.moveTo(rowHeaderWidth - 0.5, row.localTop);
        paneContext.lineTo(rowHeaderWidth - 0.5, row.localTop + row.height);
        paneContext.stroke();
        paneContext.font = `600 ${11 * zoomFactor}px ui-sans-serif, system-ui, sans-serif`;
        paneContext.fillStyle = palette.mutedText;
        paneContext.textAlign = "center";
        paneContext.textBaseline = "middle";
        paneContext.fillText(`${row.actualRow + 1}`, rowHeaderWidth / 2, row.localTop + (row.height / 2));
      }

      cornerContext.fillStyle = palette.headerSurface;
      cornerContext.fillRect(0, 0, rowHeaderWidth, headerHeight);
      cornerContext.strokeStyle = palette.border;
      cornerContext.lineWidth = 1;
      cornerContext.beginPath();
      cornerContext.moveTo(rowHeaderWidth - 0.5, 0);
      cornerContext.lineTo(rowHeaderWidth - 0.5, headerHeight);
      cornerContext.moveTo(0, headerHeight - 0.5);
      cornerContext.lineTo(rowHeaderWidth, headerHeight - 0.5);
      cornerContext.stroke();
    }
    paintedDrawingViewportRef.current = drawingViewport;
    if (shouldRepaintBody) {
      paintedBodyCanvasSignatureRef.current = nextBodyCanvasSignature;
    }
    if (shouldRepaintHeaders) {
      paintedHeaderCanvasSignatureRef.current = nextHeaderCanvasSignature;
    }
    applyCanvasViewportCompensation();
    if (canvasProfileTarget) {
      canvasProfileTarget.push({
        bodyMs: canvasProfileBodyMs,
        culledCells: canvasProfileCulledCells,
        dirtyRects: canvasProfileDirtyRects,
        lookedUpCells: canvasProfileLookedUpCells,
        paintedCells: canvasProfilePaintedCells,
        repaintBody: shouldRepaintBody,
        repaintHeaders: shouldRepaintHeaders,
        renderToLayoutMs: canvasProfileStart - xlsxGridRenderStart,
        totalMs: performance.now() - canvasProfileStart
      });
      if (canvasProfileTarget.length > 500) {
        canvasProfileTarget.splice(0, canvasProfileTarget.length - 500);
      }
    }
  }, [
    activeSheet,
    applyCanvasViewportCompensation,
    bakedCanvasDrawingSignature,
    bakedFormControlRects,
    bakedImageRects,
    bakedShapeRects,
    canvasColumnHeaderCells,
    canvasPaneAxisItems,
    canvasRowHeaderCells,
    canvasVisibleColItems,
    canvasVisibleRowItems,
    colIndexByActual,
    colPrefixSums,
    displayEffectiveColWidths,
    displayEffectiveRowHeights,
    displayHeaderHeight,
    displayRowHeaderWidth,
    displayedSelection,
    drawingViewport.left,
    drawingViewport.top,
    drawingViewport.height,
    drawingViewport.width,
    experimentalCanvas,
    frozenPaneBottom,
    frozenPaneRight,
    getCellData,
    getCanvasImage,
    getBodyBlitBufferCanvas,
    getHeaderBlitBufferCanvas,
    palette,
    resolveCellDisplayRect,
    resolveMergeAnchorCell,
    rowIndexByActual,
    rowPrefixSums,
    selectionHeaderSurface,
    shouldBakeCanvasStaticDrawings,
    stickyLeftByCol,
    stickyTopByRow,
    visibleCols,
    visibleRows,
    zoomFactor
  ]);

  const startChartMove = React.useCallback((
    event: React.PointerEvent<HTMLElement>,
    chart: XlsxChart,
    rect: XlsxImageRect
  ) => {
    if (event.button !== 0) {
      return;
    }

    event.preventDefault();
    event.stopPropagation();
    focusGrid();
    selectChart(chart.id);

    if (readOnlyRef.current || chart.editable === false) {
      return;
    }

    event.currentTarget.setPointerCapture?.(event.pointerId);
    chartInteractionRef.current = {
      baseRect: rect,
      chartId: chart.id,
      didMove: false,
      pointerId: event.pointerId,
      startClientX: event.clientX,
      startClientY: event.clientY,
      type: "move"
    };
    chartPreviewRectRef.current = { id: chart.id, rect };
    setChartPreviewRect({ id: chart.id, rect });
    setInteractionMode("select");
    document.body.style.cursor = "move";
    document.body.style.userSelect = "none";
    installChartInteractionListeners(event.pointerId);
  }, [focusGrid, selectChart]);

  const startChartResize = React.useCallback((
    event: React.PointerEvent<HTMLElement>,
    chart: XlsxChart,
    rect: XlsxImageRect,
    handle: XlsxImageResizeHandlePosition
  ) => {
    if (readOnlyRef.current || chart.editable === false || event.button !== 0) {
      return;
    }

    event.preventDefault();
    event.stopPropagation();
    focusGrid();
    selectChart(chart.id);
    event.currentTarget.setPointerCapture?.(event.pointerId);
    chartInteractionRef.current = {
      baseRect: rect,
      chartId: chart.id,
      didMove: false,
      handle,
      pointerId: event.pointerId,
      startClientX: event.clientX,
      startClientY: event.clientY,
      type: "resize"
    };
    chartPreviewRectRef.current = { id: chart.id, rect };
    setChartPreviewRect({ id: chart.id, rect });
    setInteractionMode("select");
    document.body.style.cursor = String(IMAGE_HANDLE_CURSOR[handle]);
    document.body.style.userSelect = "none";
    installChartInteractionListeners(event.pointerId);
  }, [focusGrid, selectChart]);

  const startImageMove = React.useCallback((
    event: React.PointerEvent<HTMLElement>,
    image: XlsxImage,
    rect: XlsxImageRect
  ) => {
    if (event.button !== 0) {
      return;
    }

    event.preventDefault();
    event.stopPropagation();
    focusGrid();
    selectImage(image.id);

    if (readOnlyRef.current || image.editable === false) {
      return;
    }

    event.currentTarget.setPointerCapture?.(event.pointerId);
    imageInteractionRef.current = {
      baseRect: rect,
      didMove: false,
      imageId: image.id,
      pointerId: event.pointerId,
      startClientX: event.clientX,
      startClientY: event.clientY,
      type: "move"
    };
    imagePreviewRectRef.current = { id: image.id, rect };
    setImagePreviewRect({ id: image.id, rect });
    setInteractionMode("select");
    document.body.style.cursor = "move";
    document.body.style.userSelect = "none";
    installImageInteractionListeners(event.pointerId);
  }, [focusGrid, selectImage]);

  const startImageResize = React.useCallback((
    event: React.PointerEvent<HTMLElement>,
    image: XlsxImage,
    rect: XlsxImageRect,
    handle: XlsxImageResizeHandlePosition
  ) => {
    if (readOnlyRef.current || image.editable === false || event.button !== 0) {
      return;
    }

    event.preventDefault();
    event.stopPropagation();
    focusGrid();
    selectImage(image.id);
    event.currentTarget.setPointerCapture?.(event.pointerId);
    imageInteractionRef.current = {
      baseRect: rect,
      didMove: false,
      handle,
      imageId: image.id,
      pointerId: event.pointerId,
      startClientX: event.clientX,
      startClientY: event.clientY,
      type: "resize"
    };
    imagePreviewRectRef.current = { id: image.id, rect };
    setImagePreviewRect({ id: image.id, rect });
    setInteractionMode("select");
    document.body.style.cursor = String(IMAGE_HANDLE_CURSOR[handle]);
    document.body.style.userSelect = "none";
    installImageInteractionListeners(event.pointerId);
  }, [focusGrid, selectImage]);

  const handleImageClick = React.useCallback((image: XlsxImage) => {
    if (skipNextImageClickRef.current === image.id) {
      skipNextImageClickRef.current = null;
      return;
    }

    if (image.hyperlink) {
      openHyperlink(image.hyperlink);
    }
  }, [openHyperlink]);

  const handleShapeClick = React.useCallback((shape: XlsxShape) => {
    if (shape.hyperlink) {
      openHyperlink(shape.hyperlink);
    }
  }, [openHyperlink]);

  const handleChartClick = React.useCallback((chart: XlsxChart) => {
    if (skipNextChartClickRef.current === chart.id) {
      skipNextChartClickRef.current = null;
      return;
    }

    selectChart(chart.id);
  }, [selectChart]);

  const handleChartElementPointerDown = React.useCallback((selection: XlsxChartElementSelection) => {
    skipNextChartClickRef.current = selection.chartId;
    selectChartElement(selection);
  }, [selectChartElement]);

  const handleChartElementDoubleClick = React.useCallback((selection: XlsxChartElementSelection) => {
    selectChartElement(selection);
  }, [selectChartElement]);

  if (isLoading) {
    return <>{renderLoading(loadingComponent, loadingState, palette)}</>;
  }

  if (isLoadDeferred) {
    return <>{renderDeferredLoad({ ...controller, canLoadDeferred, continueDeferredLoad, deferredLoadFileSize, isLoadDeferred }, palette)}</>;
  }

  if (error) {
    if (error instanceof XlsxFileSizeLimitExceededError) {
      return (
        <>
          {renderFileTooLarge(
            fileTooLargeState,
            {
              displayFileName,
              fileSizeBytes: error.fileSizeBytes,
              maxFileSizeBytes: error.maxFileSizeBytes
            },
            palette
          )}
        </>
      );
    }

    return <>{renderError(errorState, error, palette)}</>;
  }

  if (!activeSheet && activeTab?.kind === "chartsheet") {
    return (
      <div
        style={{
          alignItems: "stretch",
          backgroundColor: palette.canvas,
          display: "flex",
          flex: 1,
          justifyContent: "center",
          minHeight: 0,
          minWidth: 0,
          padding: 16
        }}
      >
        <div
          style={{
            backgroundColor: palette.surface,
            border: `1px solid ${palette.border}`,
            borderRadius: 12,
            display: "grid",
            gap: 16,
            gridTemplateColumns: charts.length > 1 ? "repeat(auto-fit, minmax(320px, 1fr))" : "minmax(320px, 1fr)",
            padding: 16,
            width: "100%"
          }}
        >
          {charts.length > 0 ? charts.map((chart) => {
            const chartsheetRect = { height: 320, left: 0, top: 0, width: 640 };
            return (
              <div key={chart.id} style={{ minHeight: 320, position: "relative" }}>
                  {isChartsLoading
                    ? renderChartLoadingNode(renderChartLoading, chart, chartsheetRect)
                    : (
                        <MemoChartSvg
                          chart={chart}
                          onChartElementDoubleClick={handleChartElementDoubleClick}
                          onChartElementPointerDown={handleChartElementPointerDown}
                          palette={palette}
                          rect={chartsheetRect}
                          selectedChartElement={selectedChartElement}
                        />
                      )}
              </div>
            );
          }) : (
            <div
              style={{
                alignItems: "center",
                color: palette.mutedText,
                display: "flex",
                fontSize: 13,
                justifyContent: "center",
                minHeight: 320
              }}
            >
              Chartsheet metadata is present, but no embedded chart payload was exposed for this tab.
            </div>
          )}
        </div>
      </div>
    );
  }

  if (!activeSheet) {
    return <>{renderEmpty(emptyState, palette)}</>;
  }

  const virtualRows = domExpandedWindow.rowIndices.map((index) => {
    const virtualRow = rowVirtualItemByIndex.get(index);
    return {
      end: virtualRow?.end ?? (rowPrefixSums[index + 1] ?? 0),
      index,
      key: visibleRows[index] ?? index,
      size: virtualRow?.size ?? (displayEffectiveRowHeights[index] ?? displayDefaultRowHeight),
      start: virtualRow?.start ?? (rowPrefixSums[index] ?? 0)
    };
  });
  const totalHeight = shouldUseDomVirtualizer && shouldVirtualizeRows
    ? rowVirtualizer.getTotalSize()
    : (rowPrefixSums[rowPrefixSums.length - 1] ?? 0);
  const totalWidth = totalContentWidth + displayRowHeaderWidth;
  const sheetContentHeight = displayHeaderHeight + totalHeight;
  const isLiveZooming = liveGestureZoom !== null && zoomScale === liveGestureZoom.baseZoomScale;
  const liveZoomScale = isLiveZooming
    ? Math.max(0.1, liveGestureZoom.targetZoomScale / liveGestureZoom.baseZoomScale)
    : 1;
  const liveZoomScrollLeft = scrollRef.current?.scrollLeft ?? drawingViewport.left;
  const liveZoomScrollTop = scrollRef.current?.scrollTop ?? drawingViewport.top;
  const liveZoomTranslateX = isLiveZooming
    ? (liveZoomScrollLeft + liveGestureZoom.anchor.x) * (1 - liveZoomScale)
    : 0;
  const liveZoomTranslateY = isLiveZooming
    ? (liveZoomScrollTop + liveGestureZoom.anchor.y) * (1 - liveZoomScale)
    : 0;
  const headerLabelLiveScale = isLiveZooming ? liveZoomScale : 1;
  const selectionBorderWidth = Math.max(1, zoomFactor);
  const shouldAnimateCanvasSelection = experimentalCanvas && enableCanvasSelectionAnimation;
  const canvasSelectionTransition = shouldAnimateCanvasSelection
    ? "left 120ms cubic-bezier(0.22, 1, 0.36, 1), top 120ms cubic-bezier(0.22, 1, 0.36, 1), width 120ms cubic-bezier(0.22, 1, 0.36, 1), height 120ms cubic-bezier(0.22, 1, 0.36, 1), opacity 100ms linear"
    : "none";
  const rowColSpan = renderedCols.length + 1 + (leadingColumnSpacerWidth > 0 ? 1 : 0) + (trailingColumnSpacerWidth > 0 ? 1 : 0);
  const gutterSeparatorShadow = `inset -1px 0 0 ${palette.border}, inset 0 -1px 0 ${palette.border}`;
  const maxDrawingOverlayZIndex = Math.max(
    0,
    ...shapes.map((shape) => shape.zIndex + 20),
    ...formControls.map((control) => control.zIndex + 20),
    ...images.map((image) => image.zIndex + 22),
    ...charts.map((chart) => chart.zIndex + 22)
  );
  const rowHeaderOverlayZIndex = Math.max(35, maxDrawingOverlayZIndex + 1);
  const canvasHeaderOverlayZIndex = Math.max(50, rowHeaderOverlayZIndex + 1);
  const stickyHeaderOverlayZIndex = canvasHeaderOverlayZIndex + 1;
  const cornerHeaderOverlayZIndex = canvasHeaderOverlayZIndex + 2;
  const frozenRowHeaderOverlayZIndex = rowHeaderOverlayZIndex;
  const headerCellStyle = scaleCssProperties({
    backgroundColor: palette.headerSurface,
    borderBottom: "none",
    borderRight: "none",
    boxShadow: gutterSeparatorShadow,
    color: palette.mutedText,
    fontSize: "11px",
    fontWeight: 600,
    height: HEADER_HEIGHT,
    overflow: "hidden",
    padding: "2px 4px",
    position: "sticky",
    textAlign: "center",
    top: 0,
    userSelect: "none",
    whiteSpace: "nowrap",
    zIndex: canvasHeaderOverlayZIndex
  }, zoomFactor);
  const columnResizeHandleStyle = scaleCssProperties({
    backgroundColor: "transparent",
    cursor: "col-resize",
    position: "absolute",
    right: -8,
    top: 0,
    width: 16,
    height: "100%",
    zIndex: 5
  }, zoomFactor);
  const canvasBodyViewportLayerStyle: React.CSSProperties = {
    height: 0,
    left: 0,
    overflow: "visible",
    pointerEvents: "none",
    position: "sticky",
    top: 0,
    width: 0
  };
  const canvasHeaderViewportLayerStyle: React.CSSProperties = {
    height: 0,
    left: 0,
    overflow: "visible",
    pointerEvents: "none",
    position: "sticky",
    top: 0,
    width: 0,
    zIndex: canvasHeaderOverlayZIndex
  };
  const canvasBodyBaseStyle: React.CSSProperties = {
    cursor: "cell",
    pointerEvents: "auto",
    position: "absolute",
    transformOrigin: "0 0",
    transition: "none",
    zIndex: 10
  };
  const canvasScrollBodyStyle: React.CSSProperties = {
    ...canvasBodyBaseStyle,
    display: scrollBodyCanvasWidth > 0 && scrollBodyCanvasHeight > 0 ? "block" : "none",
    left: scrollBodyCanvasLeft,
    top: scrollBodyCanvasTop
  };
  const canvasTopBodyStyle: React.CSSProperties = {
    ...canvasBodyBaseStyle,
    display: topBodyCanvasWidth > 0 && topBodyCanvasHeight > 0 ? "block" : "none",
    left: scrollBodyCanvasLeft,
    top: displayHeaderHeight,
    zIndex: 30
  };
  const canvasLeftBodyStyle: React.CSSProperties = {
    ...canvasBodyBaseStyle,
    display: leftBodyCanvasWidth > 0 && leftBodyCanvasHeight > 0 ? "block" : "none",
    left: displayRowHeaderWidth,
    top: scrollBodyCanvasTop,
    zIndex: 30
  };
  const canvasCornerBodyStyle: React.CSSProperties = {
    ...canvasBodyBaseStyle,
    display: cornerBodyCanvasWidth > 0 && cornerBodyCanvasHeight > 0 ? "block" : "none",
    left: displayRowHeaderWidth,
    top: displayHeaderHeight,
    zIndex: 31
  };
  const canvasHeaderBaseStyle: React.CSSProperties = {
    cursor: "default",
    pointerEvents: "auto",
    position: "absolute",
    transformOrigin: "0 0"
  };
  const canvasTopFrozenHeaderStyle: React.CSSProperties = {
    ...canvasHeaderBaseStyle,
    display: topFrozenHeaderCanvasWidth > 0 && drawingViewport.height > 0 ? "block" : "none",
    left: displayRowHeaderWidth,
    top: 0,
    zIndex: stickyHeaderOverlayZIndex
  };
  const canvasTopScrollHeaderStyle: React.CSSProperties = {
    ...canvasHeaderBaseStyle,
    display: topScrollHeaderCanvasWidth > 0 && drawingViewport.height > 0 ? "block" : "none",
    left: scrollBodyCanvasLeft,
    top: 0,
    zIndex: canvasHeaderOverlayZIndex
  };
  const canvasLeftFrozenHeaderStyle: React.CSSProperties = {
    ...canvasHeaderBaseStyle,
    display: leftFrozenHeaderCanvasHeight > 0 && drawingViewport.width > 0 ? "block" : "none",
    left: 0,
    top: displayHeaderHeight,
    zIndex: stickyHeaderOverlayZIndex
  };
  const canvasLeftScrollHeaderStyle: React.CSSProperties = {
    ...canvasHeaderBaseStyle,
    display: leftScrollHeaderCanvasHeight > 0 && drawingViewport.width > 0 ? "block" : "none",
    left: 0,
    top: scrollBodyCanvasTop,
    zIndex: canvasHeaderOverlayZIndex
  };
  const canvasCornerHeaderStyle: React.CSSProperties = {
    display: drawingViewport.width > 0 && drawingViewport.height > 0 ? "block" : "none",
    left: 0,
    pointerEvents: "none",
    position: "absolute",
    top: 0,
    transformOrigin: "0 0",
    zIndex: cornerHeaderOverlayZIndex
  };
  const editingOverlayRect = experimentalCanvas && editingCell
    ? resolveCellDisplayRect(editingCell)
    : null;
  const activeCellAdornment = experimentalCanvas && activeCell
    ? renderCellAdornment(activeCell)
    : null;
  const activeCellAdornmentRect = activeCell && activeCellAdornment
    ? resolveCellDisplayRect(activeCell)
    : null;
  function resolveDrawingPane(rect: XlsxImageRect) {
    return resolveFrozenDrawingPane(
      rect,
      frozenRows,
      frozenCols,
      displayActualRowHeights,
      displayActualColWidths,
      activeSheet?.freezePanes ?? null,
      stickyTopByRow,
      stickyLeftByCol,
      {
        defaultColWidth: displayDefaultColWidth,
        defaultRowHeight: displayDefaultRowHeight,
        headerHeight: displayHeaderHeight,
        rowHeaderWidth: displayRowHeaderWidth
      }
    );
  }

  function renderShapeDrawing(shape: XlsxShape, rect: XlsxImageRect, pane: FrozenDrawingPane) {
    const drawingPane = resolveDrawingPane(rect);
    if (drawingPane !== pane) {
      return null;
    }
    const hasMeasuredViewport = drawingViewport.width > 0 && drawingViewport.height > 0;
    if (pane === "scroll" && hasMeasuredViewport && !rectIntersectsViewport(rect, drawingViewport)) {
      return null;
    }

    const isFrozenDrawing = pane !== "scroll";
    const inset = shape.textBox?.insetPx;
    const groupScaleX = shape.scaleX ?? 1;
    const groupScaleY = shape.scaleY ?? 1;
    const strokeScale = Math.max(groupScaleX, groupScaleY);
    const textScale = strokeScale * zoomFactor;
    const textWidth = groupScaleX !== 0 ? rect.width / groupScaleX : rect.width;
    const textHeight = groupScaleY !== 0 ? rect.height / groupScaleY : rect.height;
    const vectorShape = resolveShapeVector(shape);
    const strokeColor = shape.stroke?.none ? "transparent" : (shape.stroke?.color ?? "transparent");
    const scaledStrokeWidth = (shape.stroke?.widthPx ?? (shape.geometry === "line" ? 2 : 1)) * strokeScale * zoomFactor;
    const headMarkerId = `${shape.id}-${pane}-head-marker`;
    const tailMarkerId = `${shape.id}-${pane}-tail-marker`;
    const headMarker = vectorShape
      ? resolveShapeLineEndMarker(
          shape.stroke?.headEndType,
          headMarkerId,
          strokeColor,
          scaledStrokeWidth,
          rect,
          vectorShape.viewBox
        )
      : null;
    const tailMarker = vectorShape
      ? resolveShapeLineEndMarker(
          shape.stroke?.tailEndType,
          tailMarkerId,
          strokeColor,
          scaledStrokeWidth,
          rect,
          vectorShape.viewBox
        )
      : null;
    const style = {
      ...buildShapeContainerStyle(shape, rect, zoomFactor),
      ...(vectorShape ? {
        backgroundColor: "transparent",
        border: "none"
      } : null)
    };

    return (
      <div
        key={`${pane}-${shape.id}`}
        onClick={() => handleShapeClick(shape)}
        style={{
          ...style,
          contain: "layout paint",
          cursor: shape.hyperlink ? "pointer" : "default",
          left: style.left,
          pointerEvents: shape.hyperlink ? "auto" : "none",
          top: style.top,
          zIndex: isFrozenDrawing ? shape.zIndex + 20 : style.zIndex
        }}
        title={shape.description}
      >
        {vectorShape ? (
          <svg
            aria-hidden="true"
            preserveAspectRatio="none"
            style={{
              height: "100%",
              inset: 0,
              position: "absolute",
              width: "100%"
            }}
            viewBox={`0 0 ${vectorShape.viewBox.width} ${vectorShape.viewBox.height}`}
          >
            {headMarker || tailMarker ? <defs>{headMarker}{tailMarker}</defs> : null}
            <path
              d={vectorShape.path}
              fill={shape.fill?.none ? "transparent" : (shape.fill?.color ?? "transparent")}
              fillOpacity={shape.fill?.opacity ?? 1}
              markerEnd={tailMarker ? `url(#${tailMarkerId})` : undefined}
              markerStart={headMarker ? `url(#${headMarkerId})` : undefined}
              stroke={strokeColor}
              strokeOpacity={shape.stroke?.opacity ?? 1}
              strokeWidth={scaledStrokeWidth}
              vectorEffect="non-scaling-stroke"
            />
          </svg>
        ) : null}
        <div
          style={{
            color: "#000000",
            display: "flex",
            flex: 1,
            flexDirection: "column",
            gap: 2 * zoomFactor,
            height: textHeight,
            justifyContent:
              shape.textBox?.verticalAlign === "middle"
                ? "center"
                : shape.textBox?.verticalAlign === "bottom"
                  ? "flex-end"
                  : "flex-start",
            paddingBottom: (inset?.bottom ?? 4) * zoomFactor,
            paddingLeft: (inset?.left ?? 6) * zoomFactor,
            paddingRight: (inset?.right ?? 6) * zoomFactor,
            paddingTop: (inset?.top ?? 4) * zoomFactor,
            pointerEvents: "none",
            position: "relative",
            transform:
              groupScaleX !== 1 || groupScaleY !== 1
                ? `scale(${groupScaleX}, ${groupScaleY})`
                : undefined,
            transformOrigin: "top left",
            width: textWidth,
            zIndex: 1
          }}
        >
          {shape.paragraphs.map((paragraph, index) => renderShapeParagraph(
            paragraph,
            index,
            shape.textBox?.horizontalAlign ?? "left",
            textScale
          ))}
        </div>
      </div>
    );
  }

  function renderFormControlDrawing(control: XlsxFormControl, rect: XlsxImageRect, pane: FrozenDrawingPane) {
    const drawingPane = resolveDrawingPane(rect);
    if (drawingPane !== pane) {
      return null;
    }

    const hasMeasuredViewport = drawingViewport.width > 0 && drawingViewport.height > 0;
    if (pane === "scroll" && hasMeasuredViewport && !rectIntersectsViewport(rect, drawingViewport)) {
      return null;
    }

    const isFrozenDrawing = pane !== "scroll";
    const controlLabel = resolveFormControlLabel(control);
    const fontSizePx = Math.max(9 * zoomFactor, ((control.fontSizePt ?? 9) * 96 / 72) * zoomFactor);
    const stroke = paletteIsDark(palette) ? "#cbd5e1" : "#475569";
    const textColor = control.textColor ?? "#000000";
    const commonStyle: React.CSSProperties = {
      alignItems: control.kind === "group-box" ? "stretch" : "center",
      color: textColor,
      contain: "layout paint",
      display: "flex",
      fontFamily: control.fontFamily ?? "Calibri, sans-serif",
      fontSize: fontSizePx,
      fontWeight: 400,
      gap: 4 * zoomFactor,
      height: rect.height,
      justifyContent:
        control.textAlign === "center"
          ? "center"
          : control.textAlign === "right"
            ? "flex-end"
            : "flex-start",
      left: rect.left,
      lineHeight: 1.2,
      overflow: "hidden",
      pointerEvents: "none",
      position: "absolute",
      top: rect.top,
      width: rect.width,
      zIndex: isFrozenDrawing ? control.zIndex + 20 : control.zIndex
    };

    if (control.kind === "group-box") {
      const hasLabel = controlLabel.length > 0;
      return (
        <div
          key={`${pane}-${control.id}`}
          style={{
            ...commonStyle,
            padding: `${hasLabel ? Math.max(7, fontSizePx * 0.5) : 0}px 0 0`,
            position: "absolute"
          }}
        >
          <div
            style={{
              border: `${Math.max(1, zoomFactor)}px solid ${stroke}`,
              borderRadius: 2 * zoomFactor,
              height: "100%",
              position: "relative",
              width: "100%"
            }}
          >
            {controlLabel ? (
              <span
                style={{
                  backgroundColor: SHEET_SURFACE,
                  left: 8 * zoomFactor,
                  maxWidth: `calc(100% - ${16 * zoomFactor}px)`,
                  padding: `0 ${4 * zoomFactor}px`,
                  position: "absolute",
                  top: -(fontSizePx * 0.6),
                  whiteSpace: "nowrap"
                }}
              >
                {controlLabel}
              </span>
            ) : null}
          </div>
        </div>
      );
    }

    let content: React.ReactNode;
    if (control.kind === "radio") {
      content = renderRadioControl(Boolean(control.checked), palette, zoomFactor);
    } else if (control.kind === "checkbox") {
      content = renderCheckboxControl(Boolean(control.checked), palette, zoomFactor);
    } else if (control.kind === "button") {
      return (
        <div
          key={`${pane}-${control.id}`}
          style={{
            ...commonStyle,
            background: "linear-gradient(180deg, #f8fafc 0%, #e2e8f0 100%)",
            border: `${Math.max(1, zoomFactor)}px solid ${stroke}`,
            borderRadius: 4 * zoomFactor,
            boxSizing: "border-box",
            justifyContent: "center",
            padding: `0 ${6 * zoomFactor}px`,
            textAlign: "center"
          }}
        >
          {controlLabel}
        </div>
      );
    } else if (control.kind === "dropdown") {
      return (
        <div
          key={`${pane}-${control.id}`}
          style={{
            ...commonStyle,
            border: `${Math.max(1, zoomFactor)}px solid ${stroke}`,
            borderRadius: 2 * zoomFactor,
            boxSizing: "border-box",
            justifyContent: "space-between",
            padding: `0 ${6 * zoomFactor}px`
          }}
        >
          <span style={{ overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{controlLabel}</span>
          <span aria-hidden="true" style={{ fontSize: fontSizePx * 0.85 }}>▼</span>
        </div>
      );
    } else if (control.kind === "editbox" || control.kind === "listbox" || control.kind === "scrollbar" || control.kind === "spinner" || control.kind === "unknown") {
      return (
        <div
          key={`${pane}-${control.id}`}
          style={{
            ...commonStyle,
            border: `${Math.max(1, zoomFactor)}px solid ${stroke}`,
            borderRadius: 2 * zoomFactor,
            boxSizing: "border-box",
            padding: `0 ${6 * zoomFactor}px`
          }}
        >
          <span style={{ overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{controlLabel}</span>
        </div>
      );
    } else {
      content = null;
    }

    return (
      <div
        key={`${pane}-${control.id}`}
        style={{
          ...commonStyle,
          padding: `0 ${Math.max(1, zoomFactor)}px`
        }}
      >
        {content}
        {controlLabel ? (
          <span
            style={{
              overflow: "hidden",
              textOverflow: "ellipsis",
              whiteSpace: "nowrap"
            }}
          >
            {controlLabel}
          </span>
        ) : null}
      </div>
    );
  }

  function renderImageDrawing(
    image: XlsxImage,
    rect: XlsxImageRect,
    pane: FrozenDrawingPane
  ) {
    const drawingPane = resolveDrawingPane(rect);
    if (drawingPane !== pane) {
      return null;
    }
    const hasMeasuredViewport = drawingViewport.width > 0 && drawingViewport.height > 0;
    if (pane === "scroll" && hasMeasuredViewport && !rectIntersectsViewport(rect, drawingViewport)) {
      return null;
    }

    const isFrozenDrawing = pane !== "scroll";
    const canEditImage = !readOnly && image.editable !== false;
    const drawingSelectionSurface = paletteIsDark(palette) ? palette.canvas : DRAWING_SELECTION_HANDLE_FILL;
    const style: React.CSSProperties = {
      contain: "layout paint",
      height: rect.height,
      left: rect.left,
      overflow: "hidden",
      pointerEvents: "none",
      position: "absolute",
      top: rect.top,
      width: rect.width,
      zIndex: isFrozenDrawing ? image.zIndex + 20 : image.zIndex
    };
    const defaultNode = (
      <img
        alt={image.description ?? image.name ?? ""}
        draggable={false}
        src={image.src}
        style={{
          display: "block",
          height: "100%",
          pointerEvents: "none",
          userSelect: "none",
          width: "100%"
        }}
      />
    );
    const selectionNode = selectedImageId === image.id ? (
      <div
        style={{
          ...style,
          contain: "layout",
          overflow: "visible",
          pointerEvents: "none",
          zIndex: isFrozenDrawing ? image.zIndex + 22 : image.zIndex + 2
        }}
      >
        {renderImageSelection
          ? renderImageSelection({
              defaultNode: (
                <div
                  style={{
                    border: `${Math.max(1, zoomFactor)}px solid ${DRAWING_SELECTION_STROKE}`,
                    boxShadow: `0 0 0 ${Math.max(1, zoomFactor)}px ${drawingSelectionSurface}`,
                    boxSizing: "border-box",
                    inset: 0,
                    pointerEvents: "none",
                    position: "absolute"
                  }}
                >
                  {canEditImage
                    ? IMAGE_HANDLE_POSITIONS.map((position) => (
                        <div
                          key={position}
                          onPointerDown={(event) => startImageResize(event, image, rect, position)}
                          style={resolveImageHandleStyle(position, DRAWING_SELECTION_STROKE, drawingSelectionSurface, zoomFactor)}
                        />
                      ))
                    : null}
                </div>
              ),
              getHandleProps: (position) => ({
                onPointerDown: (event) => {
                  if (canEditImage) {
                    startImageResize(event, image, rect, position);
                  }
                },
                style: canEditImage
                  ? resolveImageHandleStyle(position, DRAWING_SELECTION_STROKE, drawingSelectionSurface, zoomFactor)
                  : { ...resolveImageHandleStyle(position, DRAWING_SELECTION_STROKE, drawingSelectionSurface, zoomFactor), display: "none" }
              }),
              image,
              rect
            })
          : (
              <div
                style={{
                  border: `${Math.max(1, zoomFactor)}px solid ${DRAWING_SELECTION_STROKE}`,
                  boxShadow: `0 0 0 ${Math.max(1, zoomFactor)}px ${drawingSelectionSurface}`,
                  boxSizing: "border-box",
                  inset: 0,
                  pointerEvents: "none",
                  position: "absolute"
                }}
              >
                {canEditImage
                  ? IMAGE_HANDLE_POSITIONS.map((position) => (
                      <div
                        key={position}
                        onPointerDown={(event) => startImageResize(event, image, rect, position)}
                        style={resolveImageHandleStyle(position, DRAWING_SELECTION_STROKE, drawingSelectionSurface, zoomFactor)}
                      />
                    ))
                  : null}
              </div>
            )}
      </div>
    ) : null;

    return (
      <React.Fragment key={`${pane}-${image.id}`}>
        {renderImage
          ? <div style={style}>{renderImage({ defaultNode, image, rect, style })}</div>
          : <div style={style}>{defaultNode}</div>}
        <div
          onClick={() => handleImageClick(image)}
          onPointerDown={(event) => startImageMove(event, image, rect)}
          style={{
            ...style,
            background: "transparent",
            cursor: canEditImage && selectedImageId === image.id ? "move" : image.hyperlink ? "pointer" : "cell",
            pointerEvents: "auto",
            zIndex: isFrozenDrawing ? image.zIndex + 21 : image.zIndex + 1
          }}
        />
        {selectionNode}
      </React.Fragment>
    );
  }

  function renderChartDrawing(
    chart: XlsxChart,
    rect: XlsxImageRect,
    pane: FrozenDrawingPane
  ) {
    const drawingPane = resolveDrawingPane(rect);
    if (drawingPane !== pane) {
      return null;
    }
    const hasMeasuredViewport = drawingViewport.width > 0 && drawingViewport.height > 0;
    if (pane === "scroll" && hasMeasuredViewport && !rectIntersectsViewport(rect, drawingViewport)) {
      return null;
    }

    const isFrozenDrawing = pane !== "scroll";
    const canEditChart = !readOnly && chart.editable !== false;
    const drawingSelectionSurface = paletteIsDark(palette) ? palette.canvas : DRAWING_SELECTION_HANDLE_FILL;
    const style: React.CSSProperties = {
      contain: "layout paint",
      height: rect.height,
      left: rect.left,
      overflow: "hidden",
      pointerEvents: "none",
      position: "absolute",
      top: rect.top,
      width: rect.width,
      zIndex: isFrozenDrawing ? chart.zIndex + 20 : chart.zIndex
    };
    const selectionNode = selectedChartId === chart.id ? (
      <div
        style={{
          ...style,
          contain: "layout",
          overflow: "visible",
          pointerEvents: "none",
          zIndex: isFrozenDrawing ? chart.zIndex + 22 : chart.zIndex + 2
        }}
      >
        <div
          style={{
            border: `${Math.max(1, zoomFactor)}px solid ${DRAWING_SELECTION_STROKE}`,
            boxShadow: `0 0 0 ${Math.max(1, zoomFactor)}px ${drawingSelectionSurface}`,
            boxSizing: "border-box",
            inset: 0,
            pointerEvents: "none",
            position: "absolute"
          }}
        >
          {canEditChart
            ? IMAGE_HANDLE_POSITIONS.map((position) => (
                <div
                  key={position}
                  onPointerDown={(event) => startChartResize(event, chart, rect, position)}
                  style={resolveImageHandleStyle(position, DRAWING_SELECTION_STROKE, drawingSelectionSurface, zoomFactor)}
                />
              ))
            : null}
        </div>
      </div>
    ) : null;

      return (
        <React.Fragment key={`${pane}-${chart.id}`}>
          <div
            onClick={() => handleChartClick(chart)}
            onPointerDown={(event) => startChartMove(event, chart, rect)}
            style={{
              ...style,
              cursor: canEditChart && selectedChartId === chart.id ? "move" : "cell",
              pointerEvents: "auto"
            }}
          >
            {isChartsLoading
              ? renderChartLoadingNode(renderChartLoading, chart, rect)
              : (
                  <MemoChartSvg
                    chart={chart}
                    onChartElementDoubleClick={handleChartElementDoubleClick}
                    onChartElementPointerDown={handleChartElementPointerDown}
                    palette={palette}
                    rect={rect}
                    selectedChartElement={selectedChartElement}
                  />
                )}
          </div>
          {selectionNode}
        </React.Fragment>
      );
  }

  const scrollOverlayStyle: React.CSSProperties = {
    inset: 0,
    pointerEvents: "none",
    position: "absolute",
    zIndex: 20
  };
  const topOverlayStyle: React.CSSProperties = {
    height: 0,
    overflow: "visible",
    pointerEvents: "none",
    position: "sticky",
    top: 0,
    width: 0,
    zIndex: 25
  };
  const leftOverlayStyle: React.CSSProperties = {
    height: 0,
    left: 0,
    overflow: "visible",
    pointerEvents: "none",
    position: "sticky",
    width: 0,
    zIndex: 25
  };
  const cornerOverlayStyle: React.CSSProperties = {
    height: 0,
    left: 0,
    overflow: "visible",
    pointerEvents: "none",
    position: "sticky",
    top: 0,
    width: 0,
    zIndex: 26
  };
  const canvasOverlayPaneBaseStyle: React.CSSProperties = {
    overflow: "hidden",
    pointerEvents: "none",
    position: "absolute"
  };
  const canvasScrollOverlayPaneStyle: React.CSSProperties = {
    ...canvasOverlayPaneBaseStyle,
    display: scrollBodyViewportWidth > 0 && scrollBodyViewportHeight > 0 ? "block" : "none",
    height: scrollBodyViewportHeight,
    left: frozenPaneRight,
    top: frozenPaneBottom,
    width: scrollBodyViewportWidth,
    zIndex: 20
  };
  const canvasTopOverlayPaneStyle: React.CSSProperties = {
    ...canvasOverlayPaneBaseStyle,
    display: scrollBodyViewportWidth > 0 && topBodyCanvasHeight > 0 ? "block" : "none",
    height: topBodyCanvasHeight,
    left: frozenPaneRight,
    top: displayHeaderHeight,
    width: scrollBodyViewportWidth,
    zIndex: 35
  };
  const canvasLeftOverlayPaneStyle: React.CSSProperties = {
    ...canvasOverlayPaneBaseStyle,
    display: leftBodyCanvasWidth > 0 && scrollBodyViewportHeight > 0 ? "block" : "none",
    height: scrollBodyViewportHeight,
    left: displayRowHeaderWidth,
    top: frozenPaneBottom,
    width: leftBodyCanvasWidth,
    zIndex: 35
  };
  const canvasCornerOverlayPaneStyle: React.CSSProperties = {
    ...canvasOverlayPaneBaseStyle,
    display: cornerBodyCanvasWidth > 0 && cornerBodyCanvasHeight > 0 ? "block" : "none",
    height: cornerBodyCanvasHeight,
    left: displayRowHeaderWidth,
    top: displayHeaderHeight,
    width: cornerBodyCanvasWidth,
    zIndex: 36
  };
  const drawingViewportCacheSignature = `${Math.floor(drawingViewport.left / CANVAS_VIEWPORT_OVERSCAN_PX)}:${Math.floor(drawingViewport.top / CANVAS_VIEWPORT_OVERSCAN_PX)}:${drawingViewport.width}:${drawingViewport.height}`;
  const previousPaneDrawingNodes = paneDrawingNodesCacheRef.current;
  const canReusePaneDrawingNodes =
    previousPaneDrawingNodes !== null
    && previousPaneDrawingNodes.showImages === showImages
    && previousPaneDrawingNodes.chartRects === chartRects
    && previousPaneDrawingNodes.formControlRects === domFormControlRects
    && previousPaneDrawingNodes.shapeRects === domShapeRects
    && previousPaneDrawingNodes.imageRects === domImageRects
    && previousPaneDrawingNodes.selectedChartId === selectedChartId
    && previousPaneDrawingNodes.selectedImageId === selectedImageId
    && previousPaneDrawingNodes.readOnly === readOnly
    && previousPaneDrawingNodes.selectionStroke === selectionStroke
    && previousPaneDrawingNodes.renderChartLoading === renderChartLoading
    && previousPaneDrawingNodes.renderImage === renderImage
    && previousPaneDrawingNodes.renderImageSelection === renderImageSelection
      && previousPaneDrawingNodes.isChartsLoading === isChartsLoading
      && previousPaneDrawingNodes.selectedChartElement === selectedChartElement
      && previousPaneDrawingNodes.palette === palette
    && previousPaneDrawingNodes.drawingViewportSignature === drawingViewportCacheSignature;
  const paneDrawingNodes = canReusePaneDrawingNodes
    ? previousPaneDrawingNodes.value
    : (!showImages
        ? {
            corner: null,
            left: null,
            scroll: null,
            top: null
          }
        : {
            corner: (
              <>
                {chartRects.map(({ chart, rect }) => renderChartDrawing(chart, rect, "corner"))}
                {domShapeRects.map(({ shape, rect }) => renderShapeDrawing(shape, rect, "corner"))}
                {domFormControlRects.map(({ control, rect }) => renderFormControlDrawing(control, rect, "corner"))}
                {domImageRects.map(({ image, rect }) => renderImageDrawing(image, rect, "corner"))}
              </>
            ),
            left: (
              <>
                {chartRects.map(({ chart, rect }) => renderChartDrawing(chart, rect, "left"))}
                {domShapeRects.map(({ shape, rect }) => renderShapeDrawing(shape, rect, "left"))}
                {domFormControlRects.map(({ control, rect }) => renderFormControlDrawing(control, rect, "left"))}
                {domImageRects.map(({ image, rect }) => renderImageDrawing(image, rect, "left"))}
              </>
            ),
            scroll: (
              <>
                {chartRects.map(({ chart, rect }) => renderChartDrawing(chart, rect, "scroll"))}
                {domShapeRects.map(({ shape, rect }) => renderShapeDrawing(shape, rect, "scroll"))}
                {domFormControlRects.map(({ control, rect }) => renderFormControlDrawing(control, rect, "scroll"))}
                {domImageRects.map(({ image, rect }) => renderImageDrawing(image, rect, "scroll"))}
              </>
            ),
            top: (
              <>
                {chartRects.map(({ chart, rect }) => renderChartDrawing(chart, rect, "top"))}
                {domShapeRects.map(({ shape, rect }) => renderShapeDrawing(shape, rect, "top"))}
                {domFormControlRects.map(({ control, rect }) => renderFormControlDrawing(control, rect, "top"))}
                {domImageRects.map(({ image, rect }) => renderImageDrawing(image, rect, "top"))}
              </>
            )
          }) satisfies Record<FrozenDrawingPane, React.ReactNode>;

  if (!canReusePaneDrawingNodes) {
    paneDrawingNodesCacheRef.current = {
      chartRects,
      drawingViewportSignature: drawingViewportCacheSignature,
      formControlRects: domFormControlRects,
      imageRects: domImageRects,
        isChartsLoading,
        palette,
        readOnly,
        renderChartLoading,
        renderImage,
        renderImageSelection,
        selectedChartElement,
        selectedChartId,
      selectedImageId,
      selectionStroke,
      shapeRects: domShapeRects,
      showImages,
      value: paneDrawingNodes
    };
  }

  function startColumnResize(pointerId: number, actualCol: number, widthPx: number, startX: number) {
    if (!canResizeHeaders) {
      return;
    }

    resizeStateRef.current = {
      actualIndex: actualCol,
      initialPx: widthPx,
      pointerId,
      startPosition: startX,
      type: "column"
    };
    columnPreviewRef.current = { actualIndex: actualCol, size: widthPx };
    rowPreviewRef.current = null;
    const position = resolveResizeGuidePositionFromClient("column", startX);
    setResizeGuide(position === null ? null : { position, type: "column" });
    setInteractionMode("select");
    setGlobalCursor("col-resize");
    document.body.style.userSelect = "none";
  }

  function startRowResize(pointerId: number, actualRow: number, heightPx: number, startY: number) {
    if (!canResizeHeaders) {
      return;
    }

    resizeStateRef.current = {
      actualIndex: actualRow,
      initialPx: heightPx,
      pointerId,
      startPosition: startY,
      type: "row"
    };
    rowPreviewRef.current = { actualIndex: actualRow, size: heightPx };
    columnPreviewRef.current = null;
    const position = resolveResizeGuidePositionFromClient("row", startY);
    setResizeGuide(position === null ? null : { position, type: "row" });
    setInteractionMode("select");
    setGlobalCursor("row-resize");
    document.body.style.userSelect = "none";
  }

  function startCellSelection(
    target: Element,
    pointerId: number,
    anchor: XlsxCellAddress,
    axis: "cell" | "column" | "row",
    originCell: XlsxCellAddress,
    pointerOrigin: {
      contentScaleX: number;
      contentScaleY: number;
      originContentX: number;
      originContentY: number;
    },
    originOverlayRect: { height: number; left: number; top: number; width: number } | null,
    committedOnPointerDown: boolean,
    initialRange: XlsxCellRange,
    startClientX: number,
    startClientY: number
  ) {
    clearPendingSelectionDrag();
    const previewRange = normalizeRange(initialRange);
    pendingSelectionDragRef.current = {
      anchor,
      axis,
      contentScaleX: pointerOrigin.contentScaleX,
      contentScaleY: pointerOrigin.contentScaleY,
      committedOnPointerDown,
      didDrag: false,
      originCell,
      originOverlayRect,
      originContentX: pointerOrigin.originContentX,
      originContentY: pointerOrigin.originContentY,
      pointerId,
      previewRange,
      startClientX,
      startClientY
    };
    selectionPreviewRangeRef.current = previewRange;
    displayedSelectionRef.current = previewRange;
    applyPreviewOverlay(previewRange);
    installPendingSelectionDragListeners(pointerId, target);
  }

  function resolveFillRange(sourceRange: XlsxCellRange, cell: XlsxCellAddress): XlsxCellRange {
    const normalizedSource = normalizeRange(sourceRange);
    if (isCellInRange(cell, normalizedSource)) {
      return normalizedSource;
    }

    const distanceAbove = Math.max(0, normalizedSource.start.row - cell.row);
    const distanceBelow = Math.max(0, cell.row - normalizedSource.end.row);
    const distanceLeft = Math.max(0, normalizedSource.start.col - cell.col);
    const distanceRight = Math.max(0, cell.col - normalizedSource.end.col);
    const verticalDistance = Math.max(distanceAbove, distanceBelow);
    const horizontalDistance = Math.max(distanceLeft, distanceRight);

    if (verticalDistance >= horizontalDistance) {
      return normalizeRange({
        start: {
          row: distanceAbove > 0 ? cell.row : normalizedSource.start.row,
          col: normalizedSource.start.col
        },
        end: {
          row: distanceBelow > 0 ? cell.row : normalizedSource.end.row,
          col: normalizedSource.end.col
        }
      });
    }

    return normalizeRange({
      start: {
        row: normalizedSource.start.row,
        col: distanceLeft > 0 ? cell.col : normalizedSource.start.col
      },
      end: {
        row: normalizedSource.end.row,
        col: distanceRight > 0 ? cell.col : normalizedSource.end.col
      }
    });
  }

  function startFillDrag(pointerId: number, sourceRange: XlsxCellRange) {
    if (readOnly) {
      return;
    }

    cachedScrollerRectRef.current = scrollRef.current?.getBoundingClientRect() ?? null;
    const normalizedSource = normalizeRange(sourceRange);
    fillDragRef.current = {
      pointerId,
      previewRange: normalizedSource,
      sourceRange: normalizedSource
    };
    setFillPreviewRange(normalizedSource);
    setInteractionMode("fill");
    document.body.style.cursor = "crosshair";
    document.body.style.userSelect = "none";
    installFillDragListeners(pointerId, normalizedSource);
  }

  function updateFillPreview(cell: XlsxCellAddress) {
    const fillState = fillDragRef.current;
    if (!fillState) {
      return;
    }

    const nextRange = resolveFillRange(fillState.sourceRange, cell);
    fillState.previewRange = nextRange;
    applyPreviewOverlay(nextRange);
  }

  function installImageInteractionListeners(pointerId: number) {
    imageInteractionCleanupRef.current?.();

    const resolveInteractionRect = (
      interaction: NonNullable<typeof imageInteractionRef.current>,
      clientX: number,
      clientY: number
    ) => {
      const deltaX = clientX - interaction.startClientX;
      const deltaY = clientY - interaction.startClientY;
      return clampImageRect(
        interaction.type === "move"
          ? {
              ...interaction.baseRect,
              left: interaction.baseRect.left + deltaX,
              top: interaction.baseRect.top + deltaY
            }
          : resizeImageRect(interaction.baseRect, interaction.handle, deltaX, deltaY, displayImageMinSize),
        {
          contentOffsetLeft: displayRowHeaderWidth,
          contentOffsetTop: displayHeaderHeight,
          minSizePx: displayImageMinSize
        }
      );
    };

    const handlePointerMove = (event: PointerEvent) => {
      if (event.pointerId !== pointerId) {
        return;
      }

      const interaction = imageInteractionRef.current;
      if (!interaction) {
        return;
      }

      const deltaX = event.clientX - interaction.startClientX;
      const deltaY = event.clientY - interaction.startClientY;
      if (!interaction.didMove && (Math.abs(deltaX) > 3 || Math.abs(deltaY) > 3)) {
        interaction.didMove = true;
      }
      const nextRect = resolveInteractionRect(interaction, event.clientX, event.clientY);
      scheduleImagePreviewRect({ id: interaction.imageId, rect: nextRect });
    };

    const cleanup = () => {
      window.removeEventListener("pointermove", handlePointerMove);
      window.removeEventListener("pointerup", handlePointerUp);
      window.removeEventListener("pointercancel", handlePointerUp);
    };

    const handlePointerUp = (event: PointerEvent) => {
      if (event.pointerId !== pointerId) {
        return;
      }

      const interaction = imageInteractionRef.current;
      if (imagePreviewFrameRef.current !== null) {
        window.cancelAnimationFrame(imagePreviewFrameRef.current);
        imagePreviewFrameRef.current = null;
      }
      const pendingPreview = pendingImagePreviewRef.current;
      pendingImagePreviewRef.current = null;
      if (pendingPreview) {
        imagePreviewRectRef.current = pendingPreview;
        setImagePreviewRect(pendingPreview);
      }
      const finalRect = interaction
        ? resolveInteractionRect(interaction, event.clientX, event.clientY)
        : null;
      imageInteractionRef.current = null;
      imageInteractionCleanupRef.current = null;
      setInteractionMode("idle");
      document.body.style.cursor = "";
      document.body.style.userSelect = "";
      cleanup();
      if (interaction) {
        if (interaction.didMove) {
          skipNextImageClickRef.current = interaction.imageId;
        }
        if (interaction.didMove && finalRect) {
          setImageRect(interaction.imageId, toLogicalRect(finalRect));
        }
      }
      imagePreviewRectRef.current = null;
      setImagePreviewRect(null);
    };

    window.addEventListener("pointermove", handlePointerMove);
    window.addEventListener("pointerup", handlePointerUp);
    window.addEventListener("pointercancel", handlePointerUp);
    imageInteractionCleanupRef.current = cleanup;
  }

  function installChartInteractionListeners(pointerId: number) {
    chartInteractionCleanupRef.current?.();

    const resolveInteractionRect = (
      interaction: NonNullable<typeof chartInteractionRef.current>,
      clientX: number,
      clientY: number
    ) => {
      const deltaX = clientX - interaction.startClientX;
      const deltaY = clientY - interaction.startClientY;
      return clampImageRect(
        interaction.type === "move"
          ? {
              ...interaction.baseRect,
              left: interaction.baseRect.left + deltaX,
              top: interaction.baseRect.top + deltaY
            }
          : resizeImageRect(interaction.baseRect, interaction.handle, deltaX, deltaY, 48 * zoomFactor),
        {
          contentOffsetLeft: displayRowHeaderWidth,
          contentOffsetTop: displayHeaderHeight,
          minSizePx: 48 * zoomFactor
        }
      );
    };

    const handlePointerMove = (event: PointerEvent) => {
      if (event.pointerId !== pointerId) {
        return;
      }

      const interaction = chartInteractionRef.current;
      if (!interaction) {
        return;
      }

      const deltaX = event.clientX - interaction.startClientX;
      const deltaY = event.clientY - interaction.startClientY;
      if (!interaction.didMove && (Math.abs(deltaX) > 3 || Math.abs(deltaY) > 3)) {
        interaction.didMove = true;
      }
      const nextRect = resolveInteractionRect(interaction, event.clientX, event.clientY);
      scheduleChartPreviewRect({ id: interaction.chartId, rect: nextRect });
    };

    const cleanup = () => {
      window.removeEventListener("pointermove", handlePointerMove);
      window.removeEventListener("pointerup", handlePointerUp);
      window.removeEventListener("pointercancel", handlePointerUp);
    };

    const handlePointerUp = (event: PointerEvent) => {
      if (event.pointerId !== pointerId) {
        return;
      }

      const interaction = chartInteractionRef.current;
      if (chartPreviewFrameRef.current !== null) {
        window.cancelAnimationFrame(chartPreviewFrameRef.current);
        chartPreviewFrameRef.current = null;
      }
      const pendingPreview = pendingChartPreviewRef.current;
      pendingChartPreviewRef.current = null;
      if (pendingPreview) {
        chartPreviewRectRef.current = pendingPreview;
        setChartPreviewRect(pendingPreview);
      }
      const finalRect = interaction
        ? resolveInteractionRect(interaction, event.clientX, event.clientY)
        : null;
      chartInteractionRef.current = null;
      chartInteractionCleanupRef.current = null;
      setInteractionMode("idle");
      document.body.style.cursor = "";
      document.body.style.userSelect = "";
      cleanup();
      if (interaction) {
        if (interaction.didMove) {
          skipNextChartClickRef.current = interaction.chartId;
        }
        if (interaction.didMove && finalRect) {
          setChartRect(interaction.chartId, toLogicalRect(finalRect));
        }
      }
      chartPreviewRectRef.current = null;
      setChartPreviewRect(null);
    };

    window.addEventListener("pointermove", handlePointerMove);
    window.addEventListener("pointerup", handlePointerUp);
    window.addEventListener("pointercancel", handlePointerUp);
    chartInteractionCleanupRef.current = cleanup;
  }

  function resolveCurrentCell() {
    if (activeCell && rowIndexByActual.has(activeCell.row) && colIndexByActual.has(activeCell.col)) {
      return activeCell;
    }
    if (firstVisibleRow === undefined || firstVisibleCol === undefined) {
      return null;
    }
    return { row: firstVisibleRow, col: firstVisibleCol };
  }

  function ensureCellVisible(rowIndex: number, colIndex: number) {
    const scroller = scrollRef.current;
    if (!scroller) {
      return;
    }

    let nextScrollTop = scroller.scrollTop;
    let nextScrollLeft = scroller.scrollLeft;
    const rowStart = displayHeaderHeight + (rowPrefixSums[rowIndex] ?? 0);
    const rowEnd = displayHeaderHeight + (rowPrefixSums[rowIndex + 1] ?? rowStart);
    const colStart = displayRowHeaderWidth + (colPrefixSums[colIndex] ?? 0);
    const colEnd = displayRowHeaderWidth + (colPrefixSums[colIndex + 1] ?? colStart);

    if (rowEnd > frozenPaneBottom) {
      const visibleTop = scroller.scrollTop + frozenPaneBottom;
      const visibleBottom = scroller.scrollTop + scroller.clientHeight;
      if (rowStart < visibleTop) {
        nextScrollTop = rowStart - frozenPaneBottom;
      } else if (rowEnd > visibleBottom) {
        nextScrollTop = rowEnd - scroller.clientHeight;
      }
    }

    if (colEnd > frozenPaneRight) {
      const visibleLeft = scroller.scrollLeft + frozenPaneRight;
      const visibleRight = scroller.scrollLeft + scroller.clientWidth;
      if (colStart < visibleLeft) {
        nextScrollLeft = colStart - frozenPaneRight;
      } else if (colEnd > visibleRight) {
        nextScrollLeft = colEnd - scroller.clientWidth;
      }
    }

    nextScrollTop = Math.max(0, Math.min(nextScrollTop, scroller.scrollHeight - scroller.clientHeight));
    nextScrollLeft = Math.max(0, Math.min(nextScrollLeft, scroller.scrollWidth - scroller.clientWidth));

    if (nextScrollTop !== scroller.scrollTop) {
      scroller.scrollTop = nextScrollTop;
    }
    if (nextScrollLeft !== scroller.scrollLeft) {
      scroller.scrollLeft = nextScrollLeft;
    }
    syncDrawingViewport(scroller, { immediate: true });
  }

  function moveSelection(nextRowIndex: number, nextColIndex: number, extend: boolean) {
    const clampedRowIndex = Math.max(0, Math.min(nextRowIndex, visibleRows.length - 1));
    const clampedColIndex = Math.max(0, Math.min(nextColIndex, visibleCols.length - 1));
    const nextRow = visibleRows[clampedRowIndex];
    const nextCol = visibleCols[clampedColIndex];
    if (nextRow === undefined || nextCol === undefined) {
      return;
    }

    const nextRange = { start: { row: nextRow, col: nextCol }, end: { row: nextRow, col: nextCol } };
    axisSelectionRef.current = null;
    selectionPreviewRangeRef.current = null;
    displayedSelectionRef.current = nextRange;
    applyPreviewOverlay(nextRange);
    selectCell({ row: nextRow, col: nextCol }, extend ? { extend: true } : undefined);
    ensureCellVisible(clampedRowIndex, clampedColIndex);
  }

  function resolvePageRowIndex(currentRowIndex: number, direction: 1 | -1) {
    const scroller = scrollRef.current;
    const viewportHeight = Math.max(
      displayDefaultRowHeight,
      (scroller?.clientHeight ?? displayDefaultRowHeight * 20) - frozenPaneBottom
    );
    const currentOffset = rowPrefixSums[currentRowIndex] ?? 0;
    const targetOffset = direction > 0
      ? currentOffset + viewportHeight
      : currentOffset - viewportHeight;
    return Math.max(0, Math.min(findIndexForOffsetPrefix(rowPrefixSums, targetOffset), visibleRows.length - 1));
  }

  function resolvePageColIndex(currentColIndex: number, direction: 1 | -1) {
    const scroller = scrollRef.current;
    const viewportWidth = Math.max(
      displayDefaultColWidth,
      (scroller?.clientWidth ?? displayDefaultColWidth * 8) - frozenPaneRight
    );
    const currentOffset = colPrefixSums[currentColIndex] ?? 0;
    const targetOffset = direction > 0
      ? currentOffset + viewportWidth
      : currentOffset - viewportWidth;
    return Math.max(0, Math.min(findIndexForOffsetPrefix(colPrefixSums, targetOffset), visibleCols.length - 1));
  }

  function resolveColumnSelectionNavigation(eventKey: string) {
    const axisSelection = axisSelectionRef.current;
    if (axisSelection?.axis !== "column" || firstVisibleRow === undefined) {
      return null;
    }

    const startCol = Math.min(axisSelection.startCol, axisSelection.endCol);
    const endCol = Math.max(axisSelection.startCol, axisSelection.endCol);
    const targetCol = eventKey === "ArrowLeft"
      ? startCol - 1
      : eventKey === "ArrowRight"
        ? endCol + 1
        : null;
    if (targetCol === null) {
      return null;
    }

    const rowIndex = rowIndexByActual.get(firstVisibleRow);
    const colIndex = colIndexByActual.get(targetCol);
    if (rowIndex === undefined || colIndex === undefined) {
      return null;
    }

    return { colIndex, rowIndex };
  }

  function resolveRowSelectionNavigation(eventKey: string) {
    const axisSelection = axisSelectionRef.current;
    if (axisSelection?.axis !== "row" || firstVisibleCol === undefined) {
      return null;
    }

    const startRow = Math.min(axisSelection.startRow, axisSelection.endRow);
    const endRow = Math.max(axisSelection.startRow, axisSelection.endRow);
    const targetRow = eventKey === "ArrowUp"
      ? startRow - 1
      : eventKey === "ArrowDown"
        ? endRow + 1
        : null;
    if (targetRow === null) {
      return null;
    }

    const rowIndex = rowIndexByActual.get(targetRow);
    const colIndex = colIndexByActual.get(firstVisibleCol);
    if (rowIndex === undefined || colIndex === undefined) {
      return null;
    }

    return { colIndex, rowIndex };
  }

  function handleGridKeyDown(event: React.KeyboardEvent<HTMLDivElement> | KeyboardEvent) {
    if (editingCell) {
      return;
    }

    if (!readOnly && (event.metaKey || event.ctrlKey) && !event.altKey) {
      const normalizedKey = event.key.toLowerCase();
      if (normalizedKey === "z" && event.shiftKey) {
        event.preventDefault();
        redo();
        return;
      }

      if (normalizedKey === "z") {
        event.preventDefault();
        undo();
        return;
      }

      if (normalizedKey === "y") {
        event.preventDefault();
        redo();
        return;
      }
    }

    const currentCell = resolveCurrentCell();
    if (!currentCell) {
      return;
    }

    const currentRowIndex = rowIndexByActual.get(currentCell.row) ?? 0;
    const currentColIndex = colIndexByActual.get(currentCell.col) ?? 0;
    const isCommandNavigation = event.ctrlKey || event.metaKey;

    if (!readOnly && isPrintableKey(event)) {
      event.preventDefault();
      startEditing(currentCell, event.key);
      return;
    }

    switch (event.key) {
      case "ArrowDown":
        event.preventDefault();
        if (!event.shiftKey && !isCommandNavigation) {
          const target = resolveRowSelectionNavigation(event.key);
          if (target) {
            moveSelection(target.rowIndex, target.colIndex, false);
            break;
          }
        }
        moveSelection(
          isCommandNavigation
            ? (rowIndexByActual.get(resolveLastUsedVisibleIndex(visibleRows, activeSheet?.maxUsedRow ?? -1)) ?? visibleRows.length - 1)
            : currentRowIndex + 1,
          currentColIndex,
          event.shiftKey
        );
        break;
      case "ArrowUp":
        event.preventDefault();
        if (!event.shiftKey && !isCommandNavigation) {
          const target = resolveRowSelectionNavigation(event.key);
          if (target) {
            moveSelection(target.rowIndex, target.colIndex, false);
            break;
          }
        }
        moveSelection(
          isCommandNavigation
            ? (rowIndexByActual.get(resolveFirstUsedVisibleIndex(visibleRows, activeSheet?.minUsedRow ?? -1)) ?? 0)
            : currentRowIndex - 1,
          currentColIndex,
          event.shiftKey
        );
        break;
      case "ArrowLeft":
        event.preventDefault();
        if (!event.shiftKey && !isCommandNavigation) {
          const target = resolveColumnSelectionNavigation(event.key);
          if (target) {
            moveSelection(target.rowIndex, target.colIndex, false);
            break;
          }
        }
        moveSelection(
          currentRowIndex,
          isCommandNavigation
            ? (colIndexByActual.get(resolveFirstUsedVisibleIndex(visibleCols, activeSheet?.minUsedCol ?? -1)) ?? 0)
            : currentColIndex - 1,
          event.shiftKey
        );
        break;
      case "ArrowRight":
        event.preventDefault();
        if (!event.shiftKey && !isCommandNavigation) {
          const target = resolveColumnSelectionNavigation(event.key);
          if (target) {
            moveSelection(target.rowIndex, target.colIndex, false);
            break;
          }
        }
        moveSelection(
          currentRowIndex,
          isCommandNavigation
            ? (colIndexByActual.get(resolveLastUsedVisibleIndex(visibleCols, activeSheet?.maxUsedCol ?? -1)) ?? visibleCols.length - 1)
            : currentColIndex + 1,
          event.shiftKey
        );
        break;
      case "Home":
        event.preventDefault();
        moveSelection(
          isCommandNavigation
            ? (rowIndexByActual.get(resolveFirstUsedVisibleIndex(visibleRows, activeSheet?.minUsedRow ?? -1)) ?? 0)
            : currentRowIndex,
          colIndexByActual.get(resolveFirstUsedVisibleIndex(visibleCols, activeSheet?.minUsedCol ?? -1)) ?? 0,
          event.shiftKey
        );
        break;
      case "End":
        event.preventDefault();
        moveSelection(
          isCommandNavigation
            ? (rowIndexByActual.get(resolveLastUsedVisibleIndex(visibleRows, activeSheet?.maxUsedRow ?? -1)) ?? visibleRows.length - 1)
            : currentRowIndex,
          colIndexByActual.get(resolveLastUsedVisibleIndex(visibleCols, activeSheet?.maxUsedCol ?? -1)) ?? visibleCols.length - 1,
          event.shiftKey
        );
        break;
      case "PageDown":
        event.preventDefault();
        if (event.altKey) {
          moveSelection(currentRowIndex, resolvePageColIndex(currentColIndex, 1), event.shiftKey);
          break;
        }
        moveSelection(resolvePageRowIndex(currentRowIndex, 1), currentColIndex, event.shiftKey);
        break;
      case "PageUp":
        event.preventDefault();
        if (event.altKey) {
          moveSelection(currentRowIndex, resolvePageColIndex(currentColIndex, -1), event.shiftKey);
          break;
        }
        moveSelection(resolvePageRowIndex(currentRowIndex, -1), currentColIndex, event.shiftKey);
        break;
      case "Tab":
        event.preventDefault();
        if (event.shiftKey) {
          moveSelection(
            currentColIndex > 0 ? currentRowIndex : currentRowIndex - 1,
            currentColIndex > 0 ? currentColIndex - 1 : visibleCols.length - 1,
            false
          );
        } else {
          moveSelection(
            currentColIndex < visibleCols.length - 1 ? currentRowIndex : currentRowIndex + 1,
            currentColIndex < visibleCols.length - 1 ? currentColIndex + 1 : 0,
            false
          );
        }
        break;
      case "Enter":
        event.preventDefault();
        if (event.metaKey || event.ctrlKey || event.altKey) {
          break;
        }
        if (event.shiftKey) {
          moveSelection(currentRowIndex - 1, currentColIndex, false);
          break;
        }
        moveSelection(currentRowIndex + 1, currentColIndex, false);
        break;
      case "Backspace":
      case "Delete":
        if (!readOnly) {
          event.preventDefault();
          clearSelectedCells();
        }
        break;
      case "F2":
        if (!readOnly) {
          event.preventDefault();
          startEditing(currentCell);
        }
        break;
      default:
        break;
    }
  }

  gridKeyboardHandlerRef.current = handleGridKeyDown;

  const scrollerViewportProps: XlsxScrollerRenderProps["viewportProps"] = {
    ref: scrollRef,
    "aria-colcount": Math.max(activeSheet?.colCount ?? 0, displayColLimit),
    "aria-keyshortcuts": "ArrowUp ArrowDown ArrowLeft ArrowRight Home End PageUp PageDown Control+Home Control+End",
    "aria-label": activeSheet ? `${activeSheet.name} worksheet grid` : "Workbook grid",
    "aria-readonly": readOnly,
    "aria-rowcount": Math.max(activeSheet?.rowCount ?? 0, displayRowLimit),
    role: "grid",
    onScroll: handleScrollerScroll,
    onCopy: (event) => {
      if (editingCell) {
        return;
      }

      const clipboard = event.clipboardData;
      const clipboardData = getClipboardData();
      if (!clipboardData) {
        return;
      }

      event.preventDefault();
      if (clipboard) {
        clipboard.setData("text/plain", clipboardData.text);
        clipboard.setData("text/html", clipboardData.html);
        clipboard.setData(INTERNAL_CLIPBOARD_MIME, clipboardData.structured);
        return;
      }

      void copySelectionToClipboard();
    },
    onPointerDownCapture: (event) => {
      if (event.button !== 0) {
        return;
      }
      if (isInteractiveFocusTarget(event.target, event.currentTarget)) {
        return;
      }

      gridKeyboardActiveRef.current = true;
      focusGridElement(event.currentTarget);
    },
    onClickCapture: (event) => {
      if (isInteractiveFocusTarget(event.target, event.currentTarget)) {
        return;
      }

      gridKeyboardActiveRef.current = true;
      focusGridElement(event.currentTarget);
    },
    onFocus: () => {
      gridKeyboardActiveRef.current = true;
    },
    onKeyDown: handleGridKeyDown,
    onPaste: (event) => {
      if (editingCell || readOnly) {
        return;
      }

      const clipboard = event.clipboardData;
      if (!clipboard) {
        event.preventDefault();
        void pasteFromClipboard();
        return;
      }

      const structuredPayload = clipboard.getData(INTERNAL_CLIPBOARD_MIME);
      const textPayload = clipboard.getData("text/plain");
      if (!structuredPayload && !textPayload) {
        return;
      }

      event.preventDefault();
      if (structuredPayload) {
        pasteStructuredClipboardData(structuredPayload);
        return;
      }

      pasteText(textPayload);
    },
    tabIndex: 0,
    style: {
      ["--xlsx-menu-active" as string]: selectionFill,
      ["--xlsx-menu-border" as string]: palette.strongBorder,
      ["--xlsx-menu-surface" as string]: palette.surface,
      ["--xlsx-selection-header" as string]: selectionHeaderSurface,
      backgroundColor: palette.canvas,
      color: palette.text,
      cursor: resizeGuide?.type === "column" ? "col-resize" : resizeGuide?.type === "row" ? "row-resize" : undefined,
      flex: 1,
      height: "100%",
      minHeight: 0,
      minWidth: 0,
      outline: "none",
      overflow: "auto",
      width: "100%"
    }
  };

  const scrollerContent = (
    <div
          style={{
            backgroundColor: resolveSheetSurface(activeSheet, palette),
            minHeight: "100%",
            minWidth: "100%",
            position: "relative",
            width: totalWidth,
            height: sheetContentHeight
          }}
        >
          <div
            ref={wrapperRef}
            style={{
              height: sheetContentHeight,
              left: 0,
              position: "absolute",
              top: 0,
              transform: !experimentalCanvas && isLiveZooming
                ? `translate3d(${liveZoomTranslateX}px, ${liveZoomTranslateY}px, 0) scale(${liveZoomScale})`
                : undefined,
              transformOrigin: "0 0",
              transition: "none",
              willChange: !experimentalCanvas && isLiveZooming ? "transform" : undefined,
              width: totalWidth
            }}
          >
            {showImages && !experimentalCanvas ? (
              <>
                <div style={topOverlayStyle}>{paneDrawingNodes.top}</div>
                <div style={leftOverlayStyle}>{paneDrawingNodes.left}</div>
                <div style={cornerOverlayStyle}>{paneDrawingNodes.corner}</div>
                <div style={scrollOverlayStyle}>{paneDrawingNodes.scroll}</div>
              </>
            ) : null}
            {experimentalCanvas ? (
              <>
                <div style={canvasBodyViewportLayerStyle}>
                  <canvas
                    ref={scrollBodyCanvasRef}
                    onClick={handleCanvasBodyClick}
                    onDoubleClick={handleCanvasBodyDoubleClick}
                    onPointerDown={handleCanvasBodyPointerDown}
                    style={canvasScrollBodyStyle}
                  />
                  <canvas
                    ref={topBodyCanvasRef}
                    onClick={handleCanvasBodyClick}
                    onDoubleClick={handleCanvasBodyDoubleClick}
                    onPointerDown={handleCanvasBodyPointerDown}
                    style={canvasTopBodyStyle}
                  />
                  <canvas
                    ref={leftBodyCanvasRef}
                    onClick={handleCanvasBodyClick}
                    onDoubleClick={handleCanvasBodyDoubleClick}
                    onPointerDown={handleCanvasBodyPointerDown}
                    style={canvasLeftBodyStyle}
                  />
                  <canvas
                    ref={cornerBodyCanvasRef}
                    onClick={handleCanvasBodyClick}
                    onDoubleClick={handleCanvasBodyDoubleClick}
                    onPointerDown={handleCanvasBodyPointerDown}
                    style={canvasCornerBodyStyle}
                  />
                  {hasCanvasDomDrawingOverlays ? (
                    <>
                      <div style={canvasScrollOverlayPaneStyle}>
                        <div
                          ref={canvasScrollOverlayContentRef}
                          style={{
                            height: sheetContentHeight,
                            left: 0,
                            pointerEvents: "none",
                            position: "absolute",
                            top: 0,
                            transform: `translate(${-drawingViewport.left - frozenPaneRight}px, ${-drawingViewport.top - frozenPaneBottom}px)`,
                            transformOrigin: "0 0",
                            width: totalWidth
                          }}
                        >
                          {paneDrawingNodes.scroll}
                        </div>
                      </div>
                      <div style={canvasTopOverlayPaneStyle}>
                        <div
                          ref={canvasTopOverlayContentRef}
                          style={{
                            height: sheetContentHeight,
                            left: 0,
                            pointerEvents: "none",
                            position: "absolute",
                            top: 0,
                            transform: `translate(${-drawingViewport.left - frozenPaneRight}px, ${-displayHeaderHeight}px)`,
                            transformOrigin: "0 0",
                            width: totalWidth
                          }}
                        >
                          {paneDrawingNodes.top}
                        </div>
                      </div>
                      <div style={canvasLeftOverlayPaneStyle}>
                        <div
                          ref={canvasLeftOverlayContentRef}
                          style={{
                            height: sheetContentHeight,
                            left: 0,
                            pointerEvents: "none",
                            position: "absolute",
                            top: 0,
                            transform: `translate(${-displayRowHeaderWidth}px, ${-drawingViewport.top - frozenPaneBottom}px)`,
                            transformOrigin: "0 0",
                            width: totalWidth
                          }}
                        >
                          {paneDrawingNodes.left}
                        </div>
                      </div>
                      <div style={canvasCornerOverlayPaneStyle}>
                        <div
                          ref={canvasCornerOverlayContentRef}
                          style={{
                            height: sheetContentHeight,
                            left: 0,
                            pointerEvents: "none",
                            position: "absolute",
                            top: 0,
                            transform: `translate(${-displayRowHeaderWidth}px, ${-displayHeaderHeight}px)`,
                            transformOrigin: "0 0",
                            width: totalWidth
                          }}
                        >
                          {paneDrawingNodes.corner}
                        </div>
                      </div>
                    </>
                  ) : null}
                </div>
                <div style={canvasHeaderViewportLayerStyle}>
                  <canvas
                    ref={topFrozenHeaderCanvasRef}
                    onPointerLeave={handleCanvasHeaderPointerLeave}
                    onPointerMove={handleCanvasColumnHeaderPointerMove}
                    onPointerDown={handleCanvasColumnHeaderPointerDown}
                    style={canvasTopFrozenHeaderStyle}
                  />
                  <canvas
                    ref={topScrollHeaderCanvasRef}
                    onPointerLeave={handleCanvasHeaderPointerLeave}
                    onPointerMove={handleCanvasColumnHeaderPointerMove}
                    onPointerDown={handleCanvasColumnHeaderPointerDown}
                    style={canvasTopScrollHeaderStyle}
                  />
                  <canvas
                    ref={leftFrozenHeaderCanvasRef}
                    onPointerLeave={handleCanvasHeaderPointerLeave}
                    onPointerMove={handleCanvasRowHeaderPointerMove}
                    onPointerDown={handleCanvasRowHeaderPointerDown}
                    style={canvasLeftFrozenHeaderStyle}
                  />
                  <canvas
                    ref={leftScrollHeaderCanvasRef}
                    onPointerLeave={handleCanvasHeaderPointerLeave}
                    onPointerMove={handleCanvasRowHeaderPointerMove}
                    onPointerDown={handleCanvasRowHeaderPointerDown}
                    style={canvasLeftScrollHeaderStyle}
                  />
                  <canvas ref={cornerHeaderCanvasRef} style={canvasCornerHeaderStyle} />
                </div>
                {editingCell && editingOverlayRect ? (() => {
                  const editingCellStyle = getCellData(editingCell.row, editingCell.col).style;
                  const editingBackground =
                    typeof editingCellStyle.backgroundColor === "string"
                      ? editingCellStyle.backgroundColor
                      : resolveSheetSurface(activeSheet, palette);
                  const editingColor =
                    typeof editingCellStyle.color === "string"
                      ? editingCellStyle.color
                      : resolveReadableTextColor(null, editingBackground, palette);
                  return (
                  <div
                    style={{
                      left: editingOverlayRect.left,
                      position: "absolute",
                      top: editingOverlayRect.top,
                      width: editingOverlayRect.width,
                      height: editingOverlayRect.height,
                      zIndex: 28
                    }}
                  >
                    <input
                      ref={editingInputRef}
                      autoFocus
                      onBlur={handleEditingBlur}
                      onChange={(event) => setEditingValue(event.target.value)}
                      onKeyDown={(event) => {
                        event.stopPropagation();
                        if (event.key === "Enter") {
                          event.preventDefault();
                          commitEditing();
                          return;
                        }

                        if (event.key === "Escape") {
                          event.preventDefault();
                          cancelEditing();
                          return;
                        }

                        if (
                          event.key === "ArrowDown" ||
                          event.key === "ArrowLeft" ||
                          event.key === "ArrowRight" ||
                          event.key === "ArrowUp"
                        ) {
                          event.preventDefault();
                          handleEditingNavigate(event.key, event.currentTarget.value);
                        }
                      }}
                      style={{
                        backgroundColor: editingBackground,
                        border: 0,
                        boxShadow: `inset 0 0 0 ${selectionBorderWidth}px ${selectionStroke}`,
                        boxSizing: "border-box",
                        color: editingColor,
                        display: "block",
                        font: resolveCanvasFont(editingCellStyle, 12 * zoomFactor),
                        height: "100%",
                        margin: 0,
                        minHeight: 0,
                        outline: "none",
                        overflow: "hidden",
                        padding: scaleCssLengthExpression(DEFAULT_CELL_PADDING, zoomFactor),
                        width: "100%"
                      }}
                      value={editingValue}
                    />
                  </div>
                  );
                })() : null}
                {activeCellAdornment && activeCellAdornmentRect ? (
                  <div
                    style={{
                      height: activeCellAdornmentRect.height,
                      left: activeCellAdornmentRect.left,
                      pointerEvents: "none",
                      position: "absolute",
                      top: activeCellAdornmentRect.top,
                      width: activeCellAdornmentRect.width,
                      zIndex: 27
                    }}
                  >
                    <div style={{ height: "100%", pointerEvents: "auto", position: "relative", width: "100%" }}>
                      {activeCellAdornment}
                    </div>
                  </div>
                ) : null}
              </>
            ) : (
              <table
                ref={tableRef}
                style={{
                  borderCollapse: "collapse",
                  color: "#000000",
                  flex: "0 0 auto",
                  tableLayout: "fixed",
                  width: totalWidth
                }}
              >
                <colgroup>
                  <col style={{ width: displayRowHeaderWidth }} />
                  {leadingColumnSpacerWidth > 0 ? <col style={{ width: leadingColumnSpacerWidth }} /> : null}
                  {renderedCols.map((column) => (
                    <col
                      key={column.key}
                      ref={(element) => {
                        if (element) {
                          colElementRefs.current.set(column.actualCol, element);
                        } else {
                          colElementRefs.current.delete(column.actualCol);
                        }
                      }}
                      style={{ width: column.size }}
                    />
                  ))}
                  {trailingColumnSpacerWidth > 0 ? <col style={{ width: trailingColumnSpacerWidth }} /> : null}
                </colgroup>
                <thead style={{ position: "sticky", top: 0, zIndex: canvasHeaderOverlayZIndex }}>
                  <tr>
                    <th
                      style={{
                        ...headerCellStyle,
                        backgroundColor: palette.headerSurface,
                        left: 0,
                        width: displayRowHeaderWidth,
                        zIndex: cornerHeaderOverlayZIndex
                      }}
                    />
                    {leadingColumnSpacerWidth > 0 ? (
                      <th aria-hidden="true" style={{ ...headerCellStyle, padding: 0, width: leadingColumnSpacerWidth }} />
                    ) : null}
                    {renderedCols.map((column) => (
                      <th
                        data-xlsx-col-header={column.actualCol}
                        key={column.key}
                        ref={(element) => setColHeaderRef(column.actualCol, element)}
                        onPointerDown={(event) => handleColumnPointerDown(event, column.actualCol)}
                        style={{
                          ...headerCellStyle,
                          left: stickyLeftByCol.get(column.actualCol),
                          zIndex: stickyLeftByCol.has(column.actualCol) ? stickyHeaderOverlayZIndex : headerCellStyle.zIndex
                        }}
                      >
                        <div style={{ position: "relative" }}>
                          <span
                            style={{
                              display: "inline-block",
                              transform: headerLabelLiveScale !== 1 ? `scale(${1 / headerLabelLiveScale})` : undefined,
                              transformOrigin: "center center"
                            }}
                          >
                            {columnLabel(column.actualCol)}
                          </span>
                          <div
                            onPointerDown={(event) => {
                              if (!canResizeHeaders) {
                                return;
                              }

                              event.preventDefault();
                              event.stopPropagation();
                              startColumnResize(
                                event.pointerId,
                                column.actualCol,
                                column.size,
                                event.clientX
                              );
                            }}
                            style={columnResizeHandleStyle}
                          />
                        </div>
                      </th>
                    ))}
                    {trailingColumnSpacerWidth > 0 ? (
                      <th aria-hidden="true" style={{ ...headerCellStyle, padding: 0, width: trailingColumnSpacerWidth }} />
                    ) : null}
                  </tr>
                </thead>
                <tbody>
                  {virtualRows.map((virtualRow, index) => {
                    const actualRow = visibleRows[virtualRow.index];
                    if (actualRow === undefined) {
                      return null;
                    }

                    const previousEnd = index === 0 ? 0 : (virtualRows[index - 1]?.end ?? 0);
                    const gapHeight = Math.max(0, virtualRow.start - previousEnd);

                    return (
                      <React.Fragment key={`row-fragment-${virtualRow.key}`}>
                        {gapHeight > 0 ? (
                          <tr aria-hidden="true" style={{ height: gapHeight }}>
                            <td colSpan={rowColSpan} />
                          </tr>
                        ) : null}
                        <MemoGridRow
                          actualRow={actualRow}
                          editingCell={editingCell}
                          editingValue={editingValue}
                          frozenRowHeaderZIndex={frozenRowHeaderOverlayZIndex}
                          getCellData={getCellData}
                          key={virtualRow.key}
                          leadingSpacerWidth={leadingColumnSpacerWidth}
                          onCellClick={handleCellClick}
                          onCellDoubleClick={handleCellDoubleClick}
                          onCellPointerDown={handleCellPointerDown}
                          onEditingBlur={handleEditingBlur}
                          onEditingCancel={cancelEditing}
                          onEditingCommit={commitEditing}
                          onEditingNavigate={handleEditingNavigate}
                          onEditingValueChange={setEditingValue}
                          headerLabelLiveScale={headerLabelLiveScale}
                          onRowHeaderRef={setRowHeaderRef}
                          onRowPointerDown={handleRowPointerDown}
                          onRowResizePointerDown={handleRowResizePointerDown}
                          palette={palette}
                          readOnly={readOnly}
                          renderCellAdornment={renderCellAdornment}
                          rowHeight={virtualRow.size}
                          rowHeaderZIndex={rowHeaderOverlayZIndex}
                          rowHeaderWidth={displayRowHeaderWidth}
                          stickyLeftByCol={stickyLeftByCol}
                          stickyTop={stickyTopByRow.get(actualRow)}
                          trailingSpacerWidth={trailingColumnSpacerWidth}
                          visibleCols={renderedCols}
                          zoomFactor={zoomFactor}
                        />
                      </React.Fragment>
                    );
                  })}
                  {virtualRows.length > 0 && totalHeight - (virtualRows[virtualRows.length - 1]?.end ?? totalHeight) > 0 ? (
                    <tr
                      style={{
                        height: totalHeight - (virtualRows[virtualRows.length - 1]?.end ?? totalHeight)
                      }}
                    >
                      <td colSpan={rowColSpan} />
                    </tr>
                  ) : null}
                </tbody>
              </table>
            )}
            <div
              ref={selectionOverlayRef}
              style={{
                backgroundColor: selectionFill,
                boxSizing: "border-box",
                boxShadow: `inset 0 0 0 ${selectionBorderWidth}px ${selectionStroke}`,
                contain: "layout paint",
                height: resolvedSelectionOverlay?.height ?? 0,
                left: resolvedSelectionOverlay?.left ?? 0,
                opacity: resolvedSelectionOverlay ? 1 : 0,
                pointerEvents: "none",
                position: "absolute",
                top: resolvedSelectionOverlay?.top ?? 0,
                transition: canvasSelectionTransition,
                visibility: resolvedSelectionOverlay ? "visible" : "hidden",
                willChange: shouldAnimateCanvasSelection ? "left, top, width, height" : undefined,
                width: resolvedSelectionOverlay?.width ?? 0,
                zIndex: 24
              }}
            />
            <div
              ref={activeValidationOverlayRef}
              aria-hidden="true"
              style={{
                alignItems: "center",
                color: palette.mutedText,
                display: "inline-flex",
                fontSize: 10 * zoomFactor,
                fontWeight: 700,
                height: 16 * zoomFactor,
                justifyContent: "center",
                opacity: 0,
                pointerEvents: "none",
                position: "absolute",
                transform: "translateY(-50%)",
                visibility: "hidden",
                width: 12 * zoomFactor,
                zIndex: 26
              }}
            >
              ▾
            </div>
            <div
              ref={fillHandleRef}
              onPointerDown={(event) => {
                if (readOnly || event.button !== 0 || !normalizedSelection || !resolvedSelectionOverlay) {
                  return;
                }

                event.preventDefault();
                event.stopPropagation();
                startFillDrag(event.pointerId, normalizedSelection);
              }}
              style={{
                backgroundColor: selectionStroke,
                border: `${Math.max(1, zoomFactor)}px solid ${palette.surface}`,
                contain: "layout paint",
                cursor: "crosshair",
                display: !readOnly && resolvedSelectionOverlay ? "block" : "none",
                height: 8 * zoomFactor,
                left: resolvedSelectionOverlay ? resolvedSelectionOverlay.left + resolvedSelectionOverlay.width - (4 * zoomFactor) : 0,
                pointerEvents: "auto",
                position: "absolute",
                top: resolvedSelectionOverlay ? resolvedSelectionOverlay.top + resolvedSelectionOverlay.height - (4 * zoomFactor) : 0,
                transition: shouldAnimateCanvasSelection
                  ? "left 120ms cubic-bezier(0.22, 1, 0.36, 1), top 120ms cubic-bezier(0.22, 1, 0.36, 1)"
                  : "none",
                willChange: shouldAnimateCanvasSelection ? "left, top" : undefined,
                width: 8 * zoomFactor,
                zIndex: 25
              }}
            />
            {resizeGuide ? (
              <div
                aria-hidden="true"
                style={{
                  backgroundColor: selectionStroke,
                  borderRadius: Math.max(999, 3 * zoomFactor),
                  boxShadow: `0 0 0 ${Math.max(1, zoomFactor)}px ${palette.surface}`,
                  height: resizeGuide.type === "column" ? Math.max(12, 14 * zoomFactor) : Math.max(3, 2 * zoomFactor),
                  left: resizeGuide.type === "column"
                    ? resizeGuide.position - Math.max(2, 1.5 * zoomFactor)
                    : Math.max(3, 4 * zoomFactor),
                  pointerEvents: "none",
                  position: "absolute",
                  top: resizeGuide.type === "column"
                    ? Math.max(2 * zoomFactor, displayHeaderHeight - Math.max(14, 16 * zoomFactor))
                    : resizeGuide.position - Math.max(2, 1.5 * zoomFactor),
                  width: resizeGuide.type === "column" ? Math.max(3, 2 * zoomFactor) : Math.max(12, 14 * zoomFactor),
                  zIndex: 52
                }}
              />
            ) : null}
            {!renderTableHeaderMenu && openTableMenuState ? (
              <div
                ref={tableMenuRef}
                style={{
                  color: palette.text,
                  left: Math.max(displayRowHeaderWidth + (4 * zoomFactor), openTableMenuState.left),
                  position: "absolute",
                  top: openTableMenuState.top,
                  zIndex: 50
                }}
              >
                <DefaultTableHeaderMenu
                  cell={{ col: openTableMenuState.column.index + openTableMenuState.table.start.col, row: openTableMenuState.table.start.row }}
                  column={openTableMenuState.column}
                  direction={
                    sortState &&
                    sortState.tableName === openTableMenuState.table.name &&
                    sortState.columnIndex === openTableMenuState.column.index
                      ? sortState.direction
                      : null
                  }
                  sortAscending={() => {
                    sortTable(openTableMenuState.table.name, openTableMenuState.column.index, "ascending");
                    setOpenTableMenu(null);
                  }}
                  sortDescending={() => {
                    sortTable(openTableMenuState.table.name, openTableMenuState.column.index, "descending");
                    setOpenTableMenu(null);
                  }}
                  table={openTableMenuState.table}
                  triggerIcon="▾"
                  triggerProps={{ type: "button" }}
                />
              </div>
            ) : null}
          </div>
        </div>
  );

  return (
    <div style={{ backgroundColor: palette.canvas, display: "flex", flex: 1, minHeight: 0, minWidth: 0 }}>
      {renderScroller
        ? renderScroller({ children: scrollerContent, viewportProps: scrollerViewportProps })
        : <div key={activeTabIndex} {...scrollerViewportProps}>{scrollerContent}</div>}
    </div>
  );
}

function XlsxViewerInner({
  allowResizeInReadOnly = false,
  className,
  controller,
  emptyState,
  enableCanvasSelectionAnimation = true,
  enableGestureZoom = true,
  errorState,
  experimentalCanvas = true,
  fileTooLargeState,
  getCellStyle,
  height,
  isDark = false,
  loadingComponent,
  loadingState,
  renderChartLoading,
  renderImage,
  renderImageSelection,
  renderScroller,
  renderTableHeaderMenu,
  rounded = true,
  selectionColor,
  selectionFillColor,
  selectionHeaderColor,
  showImages = true,
  showDefaultToolbar = true,
  toolbar
}: XlsxViewerProps & {
  controller: XlsxViewerController;
}) {
  const palette = useViewerPalette(isDark);
  const { displayFileName, error } = controller;
  const customFileTooLarge =
    error instanceof XlsxFileSizeLimitExceededError
      ? renderCustomFileTooLarge(
          fileTooLargeState,
          {
            displayFileName,
            fileSizeBytes: error.fileSizeBytes,
            maxFileSizeBytes: error.maxFileSizeBytes
          },
          palette
        )
      : undefined;

  return (
    <ViewerAppearanceContext.Provider value={{ isDark }}>
      <ViewerContext.Provider value={controller}>
        {customFileTooLarge !== undefined ? (
          customFileTooLarge
        ) : (
          <div
            className={classNames("react-xlsx-viewer", className)}
            style={{
              blockSize: height,
              backgroundColor: palette.surface,
              borderRadius: rounded ? 12 : 0,
              color: palette.text,
              display: "flex",
              flex: "1 1 auto",
              flexDirection: "column",
              inlineSize: "100%",
              isolation: "isolate",
              maxHeight: "100%",
              maxWidth: "100%",
              minHeight: 0,
              minWidth: 0,
              overflow: "hidden",
              position: "relative",
              width: "100%"
            }}
          >
            {resolveToolbar(toolbar, showDefaultToolbar, controller, palette)}
            <div style={{ display: "flex", flex: 1, minHeight: 0, minWidth: 0 }}>
              <XlsxGrid
                allowResizeInReadOnly={allowResizeInReadOnly}
                controller={controller}
                emptyState={emptyState}
                enableCanvasSelectionAnimation={enableCanvasSelectionAnimation}
                enableGestureZoom={enableGestureZoom}
                errorState={errorState}
                experimentalCanvas={experimentalCanvas}
                fileTooLargeState={fileTooLargeState}
                getCellStyle={getCellStyle}
                loadingComponent={loadingComponent}
                loadingState={loadingState}
                palette={palette}
                renderChartLoading={renderChartLoading}
                renderImage={renderImage}
                renderImageSelection={renderImageSelection}
                renderScroller={renderScroller}
                renderTableHeaderMenu={renderTableHeaderMenu}
                selectionColor={selectionColor}
                selectionFillColor={selectionFillColor}
                selectionHeaderColor={selectionHeaderColor}
                showImages={showImages}
              />
            </div>
          </div>
        )}
      </ViewerContext.Provider>
    </ViewerAppearanceContext.Provider>
  );
}

function XlsxViewerWithInlineController(props: XlsxViewerProps) {
  const controller = useXlsxViewerController(props);
  return <XlsxViewerInner {...props} controller={controller} />;
}

function XlsxViewerProviderWithInlineController({
  children,
  isDark = false,
  ...options
}: Omit<XlsxViewerProviderProps, "controller">) {
  const controller = useXlsxViewerController(options);
  return (
    <ViewerAppearanceContext.Provider value={{ isDark }}>
      <ViewerContext.Provider value={controller}>{children}</ViewerContext.Provider>
    </ViewerAppearanceContext.Provider>
  );
}

export function XlsxViewerProvider({ children, controller, isDark = false, ...options }: XlsxViewerProviderProps) {
  if (controller) {
    return (
      <ViewerAppearanceContext.Provider value={{ isDark }}>
        <ViewerContext.Provider value={controller}>{children}</ViewerContext.Provider>
      </ViewerAppearanceContext.Provider>
    );
  }

  return (
    <XlsxViewerProviderWithInlineController {...options} isDark={isDark}>
      {children}
    </XlsxViewerProviderWithInlineController>
  );
}

export function useXlsxViewer() {
  const context = React.useContext(ViewerContext);
  if (!context) {
    throw new Error("useXlsxViewer must be used inside XlsxViewer or XlsxViewerProvider.");
  }

  return context;
}

export function useXlsxViewerSelection(): XlsxViewerSelection {
  const {
    activeCell,
    activeCellAddress,
    clearSelection,
    selectedRangeAddress,
    selectCell,
    selectRange,
    selection
  } = useXlsxViewer();

  return React.useMemo(
    () => ({
      activeCell,
      activeCellAddress,
      clearSelection,
      selectedRangeAddress,
      selectCell,
      selectRange,
      selection
    }),
    [activeCell, activeCellAddress, clearSelection, selectedRangeAddress, selectCell, selectRange, selection]
  );
}

export function useXlsxViewerZoom(): XlsxViewerZoom {
  const {
    canZoomIn,
    canZoomOut,
    defaultZoomScale,
    maxZoomScale,
    minZoomScale,
    resetZoom,
    setZoomScale,
    zoomIn,
    zoomOut,
    zoomScale
  } = useXlsxViewer();

  return React.useMemo(
    () => ({
      canZoomIn,
      canZoomOut,
      defaultZoomScale,
      maxZoomScale,
      minZoomScale,
      resetZoom,
      setZoomScale,
      zoomIn,
      zoomOut,
      zoomScale
    }),
    [
      canZoomIn,
      canZoomOut,
      defaultZoomScale,
      maxZoomScale,
      minZoomScale,
      resetZoom,
      setZoomScale,
      zoomIn,
      zoomOut,
      zoomScale
    ]
  );
}

export function useXlsxViewerEditing(): XlsxViewerEditing {
  const {
    addSheet,
    canRedo,
    canUndo,
    clearSelectedCells,
    copySelectionToClipboard,
    defineNamedRange,
    fillSelection,
    getClipboardData,
    getCellDisplayValue,
    getCellFormula,
    mergeSelection,
    pasteFromClipboard,
    pasteStructuredClipboardData,
    pasteText,
    removeActiveSheet,
    readOnly,
    redo,
    selectedCellFormula,
    selectedChartFormula,
    selectedFormula,
    selectedFormulaTarget,
    selectedValue,
    setCellFormula,
    setCellStyle,
    setCellValue,
    setRangeStyle,
    setSelectedFormula,
    setSelectedCellFormula,
    setSelectedCellStyle,
    setSelectedCellValue,
    undo,
    unmergeSelection
  } = useXlsxViewer();

  return React.useMemo(
    () => ({
      addSheet,
      canRedo,
      canUndo,
      clearSelectedCells,
      copySelectionToClipboard,
      defineNamedRange,
      fillSelection,
      getClipboardData,
      getCellDisplayValue,
      getCellFormula,
      mergeSelection,
      pasteFromClipboard,
      pasteStructuredClipboardData,
      pasteText,
      removeActiveSheet,
      readOnly,
      redo,
      selectedCellFormula,
      selectedChartFormula,
      selectedFormula,
      selectedFormulaTarget,
      selectedValue,
      setCellFormula,
      setCellStyle,
      setCellValue,
      setRangeStyle,
      setSelectedFormula,
      setSelectedCellFormula,
      setSelectedCellStyle,
      setSelectedCellValue,
      undo,
      unmergeSelection
    }),
    [
      addSheet,
      canRedo,
      canUndo,
      clearSelectedCells,
      copySelectionToClipboard,
      defineNamedRange,
      fillSelection,
      getClipboardData,
      getCellDisplayValue,
      getCellFormula,
      mergeSelection,
      pasteFromClipboard,
      pasteStructuredClipboardData,
      pasteText,
      removeActiveSheet,
      readOnly,
      redo,
      selectedCellFormula,
      selectedChartFormula,
      selectedFormula,
      selectedFormulaTarget,
      selectedValue,
      setCellFormula,
      setCellStyle,
      setCellValue,
      setRangeStyle,
      setSelectedFormula,
      setSelectedCellFormula,
      setSelectedCellStyle,
      setSelectedCellValue,
      undo,
      unmergeSelection
    ]
  );
}

export function useXlsxViewerTables(): XlsxViewerTables {
  const { sortState, sortTable, tables } = useXlsxViewer();

  return React.useMemo(
    () => ({
      sortState,
      sortTable,
      tables
    }),
    [sortState, sortTable, tables]
  );
}

export function useXlsxViewerImages(): XlsxViewerImages {
  const {
    charts,
    clearSelectedChart,
    clearSelectedChartElement,
    clearSelectedImage,
    getChartById,
    getChartSeriesFormula,
    getSheetCharts,
    getImageById,
    getSheetImages,
    images,
    isChartsLoading,
    moveChartBy,
    moveImageBy,
    readOnly,
    resizeChartBy,
    resizeImageBy,
    selectedChart,
    selectedChartElement,
    selectedChartFormula,
    selectedChartId,
    selectedImage,
    selectedImageId,
    selectChart,
    selectChartElement,
    selectImage,
    setChartSeriesFormula,
    setChartRect,
    setImageRect,
    updateChart
  } = useXlsxViewer();

  return React.useMemo(
    () => ({
      charts,
      clearSelectedChart,
      clearSelectedChartElement,
      clearSelectedImage,
      getChartById,
      getChartSeriesFormula,
      getSheetCharts,
      getImageById,
      getSheetImages,
      images,
      isChartsLoading,
      moveChartBy,
      moveImageBy,
      readOnly,
      resizeChartBy,
      resizeImageBy,
      selectedChart,
      selectedChartElement,
      selectedChartFormula,
      selectedChartId,
      selectedImage,
      selectedImageId,
      selectChart,
      selectChartElement,
      selectImage,
      setChartSeriesFormula,
      setChartRect,
      setImageRect,
      updateChart
    }),
    [
      charts,
      clearSelectedChart,
      clearSelectedChartElement,
      clearSelectedImage,
      getChartById,
      getChartSeriesFormula,
      getSheetCharts,
      getImageById,
      getSheetImages,
      images,
      isChartsLoading,
      moveChartBy,
      moveImageBy,
      readOnly,
      resizeChartBy,
      resizeImageBy,
      selectedChart,
      selectedChartElement,
      selectedChartFormula,
      selectedChartId,
      selectedImage,
      selectedImageId,
      selectChart,
      selectChartElement,
      selectImage,
      setChartSeriesFormula,
      setChartRect,
      setImageRect,
      updateChart
    ]
  );
}

export function useXlsxViewerCharts(): XlsxViewerCharts {
  const {
    activeTab,
    activeTabIndex,
    charts,
    chartsheets,
    clearSelectedChart,
    clearSelectedChartElement,
    getChartById,
    getChartSeriesFormula,
    getChartsheetById,
    getSheetCharts,
    isChartsLoading,
    moveChartBy,
    readOnly,
    resizeChartBy,
    selectChart,
    selectedChart,
    selectedChartElement,
    selectedChartFormula,
    selectedChartId,
    selectChartElement,
    setActiveTabIndex,
    setChartSeriesFormula,
    setChartRect,
    tabs,
    updateChart
  } = useXlsxViewer();

  return React.useMemo(
    () => ({
      activeTab,
      activeTabIndex,
      charts,
      chartsheets,
      clearSelectedChart,
      clearSelectedChartElement,
      getChartById,
      getChartSeriesFormula,
      getChartsheetById,
      getSheetCharts,
      isChartsLoading,
      moveChartBy,
      readOnly,
      resizeChartBy,
      selectChart,
      selectedChart,
      selectedChartElement,
      selectedChartFormula,
      selectedChartId,
      selectChartElement,
      setActiveTabIndex,
      setChartSeriesFormula,
      setChartRect,
      tabs,
      updateChart
    }),
    [
      activeTab,
      activeTabIndex,
      charts,
      chartsheets,
      clearSelectedChart,
      clearSelectedChartElement,
      getChartById,
      getChartSeriesFormula,
      getChartsheetById,
      getSheetCharts,
      isChartsLoading,
      moveChartBy,
      readOnly,
      resizeChartBy,
      selectChart,
      selectedChart,
      selectedChartElement,
      selectedChartFormula,
      selectedChartId,
      selectChartElement,
      setActiveTabIndex,
      setChartSeriesFormula,
      setChartRect,
      tabs,
      updateChart
    ]
  );
}

export function useXlsxViewerThumbnails(
  options: UseXlsxViewerThumbnailsOptions = {}
): XlsxViewerThumbnails {
  const { getSheetFormControls, getSheetImages, getSheetShapes, workbook, sheets } = useXlsxViewer();
  const { isDark } = React.useContext(ViewerAppearanceContext);
  const palette = useViewerPalette(isDark);
  const includeHeaders = options.includeHeaders ?? true;
  const resolution = options.resolution;
  const thumbnailImageCacheRef = React.useRef(new Map<string, {
    failed: boolean;
    image: HTMLImageElement;
    loaded: boolean;
    src: string;
  }>());
  const isMountedRef = React.useRef(false);
  const [thumbnailImageLoadVersion, setThumbnailImageLoadVersion] = React.useState(0);

  React.useEffect(() => {
    isMountedRef.current = true;
    return () => {
      isMountedRef.current = false;
    };
  }, []);

  const getThumbnailImage = React.useCallback((image: XlsxImage) => {
    if (typeof Image === "undefined") {
      return null;
    }

    const cacheKey = image.id;
    const cached = thumbnailImageCacheRef.current.get(cacheKey);
    if (cached && cached.src === image.src) {
      return cached.loaded && !cached.failed ? cached.image : null;
    }

    const imageElement = new Image();
    const entry = {
      failed: false,
      image: imageElement,
      loaded: false,
      src: image.src
    };
    thumbnailImageCacheRef.current.set(cacheKey, entry);
    imageElement.onload = () => {
      entry.loaded = true;
      if (isMountedRef.current) {
        setThumbnailImageLoadVersion((version) => version + 1);
      }
    };
    imageElement.onerror = () => {
      entry.failed = true;
      if (isMountedRef.current) {
        setThumbnailImageLoadVersion((version) => version + 1);
      }
    };
    imageElement.src = image.src;
    return null;
  }, []);

  const thumbnails = React.useMemo<XlsxSheetThumbnail[]>(() => {
    return sheets.map((sheet, sheetIndex) => {
      const worksheet = workbook?.getSheet(sheet.workbookSheetIndex) ?? null;
      const sheetImages = getSheetImages(sheetIndex)
        .filter((image) => image.src)
        .sort((left, right) => left.zIndex - right.zIndex);
      const sheetShapes = getSheetShapes(sheetIndex);
      const sheetFormControls = getSheetFormControls(sheetIndex).filter((control) => !control.hidden);
      const drawingBounds = [
        ...sheetImages.map((image) => image.anchor),
        ...sheetShapes.map((shape) => shape.anchor),
        ...sheetFormControls.map((control) => control.anchor)
      ].reduce<{
        maxCol: number;
        maxRow: number;
        minCol: number;
        minRow: number;
      } | null>((bounds, anchor) => {
        const drawingBound = resolveAnchoredBounds(anchor) ?? {
          maxCol: 0,
          maxRow: 0,
          minCol: 0,
          minRow: 0
        };
        return {
          maxCol: Math.max(bounds?.maxCol ?? drawingBound.maxCol, drawingBound.maxCol),
          maxRow: Math.max(bounds?.maxRow ?? drawingBound.maxRow, drawingBound.maxRow),
          minCol: Math.min(bounds?.minCol ?? drawingBound.minCol, drawingBound.minCol),
          minRow: Math.min(bounds?.minRow ?? drawingBound.minRow, drawingBound.minRow)
        };
      }, null);
      const sheetMinUsedRow = (sheet.maxUsedRow ?? -1) >= (sheet.minUsedRow ?? 0) ? (sheet.minUsedRow ?? 0) : undefined;
      const sheetMinUsedCol = (sheet.maxUsedCol ?? -1) >= (sheet.minUsedCol ?? 0) ? (sheet.minUsedCol ?? 0) : undefined;
      const effectiveMinUsedRow = Math.max(0, Math.min(sheetMinUsedRow ?? drawingBounds?.minRow ?? 0, drawingBounds?.minRow ?? sheetMinUsedRow ?? 0));
      const effectiveMinUsedCol = Math.max(0, Math.min(sheetMinUsedCol ?? drawingBounds?.minCol ?? 0, drawingBounds?.minCol ?? sheetMinUsedCol ?? 0));
      const effectiveMaxUsedRow = Math.max(sheet.maxUsedRow ?? -1, drawingBounds?.maxRow ?? -1);
      const effectiveMaxUsedCol = Math.max(sheet.maxUsedCol ?? -1, drawingBounds?.maxCol ?? -1);
      const showGridLines = sheet.showGridLines ?? true;
      const hiddenRowSet = new Set(sheet.hiddenRows ?? []);
      const hiddenColSet = new Set(sheet.hiddenCols ?? []);
      const visibleRows = buildVisibleAxisIndices(
        sheet.visibleRows ?? [],
        Math.max(THUMBNAIL_FALLBACK_ROWS, effectiveMaxUsedRow + THUMBNAIL_FALLBACK_ROWS + 1),
        effectiveMaxUsedRow,
        hiddenRowSet
      );
      const visibleCols = buildVisibleAxisIndices(
        sheet.visibleCols ?? [],
        Math.max(THUMBNAIL_FALLBACK_COLS, effectiveMaxUsedCol + THUMBNAIL_FALLBACK_COLS + 1),
        effectiveMaxUsedCol,
        hiddenColSet
      );
      const resolveColumnWidthPx = (actualCol: number) => {
        if (worksheet && !worksheet.isColumnHidden(actualCol)) {
          const width = worksheet.getColumnWidth(actualCol);
          if (width !== undefined && width !== null) {
            return Math.max(
              resolveRenderedSheetAxisPixels(resolveSheetColumnWidthPixels(width, sheet.columnWidthCharacterWidthPx), showGridLines),
              DEFAULT_COL_WIDTH / 2
            );
          }
        }

        return Math.max(
          resolveRenderedSheetAxisPixels(
            sheet.colWidthOverridesPx[actualCol] ?? sheet.defaultColWidthPx ?? DEFAULT_COL_WIDTH,
            showGridLines
          ),
          DEFAULT_COL_WIDTH / 2
        );
      };
      const resolveRowHeightPx = (actualRow: number) => {
        if (worksheet && !worksheet.isRowHidden(actualRow)) {
          const height = worksheet.getRowHeight(actualRow);
          if (height !== undefined && height !== null) {
            return Math.max(
              resolveRenderedSheetAxisPixels(resolveSheetRowHeightPixels(height), showGridLines),
              DEFAULT_ROW_HEIGHT / 1.5
            );
          }
        }

        return Math.max(
          resolveRenderedSheetAxisPixels(
            sheet.rowHeightOverridesPx[actualRow] ?? sheet.defaultRowHeightPx ?? DEFAULT_ROW_HEIGHT,
            showGridLines
          ),
          DEFAULT_ROW_HEIGHT / 1.5
        );
      };
      const previewRows = resolveThumbnailAxisIndices({
        fallbackCount: THUMBNAIL_FALLBACK_ROWS,
        getSizePx: resolveRowHeightPx,
        maxCount: THUMBNAIL_MAX_ROWS,
        maxLogicalPixels: THUMBNAIL_MAX_SOURCE_HEIGHT_PX,
        maxUsedIndex: effectiveMaxUsedRow,
        minUsedIndex: effectiveMinUsedRow,
        precomputed: visibleRows
      });
      const previewCols = resolveThumbnailAxisIndices({
        fallbackCount: THUMBNAIL_FALLBACK_COLS,
        getSizePx: resolveColumnWidthPx,
        maxCount: THUMBNAIL_MAX_COLS,
        maxLogicalPixels: THUMBNAIL_MAX_SOURCE_WIDTH_PX,
        maxUsedIndex: effectiveMaxUsedCol,
        minUsedIndex: effectiveMinUsedCol,
        precomputed: visibleCols
      });
      const rowAxis = buildThumbnailAxisItems(previewRows, resolveRowHeightPx);
      const colAxis = buildThumbnailAxisItems(previewCols, resolveColumnWidthPx);
      const rowItemByActual = new Map(rowAxis.items.map((item) => [item.actualIndex, item]));
      const colItemByActual = new Map(colAxis.items.map((item) => [item.actualIndex, item]));
      const headerHeight = includeHeaders ? HEADER_HEIGHT : 0;
      const rowHeaderWidth = includeHeaders ? ROW_HEADER_WIDTH : 0;
      const sourceWidth = rowHeaderWidth + colAxis.totalSize;
      const sourceHeight = headerHeight + rowAxis.totalSize;
      const outputSize = resolveThumbnailOutputSize(sourceWidth, sourceHeight, resolution);
      const sourceRange: XlsxCellRange = {
        start: {
          col: previewCols[0] ?? 0,
          row: previewRows[0] ?? 0
        },
        end: {
          col: previewCols[previewCols.length - 1] ?? 0,
          row: previewRows[previewRows.length - 1] ?? 0
        }
      };
      const sparklineByCell = new Map<string, XlsxSheetData["sparklines"][number]>(
        (sheet.sparklines ?? []).map((sparkline) => [`${sparkline.target.row}:${sparkline.target.col}`, sparkline])
      );
      const thumbnailImageRects = sheetImages
        .map((image) => ({
          image,
          rect: resolveThumbnailAnchoredRect(
            image.anchor,
            visibleRows,
            visibleCols,
            resolveRowHeightPx,
            resolveColumnWidthPx,
            previewRows,
            previewCols,
            {
              headerHeight,
              rowHeaderWidth
            }
          )
        }))
        .filter(({ rect }) => (
          rect.left + rect.width >= rowHeaderWidth
          && rect.top + rect.height >= headerHeight
          && rect.left <= sourceWidth
          && rect.top <= sourceHeight
        ));
      const thumbnailShapeRects = sheetShapes
        .map((shape) => ({
          rect: resolveThumbnailAnchoredRect(
            shape.anchor,
            visibleRows,
            visibleCols,
            resolveRowHeightPx,
            resolveColumnWidthPx,
            previewRows,
            previewCols,
            {
              headerHeight,
              rowHeaderWidth
            }
          ),
          shape
        }))
        .filter(({ rect }) => (
          rect.left + rect.width >= rowHeaderWidth
          && rect.top + rect.height >= headerHeight
          && rect.left <= sourceWidth
          && rect.top <= sourceHeight
        ));
      const thumbnailFormControlRects = sheetFormControls
        .map((control) => ({
          control,
          rect: resolveThumbnailAnchoredRect(
            control.anchor,
            visibleRows,
            visibleCols,
            resolveRowHeightPx,
            resolveColumnWidthPx,
            previewRows,
            previewCols,
            {
              headerHeight,
              rowHeaderWidth
            }
          )
        }))
        .filter(({ rect }) => (
          rect.left + rect.width >= rowHeaderWidth
          && rect.top + rect.height >= headerHeight
          && rect.left <= sourceWidth
          && rect.top <= sourceHeight
        ));
      const thumbnailDrawingEntries = [
        ...thumbnailShapeRects.map(({ rect, shape }) => ({
          kind: "shape" as const,
          rect,
          shape,
          zIndex: shape.zIndex
        })),
        ...thumbnailFormControlRects.map(({ control, rect }) => ({
          control,
          kind: "formControl" as const,
          rect,
          zIndex: control.zIndex
        })),
        ...thumbnailImageRects.map(({ image, rect }) => ({
          image,
          kind: "image" as const,
          rect,
          zIndex: image.zIndex
        }))
      ].sort((left, right) => left.zIndex - right.zIndex);

      const paint = (canvas: HTMLCanvasElement | null) => {
        if (!canvas) {
          return false;
        }

        const context = canvas.getContext("2d", { alpha: false });
        if (!context) {
          return false;
        }

        const dpr = typeof window === "undefined" ? 1 : Math.max(1, window.devicePixelRatio || 1);
        const outputWidth = outputSize.width;
        const outputHeight = outputSize.height;
        const scale = outputSize.scale;
        const deviceWidth = Math.max(1, Math.round(outputWidth * dpr));
        const deviceHeight = Math.max(1, Math.round(outputHeight * dpr));
        if (canvas.width !== deviceWidth) {
          canvas.width = deviceWidth;
        }
        if (canvas.height !== deviceHeight) {
          canvas.height = deviceHeight;
        }
        if (canvas.style.width !== `${outputWidth}px`) {
          canvas.style.width = `${outputWidth}px`;
        }
        if (canvas.style.height !== `${outputHeight}px`) {
          canvas.style.height = `${outputHeight}px`;
        }

        context.setTransform(dpr * scale, 0, 0, dpr * scale, 0, 0);
        context.clearRect(0, 0, Math.max(1, sourceWidth), Math.max(1, sourceHeight));
        const thumbnailSheetSurface = resolveSheetSurface(sheet, palette);
        context.fillStyle = palette.canvas;
        context.fillRect(0, 0, Math.max(1, sourceWidth), Math.max(1, sourceHeight));
        context.fillStyle = thumbnailSheetSurface;
        context.fillRect(rowHeaderWidth, headerHeight, Math.max(1, colAxis.totalSize), Math.max(1, rowAxis.totalSize));

        if (includeHeaders) {
          context.fillStyle = palette.headerSurface;
          context.fillRect(rowHeaderWidth, 0, Math.max(1, colAxis.totalSize), headerHeight);
          context.fillStyle = palette.rowHeaderSurface;
          context.fillRect(0, headerHeight, rowHeaderWidth, Math.max(1, rowAxis.totalSize));
          context.fillRect(0, 0, rowHeaderWidth, headerHeight);
        }

        if (!worksheet) {
          return true;
        }

        const conditionalFormatMetricsCache = new Map<string, number[]>();
        const cellRenderCache = new Map<string, CellRenderData>();
        const getCellData = (row: number, col: number): CellRenderData => {
          const cacheKey = `${row}:${col}`;
          const cached = cellRenderCache.get(cacheKey);
          if (cached) {
            return cached;
          }

          if (worksheet.isMergedSecondary(row, col)) {
            const mergedSecondaryData: CellRenderData = {
              isMergedSecondary: true,
              style: {},
              value: ""
            };
            cellRenderCache.set(cacheKey, mergedSecondaryData);
            return mergedSecondaryData;
          }

          const merge = worksheet.getMergeSpan(row, col) as { colSpan?: number; rowSpan?: number } | null | undefined;
          const inheritedStyle = resolveInheritedCellStyle(sheet, row, col);
          const worksheetStyle = (worksheet.getCellStyleAt(row, col) as Record<string, unknown> | null | undefined) ?? null;
          const mergedStyle = mergeResolvedCellStyle(inheritedStyle, worksheetStyle, { replaceXfSubtrees: true });
          const alignment = mergedStyle?.alignment as Record<string, unknown> | undefined;
          const sparkline = sparklineByCell.get(cacheKey) ?? null;
          const sparklineValues = sparkline
            ? (
                sparkline.range.start.row === sparkline.range.end.row
                  ? Array.from(
                      { length: Math.abs(sparkline.range.end.col - sparkline.range.start.col) + 1 },
                      (_, index) => getCellNumericValue(
                        worksheet,
                        sparkline.range.start.row,
                        Math.min(sparkline.range.start.col, sparkline.range.end.col) + index
                      )
                    )
                  : Array.from(
                      { length: Math.abs(sparkline.range.end.row - sparkline.range.start.row) + 1 },
                      (_, index) => getCellNumericValue(
                        worksheet,
                        Math.min(sparkline.range.start.row, sparkline.range.end.row) + index,
                        sparkline.range.start.col
                      )
                    )
              )
            : null;
          const checkboxState = mergedStyle?.cellControl ? getCellBooleanValue(worksheet, row, col) : null;
          const nextData: CellRenderData = {
            canvas: undefined,
            checkboxState,
            colSpan: merge?.colSpan,
            conditionalColorScale: resolveConditionalColorScaleForCell(
              row,
              col,
              worksheet,
              sheet,
              conditionalFormatMetricsCache
            ),
            conditionalDataBar: resolveConditionalDataBarForCell(
              row,
              col,
              worksheet,
              sheet,
              conditionalFormatMetricsCache
            ),
            conditionalIcon: resolveConditionalIconForCell(
              row,
              col,
              worksheet,
              sheet,
              conditionalFormatMetricsCache
            ),
            isMergedSecondary: false,
            rowSpan: merge?.rowSpan,
            sparkline: sparkline && sparklineValues ? { config: sparkline, values: sparklineValues } : null,
            style: buildCellStyle(mergedStyle, palette, sheet.themePalette, {
              showGridLines
            }),
            textRotationDeg: resolveSpreadsheetTextRotation(alignment?.textRotation),
            value: sparkline
              ? ""
              : checkboxState !== null
                ? ""
                : getCellDisplayValue(worksheet, row, col, sheet)
          };

          nextData.canvas = buildCanvasCellStyleCache(nextData.style);
          cellRenderCache.set(cacheKey, nextData);
          return nextData;
        };

        for (const rowItem of rowAxis.items) {
          for (const colItem of colAxis.items) {
            const cellData = getCellData(rowItem.actualIndex, colItem.actualIndex);
            if (cellData.isMergedSecondary) {
              continue;
            }

            const colSpan = Math.max(1, cellData.colSpan ?? 1);
            const rowSpan = Math.max(1, cellData.rowSpan ?? 1);
            let width = colItem.size;
            let height = rowItem.size;
            for (let colOffset = 1; colOffset < colSpan; colOffset += 1) {
              const nextCol = colItemByActual.get(colItem.actualIndex + colOffset);
              if (!nextCol) {
                break;
              }
              width += nextCol.size;
            }
            for (let rowOffset = 1; rowOffset < rowSpan; rowOffset += 1) {
              const nextRow = rowItemByActual.get(rowItem.actualIndex + rowOffset);
              if (!nextRow) {
                break;
              }
              height += nextRow.size;
            }

            const rect = {
              height,
              left: rowHeaderWidth + colItem.start,
              top: headerHeight + rowItem.start,
              width
            };
            const canvasCellStyle = cellData.canvas ?? buildCanvasCellStyleCache(cellData.style);
            const gradientFill = typeof cellData.style.backgroundImage === "string"
              ? resolveCanvasGradientFill(context, rect, cellData.style.backgroundImage)
              : null;
            const fillColor = cellData.conditionalColorScale?.color
              ?? (typeof cellData.style.backgroundColor === "string" ? cellData.style.backgroundColor : thumbnailSheetSurface);
            const hasExplicitCellFill =
              cellData.conditionalColorScale !== null
              || gradientFill !== null
              || (typeof cellData.style.backgroundColor === "string"
                && cellData.style.backgroundColor !== thumbnailSheetSurface);

            context.fillStyle = gradientFill ?? fillColor;
            context.fillRect(rect.left, rect.top, rect.width, rect.height);

            if (cellData.conditionalDataBar) {
              const barLeft = rect.left + 4;
              const barTop = rect.top + 4;
              const barWidth = Math.max(0, (rect.width - 8) * (cellData.conditionalDataBar.widthPercent / 100));
              const barHeight = Math.max(0, rect.height - 8);
              if (barWidth > 0 && barHeight > 0) {
                context.fillStyle = resolveCanvasDataBarFill(
                  context,
                  barLeft,
                  barTop,
                  barWidth,
                  barHeight,
                  cellData.conditionalDataBar
                );
                context.fillRect(barLeft, barTop, barWidth, barHeight);
                if (cellData.conditionalDataBar.border !== false && cellData.conditionalDataBar.borderColor) {
                  context.strokeStyle = cellData.conditionalDataBar.borderColor;
                  context.lineWidth = 1;
                  context.strokeRect(barLeft + 0.5, barTop + 0.5, Math.max(0, barWidth - 1), Math.max(0, barHeight - 1));
                }
              }
            }

            if (showGridLines && !hasExplicitCellFill) {
              context.strokeStyle = palette.border;
              context.lineWidth = 1;
              context.beginPath();
              if (colItem.start === 0) {
                context.moveTo(rect.left + 0.5, rect.top);
                context.lineTo(rect.left + 0.5, rect.top + rect.height);
              }
              if (rowItem.start === 0) {
                context.moveTo(rect.left, rect.top + 0.5);
                context.lineTo(rect.left + rect.width, rect.top + 0.5);
              }
              context.moveTo(rect.left + rect.width - 0.5, rect.top);
              context.lineTo(rect.left + rect.width - 0.5, rect.top + rect.height);
              context.moveTo(rect.left, rect.top + rect.height - 0.5);
              context.lineTo(rect.left + rect.width, rect.top + rect.height - 0.5);
              context.stroke();
            }

            if (canvasCellStyle.topBorder) {
              strokeCanvasBorderSide(context, "top", rect, canvasCellStyle.topBorder);
            }
            if (canvasCellStyle.rightBorder) {
              strokeCanvasBorderSide(context, "right", rect, canvasCellStyle.rightBorder);
            }
            if (canvasCellStyle.bottomBorder) {
              strokeCanvasBorderSide(context, "bottom", rect, canvasCellStyle.bottomBorder);
            }
            if (canvasCellStyle.leftBorder) {
              strokeCanvasBorderSide(context, "left", rect, canvasCellStyle.leftBorder);
            }

            const rawText = cellData.value ?? "";
            const shouldDrawThumbnailContent = (
              cellData.checkboxState != null
              || cellData.sparkline
              || rawText.length > 0
              || cellData.conditionalIcon
            );
            if (!shouldDrawThumbnailContent) {
              continue;
            }

            const padding = canvasCellStyle.padding;
            const contentLeft = rect.left + padding.left;
            const contentTop = rect.top + padding.top;
            const contentWidth = Math.max(0, rect.width - padding.left - padding.right);
            const contentHeight = Math.max(0, rect.height - padding.top - padding.bottom);
            if (contentWidth <= 0 || contentHeight <= 0) {
              continue;
            }

            context.save();
            context.beginPath();
            context.rect(contentLeft, contentTop, contentWidth, contentHeight);
            context.clip();
            context.font = canvasCellStyle.baseFont;
            context.fillStyle = canvasCellStyle.textColor;
            context.textBaseline = "middle";

            if (cellData.checkboxState != null) {
              const boxSize = Math.min(14, contentWidth, contentHeight);
              const boxLeft = rect.left + (rect.width - boxSize) / 2;
              const boxTop = rect.top + (rect.height - boxSize) / 2;
              context.strokeStyle = paletteIsDark(palette) ? "#cbd5e1" : "#475569";
              context.lineWidth = 1.25;
              context.strokeRect(boxLeft, boxTop, boxSize, boxSize);
              if (cellData.checkboxState) {
                context.fillStyle = paletteIsDark(palette) ? "#60a5fa" : "#2563eb";
                context.fillRect(boxLeft + 1.5, boxTop + 1.5, Math.max(0, boxSize - 3), Math.max(0, boxSize - 3));
              }
            } else if (cellData.sparkline) {
              const sparkline = cellData.sparkline.config;
              const sparklineValues = cellData.sparkline.values;
              const points = sparklineValues
                .map((value, index) => ({ index, value }))
                .filter((entry): entry is { index: number; value: number } => typeof entry.value === "number" && Number.isFinite(entry.value));
              if (points.length > 1) {
                const minValue = Math.min(...points.map((entry) => entry.value));
                const maxValue = Math.max(...points.map((entry) => entry.value));
                const sparkLeft = contentLeft + 1;
                const sparkTop = contentTop + 2;
                const sparkWidth = Math.max(1, contentWidth - 2);
                const sparkHeight = Math.max(1, contentHeight - 4);
                const xStep = points.length > 1 ? sparkWidth / (points.length - 1) : 0;
                context.strokeStyle = sparkline.color ?? "#2563eb";
                context.lineCap = "round";
                context.lineJoin = "round";
                context.lineWidth = 1.25;
                context.beginPath();
                points.forEach((entry, index) => {
                  const x = sparkLeft + index * xStep;
                  const y = sparkTop + sparkHeight - clampSparklineValue(entry.value, minValue, maxValue) * sparkHeight;
                  if (index === 0) {
                    context.moveTo(x, y);
                  } else {
                    context.lineTo(x, y);
                  }
                });
                context.stroke();
              }
            } else if (rawText.length > 0) {
              const align = canvasCellStyle.textAlign;
              context.textAlign = align;
              const textX = align === "right"
                ? contentLeft + contentWidth
                : align === "center"
                  ? contentLeft + (contentWidth / 2)
                  : contentLeft;
              const trailingInset = cellData.conditionalIcon ? 18 : 0;
              const maxTextWidth = Math.max(0, contentWidth - trailingInset);
              const singleLineHeight = resolveCanvasLineHeight(cellData.style, 12);

              if (cellData.textRotationDeg) {
                const alignedTextY = resolveCanvasTextMiddleY(cellData.style.verticalAlign, contentTop, contentHeight, singleLineHeight);
                const rotationOriginX = contentLeft + (contentWidth / 2);
                context.save();
                context.translate(rotationOriginX, alignedTextY);
                context.rotate((cellData.textRotationDeg * Math.PI) / 180);
                context.fillText(
                  rawText,
                  align === "right"
                    ? contentWidth / 2
                    : align === "center"
                      ? 0
                      : -(contentWidth / 2),
                  0
                );
                context.restore();
              } else if (canvasCellStyle.usesWrappedText || rawText.includes("\n")) {
                const lines = wrapCanvasText(context, rawText, maxTextWidth);
                const lineHeight = resolveCanvasLineHeight(cellData.style, 12);
                const textBlockHeight = lines.length * lineHeight;
                let textBlockTop = contentTop;
                if (cellData.style.verticalAlign === "middle") {
                  textBlockTop = contentTop + ((contentHeight - textBlockHeight) / 2);
                } else if (cellData.style.verticalAlign !== "top") {
                  textBlockTop = contentTop + contentHeight - textBlockHeight;
                }
                lines.forEach((line, lineIndex) => {
                  context.fillText(line, textX, textBlockTop + (lineIndex * lineHeight) + (lineHeight / 2));
                });
              } else {
                const text = canvasCellStyle.textOverflowEllipsis
                  ? truncateCanvasText(context, rawText, maxTextWidth)
                  : rawText;
                context.fillText(text, textX, resolveCanvasTextMiddleY(cellData.style.verticalAlign, contentTop, contentHeight, singleLineHeight));
              }
            }

            if (cellData.conditionalIcon) {
              const iconSize = 10;
              const iconX = rect.left + rect.width - (padding.right + iconSize + 4);
              const iconY = rect.top + (rect.height / 2);
              drawCanvasConditionalIcon(context, cellData.conditionalIcon, iconX + (iconSize / 2), iconY, iconSize);
            }

            context.restore();
          }
        }

        if (thumbnailDrawingEntries.length > 0) {
          context.save();
          context.beginPath();
          context.rect(rowHeaderWidth, headerHeight, Math.max(1, colAxis.totalSize), Math.max(1, rowAxis.totalSize));
          context.clip();
          for (const entry of thumbnailDrawingEntries) {
            if (entry.kind === "shape") {
              drawStaticShape(context, entry.shape, entry.rect, 1);
            } else if (entry.kind === "formControl") {
              drawStaticFormControl(context, entry.control, entry.rect, palette, 1, thumbnailSheetSurface);
            } else {
              const imageElement = getThumbnailImage(entry.image);
              if (imageElement) {
                context.drawImage(imageElement, entry.rect.left, entry.rect.top, entry.rect.width, entry.rect.height);
              }
            }
          }
          context.restore();
        }

        if (includeHeaders) {
          context.strokeStyle = palette.border;
          context.lineWidth = 1;
          context.font = "600 11px ui-sans-serif, system-ui, sans-serif";
          context.fillStyle = palette.mutedText;
          context.textBaseline = "middle";

          for (const colItem of colAxis.items) {
            const left = rowHeaderWidth + colItem.start;
            context.beginPath();
            context.moveTo(left + colItem.size - 0.5, 0);
            context.lineTo(left + colItem.size - 0.5, headerHeight);
            context.moveTo(left, headerHeight - 0.5);
            context.lineTo(left + colItem.size, headerHeight - 0.5);
            context.stroke();
            context.textAlign = "center";
            context.fillText(columnLabel(colItem.actualIndex), left + (colItem.size / 2), headerHeight / 2);
          }

          for (const rowItem of rowAxis.items) {
            const top = headerHeight + rowItem.start;
            context.beginPath();
            context.moveTo(0, top + rowItem.size - 0.5);
            context.lineTo(rowHeaderWidth, top + rowItem.size - 0.5);
            context.moveTo(rowHeaderWidth - 0.5, top);
            context.lineTo(rowHeaderWidth - 0.5, top + rowItem.size);
            context.stroke();
            context.textAlign = "center";
            context.fillText(`${rowItem.actualIndex + 1}`, rowHeaderWidth / 2, top + (rowItem.size / 2));
          }

          context.beginPath();
          context.moveTo(rowHeaderWidth - 0.5, 0);
          context.lineTo(rowHeaderWidth - 0.5, headerHeight);
          context.moveTo(0, headerHeight - 0.5);
          context.lineTo(rowHeaderWidth, headerHeight - 0.5);
          context.stroke();
        }

        return true;
      };

      const plan: SheetThumbnailRenderPlan = {
        aspectRatio: sourceWidth / Math.max(1, sourceHeight),
        contentHeight: rowAxis.totalSize,
        contentWidth: colAxis.totalSize,
        height: outputSize.height,
        paint,
        sourceRange,
        width: outputSize.width
      };

      return {
        ...plan,
        sheet,
        sheetIndex,
        workbookSheetIndex: sheet.workbookSheetIndex
      };
    });
  }, [
    getSheetFormControls,
    getSheetImages,
    getSheetShapes,
    getThumbnailImage,
    includeHeaders,
    palette,
    resolution,
    sheets,
    thumbnailImageLoadVersion,
    workbook
  ]);

  const paintThumbnail = React.useCallback(
    (sheetIndex: number, canvas: HTMLCanvasElement | null) => thumbnails[sheetIndex]?.paint(canvas) ?? false,
    [thumbnails]
  );

  return React.useMemo(
    () => ({
      paintThumbnail,
      thumbnails
    }),
    [paintThumbnail, thumbnails]
  );
}

export function XlsxViewer(props: XlsxViewerProps) {
  const contextController = React.useContext(ViewerContext);

  if (props.controller) {
    return <XlsxViewerInner {...props} controller={props.controller} />;
  }

  if (contextController) {
    return <XlsxViewerInner {...props} controller={contextController} />;
  }

  return <XlsxViewerWithInlineController {...props} />;
}

export function DefaultXlsxToolbar() {
  const controller = useXlsxViewer();
  const { isDark } = React.useContext(ViewerAppearanceContext);
  const palette = useViewerPalette(isDark);
  return <DefaultToolbar controller={controller} palette={palette} />;
}
