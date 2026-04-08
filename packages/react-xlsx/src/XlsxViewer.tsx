import * as React from "react";
import type { Worksheet } from "@dukelib/sheets-wasm";
import {
  elementScroll,
  observeElementOffset,
  observeElementRect,
  useVirtualizer,
  type Rect,
  type Virtualizer
} from "@tanstack/react-virtual";
import { resolveWorkbookColor, resolveWorkbookFillStyle } from "./colors";
import { useXlsxViewerController } from "./controller";
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
  XlsxChartsheet,
  XlsxCellAddress,
  XlsxCellRange,
  XlsxImage,
  XlsxImageRect,
  XlsxImageRenderProps,
  XlsxImageResizeHandlePosition,
  XlsxImageSelectionRenderProps,
  XlsxShape,
  XlsxSheetData,
  XlsxThemePalette,
  XlsxTable,
  XlsxTableColumn,
  XlsxTableHeaderMenuRenderProps,
  XlsxViewerCharts,
  XlsxViewerTables,
  XlsxViewerController,
  XlsxViewerEditing,
  XlsxViewerImages,
  XlsxViewerProps,
  XlsxViewerProviderProps,
  XlsxViewerSelection,
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
const SHEET_SURFACE = "#ffffff";
const SHEET_GRIDLINE = "#d9d9d9";
const DEFAULT_CELL_PADDING = "0 4px";
const IMAGE_HANDLE_POSITIONS: XlsxImageResizeHandlePosition[] = ["nw", "n", "ne", "e", "se", "s", "sw", "w"];

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

function observeZoomedElementRect(
  instance: Virtualizer<HTMLDivElement, Element>,
  zoomFactor: number,
  cb: (rect: Rect) => void
) {
  return observeElementRect(instance, (rect) => {
    cb({
      height: rect.height / zoomFactor,
      width: rect.width / zoomFactor
    });
  });
}

function observeZoomedElementOffset(
  instance: Virtualizer<HTMLDivElement, Element>,
  zoomFactor: number,
  cb: (offset: number, isScrolling: boolean) => void
) {
  return observeElementOffset(instance, (offset, isScrolling) => {
    cb(offset / zoomFactor, isScrolling);
  });
}

function scrollToZoomedOffset(
  offset: number,
  zoomFactor: number,
  options: Parameters<typeof elementScroll>[1],
  instance: Virtualizer<HTMLDivElement, Element>
) {
  elementScroll(offset * zoomFactor, options, instance);
}

function formatZoomScale(zoomScale: number) {
  return `${Math.round(zoomScale)}%`;
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

function isPrintableKey(event: React.KeyboardEvent) {
  return event.key.length === 1 && !event.altKey && !event.ctrlKey && !event.metaKey;
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
    rowIndexByActual?: Map<number, number>;
    rowPrefixSums?: number[];
  }
): XlsxImageRect {
  const resolveMarkerLeft = (col: number, colOffsetEmu: number) =>
    ROW_HEADER_WIDTH +
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
    HEADER_HEIGHT +
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
      left: ROW_HEADER_WIDTH + emuToPixels(anchor.positionEmu.x),
      top: HEADER_HEIGHT + emuToPixels(anchor.positionEmu.y),
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
    rowIndexByActual?: Map<number, number>;
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
  stickyLeftByCol: Map<number, number>
): FrozenDrawingPane {
  const frozenPaneBottom =
    freezePanes?.row && freezePanes.row > 0 && frozenRows.length > 0
      ? frozenRows.reduce(
          (max, row) => Math.max(max, (stickyTopByRow.get(row) ?? HEADER_HEIGHT) + (actualRowHeights[row] ?? DEFAULT_ROW_HEIGHT)),
          HEADER_HEIGHT
        )
      : null;
  const frozenPaneRight =
    freezePanes?.col && freezePanes.col > 0 && frozenCols.length > 0
      ? frozenCols.reduce(
          (max, col) => Math.max(max, (stickyLeftByCol.get(col) ?? ROW_HEADER_WIDTH) + (actualColWidths[col] ?? DEFAULT_COL_WIDTH)),
          ROW_HEADER_WIDTH
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

function buildShapeContainerStyle(shape: XlsxShape, rect: XlsxImageRect): React.CSSProperties {
  const borderWidth = shape.stroke?.none ? 0 : Math.max(0, shape.stroke?.widthPx ?? 0);
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
    borderRadius = 12;
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

function clampImageRect(rect: XlsxImageRect): XlsxImageRect {
  return {
    height: Math.max(IMAGE_MIN_SIZE_PX, rect.height),
    left: Math.max(ROW_HEADER_WIDTH, rect.left),
    top: Math.max(HEADER_HEIGHT, rect.top),
    width: Math.max(IMAGE_MIN_SIZE_PX, rect.width)
  };
}

function resolveImageHandleStyle(position: XlsxImageResizeHandlePosition, stroke: string, surface: string): React.CSSProperties {
  const offset = IMAGE_HANDLE_SIZE_PX / 2;
  const style: React.CSSProperties = {
    backgroundColor: surface,
    border: `1px solid ${stroke}`,
    borderRadius: 6,
    cursor: IMAGE_HANDLE_CURSOR[position],
    height: IMAGE_HANDLE_SIZE_PX,
    pointerEvents: "auto",
    position: "absolute",
    width: IMAGE_HANDLE_SIZE_PX
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

function resolveIsDarkMode() {
  if (typeof document === "undefined") {
    return false;
  }

  const classList = document.documentElement.classList;
  if (classList.contains("dark")) {
    return true;
  }
  if (classList.contains("light")) {
    return false;
  }

  return typeof window !== "undefined" && window.matchMedia("(prefers-color-scheme: dark)").matches;
}

function useViewerPalette() {
  const [isDarkMode, setIsDarkMode] = React.useState(resolveIsDarkMode);

  React.useEffect(() => {
    if (typeof window === "undefined") {
      return;
    }

    const update = () => setIsDarkMode(resolveIsDarkMode());
    const mediaQuery = window.matchMedia("(prefers-color-scheme: dark)");
    const observer = new MutationObserver(update);

    observer.observe(document.documentElement, {
      attributeFilter: ["class", "data-theme"],
      attributes: true
    });

    mediaQuery.addEventListener?.("change", update);
    update();

    return () => {
      observer.disconnect();
      mediaQuery.removeEventListener?.("change", update);
    };
  }, []);

  return isDarkMode ? DARK_PALETTE : LIGHT_PALETTE;
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

  const rowStyleId = sheet.rowStyleIds[row];
  if (rowStyleId !== undefined) {
    return sheet.styleById[rowStyleId] ?? null;
  }

  const colStyleId = sheet.colStyleIds[col];
  if (colStyleId !== undefined) {
    return sheet.styleById[colStyleId] ?? null;
  }

  return null;
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
    hair: "1px",
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
  const fontStyle = typeof style.fontStyle === "string" ? style.fontStyle : "normal";
  const fontWeight =
    typeof style.fontWeight === "string" || typeof style.fontWeight === "number" ? String(style.fontWeight) : "400";
  const fontSize = typeof style.fontSize === "string" ? style.fontSize : "12px";
  const fontFamily = typeof style.fontFamily === "string" ? style.fontFamily : "sans-serif";
  return `${fontStyle} ${fontWeight} ${fontSize} ${fontFamily}`;
}

function measureTextWidth(value: string, style: React.CSSProperties) {
  if (!value) {
    return 0;
  }

  if (typeof document === "undefined") {
    return value.length * 7;
  }

  textMeasureCanvas ??= document.createElement("canvas");
  const context = textMeasureCanvas.getContext("2d");
  if (!context) {
    return value.length * 7;
  }

  context.font = buildCanvasFont(style);
  return context.measureText(value).width;
}

function canCellTextOverflow(data: CellRenderData) {
  if (!data.value || data.colSpan || data.isMergedSecondary || data.style.whiteSpace === "pre-wrap") {
    return false;
  }

  const textAlign = data.style.textAlign;
  if (textAlign && textAlign !== "left" && textAlign !== "start") {
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
  overlay: Record<string, unknown> | null | undefined
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
    nextStyle.border = {
      ...(baseBorder ?? {}),
      ...(overlayBorder ?? {})
    };
  }

  const baseFill = base?.fill as Record<string, unknown> | undefined;
  const overlayFill = overlay?.fill as Record<string, unknown> | undefined;
  if (baseFill || overlayFill) {
    nextStyle.fill = overlayFill ?? baseFill;
  }

  const baseFont = base?.font as Record<string, unknown> | undefined;
  const overlayFont = overlay?.font as Record<string, unknown> | undefined;
  if (baseFont || overlayFont) {
    nextStyle.font = {
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
  close,
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
        onClick={() => {
          sortAscending();
          close();
        }}
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
        onClick={() => {
          sortDescending();
          close();
        }}
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
              onClick={download}
              style={{
                background: palette.buttonSurface,
                border: `1px solid ${palette.strongBorder}`,
                borderRadius: 8,
                color: palette.buttonText,
                cursor: "pointer",
                fontSize: 12,
                fontWeight: 500,
                padding: "6px 10px"
              }}
              type="button"
            >
              Download
            </button>
          ) : null}
        </div>
      </div>
      {tabs.length > 1 ? (
        <div
          style={{
            backgroundColor: palette.subtleSurface,
            borderBottom: `1px solid ${palette.border}`,
            display: "flex",
            gap: 6,
            overflowX: "auto",
            padding: "8px 12px"
          }}
        >
          {tabs.map((tab, index) => (
            <button
              key={tab.id}
              onClick={() => setActiveTabIndex(index)}
              style={{
                backgroundColor: index === activeTabIndex ? palette.sheetActiveSurface : palette.sheetInactiveSurface,
                border: `1px solid ${index === activeTabIndex ? palette.strongBorder : "transparent"}`,
                borderRadius: 8,
                boxShadow: index === activeTabIndex ? palette.shadow : "none",
                color: index === activeTabIndex ? palette.sheetActiveText : palette.sheetInactiveText,
                cursor: "pointer",
                fontSize: 12,
                fontWeight: 500,
                padding: "6px 12px",
                whiteSpace: "nowrap"
              }}
              type="button"
            >
              {tab.name}
            </button>
          ))}
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
};

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

function renderConditionalIcon(icon: NonNullable<CellRenderData["conditionalIcon"]>) {
  if (icon.shape === "arrow") {
    const fill = icon.color ?? "#111827";
    const stroke = icon.borderColor ?? darkenColor(fill, 0.32);
    return (
      <svg
        aria-hidden="true"
        height={14}
        style={{ display: "block" }}
        viewBox="0 0 16 16"
        width={14}
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
          fontSize: 13,
          fontWeight: 700,
          height: 14,
          justifyContent: "center",
          lineHeight: 1,
          width: 14
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
        height: 12,
        width: 12
      }}
    />
  );
}

function renderCheckboxControl(checked: boolean, palette: ViewerPalette) {
  const stroke = paletteIsDark(palette) ? "#cbd5e1" : "#475569";
  const fill = checked ? (paletteIsDark(palette) ? "#60a5fa" : "#2563eb") : "transparent";
  const check = paletteIsDark(palette) ? "#020617" : "#ffffff";
  return (
    <svg aria-hidden="true" height={14} style={{ display: "block" }} viewBox="0 0 16 16" width={14}>
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
  activeCell: XlsxCellAddress | null;
  actualRow: number;
  editingCell: XlsxCellAddress | null;
  editingValue: string;
  getCellData: (row: number, col: number) => CellRenderData;
  leadingSpacerWidth: number;
  onCellDoubleClick: (cell: XlsxCellAddress) => void;
  onCellClick: (cell: XlsxCellAddress, cellData: CellRenderData) => void;
  onCellPointerDown: (event: React.PointerEvent<HTMLTableCellElement>, cell: XlsxCellAddress, isActive: boolean) => void;
  onEditingCancel: () => void;
  onEditingCommit: () => void;
  onEditingValueChange: (value: string) => void;
  onRowPointerDown: (event: React.PointerEvent<HTMLTableCellElement>, actualRow: number) => void;
  onRowResizePointerDown: (event: React.PointerEvent<HTMLDivElement>, actualRow: number, rowHeight: number) => void;
  palette: ViewerPalette;
  readOnly: boolean;
  renderCellAdornment?: (cell: XlsxCellAddress) => React.ReactNode;
  rowHeight: number;
  rowSelected: boolean;
  stickyLeftByCol: Map<number, number>;
  stickyTop?: number;
  trailingSpacerWidth: number;
  visibleCols: RenderedColumn[];
};

function GridRow({
  activeCell,
  actualRow,
  editingCell,
  editingValue,
  getCellData,
  leadingSpacerWidth,
  onCellClick,
  onCellDoubleClick,
  onCellPointerDown,
  onEditingCancel,
  onEditingCommit,
  onEditingValueChange,
  onRowPointerDown,
  onRowResizePointerDown,
  palette,
  readOnly,
  renderCellAdornment,
  rowHeight,
  rowSelected,
  stickyLeftByCol,
  stickyTop,
  trailingSpacerWidth,
  visibleCols
}: GridRowProps) {
  const gutterSeparatorShadow = `inset -1px 0 0 ${palette.border}, inset 0 -1px 0 ${palette.border}`;

  return (
    <tr data-xlsx-row={actualRow} style={{ height: rowHeight }}>
      <td
        onPointerDown={(event) => onRowPointerDown(event, actualRow)}
        style={{
          backgroundColor: rowSelected ? "var(--xlsx-selection-header)" : palette.rowHeaderSurface,
          borderBottom: "none",
          borderRight: "none",
          boxSizing: "border-box",
          boxShadow: gutterSeparatorShadow,
          color: palette.mutedText,
          fontSize: "11px",
          height: rowHeight,
          left: 0,
          maxHeight: rowHeight,
          minWidth: ROW_HEADER_WIDTH,
          padding: "2px 4px",
          position: "sticky",
          top: stickyTop,
          textAlign: "center",
          userSelect: "none",
          width: ROW_HEADER_WIDTH,
          zIndex: stickyTop !== undefined ? 45 : 35
        }}
      >
        <div style={{ position: "relative" }}>
          {actualRow + 1}
          <div
            onPointerDown={(event) => onRowResizePointerDown(event, actualRow, rowHeight)}
            style={{
              backgroundColor: "transparent",
              bottom: -8,
              cursor: "row-resize",
              height: 16,
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
        const isActive = isSameCell(activeCell, cell);
        const isEditing = isSameCell(editingCell, cell);
        const isSpilling = Boolean(cellData.spillWidth && cellData.spillWidth > 0);
        const adornment = renderCellAdornment ? renderCellAdornment(cell) : null;
        const stickyLeft = stickyLeftByCol.get(actualCol);
        const showValidationDropdown = Boolean(
          isActive &&
          !isEditing &&
          cellData.validation?.validationType === "list" &&
          cellData.validation.showDropdown
        );
        const validationRight = adornment ? 24 : 4;
        const conditionalIconRight = validationRight + (showValidationDropdown ? 16 : 0);
        const cellStyle: React.CSSProperties = {
          ...cellData.style,
          boxSizing: "border-box",
          cursor: isEditing ? "text" : cellData.hyperlink ? "pointer" : "cell"
        };

        if (!cellData.rowSpan) {
          cellStyle.height = rowHeight;
          cellStyle.maxHeight = rowHeight;
        }

        if (isActive || isSpilling) {
          cellStyle.position = "relative";
          cellStyle.zIndex = isActive ? 3 : 2;
        }
        if (cellData.conditionalColorScale) {
          cellStyle.backgroundColor = cellData.conditionalColorScale.color;
        }
        if (cellData.conditionalColorScale || cellData.conditionalDataBar || cellData.conditionalIcon) {
          cellStyle.position = "relative";
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
          maxHeight: rowHeight,
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
        const trailingInset = (adornment ? 20 : 0) + (cellData.conditionalIcon ? 18 : 0) + (showValidationDropdown ? 18 : 0);
        if (cellData.conditionalDataBar) {
          cellContentStyle.position = "relative";
          cellContentStyle.zIndex = 1;
        }
        if (trailingInset > 0) {
          cellContentStyle.paddingRight = trailingInset + 4;
        }
        if (cellData.conditionalIcon || showValidationDropdown) {
          cellContentStyle.position = "relative";
          cellContentStyle.zIndex = 1;
        }
        if (cellData.sparkline || cellData.checkboxState !== null) {
          cellContentStyle.alignItems = "center";
          cellContentStyle.justifyContent = "center";
        }
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
            onPointerDown={(event) => onCellPointerDown(event, cell, isActive)}
            style={cellStyle}
            title={title}
          >
            {cellData.conditionalDataBar ? (
              <div
                aria-hidden="true"
                style={{
                  alignItems: "center",
                  bottom: 4,
                  display: "flex",
                  left: 4,
                  pointerEvents: "none",
                  position: "absolute",
                  right: 4,
                  top: 4,
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
                  right: conditionalIconRight,
                  top: "50%",
                  transform: "translateY(-50%)",
                  zIndex: 2
                }}
              >
                {renderConditionalIcon(cellData.conditionalIcon)}
              </div>
            ) : null}
            {showValidationDropdown ? (
              <div
                aria-hidden="true"
                style={{
                  alignItems: "center",
                  color: palette.mutedText,
                  display: "inline-flex",
                  fontSize: 10,
                  fontWeight: 700,
                  justifyContent: "center",
                  pointerEvents: "none",
                  position: "absolute",
                  right: validationRight,
                  top: "50%",
                  transform: "translateY(-50%)",
                  zIndex: 2
                }}
              >
                ▾
              </div>
            ) : null}
            {isEditing ? (
              <input
                autoFocus
                onBlur={onEditingCommit}
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
                  padding: DEFAULT_CELL_PADDING,
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
                {renderCheckboxControl(cellData.checkboxState, palette)}
              </div>
            ) : (
              <div style={cellContentStyle}>{cellData.value}</div>
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
    prev.rowSelected !== next.rowSelected ||
    prev.palette !== next.palette ||
    prev.readOnly !== next.readOnly ||
    prev.visibleCols !== next.visibleCols ||
    prev.leadingSpacerWidth !== next.leadingSpacerWidth ||
    prev.stickyLeftByCol !== next.stickyLeftByCol ||
    prev.stickyTop !== next.stickyTop ||
    prev.trailingSpacerWidth !== next.trailingSpacerWidth ||
    prev.getCellData !== next.getCellData ||
    prev.onCellClick !== next.onCellClick ||
    prev.onCellDoubleClick !== next.onCellDoubleClick ||
    prev.onCellPointerDown !== next.onCellPointerDown ||
    prev.onEditingCancel !== next.onEditingCancel ||
    prev.onEditingCommit !== next.onEditingCommit ||
    prev.onEditingValueChange !== next.onEditingValueChange ||
    prev.onRowPointerDown !== next.onRowPointerDown ||
    prev.onRowResizePointerDown !== next.onRowResizePointerDown ||
    prev.renderCellAdornment !== next.renderCellAdornment
  ) {
    return false;
  }

  const prevActiveCol = prev.activeCell?.row === prev.actualRow ? prev.activeCell.col : -1;
  const nextActiveCol = next.activeCell?.row === next.actualRow ? next.activeCell.col : -1;
  if (prevActiveCol !== nextActiveCol) {
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
  controller,
  emptyState,
  errorState,
  loadingComponent,
  loadingState,
  renderChartLoading,
  palette,
  renderImage,
  renderImageSelection,
  renderTableHeaderMenu,
  selectionColor,
  selectionFillColor,
  selectionHeaderColor,
  showImages = true
}: Pick<
  XlsxViewerProps,
  "emptyState" | "errorState" | "loadingComponent" | "loadingState" | "renderChartLoading" | "renderImage" | "renderImageSelection" | "renderTableHeaderMenu" | "selectionColor" | "selectionFillColor" | "selectionHeaderColor" | "showImages"
> & {
  controller: XlsxViewerController;
  palette: ViewerPalette;
}) {
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
    error,
    fillSelection,
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
    copySelectionToClipboard,
    pasteFromClipboard,
    pasteStructuredClipboardData,
    pasteText,
    readOnly,
    redo,
    revision,
    selectedChart,
    selectedChartId,
    selectedImage,
    selectedImageId,
    selectCell,
    selectChart,
    selectImage,
    selectRange,
    selection,
    setActiveSheetIndex,
    setChartRect,
    setImageRect,
    setCellValue,
    sheets,
    tabs,
    sortState,
    sortTable,
    tables,
    undo,
    zoomScale
  } = controller;
  const scrollRef = React.useRef<HTMLDivElement>(null);
  const wrapperRef = React.useRef<HTMLDivElement>(null);
  const tableRef = React.useRef<HTMLTableElement>(null);
  const selectionOverlayRef = React.useRef<HTMLDivElement>(null);
  const fillHandleRef = React.useRef<HTMLDivElement>(null);
  const tableMenuRef = React.useRef<HTMLDivElement>(null);
  const colElementRefs = React.useRef(new Map<number, HTMLTableColElement>());
  const rowElementRefs = React.useRef(new Map<number, HTMLTableRowElement>());
  const columnPreviewRef = React.useRef<{ actualIndex: number; size: number } | null>(null);
  const rowPreviewRef = React.useRef<{ actualIndex: number; size: number } | null>(null);
  const activeCellRef = React.useRef<XlsxCellAddress | null>(activeCell);
  const selectionRef = React.useRef<XlsxCellRange | null>(null);
  const editingCellRef = React.useRef<XlsxCellAddress | null>(null);
  const readOnlyRef = React.useRef(readOnly);
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
    | {
        anchor: XlsxCellAddress;
        axis: "cell" | "column" | "row";
        contentScaleX: number;
        contentScaleY: number;
        committedOnPointerDown: boolean;
        didDrag: boolean;
        originCell: XlsxCellAddress;
        originContentX: number;
        originContentY: number;
        previewRange: XlsxCellRange;
        pointerId: number;
        startClientX: number;
        startClientY: number;
      }
    | null
  >(null);
  const fillDragRef = React.useRef<
    | {
        pointerId: number;
        previewRange: XlsxCellRange;
        sourceRange: XlsxCellRange;
      }
    | null
  >(null);
  const selectionDragCleanupRef = React.useRef<(() => void) | null>(null);
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
  const chartPreviewRectRef = React.useRef<{ id: string; rect: XlsxImageRect } | null>(null);
  const skipNextChartClickRef = React.useRef<string | null>(null);
  const paneDrawingNodesCacheRef = React.useRef<{
    chartRects: Array<{ chart: XlsxChart; rect: XlsxImageRect }>;
    drawingViewport: DrawingViewport;
    imageRects: Array<{ image: XlsxImage; rect: XlsxImageRect }>;
    palette: ViewerPalette;
    readOnly: boolean;
    renderChartLoading: XlsxViewerProps["renderChartLoading"];
    renderImage: XlsxViewerProps["renderImage"];
    renderImageSelection: XlsxViewerProps["renderImageSelection"];
    isChartsLoading: boolean;
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
  const syncDrawingViewport = React.useCallback((scroller: HTMLDivElement | null) => {
    if (!scroller) {
      return;
    }

    setDrawingViewport((current) => {
      const next = {
        height: scroller.clientHeight / zoomFactor,
        left: scroller.scrollLeft / zoomFactor,
        top: scroller.scrollTop / zoomFactor,
        width: scroller.clientWidth / zoomFactor
      };
      return current.left === next.left
        && current.top === next.top
        && current.width === next.width
        && current.height === next.height
        ? current
        : next;
    });
  }, [zoomFactor]);
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
      activeSheet?.rowHeightOverridesPx,
      activeSheet?.rowHeights,
      activeSheet?.showGridLines,
      activeSheet?.visibleRows,
      displayRowLimit,
      worksheet,
      revision
    ]
  );
  const effectiveColWidths = React.useMemo(
    () => visibleCols.map((col) => actualColWidths[col] ?? DEFAULT_COL_WIDTH),
    [actualColWidths, visibleCols]
  );
  const effectiveRowHeights = React.useMemo(
    () => visibleRows.map((row) => actualRowHeights[row] ?? DEFAULT_ROW_HEIGHT),
    [actualRowHeights, visibleRows]
  );
  const rowIndexByActual = React.useMemo(() => new Map(visibleRows.map((row, index) => [row, index])), [visibleRows]);
  const colIndexByActual = React.useMemo(() => new Map(visibleCols.map((col, index) => [col, index])), [visibleCols]);
  const visibleRowsRef = React.useRef<number[]>(visibleRows);
  const visibleColsRef = React.useRef<number[]>(visibleCols);
  const effectiveRowHeightsRef = React.useRef<number[]>(effectiveRowHeights);
  const effectiveColWidthsRef = React.useRef<number[]>(effectiveColWidths);
  const rowPrefixSums = React.useMemo(() => buildPrefixSums(effectiveRowHeights), [effectiveRowHeights]);
  const colPrefixSums = React.useMemo(() => buildPrefixSums(effectiveColWidths), [effectiveColWidths]);
  const actualRowPrefixSums = React.useMemo(() => buildPrefixSums(actualRowHeights), [actualRowHeights]);
  const actualColPrefixSums = React.useMemo(() => buildPrefixSums(actualColWidths), [actualColWidths]);
  const stickyTopByRow = React.useMemo(
    () => buildStickyOffsets(frozenRows, actualRowHeights, HEADER_HEIGHT),
    [actualRowHeights, frozenRows]
  );
  const stickyLeftByCol = React.useMemo(
    () => buildStickyOffsets(frozenCols, actualColWidths, ROW_HEADER_WIDTH),
    [actualColWidths, frozenCols]
  );
  const frozenPaneBottom = React.useMemo(
    () => (
      frozenRows.length > 0
        ? frozenRows.reduce(
            (max, row) => Math.max(max, (stickyTopByRow.get(row) ?? HEADER_HEIGHT) + (actualRowHeights[row] ?? DEFAULT_ROW_HEIGHT)),
            HEADER_HEIGHT
          )
        : HEADER_HEIGHT
    ),
    [actualRowHeights, frozenRows, stickyTopByRow]
  );
  const frozenPaneRight = React.useMemo(
    () => (
      frozenCols.length > 0
        ? frozenCols.reduce(
            (max, col) => Math.max(max, (stickyLeftByCol.get(col) ?? ROW_HEADER_WIDTH) + (actualColWidths[col] ?? DEFAULT_COL_WIDTH)),
            ROW_HEADER_WIDTH
          )
        : ROW_HEADER_WIDTH
    ),
    [actualColWidths, frozenCols, stickyLeftByCol]
  );
  const useNativeZoomForFrozenPanes = zoomFactor !== 1 && (frozenRows.length > 0 || frozenCols.length > 0);
  const rowPrefixSumsRef = React.useRef<number[]>(rowPrefixSums);
  const colPrefixSumsRef = React.useRef<number[]>(colPrefixSums);
  const firstVisibleRow = visibleRows[0];
  const lastVisibleRow = visibleRows[visibleRows.length - 1];
  const firstVisibleCol = visibleCols[0];
  const lastVisibleCol = visibleCols[visibleCols.length - 1];
  const displayedSelection = fillPreviewRange ?? normalizedSelection;
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
  const frozenRowCount = activeSheet?.freezePanes?.row ?? 0;
  const shouldVirtualizeRows = !activeSheet?.hasVerticalMerges
    || (activeSheet.maxVerticalMergeEndRow >= 0 && activeSheet.maxVerticalMergeEndRow < frozenRowCount);
  const shouldVirtualizeCols = !activeSheet?.hasHorizontalMerges && frozenCols.length === 0;

  const rowVirtualizer = useVirtualizer({
    count: visibleRows.length,
    estimateSize: (index) => effectiveRowHeights[index] ?? DEFAULT_ROW_HEIGHT,
    getScrollElement: () => scrollRef.current,
    getItemKey: (index) => visibleRows[index] ?? index,
    observeElementOffset: (instance, cb) => observeZoomedElementOffset(instance, zoomFactor, cb),
    observeElementRect: (instance, cb) => observeZoomedElementRect(instance, zoomFactor, cb),
    overscan: 10,
    scrollToFn: (offset, options, instance) => scrollToZoomedOffset(offset, zoomFactor, options, instance)
  });
  const colVirtualizer = useVirtualizer({
    count: visibleCols.length,
    estimateSize: (index) => effectiveColWidths[index] ?? DEFAULT_COL_WIDTH,
    getScrollElement: () => scrollRef.current,
    getItemKey: (index) => visibleCols[index] ?? index,
    horizontal: true,
    observeElementOffset: (instance, cb) => observeZoomedElementOffset(instance, zoomFactor, cb),
    observeElementRect: (instance, cb) => observeZoomedElementRect(instance, zoomFactor, cb),
    overscan: 6,
    scrollToFn: (offset, options, instance) => scrollToZoomedOffset(offset, zoomFactor, options, instance)
  });
  const currentRowVirtualItems = rowVirtualizer.getVirtualItems();
  const currentColVirtualItems = colVirtualizer.getVirtualItems();
  const frozenRowVirtualIndices = React.useMemo(
    () => frozenRows
      .map((row) => rowIndexByActual.get(row))
      .filter((index): index is number => index !== undefined),
    [frozenRows, rowIndexByActual]
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
    selectionRef.current = normalizedSelection;
  }, [normalizedSelection]);

  React.useEffect(() => {
    editingCellRef.current = editingCell;
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
    if (imagePreviewFrameRef.current !== null) {
      window.cancelAnimationFrame(imagePreviewFrameRef.current);
      imagePreviewFrameRef.current = null;
    }
    if (chartPreviewFrameRef.current !== null) {
      window.cancelAnimationFrame(chartPreviewFrameRef.current);
      chartPreviewFrameRef.current = null;
    }
  }, []);

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
    if (!previewRange || selectionDragRef.current || fillDragRef.current) {
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
    effectiveRowHeightsRef.current = effectiveRowHeights;
    effectiveColWidthsRef.current = effectiveColWidths;
    rowPrefixSumsRef.current = rowPrefixSums;
    colPrefixSumsRef.current = colPrefixSums;
  }, [colPrefixSums, effectiveColWidths, effectiveRowHeights, rowPrefixSums, visibleCols, visibleRows]);

  React.useLayoutEffect(() => {
    const scroller = scrollRef.current;
    const previousZoomFactor = previousZoomFactorRef.current;
    if (!scroller || previousZoomFactor === zoomFactor) {
      previousZoomFactorRef.current = zoomFactor;
      return;
    }

    scroller.scrollLeft = (scroller.scrollLeft / previousZoomFactor) * zoomFactor;
    scroller.scrollTop = (scroller.scrollTop / previousZoomFactor) * zoomFactor;
    previousZoomFactorRef.current = zoomFactor;
    rowVirtualizer.measure();
    colVirtualizer.measure();
    syncDrawingViewport(scroller);
  }, [colVirtualizer, rowVirtualizer, syncDrawingViewport, zoomFactor]);

  React.useLayoutEffect(() => {
    syncDrawingViewport(scrollRef.current);
  }, [activeSheet, activeTabIndex, displayColLimit, displayRowLimit, syncDrawingViewport, zoomFactor]);

  React.useLayoutEffect(() => {
    const scroller = scrollRef.current;
    if (!scroller) {
      return;
    }
    syncDrawingViewport(scroller);
    if (typeof ResizeObserver === "undefined") {
      return;
    }
    const observer = new ResizeObserver(() => {
      syncDrawingViewport(scroller);
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
    if (shouldVirtualizeRows) {
      rowVirtualizer.scrollToOffset(0);
    } else if (scrollRef.current) {
      scrollRef.current.scrollTop = 0;
    }
    if (scrollRef.current) {
      scrollRef.current.scrollLeft = 0;
    }
  }, [activeSheetIndex, rowVirtualizer, shouldVirtualizeRows]);

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
    if (shouldVirtualizeRows) {
      rowVirtualizer.measure();
    }
    if (shouldVirtualizeCols) {
      colVirtualizer.measure();
    }
  }, [activeSheetIndex, colVirtualizer, revision, rowVirtualizer, shouldVirtualizeCols, shouldVirtualizeRows, visibleCols.length, visibleRows.length]);

  const handleScrollerScroll = React.useCallback((event: React.UIEvent<HTMLDivElement>) => {
    const currentScroller = event.currentTarget;
    cachedScrollerRectRef.current = null;
    syncDrawingViewport(currentScroller);
    if (
      (currentScroller.scrollHeight - (currentScroller.scrollTop + currentScroller.clientHeight)) / zoomFactor <
      OPEN_GRID_VERTICAL_EDGE_PX
    ) {
      setDisplayRowLimit((current) => current + OPEN_GRID_ROW_GROWTH);
    }

    if (
      (currentScroller.scrollWidth - (currentScroller.scrollLeft + currentScroller.clientWidth)) / zoomFactor <
      OPEN_GRID_HORIZONTAL_EDGE_PX
    ) {
      setDisplayColLimit((current) => current + OPEN_GRID_COL_GROWTH);
    }
  }, [syncDrawingViewport, zoomFactor]);

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

  const resolvePointerCellFromClient = React.useCallback((clientX: number, clientY: number): XlsxCellAddress | null => {
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

    const resolveCellFromHitTest = () => {
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
    };
    const pointerOffsetX = clientX - scrollerRect.left;
    const pointerOffsetY = clientY - scrollerRect.top;
    const localX =
      (pointerOffsetX / zoomFactor)
      + (pointerOffsetX >= frozenPaneRight * zoomFactor ? (scroller.scrollLeft / zoomFactor) : 0);
    const localY =
      (pointerOffsetY / zoomFactor)
      + (pointerOffsetY >= frozenPaneBottom * zoomFactor ? (scroller.scrollTop / zoomFactor) : 0);
    const rowContentOffset = localY - HEADER_HEIGHT;
    const colContentOffset = localX - ROW_HEADER_WIDTH;
    let geometryCell: XlsxCellAddress | null = null;

    if (localY >= HEADER_HEIGHT && localX < ROW_HEADER_WIDTH) {
      const rowIndex = findIndexForOffsetPrefix(rowPrefixSumsCurrent, rowContentOffset);
      const actualRow = visibleRowsCurrent[rowIndex];
      if (actualRow !== undefined && firstVisibleColRef.current !== undefined) {
        geometryCell = { row: actualRow, col: firstVisibleColRef.current };
      }
    } else if (localY < HEADER_HEIGHT && localX >= ROW_HEADER_WIDTH) {
      const colIndex = findIndexForOffsetPrefix(colPrefixSumsCurrent, colContentOffset);
      const actualCol = visibleColsCurrent[colIndex];
      if (actualCol !== undefined && firstVisibleRowRef.current !== undefined) {
        geometryCell = { row: firstVisibleRowRef.current, col: actualCol };
      }
    } else if (localY >= HEADER_HEIGHT && localX >= ROW_HEADER_WIDTH) {
      const rowIndex = findIndexForOffsetPrefix(rowPrefixSumsCurrent, rowContentOffset);
      const colIndex = findIndexForOffsetPrefix(colPrefixSumsCurrent, colContentOffset);
      const actualRow = visibleRowsCurrent[rowIndex];
      const actualCol = visibleColsCurrent[colIndex];
      if (actualRow !== undefined && actualCol !== undefined) {
        geometryCell = { row: actualRow, col: actualCol };
      }
    }

    if (geometryCell && !worksheet?.isMergedSecondary(geometryCell.row, geometryCell.col)) {
      return geometryCell;
    }

    return resolveCellFromHitTest() ?? geometryCell;
  }, [frozenPaneBottom, frozenPaneRight, worksheet, zoomFactor]);

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
    const visibleRowsCurrent = visibleRowsRef.current;
    const visibleColsCurrent = visibleColsRef.current;
    const rowPrefixSumsCurrent = rowPrefixSumsRef.current;
    const colPrefixSumsCurrent = colPrefixSumsRef.current;
    if (visibleRowsCurrent.length === 0 || visibleColsCurrent.length === 0) {
      return null;
    }

    const maxContentY = rowPrefixSumsCurrent[rowPrefixSumsCurrent.length - 1] ?? 0;
    const maxContentX = colPrefixSumsCurrent[colPrefixSumsCurrent.length - 1] ?? 0;
    const currentContentX = clampContentOffset(
      dragState.originContentX + ((clientX - dragState.startClientX) / dragState.contentScaleX),
      maxContentX
    );
    const currentContentY = clampContentOffset(
      dragState.originContentY + ((clientY - dragState.startClientY) / dragState.contentScaleY),
      maxContentY
    );

    if (dragState.axis === "row") {
      const rowIndex = findIndexForOffsetPrefix(rowPrefixSumsCurrent, currentContentY);
      const actualRow = visibleRowsCurrent[rowIndex];
      return actualRow !== undefined ? { row: actualRow, col: dragState.originCell.col } : null;
    }

    if (dragState.axis === "column") {
      const colIndex = findIndexForOffsetPrefix(colPrefixSumsCurrent, currentContentX);
      const actualCol = visibleColsCurrent[colIndex];
      return actualCol !== undefined ? { row: dragState.originCell.row, col: actualCol } : null;
    }

    const rowIndex = findIndexForOffsetPrefix(rowPrefixSumsCurrent, currentContentY);
    const colIndex = findIndexForOffsetPrefix(colPrefixSumsCurrent, currentContentX);
    const actualRow = visibleRowsCurrent[rowIndex];
    const actualCol = visibleColsCurrent[colIndex];
    if (actualRow === undefined || actualCol === undefined) {
      return null;
    }

    if (worksheet?.isMergedSecondary(actualRow, actualCol)) {
      return resolvePointerCellFromHitTest(clientX, clientY) ?? { row: actualRow, col: actualCol };
    }

    return { row: actualRow, col: actualCol };
  }, [resolvePointerCellFromHitTest, worksheet]);

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

    const logicalWidth = effectiveColWidths[colIndex] ?? DEFAULT_COL_WIDTH;
    const logicalHeight = effectiveRowHeights[rowIndex] ?? DEFAULT_ROW_HEIGHT;
    const contentScaleX = Math.max(0.0001, rect.width > 0 ? rect.width / logicalWidth : zoomFactor);
    const contentScaleY = Math.max(0.0001, rect.height > 0 ? rect.height / logicalHeight : zoomFactor);

    return {
      contentScaleX,
      contentScaleY,
      originContentX:
        (colPrefixSums[colIndex] ?? 0)
        + clampContentOffset((clientX - rect.left) / contentScaleX, logicalWidth),
      originContentY:
        (rowPrefixSums[rowIndex] ?? 0)
        + clampContentOffset((clientY - rect.top) / contentScaleY, logicalHeight)
    };
  }, [colIndexByActual, colPrefixSums, effectiveColWidths, effectiveRowHeights, rowIndexByActual, rowPrefixSums, zoomFactor]);

  const resolveRowPointerOrigin = React.useCallback((
    actualRow: number,
    rect: DOMRect,
    clientY: number
  ) => {
    const rowIndex = rowIndexByActual.get(actualRow);
    if (rowIndex === undefined) {
      return null;
    }

    const logicalHeight = effectiveRowHeights[rowIndex] ?? DEFAULT_ROW_HEIGHT;
    const contentScaleY = Math.max(0.0001, rect.height > 0 ? rect.height / logicalHeight : zoomFactor);
    return {
      contentScaleX: Math.max(0.0001, zoomFactor),
      contentScaleY,
      originContentX: colPrefixSums[0] ?? 0,
      originContentY:
        (rowPrefixSums[rowIndex] ?? 0)
        + clampContentOffset((clientY - rect.top) / contentScaleY, logicalHeight)
    };
  }, [colPrefixSums, effectiveRowHeights, rowIndexByActual, rowPrefixSums, zoomFactor]);

  const resolveColumnPointerOrigin = React.useCallback((
    actualCol: number,
    rect: DOMRect,
    clientX: number
  ) => {
    const colIndex = colIndexByActual.get(actualCol);
    if (colIndex === undefined) {
      return null;
    }

    const logicalWidth = effectiveColWidths[colIndex] ?? DEFAULT_COL_WIDTH;
    const contentScaleX = Math.max(0.0001, rect.width > 0 ? rect.width / logicalWidth : zoomFactor);
    return {
      contentScaleX,
      contentScaleY: Math.max(0.0001, zoomFactor),
      originContentX:
        (colPrefixSums[colIndex] ?? 0)
        + clampContentOffset((clientX - rect.left) / contentScaleX, logicalWidth),
      originContentY: rowPrefixSums[0] ?? 0
    };
  }, [colIndexByActual, colPrefixSums, effectiveColWidths, rowPrefixSums, zoomFactor]);

  const applyColumnPreview = React.useCallback((actualCol: number, widthPx: number | null) => {
    const colElement = colElementRefs.current.get(actualCol);
    if (colElement) {
      colElement.style.width = widthPx === null ? "" : `${widthPx}px`;
    }

    const baseIndex = visibleCols.indexOf(actualCol);
    const baseWidth = baseIndex >= 0 ? (effectiveColWidths[baseIndex] ?? DEFAULT_COL_WIDTH) : DEFAULT_COL_WIDTH;
    const previewWidth = widthPx ?? baseWidth;
    const baseTotalWidth = effectiveColWidths.reduce((sum, width) => sum + width, 0) + ROW_HEADER_WIDTH;
    const widthDelta = previewWidth - baseWidth;
    if (tableRef.current) {
      tableRef.current.style.width = `${baseTotalWidth + widthDelta}px`;
    }
  }, [effectiveColWidths, visibleCols]);

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
    selectionDragCleanupRef.current?.();
    fillDragCleanupRef.current?.();
    chartInteractionCleanupRef.current?.();
    imageInteractionCleanupRef.current?.();
    selectionDragCleanupRef.current = null;
    fillDragCleanupRef.current = null;
    chartInteractionCleanupRef.current = null;
    imageInteractionCleanupRef.current = null;
    selectionDragRef.current = null;
    fillDragRef.current = null;
    chartInteractionRef.current = null;
    imageInteractionRef.current = null;
    pendingResizePreviewRef.current = null;
    if (columnPreviewRef.current) {
      applyColumnPreview(columnPreviewRef.current.actualIndex, null);
    }
    if (rowPreviewRef.current) {
      applyRowPreview(rowPreviewRef.current.actualIndex, null);
    }
    columnPreviewRef.current = null;
    rowPreviewRef.current = null;
    setEditingCell(null);
    setEditingValue("");
    setFillPreviewRange(null);
    chartPreviewRectRef.current = null;
    setChartPreviewRect(null);
    imagePreviewRectRef.current = null;
    setImagePreviewRect(null);
    selectionPreviewRangeRef.current = null;
    setInteractionMode("idle");
  }, [activeSheetIndex, revision]);

  const focusGrid = React.useCallback(() => {
    scrollRef.current?.focus();
  }, []);

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
      setEditingCell(null);
      setEditingValue("");
      focusGrid();
      return;
    }

    setCellValue(editingCell, editingValue);
    setEditingCell(null);
    setEditingValue("");
    focusGrid();
  }, [editingCell, editingValue, focusGrid, readOnly, setCellValue]);

  const cancelEditing = React.useCallback(() => {
    setEditingCell(null);
    setEditingValue("");
    focusGrid();
  }, [focusGrid]);

  const [, startBatchTransition] = React.useTransition();
  const [, startSelectionTransition] = React.useTransition();
  const [asyncViewportRowBatch, setAsyncViewportRowBatch] = React.useState<WorksheetBatchWindow | null>(null);
  const viewportRequest = React.useMemo(() => {
    const firstVirtualRowIndex = currentRowVirtualItems[0]?.index;
    const lastVirtualRowIndex = currentRowVirtualItems[currentRowVirtualItems.length - 1]?.index;
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
  }, [currentRowVirtualItems, visibleRows]);

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
    cellRenderCacheRef.current.clear();
  }, [activeSheetIndex, displayColLimit, displayRowLimit, palette, revision, viewportRowBatch, worksheet]);

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

  const getCellData = React.useCallback((row: number, col: number): CellRenderData => {
    const cacheKey = `${row}:${col}`;
    const cached = cellRenderCacheRef.current.get(cacheKey);
    if (cached) {
      return cached;
    }

    const batchCoversRow = viewportRowBatch ? row >= viewportRowBatch.startRow && row <= viewportRowBatch.endRow : false;
    const batchedCell = batchCoversRow ? viewportRowBatch?.cells.get(cacheKey) : undefined;

    if (!worksheet && !batchedCell) {
      const emptyData: CellRenderData = {
        isMergedSecondary: false,
        style: {
          backgroundColor: resolveSheetSurface(activeSheet, palette),
          borderBottom: "none",
          borderRight: "none",
          boxShadow: activeSheet?.showGridLines ? buildGridlineShadow(palette.border) : undefined,
          padding: DEFAULT_CELL_PADDING,
          verticalAlign: "bottom",
          whiteSpace: "nowrap"
        },
        value: ""
      };
      cellRenderCacheRef.current.set(cacheKey, emptyData);
      return emptyData;
    }

    if (batchedCell?.isMergedSecondary) {
      const mergedSecondaryData: CellRenderData = {
        isMergedSecondary: true,
        style: {},
        value: ""
      };
      cellRenderCacheRef.current.set(cacheKey, mergedSecondaryData);
      return mergedSecondaryData;
    }

    if (!batchedCell && !batchCoversRow && worksheet?.isMergedSecondary(row, col)) {
      const mergedSecondaryData: CellRenderData = {
        isMergedSecondary: true,
        style: {},
        value: ""
      };
      cellRenderCacheRef.current.set(cacheKey, mergedSecondaryData);
      return mergedSecondaryData;
    }

    const merge = batchCoversRow
      ? batchedCell?.mergeSpan
      : (worksheet?.getMergeSpan(row, col) as { colSpan?: number; rowSpan?: number } | null | undefined);
    const inheritedStyle = resolveInheritedCellStyle(activeSheet, row, col);
    const worksheetStyle = batchCoversRow
      ? batchedCell?.style ?? null
      : (worksheet?.getCellStyleAt(row, col) as Record<string, unknown> | null | undefined) ?? null;
    const rawStyle = mergeResolvedCellStyle(inheritedStyle, worksheetStyle);
    const table = getTableAtCell(effectiveTables, row, col);
    const tableStyle = resolveTableCellStyle(table, row, col, activeSheet);
    const mergedStyle = mergeResolvedCellStyle(rawStyle, tableStyle);
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
    const nextData: CellRenderData = {
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
      isTableHeader: Boolean(table && row >= table.start.row && row < table.start.row + headerRowCount),
      rowSpan: merge?.rowSpan,
      sparkline: sparkline && sparklineValues ? { config: sparkline, values: sparklineValues } : null,
      style: buildCellStyle(mergedStyle, palette, activeSheet?.themePalette, {
        showGridLines: activeSheet?.showGridLines
      }),
      validation: resolveCellDataValidation(row, col, activeSheet),
      value: sparkline
        ? ""
        : checkboxState !== null
          ? ""
          : batchCoversRow || !worksheet
            ? batchedCell?.value ?? ""
            : getCellDisplayValue(worksheet, row, col, activeSheet)
    };

    if (canCellTextOverflow(nextData)) {
      const startColIndex = colIndexByActual.get(col);
      if (startColIndex !== undefined) {
        const horizontalPadding = getHorizontalPadding(nextData.style.padding);
        const textWidth = measureTextWidth(nextData.value, nextData.style);
        const requiredWidth = textWidth + horizontalPadding + 2;
        let availableWidth = effectiveColWidths[startColIndex] ?? DEFAULT_COL_WIDTH;

        if (requiredWidth > availableWidth) {
          for (let nextColIndex = startColIndex + 1; nextColIndex < visibleCols.length; nextColIndex += 1) {
            const nextActualCol = visibleCols[nextColIndex];
            if (nextActualCol === undefined) {
              break;
            }

            const neighborData = getCellData(row, nextActualCol);
            if (!canReceiveOverflowText(neighborData)) {
              break;
            }

            availableWidth += effectiveColWidths[nextColIndex] ?? DEFAULT_COL_WIDTH;
            if (requiredWidth <= availableWidth) {
              break;
            }
          }

          if (availableWidth > (effectiveColWidths[startColIndex] ?? DEFAULT_COL_WIDTH)) {
            nextData.spillWidth = Math.max(0, availableWidth - horizontalPadding);
          }
        }
      }
    }

    cellRenderCacheRef.current.set(cacheKey, nextData);
    return nextData;
  }, [activeSheet, colIndexByActual, effectiveColWidths, effectiveTables, palette, sparklinesByCell, viewportRowBatch, visibleCols, worksheet]);

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
      left: ROW_HEADER_WIDTH + sumPrefixRange(colPrefixSums, 0, startColIndex - 1),
      top: HEADER_HEIGHT + sumPrefixRange(rowPrefixSums, 0, startRowIndex - 1),
      width: sumPrefixRange(colPrefixSums, startColIndex, endColIndex)
    };
  }, [colIndexByActual, colPrefixSums, displayedSelection, rowIndexByActual, rowPrefixSums]);
  const resolvedSelectionOverlay = selectionOverlay;
  const virtualCols = React.useMemo<RenderedAxisItem[]>(
    () =>
      shouldVirtualizeCols
        ? colVirtualizer.getVirtualItems().map((virtualCol) => ({
            end: virtualCol.end,
            index: virtualCol.index,
            key: virtualCol.key,
            size: virtualCol.size,
            start: virtualCol.start
          }))
        : visibleCols.map((_, index) => ({
            end: colPrefixSums[index + 1] ?? 0,
            index,
            key: visibleCols[index] ?? index,
            size: effectiveColWidths[index] ?? DEFAULT_COL_WIDTH,
            start: colPrefixSums[index] ?? 0
          })),
    [colPrefixSums, colVirtualizer, effectiveColWidths, shouldVirtualizeCols, visibleCols]
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
    [effectiveColWidths, virtualCols, visibleCols]
  );
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
                : resolveImageRect(image, visibleRows, visibleCols, effectiveRowHeights, effectiveColWidths, {
                    actualColPrefixSums,
                    actualRowPrefixSums,
                    colIndexByActual,
                    colPrefixSums,
                    rowIndexByActual,
                    rowPrefixSums
                  })
          }))
        : [],
    [
      colIndexByActual,
      colPrefixSums,
      actualColPrefixSums,
      actualRowPrefixSums,
      effectiveColWidths,
      effectiveRowHeights,
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
            rect: resolveAnchoredRect(shape.anchor, visibleRows, visibleCols, effectiveRowHeights, effectiveColWidths, {
              actualColPrefixSums,
              actualRowPrefixSums,
              colIndexByActual,
              colPrefixSums,
              rowIndexByActual,
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
      effectiveColWidths,
      effectiveRowHeights,
      rowIndexByActual,
      rowPrefixSums,
      shapes,
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
                : resolveAnchoredRect(chart.anchor, visibleRows, visibleCols, effectiveRowHeights, effectiveColWidths, {
                    actualColPrefixSums,
                    actualRowPrefixSums,
                    colIndexByActual,
                    colPrefixSums,
                    rowIndexByActual,
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
      effectiveColWidths,
      effectiveRowHeights,
      rowIndexByActual,
      rowPrefixSums,
      showImages,
      visibleCols,
      visibleRows
    ]
  );

  const resolveOverlayRect = React.useCallback((range: XlsxCellRange) => {
    const normalized = normalizeRange(range);
    const startRowIndex = rowIndexByActual.get(normalized.start.row);
    const endRowIndex = rowIndexByActual.get(normalized.end.row);
    const startColIndex = colIndexByActual.get(normalized.start.col);
    const endColIndex = colIndexByActual.get(normalized.end.col);

    if (
      startRowIndex !== undefined &&
      endRowIndex !== undefined &&
      startColIndex !== undefined &&
      endColIndex !== undefined
    ) {
      let left = ROW_HEADER_WIDTH + sumPrefixRange(colPrefixSums, 0, startColIndex - 1);
      let top = HEADER_HEIGHT + sumPrefixRange(rowPrefixSums, 0, startRowIndex - 1);
      let width = sumPrefixRange(colPrefixSums, startColIndex, endColIndex);
      let height = sumPrefixRange(rowPrefixSums, startRowIndex, endRowIndex);

      const columnPreview = columnPreviewRef.current;
      if (columnPreview) {
        const previewIndex = colIndexByActual.get(columnPreview.actualIndex);
        if (previewIndex !== undefined) {
          const baseWidth = effectiveColWidths[previewIndex] ?? DEFAULT_COL_WIDTH;
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
          const baseHeight = effectiveRowHeights[previewIndex] ?? DEFAULT_ROW_HEIGHT;
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
        left: Math.max(ROW_HEADER_WIDTH, left),
        top: Math.max(HEADER_HEIGHT, top),
        width: Math.max(0, width)
      };
    }

    return null;
  }, [colIndexByActual, colPrefixSums, effectiveColWidths, effectiveRowHeights, rowIndexByActual, rowPrefixSums]);

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
    if (!cell) {
      return null;
    }

    return {
      column,
      left: cell.offsetLeft + cell.offsetWidth - 170,
      table,
      top: cell.offsetTop + cell.offsetHeight - 2
    };
  }, [effectiveTables, openTableMenu]);

  const applyPreviewOverlay = React.useCallback((range: XlsxCellRange | null) => {
    const overlay = selectionOverlayRef.current;
    if (!overlay || !range) {
      return;
    }

    const nextRect = resolveOverlayRect(range);
    if (!nextRect) {
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
      fillHandle.style.left = `${nextRect.left + nextRect.width - 4}px`;
      fillHandle.style.top = `${nextRect.top + nextRect.height - 4}px`;
    }
  }, [resolveOverlayRect]);

  const applyPreviewOverlayForCell = React.useCallback((cell: XlsxCellAddress) => {
    const rowIndex = rowIndexByActual.get(cell.row);
    const colIndex = colIndexByActual.get(cell.col);
    if (rowIndex === undefined || colIndex === undefined) {
      return;
    }

    const nextRect = {
      height: effectiveRowHeights[rowIndex] ?? DEFAULT_ROW_HEIGHT,
      left: ROW_HEADER_WIDTH + (colPrefixSums[colIndex] ?? 0),
      top: HEADER_HEIGHT + (rowPrefixSums[rowIndex] ?? 0),
      width: effectiveColWidths[colIndex] ?? DEFAULT_COL_WIDTH
    };

    const overlay = selectionOverlayRef.current;
    if (!overlay) {
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
      fillHandle.style.left = `${nextRect.left + nextRect.width - 4}px`;
      fillHandle.style.top = `${nextRect.top + nextRect.height - 4}px`;
    }
  }, [colIndexByActual, colPrefixSums, effectiveColWidths, effectiveRowHeights, rowIndexByActual, rowPrefixSums]);

  const commitSelectionRange = React.useCallback((range: XlsxCellRange) => {
    const normalized = normalizeRange(range);
    if (
      normalizedSelection &&
      rangesEqual(normalizedSelection, normalized) &&
      isSameCell(activeCell, normalized.end) &&
      selectedChartId === null &&
      selectedImageId === null
    ) {
      return;
    }

    startSelectionTransition(() => {
      selectRange(normalized);
    });
  }, [activeCell, normalizedSelection, selectRange, selectedChartId, selectedImageId, startSelectionTransition]);

  const refreshOverlayFromCurrentSelection = React.useCallback(() => {
    if (displayedSelectionRef.current) {
      applyPreviewOverlay(displayedSelectionRef.current);
    }
  }, [applyPreviewOverlay]);

  React.useEffect(() => {
    function handlePointerMove(event: PointerEvent) {
      const state = resizeStateRef.current;
      if (!state) {
        return;
      }

      pendingResizePreviewRef.current = {
        actualIndex: state.actualIndex,
        size:
          state.type === "column"
            ? Math.max(DEFAULT_COL_WIDTH / 2, state.initialPx + ((event.clientX - state.startPosition) / zoomFactor))
            : Math.max(DEFAULT_ROW_HEIGHT / 1.5, state.initialPx + ((event.clientY - state.startPosition) / zoomFactor)),
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
          applyColumnPreview(preview.actualIndex, preview.size);
        } else {
          rowPreviewRef.current = { actualIndex: preview.actualIndex, size: preview.size };
          applyRowPreview(preview.actualIndex, preview.size);
        }
        refreshOverlayFromCurrentSelection();
      });
    }

    function handlePointerUp(event: PointerEvent) {
      if (resizeStateRef.current?.pointerId === event.pointerId) {
        const resizeState = resizeStateRef.current;
        const preview = pendingResizePreviewRef.current;
        resizeStateRef.current = null;
        pendingResizePreviewRef.current = null;
        setInteractionMode("idle");
        document.body.style.cursor = "";
        document.body.style.userSelect = "";

        if (resizeFrameRef.current !== null) {
          window.cancelAnimationFrame(resizeFrameRef.current);
          resizeFrameRef.current = null;
        }

        if (resizeState.type === "column") {
          applyColumnPreview(resizeState.actualIndex, null);
          columnPreviewRef.current = null;
        } else {
          applyRowPreview(resizeState.actualIndex, null);
          rowPreviewRef.current = null;
          if (shouldVirtualizeRows) {
            rowVirtualizer.measure();
          }
        }
        if (preview && preview.actualIndex === resizeState.actualIndex && preview.type === resizeState.type) {
          if (preview.type === "column") {
            controller.resizeColumn(preview.actualIndex, preview.size);
          } else {
            controller.resizeRow(preview.actualIndex, preview.size);
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
      if (resizeFrameRef.current !== null) {
        window.cancelAnimationFrame(resizeFrameRef.current);
        resizeFrameRef.current = null;
      }
    };
  }, [applyColumnPreview, applyRowPreview, controller, refreshOverlayFromCurrentSelection, rowVirtualizer, shouldVirtualizeRows, zoomFactor]);

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

      const nextCell = resolveDraggedSelectionCell(dragState, pendingPoint.x, pendingPoint.y);
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
      if (selectionOverlayRef.current) {
        selectionOverlayRef.current.style.opacity = "0";
        selectionOverlayRef.current.style.visibility = "hidden";
      }
      return;
    }

    applyPreviewOverlay(overlayRange);
  }, [applyPreviewOverlay, displayedSelection, revision]);

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
    cell: XlsxCellAddress,
    isActive: boolean
  ) => {
    if (event.button !== 0) {
      return;
    }

    event.preventDefault();
    focusGrid();
    const anchor = event.shiftKey && selectionRef.current ? selectionRef.current.start : cell;
    const initialRange = normalizeRange({ start: anchor, end: cell });
    const committedOnPointerDown = !isActive || !editingCellRef.current;
    const pointerOrigin = resolveCellPointerOrigin(cell, event.currentTarget.getBoundingClientRect(), event.clientX, event.clientY);
    if (!pointerOrigin) {
      return;
    }

    startCellSelection(
      event.pointerId,
      anchor,
      "cell",
      cell,
      pointerOrigin,
      committedOnPointerDown,
      initialRange,
      event.clientX,
      event.clientY
    );
    applyPreviewOverlayForCell(cell);
    if (committedOnPointerDown) {
      commitSelectionRange(initialRange);
    }
  }, [applyPreviewOverlayForCell, commitSelectionRange, focusGrid, resolveCellPointerOrigin]);

  const handleRowPointerDown = React.useCallback((
    event: React.PointerEvent<HTMLTableCellElement>,
    actualRow: number
  ) => {
    if (event.button !== 0 || firstVisibleCol === undefined || lastVisibleCol === undefined) {
      return;
    }

    event.preventDefault();
    focusGrid();
    const anchorRow = event.shiftKey && selectionRef.current ? selectionRef.current.start.row : actualRow;
    const initialRange = normalizeRange({
      start: { row: anchorRow, col: firstVisibleCol },
      end: { row: actualRow, col: lastVisibleCol }
    });
    const pointerOrigin = resolveRowPointerOrigin(actualRow, event.currentTarget.getBoundingClientRect(), event.clientY);
    if (!pointerOrigin) {
      return;
    }

    startCellSelection(
      event.pointerId,
      { row: anchorRow, col: firstVisibleCol },
      "row",
      { row: actualRow, col: firstVisibleCol },
      pointerOrigin,
      true,
      initialRange,
      event.clientX,
      event.clientY
    );
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
    const anchorCol =
      event.shiftKey && normalizedSelection ? normalizedSelection.start.col : actualCol;
    const initialRange = normalizeRange({
      start: { row: firstVisibleRow, col: anchorCol },
      end: { row: lastVisibleRow, col: actualCol }
    });
    const pointerOrigin = resolveColumnPointerOrigin(actualCol, event.currentTarget.getBoundingClientRect(), event.clientX);
    if (!pointerOrigin) {
      return;
    }

    startCellSelection(
      event.pointerId,
      { row: firstVisibleRow, col: anchorCol },
      "column",
      { row: firstVisibleRow, col: actualCol },
      pointerOrigin,
      true,
      initialRange,
      event.clientX,
      event.clientY
    );
    commitSelectionRange(initialRange);
  }, [
    commitSelectionRange,
    firstVisibleRow,
    focusGrid,
    lastVisibleRow,
    normalizedSelection,
    resolveColumnPointerOrigin
  ]);

  const handleRowResizePointerDown = React.useCallback((
    event: React.PointerEvent<HTMLDivElement>,
    actualRow: number,
    rowHeight: number
  ) => {
    if (readOnlyRef.current) {
      return;
    }

    event.preventDefault();
    event.stopPropagation();
    startRowResize(event.pointerId, actualRow, rowHeight, event.clientY);
  }, []);

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
          alignItems: "center",
          background: "transparent",
          border: "none",
          color: palette.mutedText,
          cursor: "pointer",
          display: "inline-flex",
          fontSize: 10,
          height: 16,
          justifyContent: "center",
          padding: 0,
          position: "absolute",
          right: 4,
          top: 3,
          width: 16,
          zIndex: 6
        }}
        type="button"
      >
        {direction === "ascending" ? "▲" : direction === "descending" ? "▼" : "▾"}
      </button>
    );
  }, [effectiveTables, palette.mutedText, sortState]);

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

  if (isLoading) {
    return <>{renderLoading(loadingComponent, loadingState, palette)}</>;
  }

  if (isLoadDeferred) {
    return <>{renderDeferredLoad({ ...controller, canLoadDeferred, continueDeferredLoad, deferredLoadFileSize, isLoadDeferred }, palette)}</>;
  }

  if (error) {
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
                  : <MemoChartSvg chart={chart} palette={palette} rect={chartsheetRect} />}
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

  const virtualRows = !shouldVirtualizeRows
    ? visibleRows.map((actualRow, index) => ({
        end: rowPrefixSums[index + 1] ?? 0,
        index,
        key: actualRow,
        size: effectiveRowHeights[index] ?? DEFAULT_ROW_HEIGHT,
        start: rowPrefixSums[index] ?? 0
      }))
    : (() => {
        const renderedRowsByIndex = new Map<number, RenderedAxisItem>();
        currentRowVirtualItems.forEach((virtualRow) => {
          renderedRowsByIndex.set(virtualRow.index, {
            end: virtualRow.end,
            index: virtualRow.index,
            key: visibleRows[virtualRow.index] ?? virtualRow.index,
            size: virtualRow.size,
            start: virtualRow.start
          });
        });

        frozenRowVirtualIndices.forEach((index) => {
          if (renderedRowsByIndex.has(index)) {
            return;
          }
          renderedRowsByIndex.set(index, {
            end: rowPrefixSums[index + 1] ?? 0,
            index,
            key: visibleRows[index] ?? index,
            size: effectiveRowHeights[index] ?? DEFAULT_ROW_HEIGHT,
            start: rowPrefixSums[index] ?? 0
          });
        });

        return Array.from(renderedRowsByIndex.values()).sort((left, right) => left.index - right.index);
      })();
  const totalHeight = shouldVirtualizeRows
    ? rowVirtualizer.getTotalSize()
    : (rowPrefixSums[rowPrefixSums.length - 1] ?? 0);
  const totalWidth = totalContentWidth + ROW_HEADER_WIDTH;
  const sheetContentHeight = HEADER_HEIGHT + totalHeight;
  const zoomedSheetHeight = sheetContentHeight * zoomFactor;
  const zoomedSheetWidth = totalWidth * zoomFactor;
  const { fill: selectionFill, header: selectionHeaderSurface, stroke: selectionStroke } = resolveSelectionColors({
    palette,
    selectionColor,
    selectionFillColor,
    selectionHeaderColor
  });
  const selectionBorderWidth = 1.5;
  const rowColSpan = renderedCols.length + 1 + (leadingColumnSpacerWidth > 0 ? 1 : 0) + (trailingColumnSpacerWidth > 0 ? 1 : 0);
  const gutterSeparatorShadow = `inset -1px 0 0 ${palette.border}, inset 0 -1px 0 ${palette.border}`;
  const headerCellStyle: React.CSSProperties = {
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
    zIndex: 50
  };
  const columnResizeHandleStyle: React.CSSProperties = {
    backgroundColor: "transparent",
    cursor: "col-resize",
    position: "absolute",
    right: -8,
    top: 0,
    width: 16,
    height: "100%",
    zIndex: 5
  };
  function resolveDrawingPane(rect: XlsxImageRect) {
    return resolveFrozenDrawingPane(
      rect,
      frozenRows,
      frozenCols,
      actualRowHeights,
      actualColWidths,
      activeSheet?.freezePanes ?? null,
      stickyTopByRow,
      stickyLeftByCol
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
    const textScale = strokeScale;
    const textWidth = groupScaleX !== 0 ? rect.width / groupScaleX : rect.width;
    const textHeight = groupScaleY !== 0 ? rect.height / groupScaleY : rect.height;
    const vectorShape = resolveShapeVector(shape);
    const strokeColor = shape.stroke?.none ? "transparent" : (shape.stroke?.color ?? "transparent");
    const scaledStrokeWidth = (shape.stroke?.widthPx ?? (shape.geometry === "line" ? 2 : 1)) * strokeScale;
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
      ...buildShapeContainerStyle(shape, rect),
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
            gap: 2,
            height: textHeight,
            justifyContent:
              shape.textBox?.verticalAlign === "middle"
                ? "center"
                : shape.textBox?.verticalAlign === "bottom"
                  ? "flex-end"
                  : "flex-start",
            paddingBottom: inset?.bottom ?? 4,
            paddingLeft: inset?.left ?? 6,
            paddingRight: inset?.right ?? 6,
            paddingTop: inset?.top ?? 4,
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
                    border: `1.5px solid ${selectionStroke}`,
                    boxShadow: `0 0 0 1px ${palette.surface}`,
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
                          style={resolveImageHandleStyle(position, selectionStroke, palette.surface)}
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
                  ? resolveImageHandleStyle(position, selectionStroke, palette.surface)
                  : { ...resolveImageHandleStyle(position, selectionStroke, palette.surface), display: "none" }
              }),
              image,
              rect
            })
          : (
              <div
                style={{
                  border: `1.5px solid ${selectionStroke}`,
                  boxShadow: `0 0 0 1px ${palette.surface}`,
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
                        style={resolveImageHandleStyle(position, selectionStroke, palette.surface)}
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
          overflow: "visible",
          pointerEvents: "none",
          zIndex: isFrozenDrawing ? chart.zIndex + 22 : chart.zIndex + 2
        }}
      >
        <div
          style={{
            border: `1.5px solid ${selectionStroke}`,
            boxShadow: `0 0 0 1px ${palette.surface}`,
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
                  style={resolveImageHandleStyle(position, selectionStroke, palette.surface)}
                />
              ))
            : null}
        </div>
      </div>
    ) : null;

    return (
      <React.Fragment key={`${pane}-${chart.id}`}>
        <div style={style}>
          {isChartsLoading
            ? renderChartLoadingNode(renderChartLoading, chart, rect)
            : <MemoChartSvg chart={chart} palette={palette} rect={rect} />}
        </div>
        <div
          onClick={() => handleChartClick(chart)}
          onPointerDown={(event) => startChartMove(event, chart, rect)}
          style={{
            ...style,
            background: "transparent",
            cursor: canEditChart && selectedChartId === chart.id ? "move" : "cell",
            pointerEvents: "auto",
            zIndex: isFrozenDrawing ? chart.zIndex + 21 : chart.zIndex + 1
          }}
        />
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
    zIndex: 40
  };
  const leftOverlayStyle: React.CSSProperties = {
    height: 0,
    left: 0,
    overflow: "visible",
    pointerEvents: "none",
    position: "sticky",
    width: 0,
    zIndex: 40
  };
  const cornerOverlayStyle: React.CSSProperties = {
    height: 0,
    left: 0,
    overflow: "visible",
    pointerEvents: "none",
    position: "sticky",
    top: 0,
    width: 0,
    zIndex: 40
  };
  const previousPaneDrawingNodes = paneDrawingNodesCacheRef.current;
  const canReusePaneDrawingNodes =
    previousPaneDrawingNodes !== null
    && previousPaneDrawingNodes.showImages === showImages
    && previousPaneDrawingNodes.chartRects === chartRects
    && previousPaneDrawingNodes.shapeRects === shapeRects
    && previousPaneDrawingNodes.imageRects === imageRects
    && previousPaneDrawingNodes.selectedChartId === selectedChartId
    && previousPaneDrawingNodes.selectedImageId === selectedImageId
    && previousPaneDrawingNodes.readOnly === readOnly
    && previousPaneDrawingNodes.selectionStroke === selectionStroke
    && previousPaneDrawingNodes.renderChartLoading === renderChartLoading
    && previousPaneDrawingNodes.renderImage === renderImage
    && previousPaneDrawingNodes.renderImageSelection === renderImageSelection
    && previousPaneDrawingNodes.isChartsLoading === isChartsLoading
    && previousPaneDrawingNodes.palette === palette
    && previousPaneDrawingNodes.drawingViewport.left === drawingViewport.left
    && previousPaneDrawingNodes.drawingViewport.top === drawingViewport.top
    && previousPaneDrawingNodes.drawingViewport.width === drawingViewport.width
    && previousPaneDrawingNodes.drawingViewport.height === drawingViewport.height;
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
                {shapeRects.map(({ shape, rect }) => renderShapeDrawing(shape, rect, "corner"))}
                {imageRects.map(({ image, rect }) => renderImageDrawing(image, rect, "corner"))}
              </>
            ),
            left: (
              <>
                {chartRects.map(({ chart, rect }) => renderChartDrawing(chart, rect, "left"))}
                {shapeRects.map(({ shape, rect }) => renderShapeDrawing(shape, rect, "left"))}
                {imageRects.map(({ image, rect }) => renderImageDrawing(image, rect, "left"))}
              </>
            ),
            scroll: (
              <>
                {chartRects.map(({ chart, rect }) => renderChartDrawing(chart, rect, "scroll"))}
                {shapeRects.map(({ shape, rect }) => renderShapeDrawing(shape, rect, "scroll"))}
                {imageRects.map(({ image, rect }) => renderImageDrawing(image, rect, "scroll"))}
              </>
            ),
            top: (
              <>
                {chartRects.map(({ chart, rect }) => renderChartDrawing(chart, rect, "top"))}
                {shapeRects.map(({ shape, rect }) => renderShapeDrawing(shape, rect, "top"))}
                {imageRects.map(({ image, rect }) => renderImageDrawing(image, rect, "top"))}
              </>
            )
          }) satisfies Record<FrozenDrawingPane, React.ReactNode>;

  if (!canReusePaneDrawingNodes) {
    paneDrawingNodesCacheRef.current = {
      chartRects,
      drawingViewport,
      imageRects,
      isChartsLoading,
      palette,
      readOnly,
      renderChartLoading,
      renderImage,
      renderImageSelection,
      selectedChartId,
      selectedImageId,
      selectionStroke,
      shapeRects,
      showImages,
      value: paneDrawingNodes
    };
  }

  function startColumnResize(pointerId: number, actualCol: number, widthPx: number, startX: number) {
    if (readOnly) {
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
    setInteractionMode("select");
    document.body.style.cursor = "col-resize";
    document.body.style.userSelect = "none";
  }

  function startRowResize(pointerId: number, actualRow: number, heightPx: number, startY: number) {
    if (readOnly) {
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
    setInteractionMode("select");
    document.body.style.cursor = "row-resize";
    document.body.style.userSelect = "none";
  }

  function startCellSelection(
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
    committedOnPointerDown: boolean,
    initialRange: XlsxCellRange,
    startClientX: number,
    startClientY: number
  ) {
    cachedScrollerRectRef.current = scrollRef.current?.getBoundingClientRect() ?? null;
    selectionDragRef.current = {
      anchor,
      axis,
      contentScaleX: pointerOrigin.contentScaleX,
      contentScaleY: pointerOrigin.contentScaleY,
      committedOnPointerDown,
      didDrag: false,
      originCell,
      originContentX: pointerOrigin.originContentX,
      originContentY: pointerOrigin.originContentY,
      pointerId,
      previewRange: normalizeRange(initialRange),
      startClientX,
      startClientY
    };
    selectionPreviewRangeRef.current = normalizeRange(initialRange);
    displayedSelectionRef.current = selectionPreviewRangeRef.current;
    applyPreviewOverlay(selectionPreviewRangeRef.current);
    setInteractionMode("select");
    document.body.style.userSelect = "none";
    installSelectionDragListeners(pointerId);
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

    const handlePointerMove = (event: PointerEvent) => {
      if (event.pointerId !== pointerId) {
        return;
      }

      const interaction = imageInteractionRef.current;
      if (!interaction) {
        return;
      }

      const deltaX = (event.clientX - interaction.startClientX) / zoomFactor;
      const deltaY = (event.clientY - interaction.startClientY) / zoomFactor;
      if (!interaction.didMove && (Math.abs(deltaX) > 3 || Math.abs(deltaY) > 3)) {
        interaction.didMove = true;
      }
      const nextRect = clampImageRect(
        interaction.type === "move"
          ? {
              ...interaction.baseRect,
              left: interaction.baseRect.left + deltaX,
              top: interaction.baseRect.top + deltaY
            }
          : resizeImageRect(interaction.baseRect, interaction.handle, deltaX, deltaY, IMAGE_MIN_SIZE_PX)
      );

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
      const preview = pendingPreview ?? imagePreviewRectRef.current;
      imageInteractionRef.current = null;
      imageInteractionCleanupRef.current = null;
      setInteractionMode("idle");
      document.body.style.cursor = "";
      document.body.style.userSelect = "";
      cleanup();
      if (interaction && preview && preview.id === interaction.imageId) {
        if (interaction.didMove) {
          skipNextImageClickRef.current = interaction.imageId;
        }
        setImageRect(interaction.imageId, preview.rect);
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

    const handlePointerMove = (event: PointerEvent) => {
      if (event.pointerId !== pointerId) {
        return;
      }

      const interaction = chartInteractionRef.current;
      if (!interaction) {
        return;
      }

      const deltaX = (event.clientX - interaction.startClientX) / zoomFactor;
      const deltaY = (event.clientY - interaction.startClientY) / zoomFactor;
      if (!interaction.didMove && (Math.abs(deltaX) > 3 || Math.abs(deltaY) > 3)) {
        interaction.didMove = true;
      }
      const nextRect = clampImageRect(
        interaction.type === "move"
          ? {
              ...interaction.baseRect,
              left: interaction.baseRect.left + deltaX,
              top: interaction.baseRect.top + deltaY
            }
          : resizeImageRect(interaction.baseRect, interaction.handle, deltaX, deltaY, 48)
      );

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
      const preview = pendingPreview ?? chartPreviewRectRef.current;
      chartInteractionRef.current = null;
      chartInteractionCleanupRef.current = null;
      setInteractionMode("idle");
      document.body.style.cursor = "";
      document.body.style.userSelect = "";
      cleanup();
      if (interaction && preview && preview.id === interaction.chartId) {
        if (interaction.didMove) {
          skipNextChartClickRef.current = interaction.chartId;
        }
        setChartRect(interaction.chartId, preview.rect);
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

  function moveSelection(nextRowIndex: number, nextColIndex: number, extend: boolean) {
    const nextRow = visibleRows[nextRowIndex];
    const nextCol = visibleCols[nextColIndex];
    if (nextRow === undefined || nextCol === undefined) {
      return;
    }

    selectCell({ row: nextRow, col: nextCol }, extend ? { extend: true } : undefined);
  }


  return (
    <div style={{ backgroundColor: palette.canvas, display: "flex", flex: 1, minHeight: 0, minWidth: 0 }}>
      <div
        key={activeTabIndex}
        ref={scrollRef}
        onScroll={handleScrollerScroll}
        onCopy={(event) => {
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
        }}
        onKeyDown={(event) => {
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

          if (!readOnly && isPrintableKey(event)) {
            event.preventDefault();
            startEditing(currentCell, event.key);
            return;
          }

          switch (event.key) {
            case "ArrowDown":
              event.preventDefault();
              moveSelection(Math.min(currentRowIndex + 1, visibleRows.length - 1), currentColIndex, event.shiftKey);
              break;
            case "ArrowUp":
              event.preventDefault();
              moveSelection(Math.max(currentRowIndex - 1, 0), currentColIndex, event.shiftKey);
              break;
            case "ArrowLeft":
              event.preventDefault();
              moveSelection(currentRowIndex, Math.max(currentColIndex - 1, 0), event.shiftKey);
              break;
            case "ArrowRight":
              event.preventDefault();
              moveSelection(currentRowIndex, Math.min(currentColIndex + 1, visibleCols.length - 1), event.shiftKey);
              break;
            case "Tab":
              event.preventDefault();
              moveSelection(
                currentRowIndex,
                event.shiftKey ? Math.max(currentColIndex - 1, 0) : Math.min(currentColIndex + 1, visibleCols.length - 1),
                false
              );
              break;
            case "Enter":
              event.preventDefault();
              if (event.metaKey || event.ctrlKey || event.altKey) {
                break;
              }
              if (event.shiftKey) {
                moveSelection(Math.max(currentRowIndex - 1, 0), currentColIndex, false);
                break;
              }
              moveSelection(Math.min(currentRowIndex + 1, visibleRows.length - 1), currentColIndex, false);
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
        }}
        onPaste={(event) => {
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
        }}
        tabIndex={0}
        style={{
          ["--xlsx-menu-active" as string]: selectionFill,
          ["--xlsx-menu-border" as string]: palette.strongBorder,
          ["--xlsx-menu-surface" as string]: palette.surface,
          ["--xlsx-selection-header" as string]: selectionHeaderSurface,
          backgroundColor: palette.canvas,
          color: palette.text,
          flex: 1,
          height: "100%",
          minHeight: 0,
          minWidth: 0,
          outline: "none",
          overflow: "auto",
          width: "100%"
        }}
      >
        <div
          style={{
            backgroundColor: resolveSheetSurface(activeSheet, palette),
            minHeight: "100%",
            minWidth: "100%",
            position: "relative",
            width: zoomedSheetWidth,
            height: zoomedSheetHeight
          }}
        >
          <div
            ref={wrapperRef}
            style={{
              height: sheetContentHeight,
              left: 0,
              position: "absolute",
              top: 0,
              transform: useNativeZoomForFrozenPanes || zoomFactor === 1 ? undefined : `scale(${zoomFactor})`,
              transformOrigin: "top left",
              zoom: useNativeZoomForFrozenPanes ? zoomFactor : undefined,
              width: totalWidth
            }}
          >
            {showImages ? (
              <>
                <div style={topOverlayStyle}>{paneDrawingNodes.top}</div>
                <div style={leftOverlayStyle}>{paneDrawingNodes.left}</div>
                <div style={cornerOverlayStyle}>{paneDrawingNodes.corner}</div>
                <div style={scrollOverlayStyle}>{paneDrawingNodes.scroll}</div>
              </>
            ) : null}
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
                <col style={{ width: ROW_HEADER_WIDTH }} />
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
              <thead style={{ position: "sticky", top: 0, zIndex: 50 }}>
                <tr>
                  <th
                    style={{
                      ...headerCellStyle,
                      backgroundColor: palette.headerSurface,
                      left: 0,
                      width: ROW_HEADER_WIDTH,
                      zIndex: 60
                    }}
                  />
                  {leadingColumnSpacerWidth > 0 ? (
                    <th aria-hidden="true" style={{ ...headerCellStyle, padding: 0, width: leadingColumnSpacerWidth }} />
                  ) : null}
                  {renderedCols.map((column) => (
                    <th
                      data-xlsx-col-header={column.actualCol}
                      key={column.key}
                      onPointerDown={(event) => handleColumnPointerDown(event, column.actualCol)}
                      style={{
                        ...headerCellStyle,
                        left: stickyLeftByCol.get(column.actualCol),
                        backgroundColor:
                          displayedSelection &&
                          column.actualCol >= displayedSelection.start.col &&
                          column.actualCol <= displayedSelection.end.col
                            ? selectionHeaderSurface
                            : headerCellStyle.backgroundColor,
                        zIndex: stickyLeftByCol.has(column.actualCol) ? 55 : headerCellStyle.zIndex
                      }}
                    >
                      <div style={{ position: "relative" }}>
                        {columnLabel(column.actualCol)}
                        <div
                          onPointerDown={(event) => {
                            if (readOnly) {
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
                        activeCell={activeCell}
                        actualRow={actualRow}
                        editingCell={editingCell}
                        editingValue={editingValue}
                        getCellData={getCellData}
                        key={virtualRow.key}
                        leadingSpacerWidth={leadingColumnSpacerWidth}
                        onCellClick={handleCellClick}
                        onCellDoubleClick={handleCellDoubleClick}
                        onCellPointerDown={handleCellPointerDown}
                        onEditingCancel={cancelEditing}
                        onEditingCommit={commitEditing}
                        onEditingValueChange={setEditingValue}
                        onRowPointerDown={handleRowPointerDown}
                        onRowResizePointerDown={handleRowResizePointerDown}
                        palette={palette}
                        readOnly={readOnly}
                        renderCellAdornment={renderCellAdornment}
                        rowHeight={virtualRow.size}
                        rowSelected={Boolean(
                          displayedSelection &&
                            actualRow >= displayedSelection.start.row &&
                            actualRow <= displayedSelection.end.row
                        )}
                        stickyLeftByCol={stickyLeftByCol}
                        stickyTop={stickyTopByRow.get(actualRow)}
                        trailingSpacerWidth={trailingColumnSpacerWidth}
                        visibleCols={renderedCols}
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
            <div
              ref={selectionOverlayRef}
              style={{
                backgroundColor: selectionFill,
                boxSizing: "border-box",
                boxShadow: `inset 0 0 0 ${selectionBorderWidth}px ${selectionStroke}`,
                height: resolvedSelectionOverlay?.height ?? 0,
                left: resolvedSelectionOverlay?.left ?? 0,
                opacity: resolvedSelectionOverlay ? 1 : 0,
                pointerEvents: "none",
                position: "absolute",
                top: resolvedSelectionOverlay?.top ?? 0,
                transition: "none",
                visibility: resolvedSelectionOverlay ? "visible" : "hidden",
                width: resolvedSelectionOverlay?.width ?? 0,
                zIndex: 24
              }}
            />
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
                border: `1px solid ${palette.surface}`,
                cursor: "crosshair",
                display: !readOnly && resolvedSelectionOverlay ? "block" : "none",
                height: 8,
                left: resolvedSelectionOverlay ? resolvedSelectionOverlay.left + resolvedSelectionOverlay.width - 4 : 0,
                pointerEvents: "auto",
                position: "absolute",
                top: resolvedSelectionOverlay ? resolvedSelectionOverlay.top + resolvedSelectionOverlay.height - 4 : 0,
                width: 8,
                zIndex: 25
              }}
            />
            {openTableMenuState ? (
              <div
                ref={tableMenuRef}
                style={{
                  color: palette.text,
                  left: Math.max(ROW_HEADER_WIDTH + 4, openTableMenuState.left),
                  position: "absolute",
                  top: openTableMenuState.top,
                  zIndex: 50
                }}
              >
                {renderTableHeaderMenu
                  ? renderTableHeaderMenu({
                      close: () => setOpenTableMenu(null),
                      column: openTableMenuState.column,
                      direction:
                        sortState &&
                        sortState.tableName === openTableMenuState.table.name &&
                        sortState.columnIndex === openTableMenuState.column.index
                          ? sortState.direction
                          : null,
                      sortAscending: () => sortTable(openTableMenuState.table.name, openTableMenuState.column.index, "ascending"),
                      sortDescending: () => sortTable(openTableMenuState.table.name, openTableMenuState.column.index, "descending"),
                      table: openTableMenuState.table
                    })
                  : (
                    <DefaultTableHeaderMenu
                      close={() => setOpenTableMenu(null)}
                      column={openTableMenuState.column}
                      direction={
                        sortState &&
                        sortState.tableName === openTableMenuState.table.name &&
                        sortState.columnIndex === openTableMenuState.column.index
                          ? sortState.direction
                          : null
                      }
                      sortAscending={() => sortTable(openTableMenuState.table.name, openTableMenuState.column.index, "ascending")}
                      sortDescending={() => sortTable(openTableMenuState.table.name, openTableMenuState.column.index, "descending")}
                      table={openTableMenuState.table}
                    />
                  )}
              </div>
            ) : null}
          </div>
        </div>
      </div>
    </div>
  );
}

function XlsxViewerInner({
  className,
  controller,
  emptyState,
  errorState,
  height = "100%",
  loadingComponent,
  loadingState,
  renderChartLoading,
  renderImage,
  renderImageSelection,
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
  const palette = useViewerPalette();

  return (
    <ViewerContext.Provider value={controller}>
      <div
        className={classNames("react-xlsx-viewer", className)}
        style={{
          blockSize: height,
          backgroundColor: palette.surface,
          border: `1px solid ${palette.border}`,
          borderRadius: rounded ? 12 : 0,
          color: palette.text,
          display: "flex",
          flex: "1 1 auto",
          flexDirection: "column",
          inlineSize: "100%",
          maxHeight: "100%",
          maxWidth: "100%",
          minHeight: 0,
          minWidth: 0,
          overflow: "hidden",
          width: "100%"
        }}
      >
        {resolveToolbar(toolbar, showDefaultToolbar, controller, palette)}
        <div style={{ display: "flex", flex: 1, minHeight: 0, minWidth: 0 }}>
          <XlsxGrid
            controller={controller}
            emptyState={emptyState}
            errorState={errorState}
            loadingComponent={loadingComponent}
            loadingState={loadingState}
            palette={palette}
            renderChartLoading={renderChartLoading}
            renderImage={renderImage}
            renderImageSelection={renderImageSelection}
            renderTableHeaderMenu={renderTableHeaderMenu}
            selectionColor={selectionColor}
            selectionFillColor={selectionFillColor}
            selectionHeaderColor={selectionHeaderColor}
            showImages={showImages}
          />
        </div>
      </div>
    </ViewerContext.Provider>
  );
}

function XlsxViewerWithInlineController(props: XlsxViewerProps) {
  const controller = useXlsxViewerController(props);
  return <XlsxViewerInner {...props} controller={controller} />;
}

function XlsxViewerProviderWithInlineController({
  children,
  ...options
}: Omit<XlsxViewerProviderProps, "controller">) {
  const controller = useXlsxViewerController(options);
  return <ViewerContext.Provider value={controller}>{children}</ViewerContext.Provider>;
}

export function XlsxViewerProvider({ children, controller, ...options }: XlsxViewerProviderProps) {
  if (controller) {
    return <ViewerContext.Provider value={controller}>{children}</ViewerContext.Provider>;
  }

  return <XlsxViewerProviderWithInlineController {...options}>{children}</XlsxViewerProviderWithInlineController>;
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
    selectedFormula,
    selectedValue,
    setCellFormula,
    setCellValue,
    setSelectedCellFormula,
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
      selectedFormula,
      selectedValue,
      setCellFormula,
      setCellValue,
      setSelectedCellFormula,
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
      selectedFormula,
      selectedValue,
      setCellFormula,
      setCellValue,
      setSelectedCellFormula,
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
    clearSelectedImage,
    getChartById,
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
    selectedChartId,
    selectedImage,
    selectedImageId,
    selectChart,
    selectImage,
    setChartRect,
    setImageRect,
    updateChart
  } = useXlsxViewer();

  return React.useMemo(
    () => ({
      charts,
      clearSelectedChart,
      clearSelectedImage,
      getChartById,
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
      selectedChartId,
      selectedImage,
      selectedImageId,
      selectChart,
      selectImage,
      setChartRect,
      setImageRect,
      updateChart
    }),
    [
      charts,
      clearSelectedChart,
      clearSelectedImage,
      getChartById,
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
      selectedChartId,
      selectedImage,
      selectedImageId,
      selectChart,
      selectImage,
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
    getChartById,
    getChartsheetById,
    getSheetCharts,
    isChartsLoading,
    moveChartBy,
    readOnly,
    resizeChartBy,
    selectChart,
    selectedChart,
    selectedChartId,
    setActiveTabIndex,
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
      getChartById,
      getChartsheetById,
      getSheetCharts,
      isChartsLoading,
      moveChartBy,
      readOnly,
      resizeChartBy,
      selectChart,
      selectedChart,
      selectedChartId,
      setActiveTabIndex,
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
      getChartById,
      getChartsheetById,
      getSheetCharts,
      isChartsLoading,
      moveChartBy,
      readOnly,
      resizeChartBy,
      selectChart,
      selectedChart,
      selectedChartId,
      setActiveTabIndex,
      setChartRect,
      tabs,
      updateChart
    ]
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
  const palette = useViewerPalette();
  return <DefaultToolbar controller={controller} palette={palette} />;
}
