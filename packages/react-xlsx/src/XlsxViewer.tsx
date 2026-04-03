import * as React from "react";
import type { Worksheet } from "@dukelib/sheets-wasm";
import { useVirtualizer } from "@tanstack/react-virtual";
import { useXlsxViewerController } from "./controller";
import type {
  XlsxCellAddress,
  XlsxCellRange,
  XlsxViewerController,
  XlsxViewerEditing,
  XlsxViewerProps,
  XlsxViewerProviderProps,
  XlsxViewerSelection
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
const OPEN_GRID_ROW_GROWTH = 200;
const OPEN_GRID_COL_GROWTH = 24;
const OPEN_GRID_VERTICAL_EDGE_PX = 600;
const OPEN_GRID_HORIZONTAL_EDGE_PX = 480;
const SELECTION_DRAG_THRESHOLD_PX = 4;

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

const ViewerContext = React.createContext<XlsxViewerController | null>(null);

function classNames(...values: Array<string | false | null | undefined>) {
  return values.filter(Boolean).join(" ");
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

function resolveOpenGridExtent(maxUsedIndex: number, minimum: number, padding: number) {
  return Math.max(maxUsedIndex + 1 + padding, minimum);
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

function cssColor(color: Record<string, unknown> | undefined): string | null {
  if (!color?.hex) {
    return null;
  }

  const hex = String(color.hex);
  const rgb = hex.length === 8 ? hex.slice(2) : hex;
  return `#${rgb}`;
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

function mapBorder(edge: { style: string; color?: { hex?: string } }): string {
  const color = cssColor(edge.color as Record<string, unknown> | undefined) ?? "#000";
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
  const match = /^#([0-9a-f]{6})$/.exec(normalized);
  if (!match) {
    return null;
  }

  const hex = match[1];
  return [
    Number.parseInt(hex.slice(0, 2), 16),
    Number.parseInt(hex.slice(2, 4), 16),
    Number.parseInt(hex.slice(4, 6), 16)
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

function buildCellStyle(style: Record<string, unknown> | null | undefined, palette: ViewerPalette): React.CSSProperties {
  const css: React.CSSProperties = {
    backgroundColor: palette.surface,
    borderBottom: `1px solid ${palette.border}`,
    borderRight: `1px solid ${palette.border}`,
    color: palette.text,
    fontSize: "12px",
    overflow: "hidden",
    padding: "2px 4px",
    textOverflow: "ellipsis"
  };

  if (!style) {
    return css;
  }

  const fill = style.fill as Record<string, unknown> | undefined;
  let resolvedFillColor: string | null = null;
  let hasExplicitFill = false;
  if (fill) {
    const fillColor =
      fill.fillType === "solid"
        ? cssColor(fill.color as Record<string, unknown> | undefined)
        : fill.fillType === "pattern"
          ? cssColor(fill.foreground as Record<string, unknown> | undefined)
          : null;

    if (fillColor) {
      hasExplicitFill = true;
      resolvedFillColor = fillColor;
      css.backgroundColor = fillColor;
    }
  }

  const font = style.font as Record<string, unknown> | undefined;
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
    const fontColor = cssColor(font.color as Record<string, unknown> | undefined);
    if (fontColor) {
      css.color =
        paletteIsDark(palette) && !hasExplicitFill
          ? invertHexLightness(fontColor) ?? palette.text
          : fontColor;
    }
    if (typeof font.size === "number" && font.size !== 11) {
      css.fontSize = `${font.size}pt`;
    }
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
    } else {
      css.whiteSpace = "nowrap";
    }
  }

  const border = style.border as Record<string, Record<string, unknown>> | undefined;
  if (border) {
    if (border.top?.style && border.top.style !== "none") {
      css.borderTop = mapBorder(border.top as { style: string; color?: { hex?: string } });
    }
    if (border.right?.style && border.right.style !== "none") {
      css.borderRight = mapBorder(border.right as { style: string; color?: { hex?: string } });
    }
    if (border.bottom?.style && border.bottom.style !== "none") {
      css.borderBottom = mapBorder(border.bottom as { style: string; color?: { hex?: string } });
    }
    if (border.left?.style && border.left.style !== "none") {
      css.borderLeft = mapBorder(border.left as { style: string; color?: { hex?: string } });
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

function getCellDisplayValue(worksheet: Worksheet, row: number, col: number): string {
  const formatted = worksheet.getFormattedValueAt(row, col);
  if (formatted) {
    return decodeHtmlEntities(formatted);
  }

  const cellValue = worksheet.getCalculatedValueAt(row, col);
  if (cellValue.is_error) {
    return cellValue.asError() ?? "";
  }
  if (cellValue.is_empty) {
    return "";
  }

  return cellValue.toString();
}

function DefaultToolbar({ controller, palette }: { controller: XlsxViewerController; palette: ViewerPalette }) {
  const { activeSheetIndex, canDownload, displayFileName, download, sheets, setActiveSheetIndex } = controller;

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
      {sheets.length > 1 ? (
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
          {sheets.map((sheet, index) => (
            <button
              key={sheet.name}
              onClick={() => setActiveSheetIndex(index)}
              style={{
                backgroundColor: index === activeSheetIndex ? palette.sheetActiveSurface : palette.sheetInactiveSurface,
                border: `1px solid ${index === activeSheetIndex ? palette.strongBorder : "transparent"}`,
                borderRadius: 8,
                boxShadow: index === activeSheetIndex ? palette.shadow : "none",
                color: index === activeSheetIndex ? palette.sheetActiveText : palette.sheetInactiveText,
                cursor: "pointer",
                fontSize: 12,
                fontWeight: 500,
                padding: "6px 12px",
                whiteSpace: "nowrap"
              }}
              type="button"
            >
              {sheet.name}
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
  colSpan?: number;
  isMergedSecondary: boolean;
  spillWidth?: number;
  style: React.CSSProperties;
  value: string;
};

type GridRowProps = {
  activeCell: XlsxCellAddress | null;
  actualRow: number;
  editingCell: XlsxCellAddress | null;
  editingValue: string;
  getCellData: (row: number, col: number) => CellRenderData;
  onCellDoubleClick: (cell: XlsxCellAddress) => void;
  onCellPointerDown: (event: React.PointerEvent<HTMLTableCellElement>, cell: XlsxCellAddress, isActive: boolean) => void;
  onEditingCancel: () => void;
  onEditingCommit: () => void;
  onEditingValueChange: (value: string) => void;
  onRowPointerDown: (event: React.PointerEvent<HTMLTableCellElement>, actualRow: number) => void;
  onRowResizePointerDown: (event: React.PointerEvent<HTMLDivElement>, actualRow: number, rowHeight: number) => void;
  palette: ViewerPalette;
  readOnly: boolean;
  rowHeight: number;
  rowSelected: boolean;
  visibleCols: number[];
};

function GridRow({
  activeCell,
  actualRow,
  editingCell,
  editingValue,
  getCellData,
  onCellDoubleClick,
  onCellPointerDown,
  onEditingCancel,
  onEditingCommit,
  onEditingValueChange,
  onRowPointerDown,
  onRowResizePointerDown,
  palette,
  readOnly,
  rowHeight,
  rowSelected,
  visibleCols
}: GridRowProps) {
  return (
    <tr data-xlsx-row={actualRow} style={{ height: rowHeight }}>
      <td
        onPointerDown={(event) => onRowPointerDown(event, actualRow)}
        style={{
          backgroundColor: rowSelected ? "var(--xlsx-selection-header)" : palette.rowHeaderSurface,
          borderBottom: `1px solid ${palette.border}`,
          borderRight: `1px solid ${palette.strongBorder}`,
          color: palette.mutedText,
          fontSize: "11px",
          left: 0,
          minWidth: ROW_HEADER_WIDTH,
          padding: "2px 4px",
          position: "sticky",
          textAlign: "center",
          userSelect: "none",
          width: ROW_HEADER_WIDTH,
          zIndex: 1
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
      {visibleCols.map((actualCol, colIndex) => {
        const cellData = getCellData(actualRow, actualCol);
        if (cellData.isMergedSecondary) {
          return null;
        }

        const cell = { row: actualRow, col: actualCol };
        const isActive = isSameCell(activeCell, cell);
        const isEditing = isSameCell(editingCell, cell);
        const isSpilling = Boolean(cellData.spillWidth && cellData.spillWidth > 0);
        const cellStyle: React.CSSProperties = {
          ...cellData.style,
          cursor: isEditing ? "text" : "cell"
        };

        if (isActive || isSpilling) {
          cellStyle.position = "relative";
          cellStyle.zIndex = isActive ? 3 : 2;
        }
        if (isSpilling) {
          cellStyle.overflow = "visible";
          cellStyle.textOverflow = "clip";
        }
        if (isEditing) {
          cellStyle.padding = 0;
        }

        return (
          <td
            data-xlsx-cell={`${actualRow}:${actualCol}`}
            key={colIndex}
            colSpan={cellData.colSpan}
            onDoubleClick={() => {
              if (readOnly) {
                return;
              }

              onCellDoubleClick(cell);
            }}
            onPointerDown={(event) => onCellPointerDown(event, cell, isActive)}
            style={cellStyle}
            title={cellData.value}
          >
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
                  padding: "2px 4px",
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
            ) : (
              cellData.value
            )}
          </td>
        );
      })}
    </tr>
  );
}

function XlsxGrid({
  controller,
  emptyState,
  errorState,
  loadingComponent,
  loadingState,
  palette,
  selectionColor,
  selectionFillColor,
  selectionHeaderColor
}: Pick<
  XlsxViewerProps,
  "emptyState" | "errorState" | "loadingComponent" | "loadingState" | "selectionColor" | "selectionFillColor" | "selectionHeaderColor"
> & {
  controller: XlsxViewerController;
  palette: ViewerPalette;
}) {
  const {
    activeCell,
    activeSheet,
    activeSheetIndex,
    clearSelectedCells,
    error,
    fillSelection,
    getActiveWorksheet,
    getClipboardData,
    getCellDisplayValue: getControllerCellDisplayValue,
    isLoading,
    copySelectionToClipboard,
    pasteFromClipboard,
    pasteStructuredClipboardData,
    pasteText,
    readOnly,
    redo,
    revision,
    selectCell,
    selectRange,
    selection,
    setCellValue,
    undo
  } = controller;
  const scrollRef = React.useRef<HTMLDivElement>(null);
  const wrapperRef = React.useRef<HTMLDivElement>(null);
  const tableRef = React.useRef<HTMLTableElement>(null);
  const selectionOverlayRef = React.useRef<HTMLDivElement>(null);
  const fillHandleRef = React.useRef<HTMLDivElement>(null);
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
        didDrag: boolean;
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
  const [editingCell, setEditingCell] = React.useState<XlsxCellAddress | null>(null);
  const [editingValue, setEditingValue] = React.useState("");
  const [fillPreviewRange, setFillPreviewRange] = React.useState<XlsxCellRange | null>(null);
  const [selectionPreviewRange, setSelectionPreviewRange] = React.useState<XlsxCellRange | null>(null);
  const [interactionMode, setInteractionMode] = React.useState<"idle" | "fill" | "select">("idle");
  const [measuredSelectionOverlay, setMeasuredSelectionOverlay] = React.useState<{
    height: number;
    left: number;
    top: number;
    width: number;
  } | null>(null);
  const worksheet = getActiveWorksheet();
  const normalizedSelection = React.useMemo(() => (selection ? normalizeRange(selection) : null), [selection]);
  const [displayRowLimit, setDisplayRowLimit] = React.useState(() =>
    resolveOpenGridExtent(activeSheet?.maxUsedRow ?? -1, MIN_OPEN_GRID_ROWS, OPEN_GRID_ROW_PADDING)
  );
  const [displayColLimit, setDisplayColLimit] = React.useState(() =>
    resolveOpenGridExtent(activeSheet?.maxUsedCol ?? -1, MIN_OPEN_GRID_COLS, OPEN_GRID_COL_PADDING)
  );
  const visibleRows = React.useMemo(() => {
    const rows: number[] = [];
    for (let row = 0; row < displayRowLimit; row += 1) {
      if (!worksheet || !worksheet.isRowHidden(row)) {
        rows.push(row);
      }
    }
    return rows;
  }, [displayRowLimit, revision, worksheet]);
  const visibleCols = React.useMemo(() => {
    const cols: number[] = [];
    for (let col = 0; col < displayColLimit; col += 1) {
      if (!worksheet || !worksheet.isColumnHidden(col)) {
        cols.push(col);
      }
    }
    return cols;
  }, [displayColLimit, revision, worksheet]);
  const effectiveColWidths = React.useMemo(
    () =>
      visibleCols.map((col) => {
        const width = worksheet?.getColumnWidth(col);
        return width !== undefined && width !== null ? Math.max(Math.round(width * 7.5), DEFAULT_COL_WIDTH / 2) : DEFAULT_COL_WIDTH;
      }),
    [visibleCols, worksheet, revision]
  );
  const effectiveRowHeights = React.useMemo(
    () =>
      visibleRows.map((row) => {
        const height = worksheet?.getRowHeight(row);
        return height !== undefined && height !== null ? Math.max(Math.round(height * 1.33), DEFAULT_ROW_HEIGHT / 1.5) : DEFAULT_ROW_HEIGHT;
      }),
    [visibleRows, worksheet, revision]
  );
  const rowIndexByActual = React.useMemo(() => new Map(visibleRows.map((row, index) => [row, index])), [visibleRows]);
  const colIndexByActual = React.useMemo(() => new Map(visibleCols.map((col, index) => [col, index])), [visibleCols]);
  const visibleRowsRef = React.useRef<number[]>(visibleRows);
  const visibleColsRef = React.useRef<number[]>(visibleCols);
  const effectiveRowHeightsRef = React.useRef<number[]>(effectiveRowHeights);
  const effectiveColWidthsRef = React.useRef<number[]>(effectiveColWidths);
  const firstVisibleRow = visibleRows[0];
  const lastVisibleRow = visibleRows[visibleRows.length - 1];
  const firstVisibleCol = visibleCols[0];
  const lastVisibleCol = visibleCols[visibleCols.length - 1];
  const displayedSelection = fillPreviewRange ?? selectionPreviewRange ?? normalizedSelection;

  const rowVirtualizer = useVirtualizer({
    count: visibleRows.length,
    estimateSize: (index) => effectiveRowHeights[index] ?? DEFAULT_ROW_HEIGHT,
    getScrollElement: () => scrollRef.current,
    overscan: 10
  });

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
    displayedSelectionRef.current = displayedSelection;
  }, [displayedSelection]);

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
  }, [effectiveColWidths, effectiveRowHeights, visibleCols, visibleRows]);

  React.useEffect(() => {
    setDisplayRowLimit(resolveOpenGridExtent(activeSheet?.maxUsedRow ?? -1, MIN_OPEN_GRID_ROWS, OPEN_GRID_ROW_PADDING));
    setDisplayColLimit(resolveOpenGridExtent(activeSheet?.maxUsedCol ?? -1, MIN_OPEN_GRID_COLS, OPEN_GRID_COL_PADDING));
  }, [activeSheet?.maxUsedCol, activeSheet?.maxUsedRow, activeSheetIndex]);

  React.useEffect(() => {
    const selectionEnd = normalizedSelection?.end;
    const nextRowLimit = Math.max(
      resolveOpenGridExtent(activeSheet?.maxUsedRow ?? -1, MIN_OPEN_GRID_ROWS, OPEN_GRID_ROW_PADDING),
      (activeCell?.row ?? -1) + OPEN_GRID_ROW_PADDING + 1,
      (selectionEnd?.row ?? -1) + OPEN_GRID_ROW_PADDING + 1
    );
    const nextColLimit = Math.max(
      resolveOpenGridExtent(activeSheet?.maxUsedCol ?? -1, MIN_OPEN_GRID_COLS, OPEN_GRID_COL_PADDING),
      (activeCell?.col ?? -1) + OPEN_GRID_COL_PADDING + 1,
      (selectionEnd?.col ?? -1) + OPEN_GRID_COL_PADDING + 1
    );

    setDisplayRowLimit((current) => (current < nextRowLimit ? nextRowLimit : current));
    setDisplayColLimit((current) => (current < nextColLimit ? nextColLimit : current));
  }, [activeCell, activeSheet?.maxUsedCol, activeSheet?.maxUsedRow, normalizedSelection]);

  React.useEffect(() => {
    cellRenderCacheRef.current.clear();
  }, [activeSheetIndex, displayColLimit, displayRowLimit, palette, revision, worksheet]);

  React.useEffect(() => {
    rowVirtualizer.scrollToOffset(0);
    if (scrollRef.current) {
      scrollRef.current.scrollLeft = 0;
    }
  }, [activeSheetIndex]);

  React.useEffect(() => {
    rowVirtualizer.measure();
  }, [activeSheetIndex, revision, visibleRows.length]);

  React.useEffect(() => {
    const scroller = scrollRef.current;
    if (!scroller) {
      return;
    }
    const currentScroller = scroller;

    function handleScroll() {
      if (
        currentScroller.scrollHeight - (currentScroller.scrollTop + currentScroller.clientHeight) <
        OPEN_GRID_VERTICAL_EDGE_PX
      ) {
        setDisplayRowLimit((current) => current + OPEN_GRID_ROW_GROWTH);
      }

      if (
        currentScroller.scrollWidth - (currentScroller.scrollLeft + currentScroller.clientWidth) <
        OPEN_GRID_HORIZONTAL_EDGE_PX
      ) {
        setDisplayColLimit((current) => current + OPEN_GRID_COL_GROWTH);
      }
    }

    currentScroller.addEventListener("scroll", handleScroll, { passive: true });
    return () => {
      currentScroller.removeEventListener("scroll", handleScroll);
    };
  }, [activeSheetIndex]);

  const resolvePointerCellFromClient = React.useCallback((clientX: number, clientY: number): XlsxCellAddress | null => {
    const wrapper = wrapperRef.current;
    const scroller = scrollRef.current;
    const visibleRowsCurrent = visibleRowsRef.current;
    const visibleColsCurrent = visibleColsRef.current;
    const rowHeightsCurrent = effectiveRowHeightsRef.current;
    const colWidthsCurrent = effectiveColWidthsRef.current;

    if (!wrapper || !scroller || visibleRowsCurrent.length === 0 || visibleColsCurrent.length === 0) {
      return null;
    }

    const wrapperRect = wrapper.getBoundingClientRect();
    const scrollerRect = scroller.getBoundingClientRect();
    if (
      clientX < scrollerRect.left ||
      clientX > scrollerRect.right ||
      clientY < scrollerRect.top ||
      clientY > scrollerRect.bottom
    ) {
      return null;
    }

    const localX = clientX - wrapperRect.left + scroller.scrollLeft;
    const localY = clientY - wrapperRect.top + scroller.scrollTop;
    const rowContentOffset = localY - HEADER_HEIGHT;
    const colContentOffset = localX - ROW_HEADER_WIDTH;

    if (localY >= HEADER_HEIGHT && localX < ROW_HEADER_WIDTH) {
      const rowIndex = findIndexForOffset(rowHeightsCurrent, rowContentOffset);
      const actualRow = visibleRowsCurrent[rowIndex];
      if (actualRow !== undefined && firstVisibleColRef.current !== undefined) {
        return { row: actualRow, col: firstVisibleColRef.current };
      }
      return null;
    }

    if (localY < HEADER_HEIGHT && localX >= ROW_HEADER_WIDTH) {
      const colIndex = findIndexForOffset(colWidthsCurrent, colContentOffset);
      const actualCol = visibleColsCurrent[colIndex];
      if (actualCol !== undefined && firstVisibleRowRef.current !== undefined) {
        return { row: firstVisibleRowRef.current, col: actualCol };
      }
      return null;
    }

    if (localY >= HEADER_HEIGHT && localX >= ROW_HEADER_WIDTH) {
      const rowIndex = findIndexForOffset(rowHeightsCurrent, rowContentOffset);
      const colIndex = findIndexForOffset(colWidthsCurrent, colContentOffset);
      const actualRow = visibleRowsCurrent[rowIndex];
      const actualCol = visibleColsCurrent[colIndex];
      if (actualRow !== undefined && actualCol !== undefined) {
        return { row: actualRow, col: actualCol };
      }
    }

    return null;
  }, []);

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
    selectionDragCleanupRef.current = null;
    fillDragCleanupRef.current = null;
    selectionDragRef.current = null;
    fillDragRef.current = null;
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
    setSelectionPreviewRange(null);
    setInteractionMode("idle");
  }, [activeSheetIndex, revision]);

  const focusGrid = React.useCallback(() => {
    scrollRef.current?.focus();
  }, []);

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

  const getCellData = React.useCallback((row: number, col: number): CellRenderData => {
    const cacheKey = `${row}:${col}`;
    const cached = cellRenderCacheRef.current.get(cacheKey);
    if (cached) {
      return cached;
    }

    if (!worksheet) {
      const emptyData: CellRenderData = {
        isMergedSecondary: false,
        style: {
          backgroundColor: palette.surface,
          borderBottom: `1px solid ${palette.border}`,
          borderRight: `1px solid ${palette.border}`,
          padding: "2px 4px"
        },
        value: ""
      };
      cellRenderCacheRef.current.set(cacheKey, emptyData);
      return emptyData;
    }

    if (worksheet.isMergedSecondary(row, col)) {
      const mergedSecondaryData: CellRenderData = {
        isMergedSecondary: true,
        style: {},
        value: ""
      };
      cellRenderCacheRef.current.set(cacheKey, mergedSecondaryData);
      return mergedSecondaryData;
    }

    const merge = worksheet.getMergeSpan(row, col) as { colSpan?: number } | null | undefined;
    const rawStyle = worksheet.getCellStyleAt(row, col) as Record<string, unknown> | null | undefined;
    const nextData: CellRenderData = {
      colSpan: merge?.colSpan,
      isMergedSecondary: false,
      style: buildCellStyle(rawStyle, palette),
      value: getCellDisplayValue(worksheet, row, col)
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
  }, [colIndexByActual, effectiveColWidths, palette, visibleCols, worksheet]);

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
      height: sumSegment(effectiveRowHeights, startRowIndex, endRowIndex),
      left: ROW_HEADER_WIDTH + (startColIndex > 0 ? sumSegment(effectiveColWidths, 0, startColIndex - 1) : 0),
      top: HEADER_HEIGHT + (startRowIndex > 0 ? sumSegment(effectiveRowHeights, 0, startRowIndex - 1) : 0),
      width: sumSegment(effectiveColWidths, startColIndex, endColIndex)
    };
  }, [colIndexByActual, displayedSelection, effectiveColWidths, effectiveRowHeights, rowIndexByActual]);
  const resolvedSelectionOverlay = measuredSelectionOverlay ?? selectionOverlay;

  const resolveOverlayRect = React.useCallback((range: XlsxCellRange) => {
    const normalized = normalizeRange(range);
    const wrapper = wrapperRef.current;
    if (wrapper) {
      const startCell = wrapper.querySelector<HTMLElement>(
        `[data-xlsx-cell="${normalized.start.row}:${normalized.start.col}"]`
      );
      const endCell = wrapper.querySelector<HTMLElement>(
        `[data-xlsx-cell="${normalized.end.row}:${normalized.end.col}"]`
      );

      if (startCell && endCell) {
        const startLeft = startCell.offsetLeft;
        const startTop = startCell.offsetTop;
        const startRight = startLeft + startCell.offsetWidth;
        const startBottom = startTop + startCell.offsetHeight;
        const endLeft = endCell.offsetLeft;
        const endTop = endCell.offsetTop;
        const endRight = endLeft + endCell.offsetWidth;
        const endBottom = endTop + endCell.offsetHeight;

        return {
          height: Math.max(startBottom, endBottom) - Math.min(startTop, endTop),
          left: Math.min(startLeft, endLeft),
          top: Math.min(startTop, endTop),
          width: Math.max(startRight, endRight) - Math.min(startLeft, endLeft)
        };
      }
    }

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
      height: sumSegment(effectiveRowHeights, startRowIndex, endRowIndex),
      left: ROW_HEADER_WIDTH + (startColIndex > 0 ? sumSegment(effectiveColWidths, 0, startColIndex - 1) : 0),
      top: HEADER_HEIGHT + (startRowIndex > 0 ? sumSegment(effectiveRowHeights, 0, startRowIndex - 1) : 0),
      width: sumSegment(effectiveColWidths, startColIndex, endColIndex)
    };
  }, [colIndexByActual, effectiveColWidths, effectiveRowHeights, rowIndexByActual]);

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
            ? Math.max(DEFAULT_COL_WIDTH / 2, state.initialPx + (event.clientX - state.startPosition))
            : Math.max(DEFAULT_ROW_HEIGHT / 1.5, state.initialPx + (event.clientY - state.startPosition)),
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
          rowVirtualizer.measure();
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
  }, [applyColumnPreview, applyRowPreview, controller, refreshOverlayFromCurrentSelection, rowVirtualizer]);

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

    const handlePointerMove = (event: PointerEvent) => {
      if (event.pointerId !== pointerId) {
        return;
      }

      const dragState = selectionDragRef.current;
      if (!dragState) {
        return;
      }

      if (!dragState.didDrag) {
        const deltaX = Math.abs(event.clientX - dragState.startClientX);
        const deltaY = Math.abs(event.clientY - dragState.startClientY);
        if (deltaX < SELECTION_DRAG_THRESHOLD_PX && deltaY < SELECTION_DRAG_THRESHOLD_PX) {
          return;
        }

        dragState.didDrag = true;
      }

      const nextCell = resolvePointerCellFromClient(event.clientX, event.clientY);
      if (!nextCell) {
        return;
      }

      const nextRange = buildDraggedSelectionRange(dragState, nextCell);
      if (!nextRange || rangesEqual(nextRange, dragState.previewRange)) {
        return;
      }

      dragState.previewRange = nextRange;
      setSelectionPreviewRange(nextRange);
      applyPreviewOverlay(nextRange);
    };

    const finishSelectionDrag = (event: PointerEvent) => {
      if (event.pointerId !== pointerId) {
        return;
      }

      const nextCell = resolvePointerCellFromClient(event.clientX, event.clientY);
      const dragState = selectionDragRef.current;
      let nextRange = dragState?.previewRange ?? null;
      if (dragState?.didDrag && nextCell && dragState) {
        nextRange = buildDraggedSelectionRange(dragState, nextCell);
      }

      selectionDragRef.current = null;
      setSelectionPreviewRange(null);
      if (nextRange) {
        selectRange(nextRange);
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
      window.removeEventListener("pointermove", handlePointerMove);
      window.removeEventListener("pointerup", finishSelectionDrag);
      window.removeEventListener("pointercancel", finishSelectionDrag);
    };
  }

  function installFillDragListeners(pointerId: number, sourceRange: XlsxCellRange) {
    fillDragCleanupRef.current?.();

    const handlePointerMove = (event: PointerEvent) => {
      if (event.pointerId !== pointerId) {
        return;
      }

      const nextCell = resolvePointerCellFromClient(event.clientX, event.clientY);
      if (nextCell) {
        updateFillPreview(nextCell);
      }
    };

    const finishFillDrag = (event: PointerEvent) => {
      if (event.pointerId !== pointerId) {
        return;
      }

      const nextCell = resolvePointerCellFromClient(event.clientX, event.clientY);
      if (nextCell) {
        updateFillPreview(nextCell);
      }

      const nextRange = fillDragRef.current?.previewRange ?? sourceRange;
      fillDragRef.current = null;
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
      window.removeEventListener("pointermove", handlePointerMove);
      window.removeEventListener("pointerup", finishFillDrag);
      window.removeEventListener("pointercancel", finishFillDrag);
    };
  }

  React.useLayoutEffect(() => {
    if (!displayedSelection || !wrapperRef.current) {
      if (selectionOverlayRef.current) {
        selectionOverlayRef.current.style.opacity = "0";
        selectionOverlayRef.current.style.visibility = "hidden";
      }
      setMeasuredSelectionOverlay(null);
      return;
    }

    const nextOverlay = resolveOverlayRect(displayedSelection);
    if (!nextOverlay) {
      setMeasuredSelectionOverlay(null);
      return;
    }

    setMeasuredSelectionOverlay((current) => {
      if (
        current &&
        Math.abs(current.left - nextOverlay.left) < 0.5 &&
        Math.abs(current.top - nextOverlay.top) < 0.5 &&
        Math.abs(current.width - nextOverlay.width) < 0.5 &&
        Math.abs(current.height - nextOverlay.height) < 0.5
      ) {
        return current;
      }

      return nextOverlay;
    });
  }, [displayedSelection, resolveOverlayRect, revision]);

  const handleCellDoubleClick = React.useCallback((cell: XlsxCellAddress) => {
    startEditing(cell);
  }, [startEditing]);

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
    if (!isActive || !editingCellRef.current) {
      selectRange(initialRange);
    }
    startCellSelection(event.pointerId, anchor, "cell", initialRange, event.clientX, event.clientY);
  }, [focusGrid, selectRange]);

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
    selectRange(initialRange);
    startCellSelection(
      event.pointerId,
      { row: anchorRow, col: firstVisibleCol },
      "row",
      initialRange,
      event.clientX,
      event.clientY
    );
  }, [firstVisibleCol, focusGrid, lastVisibleCol, selectRange]);

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

  if (isLoading) {
    return <>{renderLoading(loadingComponent, loadingState, palette)}</>;
  }

  if (error) {
    return <>{renderError(errorState, error, palette)}</>;
  }

  if (!activeSheet) {
    return <>{renderEmpty(emptyState, palette)}</>;
  }

  const virtualRows = rowVirtualizer.getVirtualItems();
  const totalHeight = rowVirtualizer.getTotalSize();
  const totalWidth = effectiveColWidths.reduce((sum, width) => sum + width, 0) + ROW_HEADER_WIDTH;
  const { fill: selectionFill, header: selectionHeaderSurface, stroke: selectionStroke } = resolveSelectionColors({
    palette,
    selectionColor,
    selectionFillColor,
    selectionHeaderColor
  });
  const selectionBorderWidth = 1.5;
  const headerCellStyle: React.CSSProperties = {
    backgroundColor: palette.headerSurface,
    borderBottom: `2px solid ${palette.strongBorder}`,
    borderRight: `1px solid ${palette.strongBorder}`,
    color: palette.mutedText,
    fontSize: "11px",
    fontWeight: 600,
    height: HEADER_HEIGHT,
    overflow: "hidden",
    padding: "2px 4px",
    textAlign: "center",
    userSelect: "none",
    whiteSpace: "nowrap"
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
    initialRange: XlsxCellRange,
    startClientX: number,
    startClientY: number
  ) {
    selectionDragRef.current = {
      anchor,
      axis,
      didDrag: false,
      pointerId,
      previewRange: normalizeRange(initialRange),
      startClientX,
      startClientY
    };
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
    setFillPreviewRange(nextRange);
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
        key={activeSheetIndex}
        ref={scrollRef}
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
          ref={wrapperRef}
          style={{
            display: "flex",
            justifyContent: "flex-start",
            minHeight: "100%",
            minWidth: "100%",
            position: "relative",
            width: "fit-content"
          }}
        >
          <table
            ref={tableRef}
            style={{
              borderCollapse: "collapse",
              color: palette.text,
              flex: "0 0 auto",
              tableLayout: "fixed",
              width: totalWidth
            }}
          >
            <colgroup>
              <col style={{ width: ROW_HEADER_WIDTH }} />
              {visibleCols.map((_, index) => (
                <col
                  key={index}
                  ref={(element) => {
                    const actualCol = visibleCols[index];
                    if (actualCol === undefined) {
                      return;
                    }

                    if (element) {
                      colElementRefs.current.set(actualCol, element);
                    } else {
                      colElementRefs.current.delete(actualCol);
                    }
                  }}
                  style={{ width: effectiveColWidths[index] ?? DEFAULT_COL_WIDTH }}
                />
              ))}
            </colgroup>
            <thead style={{ position: "sticky", top: 0, zIndex: 2 }}>
              <tr>
                <th style={{ ...headerCellStyle, left: 0, position: "sticky", width: ROW_HEADER_WIDTH, zIndex: 3 }} />
                {visibleCols.map((actualCol, index) => (
                  <th
                    data-xlsx-col-header={actualCol}
                    key={index}
                    onPointerDown={(event) => {
                      if (event.button !== 0 || firstVisibleRow === undefined || lastVisibleRow === undefined) {
                        return;
                      }

                      event.preventDefault();
                      focusGrid();
                      const anchorCol = event.shiftKey && normalizedSelection ? normalizedSelection.start.col : actualCol;
                      const initialRange = normalizeRange({
                        start: { row: firstVisibleRow, col: anchorCol },
                        end: { row: lastVisibleRow, col: actualCol }
                      });
                      selectRange(initialRange);
                      startCellSelection(
                        event.pointerId,
                        { row: firstVisibleRow, col: anchorCol },
                        "column",
                        initialRange,
                        event.clientX,
                        event.clientY
                      );
                    }}
                    style={{
                      ...headerCellStyle,
                      backgroundColor:
                        displayedSelection &&
                        actualCol >= displayedSelection.start.col &&
                        actualCol <= displayedSelection.end.col
                          ? selectionHeaderSurface
                          : headerCellStyle.backgroundColor
                    }}
                  >
                    <div style={{ position: "relative" }}>
                      {columnLabel(actualCol)}
                      <div
                        onPointerDown={(event) => {
                          if (readOnly) {
                            return;
                          }

                          event.preventDefault();
                          event.stopPropagation();
                          startColumnResize(
                            event.pointerId,
                            actualCol,
                            effectiveColWidths[index] ?? DEFAULT_COL_WIDTH,
                            event.clientX
                          );
                        }}
                        style={columnResizeHandleStyle}
                      />
                    </div>
                  </th>
                ))}
              </tr>
            </thead>
            <tbody>
              {virtualRows[0] ? (
                <tr style={{ height: virtualRows[0].start }}>
                  <td colSpan={visibleCols.length + 1} />
                </tr>
              ) : null}
              {virtualRows.map((virtualRow) => {
                const actualRow = visibleRows[virtualRow.index];
                if (actualRow === undefined) {
                  return null;
                }

                return (
                  <GridRow
                    activeCell={activeCell}
                    actualRow={actualRow}
                    editingCell={editingCell}
                    editingValue={editingValue}
                    getCellData={getCellData}
                    key={virtualRow.key}
                    onCellDoubleClick={handleCellDoubleClick}
                    onCellPointerDown={handleCellPointerDown}
                    onEditingCancel={cancelEditing}
                    onEditingCommit={commitEditing}
                    onEditingValueChange={setEditingValue}
                    onRowPointerDown={handleRowPointerDown}
                    onRowResizePointerDown={handleRowResizePointerDown}
                    palette={palette}
                    readOnly={readOnly}
                    rowHeight={effectiveRowHeights[virtualRow.index] ?? DEFAULT_ROW_HEIGHT}
                    rowSelected={Boolean(
                      displayedSelection &&
                        actualRow >= displayedSelection.start.row &&
                        actualRow <= displayedSelection.end.row
                    )}
                    visibleCols={visibleCols}
                  />
                );
              })}
              {virtualRows.length > 0 ? (
                <tr
                  style={{
                    height: totalHeight - (virtualRows[virtualRows.length - 1]?.end ?? totalHeight)
                  }}
                >
                  <td colSpan={visibleCols.length + 1} />
                </tr>
              ) : null}
            </tbody>
          </table>
          <div
            ref={selectionOverlayRef}
            style={{
              backgroundColor: selectionFill,
              border: `${selectionBorderWidth}px solid ${selectionStroke}`,
              boxSizing: "border-box",
              height: resolvedSelectionOverlay?.height ?? 0,
              left: resolvedSelectionOverlay?.left ?? 0,
              opacity: resolvedSelectionOverlay ? 1 : 0,
              pointerEvents: "none",
              position: "absolute",
              top: resolvedSelectionOverlay?.top ?? 0,
              transition:
                interactionMode === "idle"
                  ? "top 80ms cubic-bezier(0.2, 0.8, 0.2, 1), left 80ms cubic-bezier(0.2, 0.8, 0.2, 1), width 80ms cubic-bezier(0.2, 0.8, 0.2, 1), height 80ms cubic-bezier(0.2, 0.8, 0.2, 1)"
                  : "none",
              visibility: resolvedSelectionOverlay ? "visible" : "hidden",
              width: resolvedSelectionOverlay?.width ?? 0,
              zIndex: 4
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
              zIndex: 5
            }}
          />
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
  rounded = true,
  selectionColor,
  selectionFillColor,
  selectionHeaderColor,
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
            selectionColor={selectionColor}
            selectionFillColor={selectionFillColor}
            selectionHeaderColor={selectionHeaderColor}
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
