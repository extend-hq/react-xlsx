import * as React from "react";
import type { Worksheet } from "@dukelib/sheets-wasm";
import { useVirtualizer } from "@tanstack/react-virtual";
import { useXlsxViewerController } from "./controller";
import type { XlsxViewerController, XlsxViewerProps, XlsxViewerProviderProps } from "./types";

const DEFAULT_ROW_HEIGHT = 24;
const DEFAULT_COL_WIDTH = 80;

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

function isLightColor(color: string): boolean {
  const rgb = parseHexColor(color);
  if (!rgb) {
    return false;
  }

  const [red, green, blue] = rgb.map((channel) => {
    const value = channel / 255;
    return value <= 0.04045 ? value / 12.92 : ((value + 0.055) / 1.055) ** 2.4;
  });
  const luminance = 0.2126 * red + 0.7152 * green + 0.0722 * blue;
  return luminance > 0.7;
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
  if (fill) {
    const fillColor =
      fill.fillType === "solid"
        ? cssColor(fill.color as Record<string, unknown> | undefined)
        : fill.fillType === "pattern"
          ? cssColor(fill.foreground as Record<string, unknown> | undefined)
          : null;

    if (fillColor && fillColor.toLowerCase() !== "#ffffff") {
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
      const normalized = fontColor.toLowerCase();
      const effectiveBackground = resolvedFillColor ?? palette.surface;
      css.color =
        paletteIsDark(palette) &&
        (normalized === "#000000" || normalized === "#000") &&
        !isLightColor(effectiveBackground)
          ? palette.text
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

function getCellDisplayValue(worksheet: Worksheet, row: number, col: number): string {
  const formatted = worksheet.getFormattedValueAt(row, col);
  if (formatted) {
    return formatted;
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

function renderLoading(loadingState: XlsxViewerProps["loadingState"], palette: ViewerPalette) {
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

function XlsxGrid({
  controller,
  emptyState,
  errorState,
  loadingState,
  palette
}: Pick<XlsxViewerProps, "emptyState" | "errorState" | "loadingState"> & {
  controller: XlsxViewerController;
  palette: ViewerPalette;
}) {
  const { activeSheet, activeSheetIndex, error, getActiveWorksheet, isLoading } = controller;
  const scrollRef = React.useRef<HTMLDivElement>(null);
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
  const worksheet = getActiveWorksheet();
  const rowHeights = activeSheet?.rowHeights ?? [];
  const visibleRows = activeSheet?.visibleRows ?? [];
  const visibleCols = activeSheet?.visibleCols ?? [];
  const colWidths = activeSheet?.colWidths ?? [];

  const rowVirtualizer = useVirtualizer({
    count: activeSheet?.rowCount ?? 0,
    estimateSize: (index) => rowHeights[index] ?? DEFAULT_ROW_HEIGHT,
    getScrollElement: () => scrollRef.current,
    overscan: 10
  });

  React.useEffect(() => {
    rowVirtualizer.scrollToOffset(0);
    if (scrollRef.current) {
      scrollRef.current.scrollLeft = 0;
    }
  }, [activeSheetIndex, controller.revision, rowVirtualizer]);

  React.useEffect(() => {
    function handlePointerMove(event: PointerEvent) {
      const state = resizeStateRef.current;
      if (!state) {
        return;
      }

      if (state.type === "column") {
        const delta = event.clientX - state.startPosition;
        controller.resizeColumn(state.actualIndex, state.initialPx + delta);
        return;
      }

      const delta = event.clientY - state.startPosition;
      controller.resizeRow(state.actualIndex, state.initialPx + delta);
    }

    function handlePointerUp(event: PointerEvent) {
      if (!resizeStateRef.current || resizeStateRef.current.pointerId !== event.pointerId) {
        return;
      }

      resizeStateRef.current = null;
      document.body.style.cursor = "";
      document.body.style.userSelect = "";
    }

    window.addEventListener("pointermove", handlePointerMove);
    window.addEventListener("pointerup", handlePointerUp);
    window.addEventListener("pointercancel", handlePointerUp);

    return () => {
      window.removeEventListener("pointermove", handlePointerMove);
      window.removeEventListener("pointerup", handlePointerUp);
      window.removeEventListener("pointercancel", handlePointerUp);
    };
  }, [controller]);

  if (isLoading) {
    return <>{renderLoading(loadingState, palette)}</>;
  }

  if (error) {
    return <>{renderError(errorState, error, palette)}</>;
  }

  if (!activeSheet) {
    return <>{renderEmpty(emptyState, palette)}</>;
  }

  const virtualRows = rowVirtualizer.getVirtualItems();
  const totalHeight = rowVirtualizer.getTotalSize();
  const totalWidth = colWidths.reduce((sum, width) => sum + width, 0) + 40;
  const headerCellStyle: React.CSSProperties = {
    backgroundColor: palette.headerSurface,
    borderBottom: `2px solid ${palette.strongBorder}`,
    borderRight: `1px solid ${palette.strongBorder}`,
    color: palette.mutedText,
    fontSize: "11px",
    fontWeight: 600,
    overflow: "hidden",
    padding: "2px 4px",
    textAlign: "center",
    userSelect: "none",
    whiteSpace: "nowrap"
  };
  const rowNumberStyle: React.CSSProperties = {
    backgroundColor: palette.rowHeaderSurface,
    borderBottom: `1px solid ${palette.border}`,
    borderRight: `1px solid ${palette.strongBorder}`,
    color: palette.mutedText,
    fontSize: "11px",
    left: 0,
    minWidth: 40,
    padding: "2px 4px",
    position: "sticky",
    textAlign: "center",
    userSelect: "none",
    width: 40,
    zIndex: 1
  };
  const columnResizeHandleStyle: React.CSSProperties = {
    backgroundColor: "transparent",
    cursor: "col-resize",
    position: "absolute",
    right: -3,
    top: 0,
    width: 6,
    height: "100%",
    zIndex: 5
  };
  const rowResizeHandleStyle: React.CSSProperties = {
    backgroundColor: "transparent",
    bottom: -3,
    cursor: "row-resize",
    height: 6,
    left: 0,
    position: "absolute",
    width: "100%",
    zIndex: 5
  };

  function startColumnResize(pointerId: number, actualCol: number, widthPx: number, startX: number) {
    resizeStateRef.current = {
      actualIndex: actualCol,
      initialPx: widthPx,
      pointerId,
      startPosition: startX,
      type: "column"
    };
    document.body.style.cursor = "col-resize";
    document.body.style.userSelect = "none";
  }

  function startRowResize(pointerId: number, actualRow: number, heightPx: number, startY: number) {
    resizeStateRef.current = {
      actualIndex: actualRow,
      initialPx: heightPx,
      pointerId,
      startPosition: startY,
      type: "row"
    };
    document.body.style.cursor = "row-resize";
    document.body.style.userSelect = "none";
  }

  return (
    <div style={{ backgroundColor: palette.canvas, display: "flex", flex: 1, minHeight: 0, minWidth: 0 }}>
      <div
        key={`${activeSheetIndex}-${controller.revision}`}
        ref={scrollRef}
        style={{
          backgroundColor: palette.canvas,
          color: palette.text,
          flex: 1,
          height: "100%",
          minHeight: 0,
          minWidth: 0,
          overflow: "auto",
          width: "100%"
        }}
      >
        <div
          style={{
            display: "flex",
            justifyContent: "flex-start",
            minHeight: "100%",
            minWidth: "100%",
            width: "fit-content"
          }}
        >
          <table
            style={{
              borderCollapse: "collapse",
              color: palette.text,
              flex: "0 0 auto",
              tableLayout: "fixed",
              width: totalWidth
            }}
          >
            <colgroup>
              <col style={{ width: 40 }} />
              {visibleCols.map((_, index) => (
                <col key={index} style={{ width: colWidths[index] ?? DEFAULT_COL_WIDTH }} />
              ))}
            </colgroup>
            <thead style={{ position: "sticky", top: 0, zIndex: 2 }}>
              <tr>
                <th style={{ ...headerCellStyle, left: 0, position: "sticky", width: 40, zIndex: 3 }} />
                {visibleCols.map((actualCol, index) => (
                <th key={index} style={headerCellStyle}>
                  <div style={{ position: "relative" }}>
                    {columnLabel(actualCol)}
                    <div
                      onPointerDown={(event) => {
                        event.preventDefault();
                        event.stopPropagation();
                        startColumnResize(
                          event.pointerId,
                          actualCol,
                          colWidths[index] ?? DEFAULT_COL_WIDTH,
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
                  <tr key={virtualRow.key} style={{ height: rowHeights[virtualRow.index] ?? DEFAULT_ROW_HEIGHT }}>
                    <td style={rowNumberStyle}>
                      <div style={{ position: "relative" }}>
                        {actualRow + 1}
                        <div
                          onPointerDown={(event) => {
                            event.preventDefault();
                            event.stopPropagation();
                            startRowResize(
                              event.pointerId,
                              actualRow,
                              rowHeights[virtualRow.index] ?? DEFAULT_ROW_HEIGHT,
                              event.clientY
                            );
                          }}
                          style={rowResizeHandleStyle}
                        />
                      </div>
                    </td>
                    {worksheet
                      ? visibleCols.map((actualCol, colIndex) => {
                          if (worksheet.isMergedSecondary(actualRow, actualCol)) {
                            return null;
                          }

                          const merge = worksheet.getMergeSpan(actualRow, actualCol) as
                            | { colSpan?: number }
                            | null
                            | undefined;
                          const rawStyle = worksheet.getCellStyleAt(actualRow, actualCol) as
                            | Record<string, unknown>
                            | null
                            | undefined;
                          const value = getCellDisplayValue(worksheet, actualRow, actualCol);

                          return (
                            <td
                              key={colIndex}
                              colSpan={merge?.colSpan}
                              style={buildCellStyle(rawStyle, palette)}
                              title={value}
                            >
                              {value}
                            </td>
                          );
                        })
                      : visibleCols.map((_, index) => (
                          <td
                            key={index}
                            style={{
                              backgroundColor: palette.surface,
                              borderBottom: `1px solid ${palette.border}`,
                              borderRight: `1px solid ${palette.border}`,
                              padding: "2px 4px"
                            }}
                          />
                        ))}
                  </tr>
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
  loadingState,
  rounded = true,
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
            loadingState={loadingState}
            palette={palette}
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
