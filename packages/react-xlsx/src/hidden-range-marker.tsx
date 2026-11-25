import * as React from "react";
import type { XlsxHiddenRangeMarkerRenderProps, XlsxViewerProps } from "./types";

function resizeCursor(axis: "row" | "column") {
  const transform = axis === "column" ? ' transform="rotate(90 12 12)"' : "";
  const path = "M6 10h12 M6 14h12 M12 7V3 M9 6l3-3 3 3 M12 17v4 M9 18l3 3 3-3";
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24">
    <g fill="none" stroke-linecap="round" stroke-linejoin="round"${transform}>
      <path d="${path}" stroke="white" stroke-width="3.5"/>
      <path d="${path}" stroke="#111" stroke-width="1.5"/>
    </g>
  </svg>`;
  const fallback = axis === "row" ? "row-resize" : "col-resize";
  return `url("data:image/svg+xml,${encodeURIComponent(svg)}") 12 12, ${fallback}`;
}

const ROW_RESIZE_CURSOR = resizeCursor("row");
const COLUMN_RESIZE_CURSOR = resizeCursor("column");

export function hiddenRangeEnds(hidden: Set<number>) {
  const ends = new Map<number, number>();
  for (const start of hidden) {
    if (hidden.has(start - 1)) continue;
    let end = start;
    while (hidden.has(end + 1)) end++;
    ends.set(start, end);
  }
  return ends;
}

type MarkerProps = Omit<XlsxHiddenRangeMarkerRenderProps, "defaultNode" | "unhide"> & {
  position: number | string;
  color: string;
  zoomFactor: number;
  onUnhide: (start: number, end: number) => void;
  render?: XlsxViewerProps["renderHiddenRangeMarker"];
};

export function HiddenRangeMarker({ axis, start, end, label, position, color, readOnly, canUnhide, onUnhide, render, zoomFactor }: MarkerProps) {
  const action = React.useRef(onUnhide);
  action.current = onUnhide;
  const unhide = React.useCallback(() => {
    if (canUnhide) action.current(start, end);
  }, [start, end, canUnhide]);
  const leading = start === 0;
  const row = axis === "row";
  const content = React.useMemo(() => {
    const edge = row ? (leading ? "borderTop" : "borderBottom") : (leading ? "borderLeft" : "borderRight");
    const hitSize = "max(6px, 0.8em)";
    const style: React.CSSProperties = {
      display: "block",
      boxSizing: "border-box",
      width: row ? "100%" : hitSize,
      height: row ? hitSize : "100%",
      background: "transparent",
      border: 0,
      [edge]: `3px solid ${color}`,
      borderRadius: 0,
      margin: 0,
      padding: 0,
      fontSize: "inherit",
      outlineOffset: -2
    };
    const defaultNode = canUnhide
      ? <button type="button" aria-label={label} title={label} onClick={unhide} style={{ ...style, cursor: row ? ROW_RESIZE_CURSOR : COLUMN_RESIZE_CURSOR }} />
      : <span role="img" aria-label={label} title={label} style={style} />;
    return render ? render({ axis, start, end, readOnly, canUnhide, label, unhide, defaultNode }) : defaultNode;
  }, [axis, start, end, label, leading, row, color, readOnly, canUnhide, render, unhide]);

  if (content == null || content === false) return null;
  return <span
    data-xlsx-hidden-range-marker={axis}
    onPointerDown={event => event.stopPropagation()}
    onKeyDown={event => event.stopPropagation()}
    style={{
      position: "absolute",
      display: "grid",
      placeItems: "center",
      fontSize: 10 * zoomFactor,
      zIndex: 6,
      pointerEvents: "none",
      left: row ? 0 : position,
      right: row ? 0 : undefined,
      top: row ? position : 0,
      bottom: row ? undefined : 0,
      transform: leading ? undefined : row ? "translateY(-100%)" : "translateX(-100%)"
    }}
  >
    <span style={{ display: "contents", pointerEvents: canUnhide || render ? "auto" : "none" }}>
      {content}
    </span>
  </span>;
}
