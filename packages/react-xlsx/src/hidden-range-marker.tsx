import * as React from "react";
import type { XlsxHiddenRangeMarkerRenderProps, XlsxViewerProps } from "./types";

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
  zoomFactor: number;
  onUnhide: (start: number, end: number) => void;
  render?: XlsxViewerProps["renderHiddenRangeMarker"];
};

export function HiddenRangeMarker({ axis, start, end, label, position, readOnly, onUnhide, render, zoomFactor }: MarkerProps) {
  const action = React.useRef(onUnhide);
  action.current = onUnhide;
  const unhide = React.useCallback(() => {
    if (!readOnly) action.current(start, end);
  }, [start, end, readOnly]);
  const leading = start === 0;
  const row = axis === "row";
  const content = React.useMemo(() => {
    const caret = <span aria-hidden="true" style={{
      display: "block", width: "0.7em", height: "0.4em", background: "currentColor",
      clipPath: "polygon(0 0, 100% 0, 50% 100%)",
      transform: `rotate(${row ? (leading ? 180 : 0) : (leading ? 90 : -90)}deg)`
    }} />;
    const style: React.CSSProperties = {
      display: "grid", placeItems: "center", width: row ? "1.6em" : "1em", height: row ? "1em" : "1.6em",
      background: "transparent", border: 0, borderRadius: 0, margin: 0, padding: 0, color: "inherit", fontSize: "inherit"
    };
    const defaultNode = readOnly
      ? <span role="img" aria-label={label} title={label} style={style}>{caret}</span>
      : <button type="button" aria-label={label} title={label} onClick={unhide} style={{ ...style, cursor: "pointer" }}>{caret}</button>;
    return render ? render({ axis, start, end, readOnly, label, unhide, defaultNode }) : defaultNode;
  }, [axis, start, end, label, leading, row, readOnly, render, unhide]);

  if (content == null || content === false) return null;
  return <span
    data-xlsx-hidden-range-marker={axis}
    onPointerDown={event => event.stopPropagation()}
    onKeyDown={event => event.stopPropagation()}
    style={{
      position: "absolute", display: "flex", fontSize: 10 * zoomFactor, zIndex: 6, pointerEvents: readOnly && !render ? "none" : "auto",
      left: row ? "50%" : position, top: row ? position : "50%",
      transform: row ? `translate(-50%, ${leading ? 0 : -100}%)` : `translate(${leading ? 0 : -100}%, -50%)`
    }}
  >{content}</span>;
}
