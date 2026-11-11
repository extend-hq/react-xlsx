import * as React from "react";
import { createPortal } from "react-dom";
import { VIEWPORT_CHANGE_EVENT } from "./overlay-geometry";
import type { XlsxImageRect, XlsxSelectionPopoverContext, XlsxSelectionPopoverRenderProps } from "./types";

export function SelectionPopover({ context, suspended, viewportRef, getRects, render, onApply, focusGrid }: {
  context: XlsxSelectionPopoverContext;
  suspended: boolean;
  viewportRef: React.RefObject<HTMLDivElement | null>;
  getRects: () => XlsxImageRect[];
  render: (props: XlsxSelectionPopoverRenderProps) => React.ReactNode;
  onApply: (formula: string) => void;
  focusGrid: () => void;
}) {
  const hostRef = React.useRef<HTMLDivElement>(null);
  const controller = React.useMemo(() => new AbortController(), []);
  const [closed, setClosed] = React.useState(false);
  const [rect, setRect] = React.useState<XlsxImageRect | null>(null);
  const latest = React.useRef({ getRects, onApply, context, focusGrid });
  latest.current = { getRects, onApply, context, focusGrid };
  const close = React.useCallback(() => { controller.abort(); setClosed(true); }, [controller]);
  React.useLayoutEffect(() => { if (suspended) close(); }, [close, suspended]);
  const update = React.useCallback(() => {
    const viewport = viewportRef.current;
    if (!viewport) return;
    const bounds = viewport.getBoundingClientRect();
    const win = viewport.ownerDocument.defaultView!;
    const rects = latest.current.getRects().map(r => ({
      left: Math.max(r.left, bounds.left, 0), top: Math.max(r.top, bounds.top, 0),
      right: Math.min(r.left + r.width, bounds.right, win.innerWidth),
      bottom: Math.min(r.top + r.height, bounds.bottom, win.innerHeight)
    })).filter(r => r.right > r.left && r.bottom > r.top);
    const next = rects.length ? {
      left: Math.min(...rects.map(r => r.left)), top: Math.min(...rects.map(r => r.top)),
      width: Math.max(...rects.map(r => r.right)) - Math.min(...rects.map(r => r.left)),
      height: Math.max(...rects.map(r => r.bottom)) - Math.min(...rects.map(r => r.top))
    } : null;
    setRect(previous => previous?.left === next?.left && previous?.top === next?.top
      && previous?.width === next?.width && previous?.height === next?.height ? previous : next);
  }, [viewportRef]);
  React.useLayoutEffect(update);
  React.useLayoutEffect(() => {
    const viewport = viewportRef.current;
    const owner = viewport?.ownerDocument;
    const win = owner?.defaultView;
    if (!viewport || !owner || !win) return;
    const outside = (event: Event) => {
      if (!hostRef.current?.contains(event.target as Node)) close();
    };
    const observer = new ResizeObserver(update);
    observer.observe(viewport);
    owner.addEventListener("scroll", update, true);
    viewport.addEventListener(VIEWPORT_CHANGE_EVENT, update);
    viewport.addEventListener("transitionend", update);
    owner.addEventListener("pointerdown", outside, true);
    win.addEventListener("resize", update);
    return () => {
      controller.abort();
      observer.disconnect();
      owner.removeEventListener("scroll", update, true);
      viewport.removeEventListener(VIEWPORT_CHANGE_EVENT, update);
      viewport.removeEventListener("transitionend", update);
      owner.removeEventListener("pointerdown", outside, true);
      win.removeEventListener("resize", update);
    };
  }, [close, controller, update, viewportRef]);
  const commit = React.useCallback((apply: (context: XlsxSelectionPopoverContext) => void) => {
    if (controller.signal.aborted || latest.current.context.readOnly) return false;
    apply(latest.current.context);
    close();
    latest.current.focusGrid();
    return true;
  }, [close, controller]);
  const applyFormula = React.useCallback((formula: string) => commit(() => latest.current.onApply(formula)), [commit]);
  if (closed || suspended || !rect || !viewportRef.current) return null;
  const content = render({ context, rect, signal: controller.signal, applyFormula, commit, close });
  if (content == null || content === false) return null;
  return createPortal(
    <div
      ref={hostRef}
      data-xlsx-selection-popover=""
      onPointerDown={event => event.stopPropagation()}
      onClick={event => event.stopPropagation()}
      onDoubleClick={event => event.stopPropagation()}
      onKeyDown={event => {
        event.stopPropagation();
        if (event.key === "Escape" && !event.defaultPrevented) {
          event.preventDefault();
          close();
          latest.current.focusGrid();
        }
      }}
      style={{ position: "fixed", ...rect, pointerEvents: "none", zIndex: 10000 }}
    >{content}</div>,
    viewportRef.current.ownerDocument.body
  );
}
