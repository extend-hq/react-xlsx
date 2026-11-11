import * as React from "react";
import { createPortal } from "react-dom";
import { VIEWPORT_CHANGE_EVENT, visibleOverlayRect } from "./overlay-geometry";
import type { XlsxCellEditorOverlayRenderProps, XlsxCellEditorSession, XlsxImageRect } from "./types";

export function CellEditorOverlay({ editorRef, overlayRef, clipElement, session, render, onExit }: {
  editorRef: React.RefObject<HTMLDivElement | null>;
  overlayRef: React.RefObject<HTMLDivElement | null>;
  clipElement?: HTMLElement | null;
  session: XlsxCellEditorSession;
  render: (props: XlsxCellEditorOverlayRenderProps) => React.ReactNode;
  onExit: () => void;
}) {
  const [geometry, setGeometry] = React.useState<{ owner: Document; editor: HTMLDivElement; rect: XlsxImageRect; visible: boolean } | null>(null);
  const latest = React.useRef({ onExit, session });
  latest.current = { onExit, session };
  React.useLayoutEffect(() => {
    const editor = editorRef.current;
    const owner = editor?.ownerDocument;
    const win = owner?.defaultView;
    const viewport = editor?.closest<HTMLElement>('[role="grid"]');
    if (!editor || !owner || !win || !viewport) return;
    const update = () => {
      const bounds = editor.getBoundingClientRect();
      const clipped = visibleOverlayRect(editor, clipElement);
      const rect = clipped ?? { left: bounds.left, top: bounds.top, width: bounds.width, height: bounds.height };
      const visible = clipped !== null;
      setGeometry(previous => previous && previous.editor === editor && previous.visible === visible && previous.rect.left === rect.left
        && previous.rect.top === rect.top && previous.rect.width === rect.width && previous.rect.height === rect.height
        ? previous : { owner, editor, rect, visible });
    };
    const outside = (event: Event) => {
      const target = event.target as Node | null;
      if (!latest.current.session.signal.aborted && !editor.contains(target) && !overlayRef.current?.contains(target)) {
        latest.current.onExit();
      }
    };
    update();
    const observer = new ResizeObserver(update);
    observer.observe(editor);
    observer.observe(viewport);
    owner.addEventListener("scroll", update, true);
    viewport.addEventListener(VIEWPORT_CHANGE_EVENT, update);
    owner.addEventListener("pointerdown", outside, true);
    owner.addEventListener("focusin", outside);
    win.addEventListener("resize", update);
    return () => {
      observer.disconnect();
      owner.removeEventListener("scroll", update, true);
      viewport.removeEventListener(VIEWPORT_CHANGE_EVENT, update);
      owner.removeEventListener("pointerdown", outside, true);
      owner.removeEventListener("focusin", outside);
      win.removeEventListener("resize", update);
    };
  }, [editorRef, overlayRef, clipElement, session.signal]);

  const content = geometry ? render({ ...session, rect: geometry.rect, editorElement: geometry.editor }) : null;
  if (!geometry || content == null || content === false) return null;
  return createPortal(
    <div
      ref={overlayRef}
      data-xlsx-cell-editor-overlay=""
      onPointerDown={(event) => {
        event.stopPropagation();
        const target = event.target as HTMLElement;
        if (!target.closest('input, textarea, select, [contenteditable="true"]')) event.preventDefault();
      }}
      onClick={(event) => event.stopPropagation()}
      onKeyDown={(event) => {
        event.stopPropagation();
        if (event.key === "Escape" && !event.defaultPrevented) {
          event.preventDefault();
          session.cancel();
        }
      }}
      style={{ position: "fixed", ...geometry.rect, pointerEvents: "none", visibility: geometry.visible ? "visible" : "hidden", zIndex: 10000 }}
    >
      {content}
    </div>,
    geometry.owner.body
  );
}
