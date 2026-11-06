import * as React from "react";
import type { FormulaReference } from "./formula-references";

export type EditorTextLine = { text: string; start: number };

export function CellEditor({
  value,
  references,
  inputRef,
  canvasText = false,
  canvasOffset,
  paintCanvasText,
  style,
  onChange,
  onBlur,
  onCommit,
  onCancel,
  onNavigate
}: {
  value: string;
  references: FormulaReference[];
  inputRef?: React.Ref<HTMLTextAreaElement>;
  canvasText?: boolean;
  canvasOffset?: { left: number; top: number };
  paintCanvasText?: (context: CanvasRenderingContext2D, width: number, height: number, lines: EditorTextLine[], textTop: number) => void;
  style: React.CSSProperties;
  onChange: (value: string) => void;
  onBlur: () => void;
  onCommit: () => void;
  onCancel: () => void;
  onNavigate: (key: string, value: string) => void;
}) {
  const rootRef = React.useRef<HTMLDivElement>(null);
  const mirrorRef = React.useRef<HTMLDivElement>(null);
  const textRef = React.useRef<HTMLSpanElement>(null);
  const baselineRef = React.useRef<HTMLSpanElement>(null);
  const plainTextRef = React.useRef<HTMLSpanElement>(null);
  const canvasRef = React.useRef<HTMLCanvasElement>(null);
  const textareaRef = React.useRef<HTMLTextAreaElement>(null);
  const baselineOffsetRef = React.useRef(0);
  const measuringContextRef = React.useRef<CanvasRenderingContext2D | null>(null);
  const [bounds, setBounds] = React.useState({ width: 0, height: 0 });

  React.useLayoutEffect(() => {
    const root = rootRef.current;
    const anchor = root?.parentElement;
    const viewport = root?.closest<HTMLElement>('[role="grid"]');
    if (!root || !anchor || !viewport) {
      return;
    }
    const updateBounds = () => {
      const anchorRect = anchor.getBoundingClientRect();
      const viewportRect = viewport.getBoundingClientRect();
      const anchorWidth = parseFloat(getComputedStyle(anchor).width) || anchor.offsetWidth;
      const scale = anchorWidth ? anchorRect.width / anchorWidth : 1;
      const viewportScale = viewport.offsetWidth ? viewportRect.width / viewport.offsetWidth : 1;
      const viewportRight = viewportRect.right - (viewport.offsetWidth - viewport.clientWidth - viewport.clientLeft) * viewportScale;
      const width = Math.max(anchor.clientWidth, (viewportRight - anchorRect.left) / (scale || 1));
      const height = anchorRect.height / (scale || 1);
      setBounds((previous) => previous.width === width && previous.height === height ? previous : { width, height });
    };
    updateBounds();
    const observer = new ResizeObserver(updateBounds);
    observer.observe(viewport);
    observer.observe(anchor);
    viewport.addEventListener("scroll", updateBounds);
    return () => {
      observer.disconnect();
      viewport.removeEventListener("scroll", updateBounds);
    };
  }, []);

  React.useLayoutEffect(() => {
    const text = textRef.current;
    const textarea = textareaRef.current;
    const root = rootRef.current;
    const baseline = baselineRef.current;
    const mirror = mirrorRef.current;
    if (!text || !textarea || !root || !baseline || !mirror) return;
    const computedStyle = getComputedStyle(text);
    const lineHeight = parseFloat(computedStyle.lineHeight) || parseFloat(computedStyle.fontSize) * 1.2;
    mirror.style.padding = typeof style.padding === "number" ? `${style.padding}px` : style.padding ?? "0";
    const mirrorStyle = getComputedStyle(mirror);
    const paddingTop = parseFloat(mirrorStyle.paddingTop) || 0;
    const paddingBottom = parseFloat(mirrorStyle.paddingBottom) || 0;
    const spareHeight = Math.max(0, bounds.height - paddingTop - paddingBottom - lineHeight);
    const editorTextTop = paddingTop + (style.verticalAlign === "top" ? 0 : style.verticalAlign === "middle" ? spareHeight / 2 : spareHeight);
    if (canvasText) mirror.style.paddingTop = `${editorTextTop}px`;
    const rootRect = root.getBoundingClientRect();
    const scale = rootRect.width / parseFloat(getComputedStyle(root).width) || 1;
    const textTop = (text.getBoundingClientRect().top - rootRect.top) / scale - baselineOffsetRef.current;
    textarea.style.paddingTop = `${Math.max(0, textTop)}px`;
    let baselineOffset = 0;
    if (canvasText) {
      const context = measuringContextRef.current ??= document.createElement("canvas").getContext("2d");
      if (context) {
        context.font = computedStyle.font;
        context.textBaseline = "alphabetic";
        const alphabetic = context.measureText("M");
        context.textBaseline = "middle";
        const middle = context.measureText("M");
        const browserBaseline = (baseline.getBoundingClientRect().top - rootRect.top) / scale - baselineOffsetRef.current;
        baselineOffset = editorTextTop + lineHeight / 2 + alphabetic.actualBoundingBoxAscent
          - middle.actualBoundingBoxAscent - browserBaseline;
      }
    }
    baselineOffsetRef.current = baselineOffset;
    root.style.setProperty("--xlsx-editor-baseline-offset", `${baselineOffset}px`);
    const canvas = canvasRef.current;
    if (canvas && paintCanvasText) {
      const dpr = window.devicePixelRatio || 1;
      const offsetX = ((canvasOffset?.left ?? 0) * dpr % 1) / dpr;
      const offsetY = ((canvasOffset?.top ?? 0) * dpr % 1) / dpr;
      const width = rootRect.width / scale;
      const height = rootRect.height / scale;
      canvas.width = Math.ceil((width + offsetX) * dpr);
      canvas.height = Math.ceil((height + offsetY) * dpr);
      canvas.style.width = `${canvas.width / dpr}px`;
      canvas.style.height = `${canvas.height / dpr}px`;
      canvas.style.left = `${-offsetX}px`;
      canvas.style.top = `${-offsetY}px`;
      const context = canvas.getContext("2d");
      if (context) {
        const node = plainTextRef.current?.firstChild;
        const lines: EditorTextLine[] = [];
        if (node) {
          const range = document.createRange();
          const length = node.textContent?.length ?? 0;
          let start = 0;
          while (start < length) {
            range.setStart(node, start);
            range.setEnd(node, start + 1);
            const top = range.getBoundingClientRect().top;
            let low = start + 1;
            let high = length;
            while (low < high) {
              const middle = Math.floor((low + high) / 2);
              range.setStart(node, middle);
              range.setEnd(node, middle + 1);
              if (Math.abs(range.getBoundingClientRect().top - top) < 1) low = middle + 1;
              else high = middle;
            }
            lines.push({ start, text: value.slice(start, low).replace(/\r?\n$/, "") });
            start = low;
          }
        }
        context.scale(dpr, dpr);
        context.translate(offsetX, offsetY);
        paintCanvasText(context, width, height, lines, editorTextTop);
      }
    }
  });

  const parts: React.ReactNode[] = [];
  let offset = 0;
  for (const reference of references) {
    parts.push(value.slice(offset, reference.start));
    parts.push(<span key={reference.start} style={{ color: reference.color }}>{value.slice(reference.start, reference.end)}</span>);
    offset = reference.end;
  }
  parts.push(value.slice(offset));
  return (
    <div
      ref={rootRef}
      onPointerDown={(event) => event.stopPropagation()}
      onClick={(event) => event.stopPropagation()}
      onDoubleClick={(event) => event.stopPropagation()}
      style={{
        ...style,
        boxSizing: "border-box",
        position: "absolute",
        top: 0,
        left: 0,
        minHeight: "100%",
        minWidth: "100%",
        width: "max-content",
        maxWidth: bounds.width || "100%",
        textAlign: style.textAlign ?? "left",
        whiteSpace: "pre-wrap",
        wordBreak: canvasText ? "break-all" : undefined,
        overflowWrap: "anywhere",
        padding: 0,
        lineHeight: style.lineHeight ?? "inherit",
        overflow: "hidden"
      }}
    >
      <div
        ref={mirrorRef}
        aria-hidden="true"
        data-xlsx-formula-text=""
        style={{
          boxSizing: "border-box",
          padding: style.padding,
          minHeight: bounds.height,
          display: canvasText ? "block" : "flex",
          flexDirection: "column",
          justifyContent: style.verticalAlign === "top" ? "flex-start" : style.verticalAlign === "middle" ? "center" : "flex-end",
          overflow: "hidden",
          visibility: paintCanvasText ? "hidden" : undefined,
          pointerEvents: "none"
        }}
      >
        <span ref={textRef} style={{ display: "block", transform: "translateY(var(--xlsx-editor-baseline-offset, 0px))" }}>
          <span ref={baselineRef} style={{ display: "inline-block", width: 0, height: 0, verticalAlign: "baseline" }} />
          {paintCanvasText ? <span ref={plainTextRef}>{!value || /[\r\n]$/.test(value) ? value + "\u200b" : value}</span> : <>{parts}{"\u200b"}</>}
        </span>
      </div>
      {paintCanvasText ? <canvas ref={canvasRef} aria-hidden="true" style={{ position: "absolute", pointerEvents: "none" }} /> : null}
      <textarea
        ref={(element) => {
          textareaRef.current = element;
          if (typeof inputRef === "function") inputRef(element);
          else if (inputRef) inputRef.current = element;
        }}
        aria-label="Edit cell"
        autoFocus
        autoComplete="off"
        spellCheck={false}
        rows={1}
        value={value}
        onBlur={onBlur}
        onChange={(event) => onChange(event.target.value)}
        onScroll={(event) => {
          if (mirrorRef.current) mirrorRef.current.scrollTop = event.currentTarget.scrollTop;
        }}
        onKeyDown={(event) => {
          event.stopPropagation();
          if (event.nativeEvent.isComposing) return;
          if (event.key === "Enter" && !event.altKey) {
            event.preventDefault();
            onCommit();
          } else if (event.key === "Escape") {
            event.preventDefault();
            onCancel();
          } else if (event.key === "Tab") {
            event.preventDefault();
            onNavigate(event.shiftKey ? "ArrowLeft" : "ArrowRight", event.currentTarget.value);
          } else if ((event.key === "ArrowDown" || event.key === "ArrowUp")
            && event.currentTarget.clientHeight < parseFloat(getComputedStyle(event.currentTarget).lineHeight) * 2) {
            event.preventDefault();
            onNavigate(event.key, event.currentTarget.value);
          }
        }}
        style={{
          position: "absolute",
          inset: 0,
          background: "transparent",
          border: 0,
          boxSizing: "border-box",
          color: "transparent",
          caretColor: style.color ?? "currentColor",
          font: "inherit",
          lineHeight: "inherit",
          letterSpacing: "inherit",
          textAlign: "inherit",
          height: "100%",
          width: "100%",
          minWidth: 0,
          whiteSpace: "pre-wrap",
          wordBreak: canvasText ? "break-all" : undefined,
          overflowWrap: "anywhere",
          resize: "none",
          overflow: "hidden",
          padding: style.padding,
          margin: 0,
          outline: "none",
          transform: "translateY(var(--xlsx-editor-baseline-offset, 0px))"
        }}
      />
    </div>
  );
}
