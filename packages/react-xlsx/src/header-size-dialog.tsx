import * as React from "react";
import { createPortal } from "react-dom";
import type { XlsxContextMenuContext, XlsxViewerMessages } from "./types";
import type { ViewerPalette } from "./viewer-palette";

export function HeaderSizeDialog({ axis, context, sizes, maxSize, messages, locale, palette, close, resize }: {
  axis: "row" | "column";
  context: XlsxContextMenuContext;
  sizes: readonly number[];
  maxSize: number;
  messages: XlsxViewerMessages;
  locale: string;
  palette: ViewerPalette;
  close: () => void;
  resize: (indices: readonly number[], sizePx: number) => void;
}) {
  const dialog = React.useRef<HTMLDialogElement>(null);
  const input = React.useRef<HTMLInputElement>(null);
  const titleId = React.useId();
  const coordinate = axis === "row" ? "row" : "col";
  const title = axis === "row" ? messages.rowHeight : messages.columnWidth;
  const indices = React.useMemo(() => {
    const selected = new Set<number>();
    for (const area of context.selectionRanges ?? [context.selection]) {
      for (let index = area.start[coordinate]; index <= area.end[coordinate]; index++) selected.add(index);
    }
    return [...selected];
  }, [context, coordinate]);
  const [value, setValue] = React.useState(() => {
    const size = sizes[context.cell[coordinate]];
    return Number.isFinite(size) && indices.every(index => sizes[index] === size) ? String(Math.round(size * 100) / 100) : "";
  });
  const dismiss = () => {
    dialog.current?.close();
    close();
  };
  const buttonStyle: React.CSSProperties = {
    minWidth: 60, padding: "4px 12px", borderRadius: 4, font: "inherit", cursor: "pointer",
    border: `1px solid ${palette.strongBorder}`, background: palette.buttonSurface, color: palette.buttonText
  };
  React.useLayoutEffect(() => {
    dialog.current?.showModal();
    input.current?.select();
  }, []);
  return createPortal(
    <dialog ref={dialog} aria-labelledby={titleId} lang={locale}
      onCancel={event => { event.preventDefault(); dismiss(); }}
      onKeyDown={event => event.stopPropagation()}
      style={{ margin: "auto", width: 240, maxWidth: "calc(100vw - 32px)", padding: 16, borderRadius: 8,
        border: `1px solid ${palette.border}`, background: palette.surface, color: palette.text, font: "13px/1.4 system-ui, sans-serif" }}>
      <form onSubmit={event => {
        event.preventDefault();
        const size = input.current?.valueAsNumber;
        if (size === undefined || !Number.isFinite(size)) return;
        if (indices.some(index => Math.abs(sizes[index] - size) >= 0.005)) resize(indices, size);
        dismiss();
      }}>
        <div id={titleId} style={{ fontSize: 15, fontWeight: 600, marginBottom: 12 }}>{title}</div>
        <label style={{ display: "flex", alignItems: "center", gap: 12 }}>
          {messages.pixels}
          <input ref={input} type="number" min={axis === "row" ? 16 : 30} max={Math.floor(maxSize)} step="any" required
            value={value} onChange={event => setValue(event.currentTarget.value)}
            style={{ minWidth: 0, width: "100%", padding: "4px 8px", border: `1px solid ${palette.strongBorder}`,
              borderRadius: 4, background: palette.surface, color: palette.text, font: "inherit" }} />
        </label>
        <div style={{ display: "flex", justifyContent: "flex-end", gap: 8, marginTop: 16 }}>
          <button type="button" style={buttonStyle} onClick={dismiss}>{messages.cancel}</button>
          <button type="submit" style={{ ...buttonStyle, background: palette.text, color: palette.surface }}>{messages.ok}</button>
        </div>
      </form>
    </dialog>, document.body
  );
}
