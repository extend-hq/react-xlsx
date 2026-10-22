import * as React from "react";
import { createPortal } from "react-dom";
import type {
  XlsxContextMenuAction,
  XlsxContextMenuActionId,
  XlsxContextMenuContext,
  XlsxViewerController,
  XlsxViewerProps
} from "./types";
import { isDarkViewerPalette, type ViewerPalette } from "./viewer-palette";
import { expandHighlightedRange } from "./range-highlight";

const menuStyles = `
:where(.react-xlsx-context-menu) {
  background: var(--xlsx-context-menu-background);
  color: var(--xlsx-context-menu-color);
  border: 1px solid var(--xlsx-context-menu-border);
  border-radius: 8px;
  padding: 5px;
  min-width: min(220px, calc(100vw - 16px));
  font: 13px/1.4 system-ui, sans-serif;
  box-shadow: 0 8px 28px rgb(0 0 0 / 18%);
}
:where(.react-xlsx-context-menu-item) {
  display: flex; align-items: center; justify-content: space-between; gap: 24px;
  width: 100%; border: 0; border-radius: 4px; padding: 7px 10px;
  background: transparent; color: inherit; font: inherit; text-align: left; cursor: pointer;
}
:where(.react-xlsx-context-menu-item:focus-visible),
:where(.react-xlsx-context-menu-item:hover:not(:disabled)) {
  background: var(--xlsx-context-menu-hover); outline: none;
}
:where(.react-xlsx-context-menu-item:disabled) { opacity: .45; cursor: default; }
:where(.react-xlsx-context-menu-item[data-destructive="true"]) { color: var(--xlsx-context-menu-danger); }
:where(.react-xlsx-context-menu-shortcut) { font-size: 11px; color: var(--xlsx-context-menu-muted); }
:where(.react-xlsx-context-menu-separator) { height: 1px; margin: 4px; background: var(--xlsx-context-menu-border); }
`;

export function XlsxContextMenu({
  context, position, controller, palette, closeMenu, editCell, selectExtent,
  renderContextMenu, contextMenuClassName, contextMenuStyle
}: Pick<XlsxViewerProps, "renderContextMenu" | "contextMenuClassName" | "contextMenuStyle"> & {
  context: XlsxContextMenuContext;
  position: { x: number; y: number };
  controller: XlsxViewerController;
  palette: ViewerPalette;
  closeMenu: (restoreFocus?: boolean) => void;
  editCell: () => void;
  selectExtent: (kind: "row" | "column" | "sheet") => void;
}) {
  const rootRef = React.useRef<HTMLDivElement>(null);
  const controllerRef = React.useRef(controller);
  controllerRef.current = controller;
  const aliveRef = React.useRef(true);
  const runningRef = React.useRef(false);
  const [busy, setBusy] = React.useState(false);
  const [error, setError] = React.useState<string | null>(null);
  const [location, setLocation] = React.useState(position);
  const modifier = typeof navigator !== "undefined" && /Mac|iPhone|iPad/.test(navigator.platform) ? "⌘" : "Ctrl+";
  const table = controller.tables.find(({ start, end }) =>
    context.cell.row >= start.row && context.cell.row <= end.row &&
    context.cell.col >= start.col && context.cell.col <= end.col);
  const range = expandHighlightedRange(context.selection, controller.activeSheet?.mergedRanges ?? []);
  const multiple = range.start.row !== range.end.row || range.start.col !== range.end.col;
  const actions: XlsxContextMenuAction[] = [
    ...(!controller.readOnly ? [{ id: "cut" as const, label: "Cut" }] : []),
    { id: "copy", label: "Copy", shortcut: `${modifier}C` },
    ...(!controller.readOnly ? [
      { id: "paste" as const, label: "Paste", shortcut: `${modifier}V` },
      { id: "editCell" as const, label: "Edit cell", shortcut: "F2", separatorBefore: true },
      { id: "clearContents" as const, label: "Clear contents", destructive: true },
      { id: "undo" as const, label: "Undo", shortcut: `${modifier}Z`, separatorBefore: true, disabled: !controller.canUndo },
      { id: "redo" as const, label: "Redo", shortcut: `${modifier}⇧Z`, disabled: !controller.canRedo },
      { id: "mergeAndCenter" as const, label: "Merge & Center", separatorBefore: true, disabled: !multiple },
      { id: "mergeAcross" as const, label: "Merge Across", disabled: range.start.col === range.end.col },
      { id: "mergeCells" as const, label: "Merge Cells", disabled: !multiple },
      { id: "unmergeCells" as const, label: "Unmerge Cells" },
      ...(table ? [
        { id: "sortAscending" as const, label: "Sort A to Z", separatorBefore: true },
        { id: "sortDescending" as const, label: "Sort Z to A" }
      ] : [])
    ] : []),
    { id: "selectRow", label: "Select row", separatorBefore: true },
    { id: "selectColumn", label: "Select column" },
    { id: "selectAll", label: "Select all", shortcut: `${modifier}A` }
  ];

  React.useLayoutEffect(() => {
    const root = rootRef.current;
    if (!root) return;
    const update = () => {
      const rect = root.getBoundingClientRect();
      setLocation({
        x: Math.max(8, Math.min(position.x, window.innerWidth - rect.width - 8)),
        y: Math.max(8, Math.min(position.y, window.innerHeight - rect.height - 8))
      });
    };
    update();
    const observer = new ResizeObserver(update);
    observer.observe(root);
    (root.querySelector<HTMLElement>('[role="menuitem"]:not(:disabled):not([aria-disabled="true"])') ?? root).focus();
    return () => observer.disconnect();
  }, [position]);

  React.useEffect(() => {
    aliveRef.current = true;
    const outside = (event: Event) => {
      if (!rootRef.current?.contains(event.target as Node)) closeMenu(false);
    };
    const dismiss = () => closeMenu(false);
    document.addEventListener("pointerdown", outside, true);
    document.addEventListener("focusin", outside);
    document.addEventListener("scroll", outside, true);
    window.addEventListener("resize", dismiss);
    window.addEventListener("blur", dismiss);
    return () => {
      aliveRef.current = false;
      document.removeEventListener("pointerdown", outside, true);
      document.removeEventListener("focusin", outside);
      document.removeEventListener("scroll", outside, true);
      window.removeEventListener("resize", dismiss);
      window.removeEventListener("blur", dismiss);
    };
  }, [closeMenu]);

  const runAction = async (id: XlsxContextMenuActionId): Promise<boolean> => {
    const action = actions.find((item) => item.id === id);
    if (!aliveRef.current || runningRef.current || !action || action.disabled) return false;
    const snapshot = controller;
    const isCurrent = () => aliveRef.current &&
      controllerRef.current.activeSheetIndex === snapshot.activeSheetIndex &&
      controllerRef.current.workbook === snapshot.workbook &&
      controllerRef.current.revision === snapshot.revision &&
      controllerRef.current.readOnly === snapshot.readOnly &&
      controllerRef.current.selection === snapshot.selection;
    if (!isCurrent()) return false;
    runningRef.current = true;
    setBusy(true);
    setError(null);
    try {
      switch (id) {
        case "copy":
        case "cut": {
          const copied = await controller.copySelectionToClipboard();
          if (!copied) throw new Error("Clipboard access is unavailable. Use the keyboard shortcut to copy.");
          if (!isCurrent()) return false;
          if (id === "cut") controller.clearSelectedCells();
          break;
        }
        case "paste": {
          const clipboard = navigator.clipboard;
          if (!clipboard) throw new Error("Clipboard access is unavailable. Use the keyboard shortcut to paste.");
          let structured: string | undefined;
          let text: string | undefined;
          if (clipboard.read) {
            const items = await clipboard.read();
            for (const item of items) {
              const type = ["web application/x-react-xlsx-range+json", "application/x-react-xlsx-range+json"]
                .find((type) => item.types.includes(type));
              if (type) {
                structured = await (await item.getType(type)).text();
                break;
              }
              if (text === undefined && item.types.includes("text/plain")) text = await (await item.getType("text/plain")).text();
            }
          } else {
            text = await clipboard.readText();
          }
          if (!isCurrent()) return false;
          const pasted = structured !== undefined
            ? controller.pasteStructuredClipboardData(structured)
            : text !== undefined && controller.pasteText(text);
          if (!pasted) throw new Error("No spreadsheet cells or text are available to paste.");
          break;
        }
        case "clearContents": controller.clearSelectedCells(); break;
        case "editCell": closeMenu(); editCell(); return true;
        case "undo": controller.undo(); break;
        case "redo": controller.redo(); break;
        case "mergeCells": controller.mergeSelection(); break;
        case "mergeAndCenter": controller.mergeSelection("center"); break;
        case "mergeAcross": controller.mergeSelection("across"); break;
        case "unmergeCells": controller.unmergeSelection(); break;
        case "sortAscending":
        case "sortDescending":
          if (table) controller.sortTable(table.name, context.cell.col - table.start.col, id === "sortAscending" ? "ascending" : "descending");
          break;
        case "selectRow": selectExtent("row"); break;
        case "selectColumn": selectExtent("column"); break;
        case "selectAll": selectExtent("sheet"); break;
      }
      closeMenu();
      return true;
    } catch (cause) {
      if (aliveRef.current) setError(cause instanceof Error ? cause.message : "The action could not be completed.");
      return false;
    } finally {
      runningRef.current = false;
      if (aliveRef.current) setBusy(false);
    }
  };

  const defaultNode = actions.map((action) => (
    <React.Fragment key={action.id}>
      {action.separatorBefore && <div className="react-xlsx-context-menu-separator" role="separator" />}
      <button
        className="react-xlsx-context-menu-item"
        data-action={action.id}
        data-destructive={action.destructive || undefined}
        disabled={busy || action.disabled}
        onClick={() => { void runAction(action.id); }}
        role="menuitem"
        tabIndex={-1}
        type="button"
      >
        <span>{action.label}</span>
        {action.shortcut && <span className="react-xlsx-context-menu-shortcut">{action.shortcut}</span>}
      </button>
    </React.Fragment>
  ));
  const content = renderContextMenu
    ? renderContextMenu({ context, actions, runAction, closeMenu, position: location, isDark: isDarkViewerPalette(palette), busy, error, defaultNode })
    : defaultNode;

  return createPortal(
    <div
      aria-label="Spreadsheet context menu"
      aria-busy={busy}
      className={["react-xlsx-context-menu", contextMenuClassName].filter(Boolean).join(" ")}
      data-dark={isDarkViewerPalette(palette) ? "true" : "false"}
      onContextMenu={(event) => event.preventDefault()}
      onPointerDown={(event) => event.stopPropagation()}
      onKeyDown={(event) => {
        event.stopPropagation();
        if (event.key === "Escape" || event.key === "Tab") {
          if (event.key === "Escape") event.preventDefault();
          closeMenu();
          return;
        }
        const items = Array.from(rootRef.current?.querySelectorAll<HTMLElement>('[role="menuitem"]:not(:disabled):not([aria-disabled="true"])') ?? []);
        const current = items.indexOf(document.activeElement as HTMLElement);
        let next: number | undefined;
        if (event.key === "ArrowDown") next = (current + 1) % items.length;
        if (event.key === "ArrowUp") next = (current - 1 + items.length) % items.length;
        if (event.key === "Home") next = 0;
        if (event.key === "End") next = items.length - 1;
        if (next !== undefined) {
          event.preventDefault();
          items[next]?.focus();
        }
      }}
      ref={rootRef}
      role="menu"
      style={{
        "--xlsx-context-menu-background": palette.surface,
        "--xlsx-context-menu-color": palette.text,
        "--xlsx-context-menu-border": palette.border,
        "--xlsx-context-menu-hover": palette.mutedSurface,
        "--xlsx-context-menu-muted": palette.mutedText,
        "--xlsx-context-menu-danger": palette.danger,
        ...contextMenuStyle,
        position: "fixed", left: location.x, top: location.y,
        maxHeight: "calc(100vh - 16px)", maxWidth: "calc(100vw - 16px)",
        overflow: "auto", zIndex: 10000, boxSizing: "border-box"
      } as React.CSSProperties}
      tabIndex={-1}
    >
      <style>{menuStyles}</style>
      {content}
      {error && <div role="alert" style={{ padding: "8px 10px", maxWidth: 280, color: palette.danger }}>{error}</div>}
    </div>,
    document.body
  );
}
