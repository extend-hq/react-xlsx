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
import { findIntersectingPivot } from "./pivots";

type MenuErrorKey = "clipboardCopyUnavailable" | "clipboardPasteUnavailable" | "nothingToPaste" | "actionFailed";

class MenuActionError extends Error {
  readonly key: MenuErrorKey;

  constructor(key: MenuErrorKey) {
    super(key);
    this.key = key;
  }
}

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
  const { messages } = controller;
  const rootRef = React.useRef<HTMLDivElement>(null);
  const abortController = React.useMemo(() => new AbortController(), []);
  const controllerRef = React.useRef(controller);
  controllerRef.current = controller;
  const aliveRef = React.useRef(true);
  const runningRef = React.useRef(false);
  const [busy, setBusy] = React.useState(false);
  const [errorKey, setErrorKey] = React.useState<MenuErrorKey | null>(null);
  const error = errorKey ? messages[errorKey] : null;
  const [location, setLocation] = React.useState(position);
  const modifier = typeof navigator !== "undefined" && /Mac|iPhone|iPad/.test(navigator.platform) ? "⌘" : "Ctrl+";
  const table = controller.tables.find(({ start, end }) =>
    context.cell.row >= start.row && context.cell.row <= end.row &&
    context.cell.col >= start.col && context.cell.col <= end.col);
  const range = expandHighlightedRange(context.selection, controller.activeSheet?.mergedRanges ?? []);
  const multiple = range.start.row !== range.end.row || range.start.col !== range.end.col;
  const touchesPivot = Boolean(findIntersectingPivot(controller.pivotTables, controller.activeSheet?.name, range));
  const pivot = controller.getPivotTableAtCell(context.cell);
  const actions: XlsxContextMenuAction[] = [
    ...(pivot ? [{ id: "togglePivotFieldList" as const, label: controller.pivotFieldListVisible ? messages.hidePivotFieldList : messages.showPivotFieldList, checked: controller.pivotFieldListVisible }] : []),
    ...(pivot && !controller.readOnly ? [
      { id: "refreshPivotTable" as const, label: messages.refreshPivotTable, disabled: pivot.unsupportedFeatures.length > 0 },
      { id: "showPivotDetails" as const, label: messages.showPivotDetails, disabled: pivot.unsupportedFeatures.length > 0 }
    ] : []),
    ...(!controller.readOnly ? [{ id: "cut" as const, label: messages.cut, disabled: touchesPivot }] : []),
    { id: "copy", label: messages.copy, shortcut: `${modifier}C` },
    ...(!controller.readOnly ? [
      { id: "paste" as const, label: messages.paste, shortcut: `${modifier}V`, disabled: touchesPivot },
      { id: "editCell" as const, label: messages.editCell, shortcut: "F2", separatorBefore: true, disabled: touchesPivot },
      { id: "clearContents" as const, label: messages.clearContents, destructive: true, disabled: touchesPivot },
      { id: "undo" as const, label: messages.undo, shortcut: `${modifier}Z`, separatorBefore: true, disabled: !controller.canUndo },
      { id: "redo" as const, label: messages.redo, shortcut: `${modifier}⇧Z`, disabled: !controller.canRedo },
      { id: "mergeAndCenter" as const, label: messages.mergeAndCenter, separatorBefore: true, disabled: !multiple || touchesPivot },
      { id: "mergeAcross" as const, label: messages.mergeAcross, disabled: range.start.col === range.end.col || touchesPivot },
      { id: "mergeCells" as const, label: messages.mergeCells, disabled: !multiple || touchesPivot },
      { id: "unmergeCells" as const, label: messages.unmergeCells, disabled: touchesPivot },
      ...(table ? [
        { id: "sortAscending" as const, label: messages.sortAscending, separatorBefore: true },
        { id: "sortDescending" as const, label: messages.sortDescending }
      ] : [])
    ] : []),
    { id: "selectRow", label: messages.selectRow, separatorBefore: true },
    { id: "selectColumn", label: messages.selectColumn },
    { id: "selectAll", label: messages.selectAll, shortcut: `${modifier}A` }
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
    (root.querySelector<HTMLElement>(':is([role="menuitem"],[role="menuitemcheckbox"]):not(:disabled):not([aria-disabled="true"])') ?? root).focus();
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
    document.addEventListener("wheel", outside, { capture: true, passive: true });
    document.addEventListener("gesturestart", outside, true);
    window.addEventListener("resize", dismiss);
    window.addEventListener("blur", dismiss);
    return () => {
      aliveRef.current = false;
      abortController.abort();
      document.removeEventListener("pointerdown", outside, true);
      document.removeEventListener("focusin", outside);
      document.removeEventListener("scroll", outside, true);
      document.removeEventListener("wheel", outside, true);
      document.removeEventListener("gesturestart", outside, true);
      window.removeEventListener("resize", dismiss);
      window.removeEventListener("blur", dismiss);
    };
  }, [abortController, closeMenu]);

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
    setErrorKey(null);
    try {
      switch (id) {
        case "togglePivotFieldList":
          controller.setPivotFieldListVisible(!controller.pivotFieldListVisible);
          break;
        case "refreshPivotTable":
          if (pivot) controller.refreshPivotTable(pivot.id);
          break;
        case "showPivotDetails":
          if (pivot) controller.showPivotDetails(pivot.id, context.cell.row, context.cell.col);
          break;
        case "copy":
        case "cut": {
          const copied = await controller.copySelectionToClipboard();
          if (!copied) throw new MenuActionError("clipboardCopyUnavailable");
          if (!isCurrent()) return false;
          if (id === "cut") controller.clearSelectedCells();
          break;
        }
        case "paste": {
          const clipboard = navigator.clipboard;
          if (!clipboard) throw new MenuActionError("clipboardPasteUnavailable");
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
          if (!pasted) throw new MenuActionError("nothingToPaste");
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
      if (aliveRef.current) setErrorKey(cause instanceof MenuActionError ? cause.key : "actionFailed");
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
        role={action.checked === undefined ? "menuitem" : "menuitemcheckbox"}
        aria-checked={action.checked}
        tabIndex={-1}
        type="button"
      >
        <span>{action.checked !== undefined ? <span aria-hidden="true" style={{ display: "inline-block", width: 18 }}>{action.checked ? "✓" : ""}</span> : null}{action.label}</span>
        {action.shortcut && <span className="react-xlsx-context-menu-shortcut">{action.shortcut}</span>}
      </button>
    </React.Fragment>
  ));
  const close = () => { abortController.abort(); closeMenu(); };
  const commit = (apply: (target: XlsxContextMenuContext) => void) => {
    if (!aliveRef.current || abortController.signal.aborted || controllerRef.current.readOnly
      || controllerRef.current.workbook !== controller.workbook || controllerRef.current.activeSheetIndex !== context.sheetIndex
      || controllerRef.current.selection !== controller.selection || controllerRef.current.revision !== controller.revision) return false;
    apply(context);
    close();
    return true;
  };
  const content = renderContextMenu
    ? renderContextMenu({ context, actions, runAction, closeMenu: close, signal: abortController.signal, commit, position: location, isDark: isDarkViewerPalette(palette), busy, error, defaultNode })
    : defaultNode;

  return createPortal(
    <div
      lang={controller.locale}
      aria-label={messages.contextMenu}
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
        const items = Array.from(rootRef.current?.querySelectorAll<HTMLElement>(':is([role="menuitem"],[role="menuitemcheckbox"]):not(:disabled):not([aria-disabled="true"])') ?? []);
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
