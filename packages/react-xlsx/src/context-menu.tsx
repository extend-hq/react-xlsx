import { EXCEL_DRAWING_CLIPBOARD_TYPE, INTERNAL_DRAWING_CLIPBOARD_TYPE } from "./drawing-clipboard";
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
import { chartTypeOptions, chartTypeUnavailable, matchesChartType } from "./chart-types";
import { ContextMenuItems, handleMenuNavigation } from "./context-menu-items";

type MenuErrorKey = "clipboardCopyUnavailable" | "clipboardPasteUnavailable" | "nothingToPaste" | "actionFailed";

class MenuActionError extends Error {
  readonly key: MenuErrorKey;

  constructor(key: MenuErrorKey) {
    super(key);
    this.key = key;
  }
}

export const menuStyles = `
:where(.react-xlsx-context-menu) {
  background: var(--xlsx-context-menu-background);
  color: var(--xlsx-context-menu-color);
  border: 1px solid var(--xlsx-context-menu-border);
  border-radius: var(--xlsx-menu-radius, 3px);
  padding: var(--xlsx-menu-padding, 4px);
  min-width: min(220px, calc(100vw - 16px));
  font: 13px/1.4 system-ui, sans-serif;
  box-shadow: 0 8px 28px rgb(0 0 0 / 18%);
}
:where(.react-xlsx-context-menu-item) {
  display: flex; align-items: center; justify-content: space-between; gap: 24px;
  width: 100%; border: 0; border-radius: max(0px, calc(var(--xlsx-menu-radius, 3px) - var(--xlsx-menu-padding, 4px)));
  padding: 7px 10px;
  background: transparent; color: inherit; font: inherit; text-align: left; cursor: pointer;
}
:where(.react-xlsx-context-menu-item[aria-pressed="true"]) { background: var(--xlsx-menu-active, var(--xlsx-context-menu-hover)); }
:where(.react-xlsx-context-menu-item:focus-visible),
:where(.react-xlsx-context-menu-item:hover:not(:disabled)) {
  background: var(--xlsx-context-menu-hover); outline: none;
}
:where(.react-xlsx-context-menu-item:disabled) { opacity: .45; cursor: default; }
:where(.react-xlsx-context-menu-item[data-destructive="true"]) { color: var(--xlsx-context-menu-danger); }
:where(.react-xlsx-context-menu-shortcut) { font-size: 11px; color: var(--xlsx-context-menu-muted); }
:where(.react-xlsx-context-menu-separator) { height: 1px; margin: 4px; background: var(--xlsx-context-menu-border); }
`;

export function resolveMenuPaletteStyle(palette: ViewerPalette): React.CSSProperties {
  return {
    "--xlsx-context-menu-background": palette.surface,
    "--xlsx-context-menu-color": palette.text,
    "--xlsx-context-menu-border": palette.border,
    "--xlsx-context-menu-hover": palette.mutedSurface,
    "--xlsx-context-menu-muted": palette.mutedText,
    "--xlsx-context-menu-danger": palette.danger
  } as React.CSSProperties;
}

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
  const [errorMessage, setErrorMessage] = React.useState<string | null>(null);
  const error = errorMessage ?? (errorKey ? messages[errorKey] : null);
  const [location, setLocation] = React.useState(position);
  const modifier = typeof navigator !== "undefined" && /Mac|iPhone|iPad/.test(navigator.platform) ? "⌘" : "Ctrl+";
  const table = controller.tables.find(({ start, end }) =>
    context.cell.row >= start.row && context.cell.row <= end.row &&
    context.cell.col >= start.col && context.cell.col <= end.col);
  const range = expandHighlightedRange(context.selection, controller.activeSheet?.mergedRanges ?? []);
  const multiple = range.start.row !== range.end.row || range.start.col !== range.end.col;
  const touchesPivot = Boolean(findIntersectingPivot(controller.pivotTables, controller.activeSheet?.name, range));
  const pivot = controller.getPivotTableAtCell(context.cell);
  const cellActions: XlsxContextMenuAction[] = [
    ...(!controller.readOnly ? [{ id: "cut" as const, label: messages.cut, shortcut: `${modifier}X`, disabled: touchesPivot }] : []),
    { id: "copy", label: messages.copy, shortcut: `${modifier}C` },
    ...(!controller.readOnly ? [
      { id: "undo" as const, label: messages.undo, shortcut: `${modifier}Z`, disabled: !controller.canUndo },
      { id: "redo" as const, label: messages.redo, shortcut: `${modifier}⇧Z`, disabled: !controller.canRedo },
      { id: "paste" as const, label: messages.paste, shortcut: `${modifier}V`, separatorBefore: true, disabled: touchesPivot }
    ] : []),
    ...(pivot ? [{ id: "togglePivotFieldList" as const, label: controller.pivotFieldListVisible ? messages.hidePivotFieldList : messages.showPivotFieldList, checked: controller.pivotFieldListVisible }] : []),
    ...(pivot && !controller.readOnly ? [
      { id: "refreshPivotTable" as const, label: messages.refreshPivotTable, disabled: pivot.unsupportedFeatures.length > 0 },
      { id: "showPivotDetails" as const, label: messages.showPivotDetails, disabled: pivot.unsupportedFeatures.length > 0 }
    ] : []),
    ...(!controller.readOnly ? [
      { id: "editCell" as const, label: messages.editCell, shortcut: "F2", separatorBefore: true, disabled: touchesPivot },
      { id: "clearContents" as const, label: messages.clearContents, destructive: true, disabled: touchesPivot },
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

  const drawing = context.chart ?? context.image;
  const drawingDisabled = controller.readOnly || drawing?.editable === false;
  const drawings = [...controller.charts, ...controller.images, ...controller.shapes, ...controller.formControls];
  const minOrder = Math.min(...drawings.map(item => item.zIndex));
  const maxOrder = Math.max(...drawings.map(item => item.zIndex));
  const actions: XlsxContextMenuAction[] = drawing ? [
    { id: "cut", label: messages.cut, shortcut: `${modifier}X`, disabled: drawingDisabled },
    { id: "copy", label: messages.copy, shortcut: `${modifier}C` },
    { id: "undo", label: messages.undo, shortcut: `${modifier}Z`, disabled: controller.readOnly || !controller.canUndo },
    { id: "redo", label: messages.redo, shortcut: `${modifier}⇧Z`, disabled: controller.readOnly || !controller.canRedo },
    ...(context.chart ? [{ id: "changeChartType" as const, label: "Change Chart Type", separatorBefore: true, disabled: drawingDisabled,
      children: [...new Set(chartTypeOptions.map(option => option.family))].map(family => ({
        id: `chartFamily:${family}` as const, label: family,
        children: chartTypeOptions.filter(option => option.family === family).map(option => ({
          id: `chartType:${option.id}` as const, label: option.label,
          checked: matchesChartType(context.chart!, option), disabled: chartTypeUnavailable(context.chart!, option)
        }))
      }))
    }] : []),
    { id: "bringToFront", label: "Bring to Front", separatorBefore: !context.chart, disabled: drawingDisabled || drawing.zIndex >= maxOrder,
      children: [
        { id: "bringToFront", label: "Bring to Front", disabled: drawingDisabled || drawing.zIndex >= maxOrder },
        { id: "bringForward", label: "Bring Forward", disabled: drawingDisabled || drawing.zIndex >= maxOrder }
      ] },
    { id: "sendToBack", label: "Send to Back", disabled: drawingDisabled || drawing.zIndex <= minOrder,
      children: [
        { id: "sendToBack", label: "Send to Back", disabled: drawingDisabled || drawing.zIndex <= minOrder },
        { id: "sendBackward", label: "Send Backward", disabled: drawingDisabled || drawing.zIndex <= minOrder }
      ] },

  ] : cellActions;
  const findAction = (items: XlsxContextMenuAction[], id: XlsxContextMenuActionId): XlsxContextMenuAction | undefined => {
    for (const item of items) {
      if (item.disabled) continue;
      if (item.children) { const found = findAction(item.children, id); if (found) return found; }
      else if (item.id === id) return item;
    }
    return undefined;
  };

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
    const action = findAction(actions, id);
    if (!aliveRef.current || runningRef.current || !action || action.disabled) return false;
    const snapshot = controller;
    const isCurrent = () => aliveRef.current &&
      controllerRef.current.activeSheetIndex === snapshot.activeSheetIndex &&
      controllerRef.current.workbook === snapshot.workbook &&
      controllerRef.current.revision === snapshot.revision &&
      controllerRef.current.readOnly === snapshot.readOnly &&
      controllerRef.current.selection === snapshot.selection &&
      controllerRef.current.selectedChartId === snapshot.selectedChartId &&
      controllerRef.current.selectedImageId === snapshot.selectedImageId;
    if (!isCurrent()) return false;
    runningRef.current = true;
    setBusy(true);
    setErrorKey(null);
    setErrorMessage(null);
    try {
      if (id.startsWith("chartType:") && context.chart) {
        if (!await controller.changeChartType(context.chart.id, id.slice("chartType:".length))) throw new MenuActionError("actionFailed");
        closeMenu();
        return true;
      }
      switch (id) {
        case "bringToFront": case "bringForward": case "sendBackward": case "sendToBack":
          if (!drawing || !await controller.arrangeDrawing(context.chart ? "chart" : "image", drawing.id, id)) throw new MenuActionError("actionFailed");
          break;
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
          if (drawing) {
            if (!await controller.copyDrawingToClipboard(context.chart ? "chart" : "image", drawing.id, id === "cut")) throw new MenuActionError("clipboardCopyUnavailable");
            break;
          }
          const copied = await controller.copySelectionToClipboard();
          if (!copied) throw new MenuActionError("clipboardCopyUnavailable");
          if (!isCurrent()) return false;
          if (id === "cut") controller.clearSelectedCells();
          break;
        }
        case "paste": {
          if (controller.canReadNativeDrawingClipboard) {
            if (!await controller.pasteFromClipboard()) throw new MenuActionError("nothingToPaste");
            break;
          }
          const clipboard = navigator.clipboard;
          if (!clipboard) throw new MenuActionError("clipboardPasteUnavailable");
          let imageOnly = false;
          let hasDrawing = false;
          let structured: string | undefined;
          let text: string | undefined;
          if (clipboard.read) {
            const items = await clipboard.read();
            imageOnly = items.some(item => item.types.includes("image/png") && !item.types.includes("text/plain") && !item.types.includes("text/html"));
            for (const item of items) {
              if ([INTERNAL_DRAWING_CLIPBOARD_TYPE, EXCEL_DRAWING_CLIPBOARD_TYPE].some(type => item.types.includes(type))) { hasDrawing = true; break; }
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
          if (imageOnly || hasDrawing || text?.startsWith("react-xlsx-drawing:")) {
            if (!await controller.pasteDrawingFromClipboard(hasDrawing ? undefined : text)) throw new MenuActionError("nothingToPaste");
            break;
          }
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
      if (aliveRef.current) {
        setErrorKey(cause instanceof MenuActionError ? cause.key : "actionFailed");
        if (!(cause instanceof MenuActionError) && cause instanceof Error) setErrorMessage(cause.message);
      }
      return false;
    } finally {
      runningRef.current = false;
      if (aliveRef.current) setBusy(false);
    }
  };

  const defaultNode = <ContextMenuItems actions={actions} busy={busy} runAction={runAction} contextMenuClassName={contextMenuClassName} contextMenuStyle={contextMenuStyle} />;
  const close = () => { abortController.abort(); closeMenu(); };
  const commit = (apply: (target: XlsxContextMenuContext) => void) => {
    if (!aliveRef.current || abortController.signal.aborted || controllerRef.current.readOnly
      || controllerRef.current.workbook !== controller.workbook || controllerRef.current.activeSheetIndex !== context.sheetIndex
      || controllerRef.current.selection !== controller.selection || controllerRef.current.revision !== controller.revision
      || controllerRef.current.selectedChartId !== controller.selectedChartId || controllerRef.current.selectedImageId !== controller.selectedImageId) return false;
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
      data-context-kind={context.kind}
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
        handleMenuNavigation(event);
      }}
      ref={rootRef}
      role="menu"
      style={{
        ...resolveMenuPaletteStyle(palette),
        ...contextMenuStyle,
        position: "fixed", left: location.x, top: location.y,
        maxHeight: "calc(100vh - 16px)", maxWidth: "calc(100vw - 16px)",
        overflow: renderContextMenu ? "visible" : "auto", zIndex: 10000, boxSizing: "border-box"
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
