import * as React from "react";
import { parseXlsxRange } from "./range-reference";
import type { XlsxViewerController } from "./types";
import type { XlsxPivotTable, XlsxPivotSourceField, XlsxPivotDrillDown } from "./pivot-types";

export type XlsxPivotFieldArea = "rows" | "columns" | "values" | "reportFilters";
export interface XlsxPivotTableEditorOptions {
  autoSelect?: boolean;
  initialOpen?: boolean;
  enabled?: boolean;
}
export interface XlsxPivotTableEditor {
  viewer: XlsxViewerController;
  readOnly: boolean;
  open: boolean;
  setOpen: (open: boolean) => void;
  draft: XlsxPivotTable | null;
  dirty: boolean;
  stale: boolean;
  sourceNeedsLoad: boolean;
  error: string | null;
  status: string | null;
  edit: (change: (pivot: XlsxPivotTable) => XlsxPivotTable) => void;
  select: (id: string) => void;
  create: () => void;
  apply: () => void;
  reset: () => void;
  remove: () => void;
  refresh: (all?: boolean) => void;
  readFields: () => void;
  addField: (area: XlsxPivotFieldArea, field: number) => void;
  removeField: (area: XlsxPivotFieldArea, index: number) => void;
  moveField: (fromArea: XlsxPivotFieldArea, fromIndex: number, toArea: XlsxPivotFieldArea, toIndex?: number) => void;
  details: XlsxPivotDrillDown | null;
  detailContext: { id: string; row: number; col: number } | null;
  showDetails: () => void;
  clearDetails: () => void;
  run: (action: () => void) => void;
}
export interface XlsxPivotTableControlsRenderProps {
  editor: XlsxPivotTableEditor;
  controller: XlsxViewerController;
  defaultNode: React.ReactNode;
  isDark: boolean;
}

function address(row: number, col: number) {
  let letters = "";
  for (let n = col + 1; n > 0; n = Math.floor((n - 1) / 26)) letters = String.fromCharCode(65 + (n - 1) % 26) + letters;
  return `${letters}${row + 1}`;
}
function newAxis(field: number): XlsxPivotTable["rows"][number] {
  return { field, sort: "ascending", subtotals: true, collapsed: [] };
}
function newValue(field: XlsxPivotSourceField): XlsxPivotTable["values"][number] {
  const aggregation = field.dataType === "number" ? "sum" : "count";
  return { field: field.index, name: `${aggregation === "sum" ? "Sum" : "Count"} of ${field.name}`, aggregation, showAs: "normal" };
}
function copyPivot(pivot: XlsxPivotTable) { return structuredClone(pivot); }
type DetailRows = XlsxPivotDrillDown;
type DetailContext = { id: string; row: number; col: number };

export function usePivotTableEditor(viewer: XlsxViewerController, options: XlsxPivotTableEditorOptions = {}): XlsxPivotTableEditor {
  const open = viewer.pivotFieldListVisible;
  const setOpen = viewer.setPivotFieldListVisible;
  const [draft, setDraft] = React.useState<XlsxPivotTable | null>(null);
  const [dirty, setDirty] = React.useState(false);
  const [error, setError] = React.useState<string | null>(null);
  const [status, setStatus] = React.useState<string | null>(null);
  const [details, setDetails] = React.useState<DetailRows | null>(null);
  const [detailContext, setDetailContext] = React.useState<DetailContext | null>(null);
  const [loadedSource, setLoadedSource] = React.useState("");
  const [savedDefinition, setSavedDefinition] = React.useState("");
  const [pendingFocus, setPendingFocus] = React.useState<{ sheet: string; row: number; col: number } | null>(null);
  const sourceNeedsLoad = Boolean(draft && JSON.stringify(draft.source) !== loadedSource);
  const stale = Boolean(dirty && draft?.id && savedDefinition !== JSON.stringify(viewer.pivotTables.find(pivot => pivot.id === draft.id) ?? null));
  const run = React.useCallback((action: () => void) => {
    setError(null);
    setStatus(null);
    try { action(); } catch (cause) { setError(cause instanceof Error ? cause.message : String(cause)); setOpen(true); }
  }, []);

  React.useEffect(() => {
    if (options.enabled === false) return;
    setDraft(null); setDirty(false); setError(null); setStatus(null); setDetails(null); setDetailContext(null); setLoadedSource(""); setSavedDefinition(""); setPendingFocus(null);
    if (options.initialOpen !== undefined) setOpen(options.initialOpen);
  }, [viewer.src, viewer.file]);

  React.useEffect(() => {
    if (options.enabled === false || dirty || !draft?.id) return;
    const current = viewer.pivotTables.find(pivot => pivot.id === draft.id);
    setDraft(current ? copyPivot(current) : null);
    if (current) { setLoadedSource(JSON.stringify(current.source)); setSavedDefinition(JSON.stringify(current)); }
  }, [viewer.pivotTables]);

  React.useEffect(() => {
    if (options.enabled === false || options.autoSelect === false || dirty || pendingFocus || !viewer.activeCell || !viewer.activeSheet) return;
    const current = viewer.getPivotTableAtCell(viewer.activeCell, viewer.activeSheet.name);
    if (!current || current.id === draft?.id) return;
    setDraft(copyPivot(current)); setLoadedSource(JSON.stringify(current.source)); setSavedDefinition(JSON.stringify(current)); setError(null); setStatus(null); setDetails(null);
  }, [viewer.activeCell?.row, viewer.activeCell?.col, viewer.activeSheet?.name, viewer.pivotTables, options.autoSelect, options.enabled]);

  React.useEffect(() => {
    if (!pendingFocus || viewer.activeSheet?.name.toLowerCase() !== pendingFocus.sheet.toLowerCase()) return;
    viewer.selectCell({ row: pendingFocus.row, col: pendingFocus.col });
    setPendingFocus(null);
  }, [pendingFocus, viewer.activeSheet?.name, viewer.selectCell]);

  const focus = (pivot: XlsxPivotTable) => {
    const sheetIndex = viewer.sheets.findIndex(sheet => sheet.name.toLowerCase() === pivot.target.sheet.toLowerCase());
    if (sheetIndex >= 0) viewer.setActiveSheetIndex(sheetIndex);
    const parsed = parseXlsxRange(pivot.outputRange || pivot.target.cell);
    if (parsed) setPendingFocus({ sheet: pivot.target.sheet, ...parsed.range.start });
  };
  const select = (id: string) => run(() => {
    if (dirty) throw new Error("Apply or discard your changes before switching PivotTables.");
    const pivot = viewer.pivotTables.find(item => item.id === id);
    if (!pivot) return;
    setDraft(copyPivot(pivot)); setLoadedSource(JSON.stringify(pivot.source)); setSavedDefinition(JSON.stringify(pivot)); setDirty(false); setDetails(null); setOpen(true); focus(pivot);
  });
  const edit = (change: (pivot: XlsxPivotTable) => XlsxPivotTable) => {
    if (options.enabled === false || viewer.readOnly || draft?.unsupportedFeatures.length) return;
    setDraft(current => current ? change(current) : current); setDirty(true); setStatus(null); setError(null); setDetails(null);
  };
  const create = () => run(() => {
    if (dirty) throw new Error("Apply or discard your current changes before creating another PivotTable.");
    if (!viewer.activeSheet || viewer.readOnly) return;
    const sheet = viewer.activeSheet;
    const selected = viewer.selection;
    const sourceRange = selected && (selected.start.row !== selected.end.row || selected.start.col !== selected.end.col)
      ? `${address(selected.start.row, selected.start.col)}:${address(selected.end.row, selected.end.col)}`
      : `${address(sheet.minUsedRow, sheet.minUsedCol)}:${address(sheet.maxUsedRow, sheet.maxUsedCol)}`;
    let index = 1;
    while (viewer.pivotTables.some(pivot => pivot.name === `PivotTable${index}`)) index++;
    const next: XlsxPivotTable = {
      id: "", name: `PivotTable${index}`, source: { sheet: sheet.name, range: sourceRange },
      target: { sheet: sheet.name, cell: address(sheet.minUsedRow, Math.min(sheet.maxUsedCol + 3, 16_383)) },
      rows: [], columns: [], values: [], filters: [], reportFilters: [], calculatedItems: [],
      layout: "compact", repeatLabels: false, rowGrandTotals: true, columnGrandTotals: true,
      subtotalsPosition: "bottom", styleName: "PivotStyleLight16", fields: [], unsupportedFeatures: [], refreshOnLoad: false, valuesOnRows: false,
    };
    setDraft(next); setDirty(true); setOpen(true); setDetails(null); setSavedDefinition("");
    const fields = viewer.getPivotSourceFields(next.source);
    next.fields = fields;
    setLoadedSource(JSON.stringify(next.source));
    setDraft({ ...next });
  });
  const readFields = () => run(() => {
    if (!draft) return;
    const fields = viewer.getPivotSourceFields(draft.source);
    setLoadedSource(JSON.stringify(draft.source));
    const sameSchema = fields.length === draft.fields.length && fields.every((field, index) => field.name === draft.fields[index]?.name);
    edit(pivot => ({ ...pivot, fields, ...(sameSchema ? {} : { rows: [], columns: [], values: [], filters: [], reportFilters: [], calculatedItems: [] }) }));
    setStatus(sameSchema ? "Source fields updated." : "Fields loaded. Choose the fields to summarize.");
  });
  const apply = () => run(() => {
    if (!draft) return;
    if (stale) throw new Error("This PivotTable changed while you were editing. Discard your draft to load its current definition.");
    if (sourceNeedsLoad) throw new Error("Load the source fields on the Source tab before applying these changes.");
    const result = viewer.upsertPivotTable(draft);
    setDraft(copyPivot(result)); setLoadedSource(JSON.stringify(result.source)); setSavedDefinition(JSON.stringify(result)); setDirty(false); setStatus("PivotTable updated."); setDetails(null); focus(result);
  });
  const refresh = (all = false) => run(() => {
    if (dirty) throw new Error("Apply or discard your field changes before refreshing.");
    if (all) {
      const refreshed = viewer.refreshAllPivotTables();
      setStatus(`${refreshed.length} PivotTable${refreshed.length === 1 ? "" : "s"} refreshed.`);
    } else if (draft?.id) {
      const refreshed = viewer.refreshPivotTable(draft.id);
      setDraft(copyPivot(refreshed)); setLoadedSource(JSON.stringify(refreshed.source)); setSavedDefinition(JSON.stringify(refreshed)); setDirty(false); setStatus("PivotTable refreshed from its source.");
    }
    setDetails(null);
  });
  const remove = () => run(() => {
    if (dirty) throw new Error("Apply or discard your changes before removing the report.");
    if (!draft?.id) return;
    if (viewer.removePivotTable(draft.id)) { setDraft(null); setDirty(false); setStatus("PivotTable removed. Undo restores it."); setDetails(null); }
  });
  const reset = () => run(() => {
    const saved = viewer.pivotTables.find(pivot => pivot.id === draft?.id);
    setDraft(saved ? copyPivot(saved) : null); setLoadedSource(saved ? JSON.stringify(saved.source) : ""); setSavedDefinition(saved ? JSON.stringify(saved) : ""); setDirty(false); setDetails(null);
  });
  const showDetails = () => run(() => {
    if (!draft?.id || !viewer.activeCell) return;
    if (viewer.activeSheet?.name.toLowerCase() !== draft.target.sheet.toLowerCase()) throw new Error("Select a value cell inside this PivotTable first.");
    setDetails(viewer.pivotDrillDown(draft.id, viewer.activeCell.row, viewer.activeCell.col)); setDetailContext({ id: draft.id, row: viewer.activeCell.row, col: viewer.activeCell.col }); setOpen(true);
  });
  const uniqueValue = (pivot: XlsxPivotTable, field: XlsxPivotSourceField) => {
    const value = newValue(field);
    let name = value.name;
    for (let index = 2; pivot.values.some(item => item.name.toLocaleLowerCase() === name.toLocaleLowerCase()); index++) name = `${value.name} ${index}`;
    return { ...value, name };
  };
  const replaceValues = (pivot: XlsxPivotTable, values: XlsxPivotTable["values"]): XlsxPivotTable => {
    const mapIndex = (index: number) => values.indexOf(pivot.values[index]);
    const updateAxis = (axis: XlsxPivotTable["rows"][number]): XlsxPivotTable["rows"][number] => {
      if (axis.sort !== "valueAscending" && axis.sort !== "valueDescending") return axis;
      const index = mapIndex(axis.sortMeasure ?? 0);
      return index < 0 ? { ...axis, sort: "ascending", sortMeasure: null } : { ...axis, sortMeasure: index };
    };
    return {
      ...pivot, values,
      rows: pivot.rows.map(updateAxis), columns: pivot.columns.map(updateAxis),
      filters: pivot.filters.flatMap(filter => {
        if (filter.kind !== "value" && filter.kind !== "top10") return [filter];
        const measure = mapIndex(filter.measure);
        return measure < 0 ? [] : [{ ...filter, measure }];
      }),
      ...(pivot.charts ? { charts: pivot.charts.map(chart => chart.measure == null ? chart : { ...chart, measure: mapIndex(chart.measure) < 0 ? null : mapIndex(chart.measure) }) } : {}),
    };
  };
  const addField = (area: XlsxPivotFieldArea, field: number) => edit(pivot => {
    const source = pivot.fields.find(item => item.index === field);
    if (!source) return pivot;
    if (area === "values") return { ...pivot, values: [...pivot.values, uniqueValue(pivot, source)] };
    const next = { ...pivot, rows: pivot.rows.filter(item => item.field !== field), columns: pivot.columns.filter(item => item.field !== field), reportFilters: pivot.reportFilters.filter(item => item !== field) };
    return area === "reportFilters" ? { ...next, reportFilters: [...next.reportFilters, field] } : { ...next, [area]: [...next[area], newAxis(field)] };
  });
  const removeField = (area: XlsxPivotFieldArea, index: number) => edit(pivot => area === "values"
    ? replaceValues(pivot, pivot.values.filter((_, position) => position !== index))
    : { ...pivot, [area]: pivot[area].filter((_, position) => position !== index) });
  const moveField = (fromArea: XlsxPivotFieldArea, fromIndex: number, toArea: XlsxPivotFieldArea, toIndex?: number) => {
    if (fromArea === "values" && toArea !== "values" && draft?.values[fromIndex]?.formula?.trim()) {
      run(() => { throw new Error("Calculated fields can only be placed in Values."); });
      return;
    }
    edit(pivot => {
      const entry = pivot[fromArea][fromIndex];
      if (entry === undefined || (fromArea === "values" && toArea !== "values" && typeof entry === "object" && "formula" in entry && entry.formula?.trim())) return pivot;
      if (fromArea === toArea) {
        const entries = [...pivot[fromArea]];
        entries.splice(fromIndex, 1);
        entries.splice(Math.max(0, Math.min(toIndex ?? entries.length, entries.length)), 0, entry);
        return fromArea === "values" ? replaceValues(pivot, entries as XlsxPivotTable["values"]) : { ...pivot, [fromArea]: entries };
      }
      const field = typeof entry === "number" ? entry : entry.field;
      const source = pivot.fields.find(item => item.index === field);
      if (!source) return pivot;
      let next = fromArea === "values" ? replaceValues(pivot, pivot.values.filter((_, index) => index !== fromIndex)) : { ...pivot, [fromArea]: pivot[fromArea].filter((_, index) => index !== fromIndex) };
      if (toArea === "values") {
        const values = [...next.values];
        values.splice(Math.max(0, Math.min(toIndex ?? values.length, values.length)), 0, uniqueValue(next, source));
        return replaceValues(next, values);
      }
      next = { ...next, rows: next.rows.filter(item => item.field !== field), columns: next.columns.filter(item => item.field !== field), reportFilters: next.reportFilters.filter(item => item !== field) };
      const entries = [...next[toArea]];
      const added = toArea === "reportFilters" ? field : fromArea === "rows" || fromArea === "columns" ? entry : newAxis(field);
      entries.splice(Math.max(0, Math.min(toIndex ?? entries.length, entries.length)), 0, added as never);
      return { ...next, [toArea]: entries };
    });
  };
  return { viewer, readOnly: viewer.readOnly || Boolean(draft?.unsupportedFeatures.length), open, setOpen, draft, dirty, edit, select, create, apply, reset, remove, refresh, readFields, addField, removeField, moveField, details, detailContext, showDetails, clearDetails: () => { setDetails(null); setDetailContext(null); }, sourceNeedsLoad, stale, error, status, run };
}
