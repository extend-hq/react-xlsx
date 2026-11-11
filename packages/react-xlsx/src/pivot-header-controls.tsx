import * as React from "react";
import { createPortal } from "react-dom";
import { VIEWPORT_CHANGE_EVENT, visibleOverlayRect } from "./overlay-geometry";
import type { XlsxPivotAxisField, XlsxPivotFilter, XlsxPivotHeader, XlsxPivotHeaderField, XlsxPivotMember, XlsxPivotTable } from "./pivot-types";
import type { XlsxCellAddress, XlsxImageRect } from "./types";

export interface XlsxPivotTableHeaderMenuRenderProps {
  pivot: XlsxPivotTable;
  header: XlsxPivotHeader;
  field: XlsxPivotHeaderField;
  setField: (depth: number) => void;
  members: XlsxPivotMember[];
  filters: XlsxPivotFilter[];
  sort: XlsxPivotAxisField["sort"] | null;
  readOnly: boolean;
  isOpen: boolean;
  setOpen: (open: boolean) => void;
  close: () => void;
  error: string | null;
  triggerProps: React.ComponentPropsWithRef<"button">;
  popupProps: React.ComponentPropsWithRef<"div">;
  defaultNode: React.ReactNode;
  sortBy: (direction: XlsxPivotAxisField["sort"], measure?: number) => void;
  setMembers: (keys: string[] | null) => void;
  setFilter: (filter: XlsxPivotFilter) => void;
  clearFilters: () => void;
}

export interface PivotHeaderControlsProps {
  pivots: XlsxPivotTable[];
  sheetName?: string;
  rectForCell: (cell: XlsxCellAddress) => XlsxImageRect | null;
  viewport: XlsxImageRect;
  viewportElement?: HTMLElement | null;
  experimentalCanvas?: boolean;
  isDark?: boolean;
  stickyTop: Map<number, number>;
  stickyLeft: Map<number, number>;
  frozenRight: number;
  frozenBottom: number;
  zoom: number;
  readOnly: boolean;
  onChange: (pivot: XlsxPivotTable) => unknown;
  colors: { surface: string; text: string; border: string; mutedText?: string; headerSurface?: string; headerText?: string };
  renderMenu?: (props: XlsxPivotTableHeaderMenuRenderProps) => React.ReactNode;
  openLabel?: string;
  renderControl?: (cell: XlsxCellAddress, content: React.ReactNode) => React.ReactNode;
}

type Options = ReadonlyArray<readonly [string, string]>;
type MenuContext = Omit<XlsxPivotTableHeaderMenuRenderProps, "defaultNode">;
const OPERATORS: Options = [["equal", "Equals"], ["notEqual", "Does not equal"], ["contains", "Contains"], ["notContains", "Does not contain"], ["beginsWith", "Begins with"], ["notBeginsWith", "Does not begin with"], ["endsWith", "Ends with"], ["notEndsWith", "Does not end with"], ["greaterThan", "Greater than"], ["greaterThanOrEqual", "Greater than or equal to"], ["lessThan", "Less than"], ["lessThanOrEqual", "Less than or equal to"], ["between", "Between"], ["notBetween", "Not between"]];
const stack: React.CSSProperties = { display: "flex", flexDirection: "column", gap: 8, minWidth: 0 };
const row: React.CSSProperties = { display: "flex", alignItems: "center", gap: 6 };
const control: React.CSSProperties = { boxSizing: "border-box", width: "100%", minWidth: 0, minHeight: 30, padding: "4px 7px", font: "inherit", color: "inherit", background: "var(--pivot-menu-surface)", border: "1px solid var(--pivot-menu-border)", borderRadius: 3 };
const subdued: React.CSSProperties = { fontSize: 11, opacity: 0.8, lineHeight: 1.45, margin: 0 };

function groupUnit(axis: XlsxPivotAxisField) { return axis.grouping ? axis.grouping.kind === "date" ? axis.grouping.unit : axis.grouping.kind : null; }
function axisFor(pivot: XlsxPivotTable, header: XlsxPivotHeader, field: XlsxPivotHeaderField) {
  const axes = header.axis === "reportFilters" ? [...pivot.rows, ...pivot.columns] : pivot[header.axis];
  const candidate = header.axis === "reportFilters" ? undefined : axes[field.depth];
  return candidate?.field === field.field ? candidate : axes.find(axis => axis.field === field.field && groupUnit(axis) === (field.groupingUnit ?? null));
}
function scoped(filter: XlsxPivotFilter, field: XlsxPivotHeaderField) { return filter.field === field.field && (filter.groupingUnit ?? null) === (field.groupingUnit ?? null); }
function filterDefault(field: XlsxPivotHeaderField, kind: XlsxPivotFilter["kind"]): XlsxPivotFilter { return { field: field.field, groupingUnit: field.groupingUnit, kind, values: [], operator: "equal", value: "", value2: "", measure: 0, count: 10, percent: false, top: true }; }

function Button({ children, style, ...props }: React.ButtonHTMLAttributes<HTMLButtonElement>) {
  return <button type="button" {...props} style={{ ...control, width: "auto", cursor: props.disabled ? "default" : "pointer", opacity: props.disabled ? 0.5 : 1, ...style }}>{children}</button>;
}
function Choice({ label, value, options, onChange }: { label: string; value: string; options: Options; onChange: (value: string) => void }) {
  return <label style={{ ...stack, gap: 3 }}><span>{label}</span><select value={value} onChange={event => onChange(event.target.value)} style={control}>{options.map(([key, label]) => <option key={key} value={key}>{label}</option>)}</select></label>;
}
function Input({ label, value, onChange, ...props }: Omit<React.InputHTMLAttributes<HTMLInputElement>, "value" | "onChange"> & { label: string; value: string; onChange: (value: string) => void }) {
  return <label style={{ ...stack, gap: 3 }}><span>{label}</span><input {...props} value={value} onChange={event => onChange(event.target.value)} style={control} /></label>;
}

function DefaultPivotHeaderPopup({ menu, anchor, colors, popupRef }: { menu: MenuContext; anchor: DOMRect; colors: PivotHeaderControlsProps["colors"]; popupRef: React.RefObject<HTMLDivElement | null> }) {
  const [search, setSearch] = React.useState("");
  const existingMembers = menu.filters.find(filter => filter.kind === "members");
  const [selected, setSelected] = React.useState(() => new Set(existingMembers?.values ?? menu.members.map(member => member.key)));
  const [mode, setMode] = React.useState<XlsxPivotFilter["kind"]>("members");
  const [draft, setDraft] = React.useState<XlsxPivotFilter>(() => filterDefault(menu.field, "label"));
  const [sortMeasure, setSortMeasure] = React.useState("0");
  const [position, setPosition] = React.useState({ top: anchor.bottom + 4, left: anchor.left, width: 308, maxHeight: 440 });
  const titleId = React.useId();
  const searchRef = React.useRef<HTMLInputElement>(null);
  React.useEffect(() => { searchRef.current?.focus(); }, []);
  React.useLayoutEffect(() => {
    const padding = 8;
    const width = Math.max(0, Math.min(308, window.innerWidth - padding * 2));
    const maxHeight = Math.max(0, Math.min(480, window.innerHeight - padding * 2));
    const height = Math.min(popupRef.current?.getBoundingClientRect().height ?? 440, maxHeight);
    const below = anchor.bottom + 4;
    const top = below + height <= window.innerHeight - padding ? below : Math.max(padding, anchor.top - height - 4);
    const left = Math.max(padding, Math.min(anchor.right - width, window.innerWidth - width - padding));
    setPosition(current => current.top === top && current.left === left && current.width === width && current.maxHeight === maxHeight ? current : { top, left, width, maxHeight });
  }, [anchor, mode, menu.error, menu.field.depth, popupRef]);
  const visible = menu.members.filter(member => member.caption.toLocaleLowerCase().includes(search.toLocaleLowerCase()));
  const all = menu.members.length > 0 && menu.members.every(member => selected.has(member.key));
  const indeterminate = !all && menu.members.some(member => selected.has(member.key));
  const setKind = (kind: XlsxPivotFilter["kind"]) => { setMode(kind); if (kind !== "members") setDraft(menu.filters.find(filter => filter.kind === kind) ?? filterDefault(menu.field, kind)); };
  const apply = () => {
    if (mode === "members") menu.setMembers(all ? null : [...selected]);
    else menu.setFilter(draft);
  };
  const trapKeys = (event: React.KeyboardEvent) => {
    event.stopPropagation();
    if (event.key === "Escape") { event.preventDefault(); menu.close(); return; }
    if (event.key !== "Tab") return;
    const targets = popupRef.current?.querySelectorAll<HTMLElement>('button:not(:disabled), input:not(:disabled), select:not(:disabled), [tabindex="0"]');
    if (!targets?.length) return;
    const first = targets[0]; const last = targets[targets.length - 1];
    if (event.shiftKey && document.activeElement === first) { event.preventDefault(); last.focus(); }
    else if (!event.shiftKey && document.activeElement === last) { event.preventDefault(); first.focus(); }
  };
  return <div ref={popupRef} role="dialog" aria-labelledby={titleId} onKeyDown={trapKeys} onPointerDown={event => event.stopPropagation()} onClick={event => event.stopPropagation()} style={{ position: "fixed", ...position, boxSizing: "border-box", overflowY: "auto", zIndex: 1000, pointerEvents: "auto", color: colors.text, background: colors.surface, border: `1px solid ${colors.border}`, borderRadius: 4, boxShadow: "0 4px 18px #0003", padding: 12, font: "12px/1.4 system-ui, sans-serif", ...({ "--pivot-menu-surface": colors.surface, "--pivot-menu-border": colors.border } as React.CSSProperties) }}>
    <div style={{ ...stack, gap: 10 }}>
      <div style={{ ...row, justifyContent: "space-between" }}><h3 id={titleId} style={{ fontSize: 13, margin: 0, overflowWrap: "anywhere" }}>{menu.field.name}</h3><Button aria-label="Close PivotTable filter" onClick={menu.close} style={{ border: 0, fontSize: 17, padding: "1px 5px" }}>×</Button></div>
      {menu.header.fields.length > 1 ? <Choice label="Select field" value={String(menu.field.depth)} options={menu.header.fields.map(field => [String(field.depth), field.name])} onChange={depth => menu.setField(Number(depth))} /> : null}
      {menu.error ? <p role="alert" style={{ margin: 0, borderInlineStart: "2px solid #c7372c", paddingInlineStart: 6 }}>{menu.error}</p> : null}
      {menu.readOnly ? <p style={subdued}>{menu.pivot.unsupportedFeatures.length ? "This PivotTable contains features preserved for Excel. Sorting and filtering are unavailable." : "Enable workbook editing to change sorting or filters."}</p> : null}
      <fieldset disabled={menu.readOnly} style={{ ...stack, margin: 0, padding: 0, border: 0, minWidth: 0 }}>
        {menu.sort !== null ? <><div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 6 }}><Button aria-pressed={menu.sort === "ascending"} onClick={() => menu.sortBy("ascending")}>Sort A to Z</Button><Button aria-pressed={menu.sort === "descending"} onClick={() => menu.sortBy("descending")}>Sort Z to A</Button></div><details><summary style={{ cursor: "pointer" }}>More sort options</summary><div style={{ ...stack, paddingTop: 8 }}><Button onClick={() => menu.sortBy("manual")}>Use source order</Button>{menu.pivot.values.length ? <><Choice label="Sort by value" value={sortMeasure} options={menu.pivot.values.map((value, index) => [String(index), value.name])} onChange={setSortMeasure} /><div style={row}><Button onClick={() => menu.sortBy("valueAscending", Number(sortMeasure))}>Smallest first</Button><Button onClick={() => menu.sortBy("valueDescending", Number(sortMeasure))}>Largest first</Button></div></> : null}</div></details></> : null}
        <Button disabled={!menu.filters.length} onClick={menu.clearFilters}>Clear filters from “{menu.field.name}”</Button>
      </fieldset>
      <Choice label="Filter" value={mode} options={[["members", "Select items"], ["label", "Label filter"], ...(menu.pivot.values.length ? [["value", "Value filter"], ["top10", "Top / bottom filter"]] as Options : [])]} onChange={kind => setKind(kind as XlsxPivotFilter["kind"])} />
      {mode === "members" ? <>
        <label style={{ ...stack, gap: 3 }}><span>Search items</span><input ref={searchRef} type="search" value={search} onChange={event => setSearch(event.target.value)} style={control} /></label>
        <fieldset disabled={menu.readOnly} style={{ margin: 0, padding: 0, border: 0, minWidth: 0 }}>
          <label style={{ ...row, padding: "3px 0 6px" }}><input type="checkbox" checked={all} ref={node => { if (node) node.indeterminate = indeterminate; }} onChange={event => setSelected(new Set(event.target.checked ? menu.members.map(member => member.key) : []))} />Select all</label>
          <div role="group" aria-label={`${menu.field.name} items`} style={{ border: `1px solid ${colors.border}`, maxHeight: 180, overflowY: "auto", padding: "4px 6px" }}>{visible.slice(0, 300).map(member => <label key={member.key} style={{ ...row, alignItems: "start", padding: "3px 0", overflowWrap: "anywhere" }}><input type="checkbox" style={{ marginTop: 2, flexShrink: 0 }} checked={selected.has(member.key)} onChange={event => setSelected(previous => { const next = new Set(previous); if (event.target.checked) next.add(member.key); else next.delete(member.key); return next; })} />{member.caption || "(blank)"}</label>)}{!visible.length ? <p style={subdued}>No matching items.</p> : null}{visible.length > 300 ? <p style={subdued}>Search to narrow {visible.length.toLocaleString()} items.</p> : null}</div>
        </fieldset>
        <p style={subdued}>{selected.size.toLocaleString()} selected</p>
      </> : <fieldset disabled={menu.readOnly} style={{ ...stack, margin: 0, padding: 0, border: 0, minWidth: 0 }}>
        {mode === "label" || mode === "value" ? <><Choice label="Condition" value={draft.operator} options={mode === "value" ? OPERATORS.filter(([operator]) => !/contains|begins|ends/i.test(operator)) : OPERATORS} onChange={operator => setDraft(current => ({ ...current, operator: operator as XlsxPivotFilter["operator"] }))} /><Input label="Value" type={mode === "value" ? "number" : "text"} value={draft.value} onChange={value => setDraft(current => ({ ...current, value }))} />{draft.operator === "between" || draft.operator === "notBetween" ? <Input label="And" type={mode === "value" ? "number" : "text"} value={draft.value2} onChange={value2 => setDraft(current => ({ ...current, value2 }))} /> : null}</> : null}
        {mode === "top10" ? <><Choice label="Show" value={draft.top ? "top" : "bottom"} options={[["top", "Top"], ["bottom", "Bottom"]]} onChange={top => setDraft(current => ({ ...current, top: top === "top" }))} /><Input label="Amount" type="number" min={1} value={String(draft.count)} onChange={count => setDraft(current => ({ ...current, count: Number(count) }))} /><Choice label="Unit" value={draft.percent ? "percent" : "items"} options={[["items", "Items"], ["percent", "Percent"]]} onChange={unit => setDraft(current => ({ ...current, percent: unit === "percent" }))} /></> : null}
        {mode === "value" || mode === "top10" ? <Choice label="By value" value={String(draft.measure)} options={menu.pivot.values.map((value, index) => [String(index), value.name])} onChange={measure => setDraft(current => ({ ...current, measure: Number(measure) }))} /> : null}
      </fieldset>}
      <div style={{ ...row, justifyContent: "flex-end", borderTop: `1px solid ${colors.border}`, paddingTop: 9 }}><Button onClick={menu.close}>Cancel</Button><Button disabled={menu.readOnly || (mode === "members" && selected.size === 0)} onClick={apply} style={{ fontWeight: 600 }}>Apply</Button></div>
    </div>
  </div>;
}

export function PivotHeaderControls({ pivots, sheetName, rectForCell, viewport, viewportElement, experimentalCanvas = true, isDark = false, stickyTop, stickyLeft, frozenRight, frozenBottom, zoom, readOnly, onChange, colors, renderMenu, openLabel, renderControl }: PivotHeaderControlsProps) {
  const [open, setOpen] = React.useState<{ key: string; depth: number } | null>(null);
  const [error, setError] = React.useState<string | null>(null);
  const [anchor, setAnchor] = React.useState<DOMRect | null>(null);
  const triggers = React.useRef(new Map<string, HTMLButtonElement>());
  const triggerCells = React.useRef(new Map<string, { frozenX: boolean; frozenY: boolean }>());
  const [measuredViewport, setMeasuredViewport] = React.useState<XlsxImageRect | null>(null);
  const popupRef = React.useRef<HTMLDivElement>(null);
  const dialogId = React.useId();
  const size = 16 * zoom;
  const liveViewport = measuredViewport ?? viewport;
  const clipTriggers = React.useRef(() => {});
  clipTriggers.current = () => {
    if (!viewportElement || renderControl) return;
    const bounds = viewportElement.getBoundingClientRect();
    const scaleX = bounds.width / (viewportElement.offsetWidth || bounds.width || 1);
    const scaleY = bounds.height / (viewportElement.offsetHeight || bounds.height || 1);
    const left = bounds.left + viewportElement.clientLeft * scaleX;
    const top = bounds.top + viewportElement.clientTop * scaleY;
    const rowHeader = stickyLeft.size ? Math.min(...stickyLeft.values()) : frozenRight;
    const columnHeader = stickyTop.size ? Math.min(...stickyTop.values()) : frozenBottom;
    for (const [key, trigger] of triggers.current) {
      const cell = triggerCells.current.get(key);
      if (!cell) continue;
      const rect = trigger.getBoundingClientRect();
      const minX = left + (cell.frozenX ? rowHeader : frozenRight) * scaleX;
      const minY = top + (cell.frozenY ? columnHeader : frozenBottom) * scaleY;
      const maxX = left + (cell.frozenX ? frozenRight : viewportElement.clientWidth) * scaleX;
      const maxY = top + (cell.frozenY ? frozenBottom : viewportElement.clientHeight) * scaleY;
      trigger.style.visibility = rect.left >= minX - 0.25 && rect.top >= minY - 0.25 && rect.right <= maxX + 0.25 && rect.bottom <= maxY + 0.25 ? "visible" : "hidden";
    }
  };
  React.useLayoutEffect(() => {
    if (!viewportElement || renderControl) { setMeasuredViewport(null); return; }
    let frame: number | null = null;
    const measure = () => {
      const next = { left: viewportElement.scrollLeft, top: viewportElement.scrollTop, width: viewportElement.clientWidth, height: viewportElement.clientHeight };
      setMeasuredViewport(current => current && current.left === next.left && current.top === next.top && current.width === next.width && current.height === next.height ? current : next);
      clipTriggers.current();
      if (frame !== null) cancelAnimationFrame(frame);
      frame = requestAnimationFrame(() => { frame = null; clipTriggers.current(); });
    };
    measure();
    viewportElement.addEventListener("scroll", measure, { passive: true });
    const observer = typeof ResizeObserver === "undefined" ? null : new ResizeObserver(measure);
    observer?.observe(viewportElement);
    return () => { viewportElement.removeEventListener("scroll", measure); observer?.disconnect(); if (frame !== null) cancelAnimationFrame(frame); };
  }, [viewportElement, Boolean(renderControl)]);
  React.useLayoutEffect(() => { clipTriggers.current(); });
  React.useEffect(() => { setOpen(null); setError(null); }, [sheetName]);
  React.useLayoutEffect(() => {
    if (!open) return;
    const trigger = triggers.current.get(open.key);
    if (!trigger || !visibleOverlayRect(trigger)) { setOpen(null); return; }
    const next = trigger.getBoundingClientRect();
    setAnchor(current => current && current.x === next.x && current.y === next.y && current.width === next.width && current.height === next.height ? current : next);
  });
  React.useLayoutEffect(() => {
    if (!open) return;
    const measure = () => {
      const trigger = triggers.current.get(open.key);
      if (!trigger || !visibleOverlayRect(trigger)) { setOpen(null); return; }
      const next = trigger.getBoundingClientRect();
      setAnchor(current => current && current.x === next.x && current.y === next.y && current.width === next.width && current.height === next.height ? current : next);
    };
    measure();
    window.addEventListener("resize", measure);
    window.addEventListener("scroll", measure, true);
    viewportElement?.addEventListener(VIEWPORT_CHANGE_EVENT, measure);
    const outside = (event: PointerEvent) => {
      const target = event.target as Node | null;
      if (target && (popupRef.current?.contains(target) || triggers.current.get(open.key)?.contains(target))) return;
      setOpen(null);
    };
    document.addEventListener("pointerdown", outside, true);
    return () => { window.removeEventListener("resize", measure); window.removeEventListener("scroll", measure, true); viewportElement?.removeEventListener(VIEWPORT_CHANGE_EVENT, measure); document.removeEventListener("pointerdown", outside, true); };
  }, [open, pivots, viewportElement, viewport.left, viewport.top, viewport.width, viewport.height]);
  const close = () => { const key = open?.key; setOpen(null); setError(null); if (key) requestAnimationFrame(() => triggers.current.get(key)?.focus()); };
  return <>{pivots.filter(pivot => pivot.target.sheet.toLowerCase() === sheetName?.toLowerCase()).flatMap(pivot => (pivot.headers ?? []).flatMap(header => {
    if (!header.fields.length) return [];
    const rect = rectForCell(header);
    if (!rect) return [];
    const frozenTop = stickyTop.get(header.row); const frozenLeft = stickyLeft.get(header.col);
    const top = renderControl || frozenTop === undefined ? rect.top : frozenTop + liveViewport.top;
    const left = renderControl || frozenLeft === undefined ? rect.left : frozenLeft + liveViewport.left;
    const width = Math.min(size, rect.width); const height = Math.min(size, rect.height);
    const x = left + Math.max(0, rect.width - size - (experimentalCanvas ? 1 : 4) * zoom);
    const y = top + (experimentalCanvas ? Math.max(0, rect.height - size - zoom) : 3 * zoom);
    if (!renderControl && (x < liveViewport.left + (frozenLeft === undefined ? frozenRight : 0) || y < liveViewport.top + (frozenTop === undefined ? frozenBottom : 0) || x + width > liveViewport.left + liveViewport.width || y + height > liveViewport.top + liveViewport.height)) return [];
    const key = `${pivot.id}:${header.row}:${header.col}:${header.axis}`;
    const isOpen = open?.key === key;
    const field = header.fields.find(field => isOpen && field.depth === open.depth) ?? header.fields[0];
    const axis = axisFor(pivot, header, field);
    const members = axis?.items?.length ? axis.items : pivot.fields.find(source => source.index === field.field)?.items ?? [];
    const filters = pivot.filters.filter(filter => scoped(filter, field));
    const filtered = header.fields.some(field => pivot.filters.some(filter => scoped(filter, field)));
    const label = openLabel?.replace("{column}", header.caption || field.name) ?? `Sort and filter ${header.caption || field.name}`;
    const locked = readOnly || pivot.unsupportedFeatures.length > 0;
    const direction = axis?.sort === "ascending" || axis?.sort === "valueAscending" ? "ascending" : axis?.sort === "descending" || axis?.sort === "valueDescending" ? "descending" : null;
    const foreground = isDark ? colors.headerText ?? colors.text : "#667174";
    const triggerContent = experimentalCanvas ? <svg aria-hidden="true" viewBox="0 0 16 16" width={size} height={size} style={{ display: "block", flexShrink: 0 }}>
      <rect width="16" height="16" fill={isDark ? colors.headerSurface ?? colors.surface : "#f8fafb"} /><rect x=".5" y=".5" width="15" height="15" fill="none" stroke={isDark ? colors.mutedText ?? colors.border : "#969c9e"} />
      {filtered ? <path d={direction ? "M2 3h6L6 7v5l-2 1V7z" : "M3 3h10L9 8v4l-2 1V8z"} fill={foreground} /> : <path d={direction ? "M2 10h5l-2.5 3z" : "M4 6h8l-4 4z"} fill={foreground} />}
      {direction ? <><rect x="10" y="5" width="2" height="6" fill={foreground} /><path d={direction === "ascending" ? "M11 3L7.5 6h7z" : "M11 13L7.5 10h7z"} fill={foreground} /></> : null}
    </svg> : filtered ? <svg aria-hidden="true" viewBox="0 0 16 16" width={11 * zoom} height={11 * zoom}><path d="M2 3h12L9.5 8v4l-3 1V8z" fill="currentColor" /></svg> : <span aria-hidden="true">{direction === "ascending" ? "▲" : direction === "descending" ? "▼" : "▾"}</span>;
    const change = (next: XlsxPivotTable) => {
      if (locked) { setError("This PivotTable cannot be edited."); return; }
      try { if (onChange(next) === false) throw new Error("The PivotTable could not be updated."); close(); }
      catch (cause) { setError(cause instanceof Error ? cause.message : String(cause)); }
    };
    const context: MenuContext = {
      pivot, header, field, members, filters, sort: header.axis === "reportFilters" ? null : axis?.sort ?? null, readOnly: locked, isOpen,
      setField: depth => { if (header.fields.some(field => field.depth === depth)) { setOpen({ key, depth }); setError(null); } },
      setOpen: next => { if (next) { setOpen({ key, depth: field.depth }); setAnchor(triggers.current.get(key)?.getBoundingClientRect() ?? null); setError(null); } else setOpen(null); },
      close, error: isOpen ? error : null,
      popupProps: { ref: popupRef, id: dialogId, onPointerDown: event => event.stopPropagation(), onClick: event => event.stopPropagation(), onKeyDown: event => { event.stopPropagation(); if (event.key === "Escape") { event.preventDefault(); close(); } } },
      sortBy: (direction, measure) => {
        if (header.axis === "reportFilters" || !axis) { setError("This field does not have a report sort order."); return; }
        change({ ...pivot, [header.axis]: pivot[header.axis].map((item, depth) => depth === field.depth ? { ...item, sort: direction, sortMeasure: measure ?? item.sortMeasure ?? 0 } : item) });
      },
      setMembers: keys => change({ ...pivot, filters: [...pivot.filters.filter(filter => !scoped(filter, field) || filter.kind !== "members"), ...(keys === null ? [] : [{ ...filterDefault(field, "members"), values: [...keys] }])] }),
      setFilter: filter => change({ ...pivot, filters: [...pivot.filters.filter(item => !scoped(item, field) || item.kind !== filter.kind), { ...filter, field: field.field, groupingUnit: field.groupingUnit }] }),
      clearFilters: () => change({ ...pivot, filters: pivot.filters.filter(filter => !scoped(filter, field)) }),
      triggerProps: {
        type: "button", "aria-label": label, "aria-haspopup": "dialog", "aria-expanded": isOpen, "aria-controls": isOpen ? dialogId : undefined, title: label,
        ref: node => { if (node) { triggers.current.set(key, node); triggerCells.current.set(key, { frozenX: frozenLeft !== undefined, frozenY: frozenTop !== undefined }); } else { triggers.current.delete(key); triggerCells.current.delete(key); } },
        onClick: event => { event.stopPropagation(); setError(null); if (isOpen) setOpen(null); else { setAnchor(event.currentTarget.getBoundingClientRect()); setOpen({ key, depth: field.depth }); } },
        onPointerDown: event => event.stopPropagation(),
        onKeyDown: event => { event.stopPropagation(); if (event.key === "Escape") close(); if (event.key === "ArrowDown" || event.key === "ArrowUp") { event.preventDefault(); setAnchor(event.currentTarget.getBoundingClientRect()); setOpen({ key, depth: field.depth }); setError(null); } },
        style: { position: "absolute", left: x, top: y, width, height, padding: 0, display: "inline-flex", alignItems: "center", justifyContent: "center", overflow: "hidden", pointerEvents: "auto", background: "transparent", color: colors.mutedText ?? colors.text, border: "none", borderRadius: 0, cursor: "pointer", zIndex: 26, font: `${10 * zoom}px/1 system-ui` },
        children: triggerContent,
      },
    };
    const defaultNode = <><button {...context.triggerProps} data-xlsx-pivot-header="" />{isOpen && anchor && typeof document !== "undefined" ? createPortal(<div id={dialogId}><DefaultPivotHeaderPopup key={`${key}:${field.depth}`} menu={context} anchor={anchor} colors={colors} popupRef={popupRef} /></div>, document.body) : null}</>;
    const content = renderMenu ? renderMenu({ ...context, defaultNode }) : defaultNode;
    return [<React.Fragment key={key}>{renderControl ? renderControl(header, content) : content}</React.Fragment>];
  }))}</>;
}
