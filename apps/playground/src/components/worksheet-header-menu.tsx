import * as React from "react";
import { ArrowDownAZ, ArrowUpAZ, Check, ChevronDown, Filter, X } from "lucide-react";
import type { XlsxPivotFilter, XlsxPivotTableHeaderMenuRenderProps, XlsxTableHeaderMenuRenderProps } from "@extend-ai/react-xlsx";
import { Button } from "@/components/ui/button";
import { Checkbox } from "@/components/ui/checkbox";
import { Field, FieldLabel } from "@/components/ui/field";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { Menu, MenuGroup, MenuGroupLabel, MenuItem, MenuPopup, MenuTrigger } from "@/components/ui/menu";
import { Popover, PopoverPopup, PopoverTitle, PopoverTrigger } from "@/components/ui/popover";
import { Select, SelectItem, SelectPopup, SelectTrigger, SelectValue } from "@/components/ui/select";

type Options = ReadonlyArray<readonly [string, string]>;
const FILTER_OPERATORS: Options = [
  ["equal", "Equals"], ["notEqual", "Does not equal"], ["contains", "Contains"], ["notContains", "Does not contain"],
  ["beginsWith", "Begins with"], ["notBeginsWith", "Does not begin with"], ["endsWith", "Ends with"], ["notEndsWith", "Does not end with"],
  ["greaterThan", "Greater than"], ["greaterThanOrEqual", "Greater than or equal to"], ["lessThan", "Less than"],
  ["lessThanOrEqual", "Less than or equal to"], ["between", "Between"], ["notBetween", "Not between"],
];
const PopupContainer = React.createContext<HTMLDivElement | null>(null);

function HeaderChoice({ label, value, options, onChange, disabled }: { label: string; value: string; options: Options; onChange: (value: string) => void; disabled?: boolean }) {
  const container = React.useContext(PopupContainer);
  const items = options.map(([value, label]) => ({ value, label }));
  return <Field className="gap-1.5" disabled={disabled}>
    <FieldLabel className="text-xs font-normal">{label}</FieldLabel>
    <Select items={items} value={value} disabled={disabled} onValueChange={next => { if (next !== null) onChange(next); }}>
      <SelectTrigger size="sm" className="min-w-0 text-xs"><SelectValue /></SelectTrigger>
      <SelectPopup portalProps={{ container }} alignItemWithTrigger={false}>{items.map(item => <SelectItem key={item.value} value={item.value} className="text-xs">{item.label}</SelectItem>)}</SelectPopup>
    </Select>
  </Field>;
}

function HeaderInput({ label, value, onChange, ...props }: Omit<React.ComponentProps<typeof Input>, "value" | "onChange"> & { label: string; value: string; onChange: (value: string) => void }) {
  return <Field className="gap-1.5"><FieldLabel className="text-xs font-normal">{label}</FieldLabel><Input {...props} className="h-8 text-xs" value={value} onChange={event => onChange(event.target.value)} /></Field>;
}

function headerStyle(style: React.CSSProperties | undefined): React.CSSProperties {
  const { background: _background, backgroundColor: _backgroundColor, color: _color, border: _border, borderRadius: _borderRadius, font: _font, ...position } = style ?? {};
  return position;
}

export function WorksheetHeaderMenu({ menu }: { menu: XlsxTableHeaderMenuRenderProps }) {
  return menu.pivot ? <PivotHeaderMenu menu={menu.pivot} /> : <TableHeaderMenu menu={menu} />;
}

function TableHeaderMenu({ menu }: { menu: XlsxTableHeaderMenuRenderProps }) {
  const [open, setOpen] = React.useState(false);
  return <Menu open={open} onOpenChange={setOpen}>
    <MenuTrigger {...menu.triggerProps} onKeyDown={event => { menu.triggerProps.onKeyDown?.(event); if (event.key === "Escape") setOpen(false); }} style={headerStyle(menu.triggerProps.style)} render={<Button size="icon-xs" variant="outline" className="playground-header-menu-trigger rounded-sm p-0 [&_svg]:size-3" />}>
      {menu.direction === "ascending" ? <ArrowDownAZ /> : menu.direction === "descending" ? <ArrowUpAZ /> : <ChevronDown />}
    </MenuTrigger>
    <MenuPopup align="end" className="w-56" onPointerDown={event => event.stopPropagation()} onKeyDown={event => { event.stopPropagation(); if (event.key === "Escape") setOpen(false); }}>
      <MenuGroup>
      <MenuGroupLabel>{menu.column.name}</MenuGroupLabel>
      <MenuItem onClick={() => { menu.sortAscending(); setOpen(false); }}><ArrowDownAZ />Sort A to Z{menu.direction === "ascending" ? <Check className="ms-auto" /> : null}</MenuItem>
      <MenuItem onClick={() => { menu.sortDescending(); setOpen(false); }}><ArrowUpAZ />Sort Z to A{menu.direction === "descending" ? <Check className="ms-auto" /> : null}</MenuItem>
      </MenuGroup>
    </MenuPopup>
  </Menu>;
}

function PivotHeaderMenu({ menu }: { menu: XlsxPivotTableHeaderMenuRenderProps }) {
  const [container, setContainer] = React.useState<HTMLDivElement | null>(null);
  const localPopup = React.useRef<HTMLDivElement | null>(null);
  const popupRef = menu.popupProps.ref;
  const assignPopup = React.useCallback((node: HTMLDivElement | null) => {
    const previous = localPopup.current;
    localPopup.current = node;
    setContainer(node);
    if (typeof popupRef === "function") popupRef(node);
    else if (popupRef && (node || popupRef.current === previous)) popupRef.current = node;
  }, [popupRef]);
  const filtered = menu.header.fields.some(field => menu.pivot.filters.some(filter => filter.field === field.field && (filter.groupingUnit ?? null) === (field.groupingUnit ?? null)));
  return <Popover open={menu.isOpen} onOpenChange={menu.setOpen}>
    <PopoverTrigger {...menu.triggerProps} style={headerStyle(menu.triggerProps.style)} render={<Button size="icon-xs" variant="outline" className="playground-header-menu-trigger rounded-sm p-0 [&_svg]:size-3" />}>
      {filtered ? <Filter /> : <ChevronDown />}
    </PopoverTrigger>
    <PopoverPopup {...menu.popupProps} ref={assignPopup} align="end" positionMethod="fixed" positionerClassName="transition-none" className="w-80 max-w-[calc(100vw-1rem)]" finalFocus={false}>
      <PopupContainer.Provider value={container}>
        {menu.isOpen ? <PivotHeaderOptions key={`${menu.pivot.id}:${menu.header.row}:${menu.header.col}:${menu.field.depth}`} menu={menu} /> : null}
      </PopupContainer.Provider>
    </PopoverPopup>
  </Popover>;
}

function PivotHeaderOptions({ menu }: { menu: XlsxPivotTableHeaderMenuRenderProps }) {
  const [search, setSearch] = React.useState("");
  const [selected, setSelected] = React.useState(() => new Set(menu.filters.find(filter => filter.kind === "members")?.values ?? menu.members.map(member => member.key)));
  const [mode, setMode] = React.useState<XlsxPivotFilter["kind"]>("members");
  const filterDefault = (kind: XlsxPivotFilter["kind"]): XlsxPivotFilter => ({ field: menu.field.field, ...(menu.field.groupingUnit ? { groupingUnit: menu.field.groupingUnit } : {}), kind, values: [], operator: "equal", value: "", value2: "", measure: 0, count: 10, percent: false, top: true });
  const [draft, setDraft] = React.useState<XlsxPivotFilter>(() => filterDefault("label"));
  const axes = menu.header.axis === "reportFilters" ? [] : menu.pivot[menu.header.axis];
  const [sortMeasure, setSortMeasure] = React.useState(String(axes[menu.field.depth]?.sortMeasure ?? 0));
  const [moreSort, setMoreSort] = React.useState(false);
  const visible = menu.members.filter(member => member.caption.toLocaleLowerCase().includes(search.toLocaleLowerCase()));
  const all = menu.members.length > 0 && menu.members.every(member => selected.has(member.key));
  const some = !all && menu.members.some(member => selected.has(member.key));
  const setKind = (kind: XlsxPivotFilter["kind"]) => { setMode(kind); if (kind !== "members") setDraft(menu.filters.find(filter => filter.kind === kind) ?? filterDefault(kind)); };
  const values = menu.pivot.values.map((value, index) => [String(index), value.name] as const);
  return <div className="flex min-w-0 flex-col gap-3 text-xs">
    <div className="flex items-center justify-between gap-3">
      <PopoverTitle className="min-w-0 break-words text-sm">{menu.field.name}</PopoverTitle>
      <Button size="icon-xs" variant="ghost" aria-label="Close PivotTable filter" onClick={menu.close}><X /></Button>
    </div>
    {menu.header.fields.length > 1 ? <HeaderChoice label="Select field" value={String(menu.field.depth)} options={menu.header.fields.map(field => [String(field.depth), field.name])} onChange={depth => menu.setField(Number(depth))} /> : null}
    {menu.error ? <p role="alert" className="rounded-md border border-destructive/30 bg-destructive/5 p-2 text-destructive-foreground">{menu.error}</p> : null}
    {menu.readOnly ? <p className="leading-relaxed text-muted-foreground">{menu.pivot.unsupportedFeatures.length ? "This PivotTable contains features preserved for Excel. Sorting and filtering are unavailable." : "Enable workbook editing to change sorting or filters."}</p> : null}
    {menu.sort !== null ? <div className="flex flex-col gap-2">
      <div className="grid grid-cols-2 gap-2">
        <Button size="sm" variant="outline" disabled={menu.readOnly} aria-pressed={menu.sort === "ascending"} onClick={() => menu.sortBy("ascending")}><ArrowDownAZ />Sort A to Z</Button>
        <Button size="sm" variant="outline" disabled={menu.readOnly} aria-pressed={menu.sort === "descending"} onClick={() => menu.sortBy("descending")}><ArrowUpAZ />Sort Z to A</Button>
      </div>
      <Button size="xs" variant="ghost" className="justify-start" aria-expanded={moreSort} onClick={() => setMoreSort(!moreSort)}><ChevronDown className={moreSort ? "rotate-180" : undefined} />More sort options</Button>
      {moreSort ? <div className="flex flex-col gap-2 rounded-md border p-2">
        <Button size="sm" variant="ghost" disabled={menu.readOnly} onClick={() => menu.sortBy("manual")}>Use source order</Button>
        {values.length ? <><HeaderChoice label="Sort by value" value={sortMeasure} options={values} onChange={setSortMeasure} disabled={menu.readOnly} /><div className="grid grid-cols-2 gap-2"><Button size="sm" variant="outline" disabled={menu.readOnly} onClick={() => menu.sortBy("valueAscending", Number(sortMeasure))}>Smallest first</Button><Button size="sm" variant="outline" disabled={menu.readOnly} onClick={() => menu.sortBy("valueDescending", Number(sortMeasure))}>Largest first</Button></div></> : null}
      </div> : null}
    </div> : null}
    <Button size="sm" variant="ghost" className="justify-start whitespace-normal text-left" disabled={menu.readOnly || !menu.filters.length} onClick={menu.clearFilters}>Clear filters from “{menu.field.name}”</Button>
    <HeaderChoice label="Filter" value={mode} options={[["members", "Select items"], ["label", "Label filter"], ...(values.length ? [["value", "Value filter"], ["top10", "Top / bottom filter"]] as Options : [])]} onChange={kind => setKind(kind as XlsxPivotFilter["kind"])} />
    {mode === "members" ? <>
      <HeaderInput label="Search items" type="search" value={search} onChange={setSearch} />
      <Label className="flex items-center gap-2 text-xs font-normal"><Checkbox disabled={menu.readOnly} checked={all} indeterminate={some} onCheckedChange={checked => setSelected(new Set(checked ? menu.members.map(member => member.key) : []))} />Select all</Label>
      <div role="group" aria-label={`${menu.field.name} items`} className="max-h-44 overflow-y-auto rounded-md border p-2">
        {visible.slice(0, 300).map(member => <Label key={member.key} className="flex items-start gap-2 py-1 text-xs font-normal"><Checkbox disabled={menu.readOnly} checked={selected.has(member.key)} onCheckedChange={checked => setSelected(previous => { const next = new Set(previous); if (checked) next.add(member.key); else next.delete(member.key); return next; })} /><span className="min-w-0 break-words">{member.caption || "(blank)"}</span></Label>)}
        {!visible.length ? <p className="text-muted-foreground">No matching items.</p> : null}
        {visible.length > 300 ? <p className="text-muted-foreground">Search to narrow {visible.length.toLocaleString()} items.</p> : null}
      </div>
      <p className="text-muted-foreground">{selected.size.toLocaleString()} selected</p>
    </> : <div className="flex flex-col gap-3">
      {mode === "label" || mode === "value" ? <>
        <HeaderChoice label="Condition" value={draft.operator} disabled={menu.readOnly} options={mode === "value" ? FILTER_OPERATORS.filter(([operator]) => !/contains|begins|ends/i.test(operator)) : FILTER_OPERATORS} onChange={operator => setDraft(current => ({ ...current, operator: operator as XlsxPivotFilter["operator"] }))} />
        <HeaderInput label="Value" disabled={menu.readOnly} type={mode === "value" ? "number" : "text"} value={draft.value} onChange={value => setDraft(current => ({ ...current, value }))} />
        {draft.operator === "between" || draft.operator === "notBetween" ? <HeaderInput label="And" disabled={menu.readOnly} type={mode === "value" ? "number" : "text"} value={draft.value2} onChange={value2 => setDraft(current => ({ ...current, value2 }))} /> : null}
      </> : null}
      {mode === "top10" ? <><HeaderChoice label="Show" disabled={menu.readOnly} value={draft.top ? "top" : "bottom"} options={[["top", "Top"], ["bottom", "Bottom"]]} onChange={top => setDraft(current => ({ ...current, top: top === "top" }))} /><HeaderInput label="Amount" disabled={menu.readOnly} type="number" min={1} value={String(draft.count)} onChange={count => setDraft(current => ({ ...current, count: Number(count) }))} /><HeaderChoice label="Unit" disabled={menu.readOnly} value={draft.percent ? "percent" : "items"} options={[["items", "Items"], ["percent", "Percent"]]} onChange={unit => setDraft(current => ({ ...current, percent: unit === "percent" }))} /></> : null}
      {mode === "value" || mode === "top10" ? <HeaderChoice label="By value" disabled={menu.readOnly} value={String(draft.measure)} options={values} onChange={measure => setDraft(current => ({ ...current, measure: Number(measure) }))} /> : null}
    </div>}
    <div className="flex justify-end gap-2 border-t pt-3">
      <Button size="sm" variant="outline" onClick={menu.close}>Cancel</Button>
      <Button size="sm" disabled={menu.readOnly || (mode === "members" && selected.size === 0)} onClick={() => { if (mode === "members") menu.setMembers(all ? null : [...selected]); else menu.setFilter(draft); }}>Apply</Button>
    </div>
  </div>;
}
