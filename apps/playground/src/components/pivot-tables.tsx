import * as React from "react";
import {
  ArrowDown, ArrowUp, ChartColumn, Check, ChevronDown, ChevronRight, Download, Filter, GripVertical,
  LayoutList, ListFilter, Plus, RefreshCw, Search, Settings2, Sigma, Table2, Trash2, X,
} from "lucide-react";
import { useXlsxViewer, useXlsxPivotTableEditor, type XlsxPivotTable, type XlsxPivotTableEditor } from "@extend-ai/react-xlsx";
import { RibbonButton, RibbonGroup } from "@/components/ribbon";
import { Button } from "@/components/ui/button";
import { Checkbox } from "@/components/ui/checkbox";
import { Field, FieldDescription, FieldLabel } from "@/components/ui/field";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { Select, SelectItem, SelectPopup, SelectTrigger, SelectValue } from "@/components/ui/select";
import { Tabs, TabsList, TabsPanel, TabsTab } from "@/components/ui/tabs";
import { cn } from "@/lib/utils";
import { Group, Panel, Separator } from "react-resizable-panels";
import { useMediaQuery } from "@/hooks/use-media-query";
import { PivotVisualFilters } from "@/components/pivot-visual-filters";

type AxisField = XlsxPivotTable["rows"][number];
type DataField = XlsxPivotTable["values"][number];
type PivotFilter = XlsxPivotTable["filters"][number];
type SourceField = XlsxPivotTable["fields"][number];
type Grouping = NonNullable<AxisField["grouping"]>;
type Area = "rows" | "columns" | "values" | "reportFilters";
type DetailRows = { fields: string[]; rows: (string | number | boolean | null)[][] };
type Options = ReadonlyArray<readonly [string, string]>;

const AGGREGATIONS: Options = [
  ["sum", "Sum"], ["count", "Count"], ["average", "Average"], ["max", "Max"], ["min", "Min"],
  ["product", "Product"], ["countNums", "Count Numbers"], ["stdDev", "StdDev"], ["stdDevp", "StdDevp"],
  ["var", "Var"], ["varp", "Varp"],
];
const SHOW_AS: Options = [
  ["normal", "No Calculation"], ["percentOfGrandTotal", "% of Grand Total"],
  ["percentOfColumnTotal", "% of Column Total"], ["percentOfRowTotal", "% of Row Total"],
  ["percent", "% Of"], ["percentOfParentRowTotal", "% of Parent Row Total"],
  ["percentOfParentColumnTotal", "% of Parent Column Total"], ["difference", "Difference From"],
  ["percentDifference", "% Difference From"], ["runningTotal", "Running Total In"],
  ["percentRunningTotal", "% Running Total In"], ["rankAscending", "Rank Smallest to Largest"],
  ["rankDescending", "Rank Largest to Smallest"], ["index", "Index"],
];
const OPERATORS: Options = [
  ["equal", "Equals"], ["notEqual", "Does Not Equal"], ["contains", "Contains"], ["notContains", "Does Not Contain"],
  ["beginsWith", "Begins With"], ["notBeginsWith", "Does Not Begin With"], ["endsWith", "Ends With"], ["notEndsWith", "Does Not End With"], ["greaterThan", "Greater Than"],
  ["greaterThanOrEqual", "Greater Than or Equal To"], ["lessThan", "Less Than"],
  ["lessThanOrEqual", "Less Than or Equal To"], ["between", "Between"], ["notBetween", "Not Between"],
];
const AREA_LABELS: Record<Area, string> = { rows: "Rows", columns: "Columns", values: "Values", reportFilters: "Filters" };
const AREAS: Area[] = ["reportFilters", "columns", "rows", "values"];
const STYLES: Options = [
  ["PivotStyleLight16", "Light 16"], ["PivotStyleLight1", "Light 1"], ["PivotStyleLight8", "Light 8"],
  ["PivotStyleMedium9", "Medium 9"], ["PivotStyleMedium2", "Medium 2"], ["PivotStyleMedium4", "Medium 4"],
  ["PivotStyleDark1", "Dark 1"], ["PivotStyleDark2", "Dark 2"],
];
const CHART_TYPES: Options = [
  ["ColumnClustered", "Clustered Column"], ["ColumnStacked", "Stacked Column"], ["ColumnPercentStacked", "100% Stacked Column"],
  ["BarClustered", "Clustered Bar"], ["BarStacked", "Stacked Bar"], ["BarPercentStacked", "100% Stacked Bar"],
  ["Line", "Line"], ["LineStacked", "Stacked Line"],
  ["Area", "Area"], ["AreaStacked", "Stacked Area"], ["AreaPercentStacked", "100% Stacked Area"], ["Pie", "Pie"], ["Doughnut", "Doughnut"],
];

function newFilter(field: number): PivotFilter {
  return { field, kind: "members", values: [], operator: "equal", value: "", value2: "", measure: 0, top: true, count: 10, percent: false };
}

function fieldName(pivot: XlsxPivotTable, field: number) {
  return pivot.fields.find(item => item.index === field)?.name ?? `Field ${field + 1}`;
}

const PivotContext = React.createContext<XlsxPivotTableEditor | null>(null);

function usePivotPanel() {
  const value = React.useContext(PivotContext);
  if (!value) throw new Error("Pivot panel requires its provider.");
  return value;
}

export function PivotPlaygroundProvider({ children }: { children: React.ReactNode }) {
  const editor = useXlsxPivotTableEditor();
  return <PivotContext.Provider value={editor}>{children}</PivotContext.Provider>;
}

export function PivotPlaygroundLayout({ children }: { children: React.ReactNode }) {
  const editor = usePivotPanel();
  const wide = useMediaQuery("(min-width: 800px)");
  const width = React.useRef(400);
  const height = React.useRef(360);
  return <main className="relative flex min-h-0 min-w-0 flex-1 bg-background">
    <Group className="h-full min-h-0 w-full min-w-0" orientation={wide ? "horizontal" : "vertical"}>
      <Panel className="h-full min-h-0 min-w-0" defaultSize="70%" id="pivot-workbook" minSize={wide ? 300 : 80}>
        {children}
      </Panel>
      {editor.open ? <>
        <Separator aria-label="Resize PivotTable field list" className={cn("group relative z-10 flex shrink-0 items-center justify-center bg-border outline-none transition-colors hover:bg-brand/40 focus-visible:bg-brand/50", wide ? "w-1 cursor-col-resize" : "h-1 cursor-row-resize")}>
          <span aria-hidden className={cn("rounded-full bg-muted-foreground/40 group-hover:bg-brand group-focus-visible:bg-brand", wide ? "h-8 w-0.5" : "h-0.5 w-8")} />
        </Separator>
        <Panel className="h-full min-h-0 min-w-0" defaultSize={wide ? width.current : height.current} groupResizeBehavior="preserve-pixel-size" id="pivot-field-list" maxSize={wide ? "65%" : "80%"} minSize={wide ? 280 : 180} onResize={size => { if (size.inPixels > 0) { if (wide) width.current = size.inPixels; else height.current = size.inPixels; } }}>
          <PivotTablePanel />
        </Panel>
      </> : null}
    </Group>
  </main>;
}

export function InsertPivotButton() {
  const viewer = useXlsxViewer();
  const panel = usePivotPanel();
  return <RibbonButton disabled={!viewer.activeSheet || viewer.readOnly || panel.dirty} label="PivotTable" onClick={panel.create} showLabel><Table2 /></RibbonButton>;
}

export function PivotTableRibbon({ onLoadSample }: { onLoadSample: () => void }) {
  const viewer = useXlsxViewer();
  const panel = usePivotPanel();
  const editable = !viewer.readOnly && !panel.dirty && Boolean(panel.draft?.id) && !panel.draft?.unsupportedFeatures.length;
  return <>
    <RibbonGroup><InsertPivotButton /><RibbonButton active={panel.open} label="Field List" onClick={() => panel.setOpen(!panel.open)} showLabel><LayoutList /></RibbonButton></RibbonGroup>
    <RibbonGroup>
      <RibbonButton disabled={!editable} label="Refresh" onClick={() => panel.refresh()} showLabel><RefreshCw /></RibbonButton>
      <RibbonButton disabled={viewer.readOnly || viewer.pivotTables.length === 0 || panel.dirty} label="Refresh All" onClick={() => { panel.setOpen(true); panel.refresh(true); }} showLabel><RefreshCw /></RibbonButton>
      <RibbonButton disabled={!editable || !viewer.activeCell} label="Show Details" onClick={panel.showDetails} showLabel><Search /></RibbonButton>
      <RibbonButton disabled={!editable || !panel.draft?.values.length} label="PivotChart" onClick={() => panel.run(() => { if (panel.draft?.id) viewer.createPivotChart(panel.draft.id); })} showLabel><ChartColumn /></RibbonButton>
    </RibbonGroup>
    <RibbonGroup><RibbonButton label="Pivot workbook" onClick={onLoadSample} showLabel><Table2 /></RibbonButton></RibbonGroup>
    {panel.dirty ? <span className="px-2 text-xs text-muted-foreground">Apply field changes to update the report.</span> : null}
  </>;
}

export function PivotRefreshAllButton() {
  const viewer = useXlsxViewer();
  const panel = usePivotPanel();
  return <RibbonButton disabled={viewer.readOnly || !viewer.pivotTables.length || panel.dirty} label="Refresh PivotTables" onClick={() => { panel.setOpen(true); panel.refresh(true); }} showLabel><RefreshCw /></RibbonButton>;
}

function Choice({ label, value, options, onChange, description, disabled }: {
  label: string; value: string; options: Options; onChange: (value: string) => void; description?: string; disabled?: boolean;
}) {
  const items = options.map(([value, label]) => ({ value, label }));
  return <Field className="w-full gap-1.5" disabled={disabled}>
    <FieldLabel className="text-xs font-normal">{label}</FieldLabel>
    <Select disabled={disabled} items={items} onValueChange={value => { if (value !== null) onChange(value); }} value={value}>
      <SelectTrigger className="min-w-0 text-xs" size="sm"><SelectValue placeholder="Choose…" /></SelectTrigger>
      <SelectPopup>{items.map(item => <SelectItem className="text-xs" key={item.value} value={item.value}>{item.label}</SelectItem>)}</SelectPopup>
    </Select>
    {description ? <FieldDescription>{description}</FieldDescription> : null}
  </Field>;
}

function TextField({ label, value, onChange, description, ...props }: {
  label: string; value: string; onChange: (value: string) => void; description?: string;
} & Omit<React.ComponentProps<typeof Input>, "value" | "onChange" | "size">) {
  return <Field className="w-full gap-1.5"><FieldLabel className="text-xs font-normal">{label}</FieldLabel><Input {...props} onChange={event => onChange(event.target.value)} size="sm" type={props.type ?? "text"} value={value} />{description ? <FieldDescription>{description}</FieldDescription> : null}</Field>;
}

function Toggle({ label, checked, onChange, disabled }: { label: string; checked: boolean; onChange: (checked: boolean) => void; disabled?: boolean }) {
  return <Label className="cursor-pointer gap-2 py-0.5 text-xs font-normal"><Checkbox checked={checked} disabled={disabled} onCheckedChange={onChange} />{label}</Label>;
}

function Section({ title, children, open = false }: { title: string; children: React.ReactNode; open?: boolean }) {
  return <details className="group border-b last:border-0" open={open || undefined}><summary className="flex cursor-pointer list-none items-center gap-2 px-4 py-3 text-xs font-semibold outline-none focus-visible:ring-2 focus-visible:ring-inset focus-visible:ring-ring [&::-webkit-details-marker]:hidden"><ChevronRight className="size-3.5 shrink-0 group-open:hidden" /><ChevronDown className="hidden size-3.5 shrink-0 group-open:block" />{title}</summary><div className="flex flex-col gap-3 px-4 pb-4">{children}</div></details>;
}

export function PivotTablePanel() {
  const viewer = useXlsxViewer();
  const panel = usePivotPanel();
  const [tab, setTab] = React.useState("fields");
  const [deletePending, setDeletePending] = React.useState(false);
  const draft = panel.draft;
  const unsupported = Boolean(draft?.unsupportedFeatures.length);
  React.useEffect(() => { setDeletePending(false); setTab("fields"); }, [draft?.id]);
  if (!panel.open) return null;
  return <aside aria-label="PivotTable Fields" className="flex h-full min-h-0 w-full min-w-0 flex-col bg-background">
    <div className="flex items-center gap-2 border-b px-4 py-3"><Table2 className="size-4 text-brand" /><h2 className="flex-1 text-sm font-semibold">PivotTable Fields</h2><Button aria-label="Close PivotTable Fields" onClick={() => panel.setOpen(false)} size="icon-xs" type="button" variant="ghost"><X /></Button></div>
    {viewer.pivotTables.length > 0 ? <div className="border-b px-4 py-3"><Choice disabled={panel.dirty} label="PivotTable" onChange={panel.select} options={viewer.pivotTables.map(pivot => [pivot.id, `${pivot.name} · ${pivot.target.sheet}`])} value={draft?.id ?? ""} />{panel.dirty ? <p className="mt-1.5 text-xs text-muted-foreground">Apply or discard changes before switching reports.</p> : null}</div> : null}
    {panel.error ? <div className="m-3 rounded-md border border-destructive/30 bg-destructive/5 p-3 text-xs text-destructive-foreground" role="alert">{panel.error}</div> : null}
    {panel.stale ? <div className="border-b bg-amber-500/5 px-4 py-3 text-xs" role="alert">This report changed while you were editing. Discard your draft to load the current report before making further changes.</div> : null}
    {panel.status ? <div aria-live="polite" className="flex items-start gap-2 border-b px-4 py-2.5 text-xs text-muted-foreground"><Check className="mt-0.5 size-3 shrink-0 text-brand" />{panel.status}</div> : null}
    {draft && unsupported ? <div className="border-b bg-amber-500/5 px-4 py-3 text-xs leading-relaxed"><p className="font-medium">This PivotTable is preserved for Excel.</p><p className="mt-1 text-muted-foreground">Its saved results remain visible. Editing and refresh are unavailable for these features:</p><ul className="mt-1 list-disc ps-4 text-muted-foreground">{draft.unsupportedFeatures.map(feature => <li key={feature}>{feature}</li>)}</ul></div> : null}
    {draft?.refreshError ? <div className="border-b bg-amber-500/5 px-4 py-3 text-xs" role="alert"><p className="font-medium">Refresh on open did not complete.</p><p className="mt-1 text-muted-foreground">{draft.refreshError} The saved results are shown.</p></div> : null}
    {viewer.readOnly ? <p className="border-b px-4 py-3 text-xs text-muted-foreground">Read only. Turn off Read only to change PivotTables.</p> : null}
    {panel.details ? <DrillDownView detail={panel.details} onBack={panel.clearDetails} /> : !draft ? <div className="flex flex-1 flex-col items-start justify-center gap-3 p-6"><Table2 className="size-7 text-brand" /><h3 className="text-sm font-semibold">Summarize your worksheet</h3><p className="text-xs leading-relaxed text-muted-foreground">Select data with a header row, then choose PivotTable. Arrange fields to compare categories, calculate totals, and explore the underlying records.</p><Button disabled={!viewer.activeSheet || viewer.readOnly} onClick={panel.create} size="sm" type="button"><Plus />Create PivotTable</Button>{viewer.pivotTables.length ? <p className="text-xs text-muted-foreground">Or choose an existing report above.</p> : null}</div> : <>
      <Tabs className="min-h-0 flex-1 gap-0" onValueChange={value => setTab(String(value))} value={tab}>
        <TabsList className="mx-4 my-2 w-auto" size="sm"><TabsTab value="fields">Fields</TabsTab><TabsTab value="design">Design</TabsTab><TabsTab value="source">Source</TabsTab></TabsList>
        <div className="min-h-0 flex-1 overflow-y-auto overscroll-contain">
          <fieldset className="min-w-0" disabled={viewer.readOnly || unsupported}>
            <TabsPanel value="fields"><FieldList pivot={draft} /><FieldAreas pivot={draft} /><Section title="Visual filters"><PivotVisualFilters disabled={viewer.readOnly || unsupported} onChange={next => panel.edit(() => next)} pivot={draft} /></Section><CalculatedFields pivot={draft} /></TabsPanel>
            <TabsPanel value="design"><PivotDesign pivot={draft} /></TabsPanel>
            <TabsPanel value="source"><PivotSource pivot={draft} /></TabsPanel>
          </fieldset>
        </div>
      </Tabs>
      <div className="flex shrink-0 flex-col gap-2 border-t bg-chrome/50 px-4 py-3">
        {panel.sourceNeedsLoad ? <button className="text-left text-xs text-brand underline underline-offset-2" onClick={() => setTab("source")} type="button">Load fields from the changed source before applying.</button> : null}
        <div className="flex items-center gap-2"><Button className="flex-1" disabled={viewer.readOnly || unsupported || !panel.dirty || !(draft.values.length || draft.rows.length || draft.columns.length) || panel.sourceNeedsLoad || panel.stale} onClick={panel.apply} size="sm" type="button">{draft.id ? "Apply Changes" : "Create PivotTable"}</Button><Button disabled={!panel.dirty} onClick={panel.reset} size="sm" type="button" variant="outline">Discard</Button></div>
        {draft.id ? <div className="flex items-center justify-between"><Button disabled={viewer.readOnly || unsupported || panel.dirty} onClick={() => panel.refresh()} size="xs" type="button" variant="ghost"><RefreshCw />Refresh</Button>{deletePending ? <div className="flex items-center gap-1"><Button onClick={() => setDeletePending(false)} size="xs" type="button" variant="ghost">Cancel</Button><Button disabled={viewer.readOnly} onClick={() => { panel.remove(); setDeletePending(false); }} size="xs" type="button" variant="destructive">Remove report</Button></div> : <Button disabled={viewer.readOnly || unsupported || panel.dirty} onClick={() => setDeletePending(true)} size="xs" type="button" variant="ghost"><Trash2 />Remove</Button>}</div> : null}
      </div>
    </>}
  </aside>;
}

function FieldList({ pivot }: { pivot: XlsxPivotTable }) {
  const panel = usePivotPanel();
  const [query, setQuery] = React.useState("");
  const [area, setArea] = React.useState<Area>("rows");
  const selected = new Set([...pivot.rows.map(field => field.field), ...pivot.columns.map(field => field.field), ...pivot.reportFilters, ...pivot.values.filter(field => !field.formula?.trim()).map(field => field.field)]);
  const add = (field: SourceField) => panel.addField(area, field.index);
  return <div className="flex flex-col gap-2 border-b px-4 pb-3 pt-1">
    <p className="text-xs text-muted-foreground">Add fields to an area, or drag them into place.</p>
    <Input aria-label="Search pivot fields" onChange={event => setQuery(event.target.value)} placeholder="Search fields" size="sm" type="search" value={query} />
    <div className="flex items-end gap-2"><Choice label="Add fields to" onChange={value => setArea(value as Area)} options={AREAS.map(area => [area, AREA_LABELS[area]])} value={area} /><span className="shrink-0 pb-1.5 text-[11px] text-muted-foreground">{pivot.fields.length} fields</span></div>
    <div className="max-h-44 overflow-y-auto rounded-md border py-1">
      {pivot.fields.filter(field => field.name.toLowerCase().includes(query.toLowerCase())).map(field => <div className="flex items-center gap-1.5 px-2 py-1 hover:bg-accent/60" draggable key={field.index} onDragStart={event => { event.dataTransfer.setData("application/x-pivot-field", String(field.index)); event.dataTransfer.effectAllowed = "copy"; }}>
        <GripVertical aria-hidden className="size-3 text-muted-foreground/60" /><span className="flex-1 truncate text-xs" title={field.name}>{field.name}</span>{selected.has(field.index) ? <Check aria-label="In report" className="size-3 text-brand" /> : null}<Button aria-label={`Add ${field.name} to ${AREA_LABELS[area]}`} disabled={area !== "values" && (pivot.rows.some(item => item.field === field.index) || pivot.columns.some(item => item.field === field.index) || pivot.reportFilters.includes(field.index))} onClick={() => add(field)} size="icon-xs" type="button" variant="ghost"><Plus /></Button>
      </div>)}
      {pivot.fields.length === 0 ? <p className="px-3 py-4 text-xs text-muted-foreground">Choose a source range and load its fields on the Source tab.</p> : null}
    </div>
  </div>;
}

function FieldAreas({ pivot }: { pivot: XlsxPivotTable }) {
  const viewer = useXlsxViewer();
  const panel = usePivotPanel();
  const [expanded, setExpanded] = React.useState<string | null>(null);
  const [dropArea, setDropArea] = React.useState<Area | null>(null);
  const move = (area: Area, index: number, offset: number) => panel.moveField(area, index, area, index + offset);
  const remove = (area: Area, index: number) => panel.removeField(area, index);
  const acceptDrop = (event: React.DragEvent, area: Area, before?: number) => {
    setDropArea(null);
    if (viewer.readOnly || pivot.unsupportedFeatures.length) return;
    const rawEntry = event.dataTransfer.getData("application/x-pivot-entry");
    const rawField = event.dataTransfer.getData("application/x-pivot-field");
    if (!rawEntry && !rawField) return;
    event.preventDefault(); event.stopPropagation();
    if (rawEntry) {
      try {
        const entry = JSON.parse(rawEntry) as { area: Area; index: number };
        if (AREAS.includes(entry.area) && Number.isInteger(entry.index)) panel.moveField(entry.area, entry.index, area, before !== undefined && entry.area === area && entry.index < before ? before - 1 : before);
      } catch { return; }
    } else {
      const field = pivot.fields.find(item => item.index === Number(rawField));
      if (field) panel.addField(area, field.index);
    }
    setExpanded(null);
  };
  return <div className="flex flex-col border-b">
    {AREAS.map(area => <section aria-label={AREA_LABELS[area]} className={cn("border-b px-4 py-3 last:border-0", dropArea === area && "bg-brand/5 ring-1 ring-inset ring-brand/40")} key={area}
      onDragLeave={event => { if (!event.currentTarget.contains(event.relatedTarget as Node)) setDropArea(null); }}
      onDragOver={event => { if (!viewer.readOnly && (event.dataTransfer.types.includes("application/x-pivot-field") || event.dataTransfer.types.includes("application/x-pivot-entry"))) { event.preventDefault(); event.dataTransfer.dropEffect = event.dataTransfer.types.includes("application/x-pivot-entry") ? "move" : "copy"; setDropArea(area); } }}
      onDrop={event => acceptDrop(event, area)}>
      <h3 className="mb-2 flex items-center gap-1.5 text-xs font-semibold">{area === "values" ? <Sigma className="size-3.5" /> : area === "reportFilters" ? <Filter className="size-3.5" /> : <LayoutList className="size-3.5" />}{AREA_LABELS[area]}<span className="ms-auto text-[11px] font-normal text-muted-foreground">{pivot[area].length}</span></h3>
      {pivot[area].length === 0 ? <div className="rounded-md border border-dashed px-3 py-2.5 text-[11px] text-muted-foreground">Drop fields here or add from the list.</div> : <div className="space-y-1.5">{pivot[area].map((entry, index) => {
        const fieldIndex = typeof entry === "number" ? entry : entry.field;
        const title = area === "values" ? (entry as DataField).name || fieldName(pivot, fieldIndex) : (typeof entry === "object" ? entry.name : undefined) || fieldName(pivot, fieldIndex);
        const key = `${area}:${index}`;
        return <div className="rounded-md border bg-chrome/30" key={key} onDrop={event => acceptDrop(event, area, index)}>
          <div className="flex items-center gap-0.5 px-1.5 py-1" draggable={!viewer.readOnly} onDragStart={event => { event.dataTransfer.setData("application/x-pivot-entry", JSON.stringify({ area, index })); event.dataTransfer.effectAllowed = "move"; }}>
            <GripVertical aria-hidden className="size-3 shrink-0 cursor-grab text-muted-foreground/60" />
            <button aria-expanded={expanded === key} aria-label={`${title} settings`} className="flex min-w-0 flex-1 items-center gap-1.5 rounded px-1 py-1 text-left text-xs outline-none hover:bg-accent focus-visible:ring-2 focus-visible:ring-ring" onClick={() => setExpanded(expanded === key ? null : key)} type="button"><Settings2 className="size-3 shrink-0 text-muted-foreground" /><span className="truncate">{title}</span><ChevronDown className={cn("ms-auto size-3 shrink-0 text-muted-foreground", expanded === key && "rotate-180")} /></button>
            <Button aria-label={`Move ${title} up`} disabled={index === 0} onClick={() => move(area, index, -1)} size="icon-xs" type="button" variant="ghost"><ArrowUp /></Button><Button aria-label={`Move ${title} down`} disabled={index === pivot[area].length - 1} onClick={() => move(area, index, 1)} size="icon-xs" type="button" variant="ghost"><ArrowDown /></Button><Button aria-label={`Remove ${title} from ${AREA_LABELS[area]}`} onClick={() => remove(area, index)} size="icon-xs" type="button" variant="ghost"><X /></Button>
          </div>
          {expanded === key ? <div className="flex flex-col gap-3 border-t bg-background px-3 py-3">
            {area === "values" ? <ValueSettings field={entry as DataField} onChange={field => panel.edit(current => ({ ...current, values: current.values.map((value, i) => i === index ? field : value) }))} pivot={pivot} /> : <>
              {area !== "reportFilters" ? <AxisSettings field={entry as AxisField} onAddLevel={unit => panel.edit(current => ({ ...current, [area]: [...current[area], { ...entry as AxisField, name: unit[0].toUpperCase() + unit.slice(1), grouping: { kind: "date", unit, interval: 1, groups: [] }, items: [], collapsed: [] }] }))} onChange={field => panel.edit(current => ({ ...current, [area]: current[area].map((value, i) => i === index ? field : value) }))} pivot={pivot} /> : null}
              <FilterSettings axis={area === "rows" || area === "columns" ? entry as AxisField : undefined} field={fieldIndex} pivot={pivot} />
              <Choice label="Move to" onChange={value => { panel.moveField(area, index, value as Area); setExpanded(null); }} options={AREAS.filter(target => target !== "values").map(target => [target, AREA_LABELS[target]])} value={area} />
            </>}
          </div> : null}
        </div>;
      })}</div>}
    </section>)}
  </div>;
}

function ValueSettings({ pivot, field, onChange }: { pivot: XlsxPivotTable; field: DataField; onChange: (field: DataField) => void }) {
  const needsBase = ["percent", "difference", "percentDifference", "runningTotal", "percentRunningTotal", "rankAscending", "rankDescending"].includes(field.showAs);
  const needsItem = ["percent", "difference", "percentDifference"].includes(field.showAs);
  const axisIndexes = [...pivot.rows, ...pivot.columns].map(field => field.field);
  const baseField = pivot.fields.find(item => item.index === field.baseField);
  const baseAxes = [...pivot.rows, ...pivot.columns].filter(item => item.field === field.baseField);
  const baseAxis = baseAxes.find(item => !field.baseGroupingUnit || item.grouping?.unit === field.baseGroupingUnit);
  return <>
    <TextField label="Custom name" onChange={name => onChange({ ...field, name })} value={field.name} />
    <Choice label="Summarize values by" onChange={aggregation => onChange({ ...field, aggregation: aggregation as DataField["aggregation"] })} options={AGGREGATIONS} value={field.aggregation} />
    <Choice label="Show values as" onChange={showAs => onChange({ ...field, showAs: showAs as DataField["showAs"], baseField: field.baseField ?? axisIndexes[0] })} options={SHOW_AS} value={field.showAs} />
    {needsBase ? <Choice label="Base field" onChange={value => onChange({ ...field, baseField: Number(value), baseItem: null, baseGroupingUnit: null })} options={pivot.fields.filter(item => axisIndexes.includes(item.index)).map(item => [String(item.index), item.name])} value={field.baseField == null ? "" : String(field.baseField)} /> : null}
    {needsBase && baseAxes.length > 1 ? <Choice label="Base grouping level" onChange={baseGroupingUnit => onChange({ ...field, baseGroupingUnit, baseItem: null })} options={baseAxes.filter(item => item.grouping?.kind === "date").map(item => [item.grouping!.unit, item.name || item.grouping!.unit])} value={field.baseGroupingUnit ?? baseAxes[0].grouping?.unit ?? ""} /> : null}
    {needsItem ? <Choice label="Base item" onChange={baseItem => onChange({ ...field, baseItem })} options={[["previous", "(previous)"], ["next", "(next)"], ...((baseAxis?.items?.length ? baseAxis.items : baseField?.items)?.map(item => [item.key, item.caption] as const) ?? [])]} value={field.baseItem ?? "previous"} /> : null}
    {field.formula?.trim() ? <TextField label="Calculated source field name" onChange={sourceName => onChange({ ...field, sourceName: sourceName || null })} value={field.sourceName ?? field.name} /> : null}
    <TextField label="Number format" onChange={numberFormat => onChange({ ...field, numberFormat: numberFormat || null })} placeholder="#,##0.00" value={field.numberFormat ?? ""} />
    <TextField description="Use field names (quote names with spaces), + - * / ^ %, parentheses, and SUM, PRODUCT, AVERAGE, MIN, MAX, ABS, ROUND, or SQRT." label="Calculated field formula" onChange={formula => onChange({ ...field, formula: formula || null })} placeholder="Leave blank for a source field" value={field.formula ?? ""} />
  </>;
}

function AxisSettings({ pivot, field, onChange, onAddLevel }: { pivot: XlsxPivotTable; field: AxisField; onChange: (field: AxisField) => void; onAddLevel: (unit: Grouping["unit"]) => void }) {
  const source = pivot.fields.find(item => item.index === field.field);
  const availableUnits = (["years", "quarters", "months", "days", "hours", "minutes", "seconds"] as const).filter(unit => ![...pivot.rows, ...pivot.columns].some(item => item.field === field.field && item.grouping?.kind === "date" && item.grouping.unit === unit));
  return <>
    <TextField label="Custom name" onChange={name => onChange({ ...field, name: name || null })} placeholder={source?.name} value={field.name ?? ""} />
    <Choice label="Sort" onChange={sort => onChange({ ...field, sort: sort as AxisField["sort"], sortMeasure: field.sortMeasure ?? 0 })} options={[["ascending", "Ascending (A to Z)"], ["descending", "Descending (Z to A)"], ["manual", "Source order"], ["valueAscending", "Smallest to Largest by Value"], ["valueDescending", "Largest to Smallest by Value"]]} value={field.sort} />
    {field.sort === "valueAscending" || field.sort === "valueDescending" ? <Choice label="Sort by value" onChange={value => onChange({ ...field, sortMeasure: Number(value) })} options={pivot.values.map((item, index) => [String(index), item.name])} value={String(field.sortMeasure ?? 0)} /> : null}
    <Toggle checked={field.subtotals} label="Show subtotals for this field" onChange={subtotals => onChange({ ...field, subtotals })} />
    {field.subtotals ? <details className="group"><summary className="flex cursor-pointer list-none items-center gap-1 text-xs [&::-webkit-details-marker]:hidden"><ChevronRight className="size-3 group-open:rotate-90" />Custom subtotal functions{field.subtotalFunctions?.length ? ` (${field.subtotalFunctions.length})` : ""}</summary><div className="mt-2 grid grid-cols-2 gap-1">{AGGREGATIONS.map(([value, label]) => <Toggle checked={field.subtotalFunctions?.some(item => item === value) ?? false} key={value} label={label} onChange={checked => onChange({ ...field, subtotalFunctions: checked ? [...(field.subtotalFunctions ?? []), value as DataField["aggregation"]] : field.subtotalFunctions?.filter(item => item !== value) })} />)}</div><p className="mt-2 text-[11px] text-muted-foreground">Leave unchecked to use each value field’s summary function.</p></details> : null}
    <Toggle checked={field.showItemsWithNoData ?? false} label="Show items with no data" onChange={showItemsWithNoData => onChange({ ...field, showItemsWithNoData })} />
    <GroupingSettings field={field} onChange={onChange} source={source} />
    {field.grouping?.kind === "date" && availableUnits.length ? <Choice label="Add date grouping level" onChange={value => onAddLevel(value as Grouping["unit"])} options={availableUnits.map(unit => [unit, unit[0].toUpperCase() + unit.slice(1)])} value="" /> : null}
    <details className="group"><summary className="flex cursor-pointer list-none items-center gap-1 text-xs [&::-webkit-details-marker]:hidden"><ChevronRight className="size-3 group-open:rotate-90" />Collapse members{field.collapsed.length ? ` (${field.collapsed.length})` : ""}</summary><div className="mt-2">{field.grouping && !field.items?.length ? <p className="text-[11px] text-muted-foreground">Apply the grouping first to choose groups to collapse.</p> : <MemberChecklist items={field.items?.length ? field.items : source?.items ?? []} onChange={collapsed => onChange({ ...field, collapsed })} selected={field.collapsed} />}</div></details>
  </>;
}

function GroupingSettings({ field, onChange, source }: { field: AxisField; onChange: (field: AxisField) => void; source?: SourceField }) {
  const grouping = field.grouping;
  const update = (patch: Partial<Grouping>) => onChange({ ...field, grouping: { ...grouping!, ...patch }, items: [], collapsed: [] });
  return <>
    <Choice label="Grouping" onChange={kind => onChange({ ...field, grouping: kind === "none" ? null : { kind: kind as Grouping["kind"], unit: "months", interval: 1, groups: [] }, items: [], collapsed: [] })} options={[["none", "No grouping"], ["number", "Numeric intervals"], ["date", "Dates"], ["manual", "Named groups"]]} value={grouping?.kind ?? "none"} />
    {grouping?.kind === "date" ? <Choice label="Group dates by" onChange={unit => update({ unit: unit as Grouping["unit"] })} options={["years", "quarters", "months", "days", "hours", "minutes", "seconds"].map(unit => [unit, unit[0].toUpperCase() + unit.slice(1)])} value={grouping.unit} /> : null}
    {grouping && grouping.kind !== "manual" ? <div className="flex flex-col gap-2"><TextField label={grouping.kind === "date" ? "Interval" : "By"} min={0.000001} onChange={value => update({ interval: Number(value) })} step="any" type="number" value={String(grouping.interval)} /><div className="grid grid-cols-2 gap-2"><TextField label="Starting at" onChange={value => update({ start: value === "" ? null : Number(value) })} placeholder="Automatic" step="any" type="number" value={grouping.start == null ? "" : String(grouping.start)} /><TextField label="Ending at" onChange={value => update({ end: value === "" ? null : Number(value) })} placeholder="Automatic" step="any" type="number" value={grouping.end == null ? "" : String(grouping.end)} /></div>{grouping.kind === "date" ? <p className="text-[11px] text-muted-foreground">Optional bounds use Excel date serial values.</p> : null}</div> : null}
    {grouping?.kind === "manual" ? <div className="flex flex-col gap-2">{grouping.groups.map((group, index) => <div className="space-y-2 rounded-md border p-2" key={index}><div className="flex items-end gap-1"><TextField label="Group name" onChange={name => update({ groups: grouping.groups.map((entry, i) => i === index ? { ...entry, name } : entry) })} value={group.name} /><Button aria-label={`Remove group ${group.name}`} onClick={() => update({ groups: grouping.groups.filter((_, i) => i !== index) })} size="icon-sm" type="button" variant="ghost"><X /></Button></div><MemberChecklist items={source?.items ?? []} onChange={members => update({ groups: grouping.groups.map((entry, i) => i === index ? { ...entry, members } : entry) })} selected={group.members} /></div>)}<Button onClick={() => update({ groups: [...grouping.groups, { name: `Group${grouping.groups.length + 1}`, members: [] }] })} size="xs" type="button" variant="outline"><Plus />Add group</Button></div> : null}
  </>;
}

function MemberChecklist({ items, selected, onChange }: { items: SourceField["items"]; selected: string[]; onChange: (values: string[]) => void }) {
  const [search, setSearch] = React.useState("");
  const visible = items.filter(item => item.caption.toLowerCase().includes(search.toLowerCase()));
  return <div className="space-y-1.5"><Input aria-label="Search field members" onChange={event => setSearch(event.target.value)} placeholder="Search items" size="sm" type="search" value={search} /><div className="flex items-center gap-2 text-[11px]"><button className="rounded text-brand outline-none hover:underline focus-visible:ring-2 focus-visible:ring-ring" onClick={() => onChange(items.map(item => item.key))} type="button">Select all</button><button className="rounded text-muted-foreground outline-none hover:underline focus-visible:ring-2 focus-visible:ring-ring" onClick={() => onChange([])} type="button">Clear</button><span className="ms-auto text-muted-foreground">{selected.length} selected</span></div><div className="max-h-36 space-y-0.5 overflow-y-auto rounded border px-2 py-1.5">{visible.slice(0, 300).map(item => <Toggle checked={selected.includes(item.key)} key={item.key} label={item.caption || "(blank)"} onChange={checked => onChange(checked ? [...selected, item.key] : selected.filter(key => key !== item.key))} />)}{visible.length > 300 ? <p className="py-1 text-[11px] text-muted-foreground">Search to narrow {visible.length.toLocaleString()} items.</p> : null}{!visible.length ? <p className="py-2 text-[11px] text-muted-foreground">No matching items</p> : null}</div></div>;
}

function FilterSettings({ pivot, field, axis }: { pivot: XlsxPivotTable; field: number; axis?: AxisField }) {
  const panel = usePivotPanel();
  const source = pivot.fields.find(item => item.index === field);
  const members = axis?.items?.length ? axis.items : source?.items ?? [];
  const groupingUnit = axis?.grouping ? axis.grouping.kind === "date" ? axis.grouping.unit : axis.grouping.kind : null;
  const filters = pivot.filters.map((filter, index) => ({ filter, index })).filter(item => item.filter.field === field && (item.filter.groupingUnit ?? null) === groupingUnit);
  const kinds: Options = [["members", "Select items"], ["label", "Label filter"], ["value", "Value filter"], ["top10", "Top / bottom filter"]];
  const available = kinds.filter(([kind]) => !filters.some(item => item.filter.kind === kind));
  const update = (index: number, next: PivotFilter | null) => panel.edit(current => ({
    ...current,
    filters: current.filters.flatMap((filter, i) => i === index ? next ? [next] : [] : [filter]),
  }));
  return <div className="flex flex-col gap-3 border-t pt-3">
    <h4 className="flex items-center gap-1.5 text-xs font-medium"><Filter className="size-3" />Field filters</h4>
    {!filters.length ? <p className="text-[11px] text-muted-foreground">All items are included.</p> : null}
    {filters.map(({ filter, index }) => <div className="flex flex-col gap-2 rounded-md border p-2" key={index}>
      <div className="flex items-center gap-2"><h5 className="flex-1 text-xs font-medium">{kinds.find(([kind]) => kind === filter.kind)?.[1]}</h5><Button aria-label={`Remove ${filter.kind} filter`} onClick={() => update(index, null)} size="icon-xs" type="button" variant="ghost"><X /></Button></div>
      {filter.kind === "members" ? <><MemberChecklist items={members} onChange={values => update(index, { ...filter, values })} selected={filter.values} /><p className="text-[11px] text-muted-foreground">Only selected items are included.</p></> : null}
      {["label", "value"].includes(filter.kind) ? <><Choice label="Condition" onChange={operator => update(index, { ...filter, operator: operator as PivotFilter["operator"] })} options={filter.kind === "value" ? OPERATORS.filter(([operator]) => !["contains", "notContains", "beginsWith", "notBeginsWith", "endsWith", "notEndsWith"].includes(operator)) : OPERATORS} value={filter.operator} /><TextField label="Value" onChange={value => update(index, { ...filter, value })} type={filter.kind === "value" ? "number" : "text"} value={filter.value} />{["between", "notBetween"].includes(filter.operator) ? <TextField label="And" onChange={value2 => update(index, { ...filter, value2 })} type={filter.kind === "value" ? "number" : "text"} value={filter.value2} /> : null}</> : null}
      {filter.kind === "top10" ? <div className="grid grid-cols-2 gap-2"><Choice label="Show" onChange={value => update(index, { ...filter, top: value === "top" })} options={[["top", "Top"], ["bottom", "Bottom"]]} value={filter.top ? "top" : "bottom"} /><TextField label="Amount" min={1} onChange={value => update(index, { ...filter, count: Number(value) })} type="number" value={String(filter.count)} /><Choice label="Unit" onChange={value => update(index, { ...filter, percent: value === "percent" })} options={[["items", "Items"], ["percent", "Percent"]]} value={filter.percent ? "percent" : "items"} /></div> : null}
      {["value", "top10"].includes(filter.kind) ? <Choice label="By value" onChange={value => update(index, { ...filter, measure: Number(value) })} options={pivot.values.map((value, i) => [String(i), value.name])} value={String(filter.measure)} /> : null}
    </div>)}
    {available.length ? <Choice label="Add a filter" onChange={kind => panel.edit(current => ({ ...current, filters: [...current.filters, { ...newFilter(field), groupingUnit, kind: kind as PivotFilter["kind"], values: kind === "members" ? members.map(item => item.key) : [] }] }))} options={available} value="" /> : null}
  </div>;
}

function CalculatedFields({ pivot }: { pivot: XlsxPivotTable }) {
  const panel = usePivotPanel();
  return <Section title="Calculated fields and items">
    <p className="text-xs text-muted-foreground">Add a value calculation, or add a calculated member to a row or column field.</p>
    <Button disabled={!pivot.fields.length} onClick={() => panel.edit(current => ({ ...current, values: [...current.values, { field: current.fields[0].index, name: `Calculated ${current.values.length + 1}`, aggregation: "sum", showAs: "normal", formula: "0" }] }))} size="sm" type="button" variant="outline"><Sigma />Add calculated field</Button>
    <p className="text-[11px] text-muted-foreground">Edit its formula in Values → field settings.</p>
    {pivot.calculatedItems.map((item, index) => <div className="space-y-2 rounded-md border p-3" key={index}><div className="flex items-center justify-between"><span className="text-xs font-medium">Calculated item {index + 1}</span><Button aria-label={`Remove calculated item ${item.name}`} onClick={() => panel.edit(current => ({ ...current, calculatedItems: current.calculatedItems.filter((_, i) => i !== index) }))} size="icon-xs" type="button" variant="ghost"><X /></Button></div><Choice label="Field" onChange={value => panel.edit(current => ({ ...current, calculatedItems: current.calculatedItems.map((entry, i) => i === index ? { ...entry, field: Number(value) } : entry) }))} options={pivot.fields.map(field => [String(field.index), field.name])} value={String(item.field)} /><TextField label="Name" onChange={name => panel.edit(current => ({ ...current, calculatedItems: current.calculatedItems.map((entry, i) => i === index ? { ...entry, name } : entry) }))} value={item.name} /><TextField label="Formula" onChange={formula => panel.edit(current => ({ ...current, calculatedItems: current.calculatedItems.map((entry, i) => i === index ? { ...entry, formula } : entry) }))} value={item.formula} /></div>)}
    <Button disabled={!pivot.fields.length} onClick={() => panel.edit(current => ({ ...current, calculatedItems: [...current.calculatedItems, { field: current.rows[0]?.field ?? current.fields[0].index, name: `Calculated item ${current.calculatedItems.length + 1}`, formula: "0" }] }))} size="sm" type="button" variant="outline"><Plus />Add calculated item</Button>
  </Section>;
}

function PivotDesign({ pivot }: { pivot: XlsxPivotTable }) {
  const panel = usePivotPanel();
  const update = (patch: Partial<XlsxPivotTable>) => panel.edit(current => ({ ...current, ...patch }));
  return <>
    <Section open title="Report layout"><Choice label="Report form" onChange={layout => update({ layout: layout as XlsxPivotTable["layout"] })} options={[["compact", "Compact Form"], ["outline", "Outline Form"], ["tabular", "Tabular Form"]]} value={pivot.layout} /><Choice label="Multiple value fields" onChange={value => update({ valuesOnRows: value === "rows" })} options={[["columns", "Show values in columns"], ["rows", "Show values in rows"]]} value={pivot.valuesOnRows ? "rows" : "columns"} /><Toggle checked={pivot.repeatLabels} label="Repeat all item labels" onChange={repeatLabels => update({ repeatLabels })} /></Section>
    <Section open title="Subtotals and grand totals"><Choice label="Subtotals" onChange={subtotalsPosition => update({ subtotalsPosition: subtotalsPosition as XlsxPivotTable["subtotalsPosition"] })} options={[["top", "Show at Top of Group"], ["bottom", "Show at Bottom of Group"], ["none", "Do Not Show Subtotals"]]} value={pivot.subtotalsPosition} /><Toggle checked={pivot.rowGrandTotals} label="Show grand totals for rows" onChange={rowGrandTotals => update({ rowGrandTotals })} /><Toggle checked={pivot.columnGrandTotals} label="Show grand totals for columns" onChange={columnGrandTotals => update({ columnGrandTotals })} /></Section>
    <Section open title="Style"><Toggle checked={pivot.preserveFormatting !== false} label="Preserve cell formatting on update" onChange={preserveFormatting => update({ preserveFormatting })} /><Choice label="Excel PivotTable style" onChange={styleName => update({ styleName })} options={STYLES.some(([name]) => name === pivot.styleName) ? STYLES : [[pivot.styleName, pivot.styleName], ...STYLES]} value={pivot.styleName} /><p className="text-xs text-muted-foreground">The selected style is saved with the workbook. Value number formats are set individually in field settings.</p></Section>
    <Section title="PivotChart"><PivotChartSettings pivot={pivot} /></Section>
  </>;
}

function PivotChartSettings({ pivot }: { pivot: XlsxPivotTable }) {
  const viewer = useXlsxViewer();
  const panel = usePivotPanel();
  const [chartType, setChartType] = React.useState<NonNullable<XlsxPivotTable["charts"]>[number]["chartType"]>("ColumnClustered");
  const [measure, setMeasure] = React.useState("all");
  const [title, setTitle] = React.useState(pivot.name);
  const [sheet, setSheet] = React.useState(pivot.target.sheet);
  const [cell, setCell] = React.useState("");
  const [created, setCreated] = React.useState(false);
  const create = () => panel.run(() => {
    viewer.createPivotChart(pivot.id, { chartType, title, sheet, cell: cell.trim() || undefined, measure: measure === "all" ? null : Number(measure) });
    setCreated(true);
  });
  return <div className="flex flex-col gap-3">
    <Choice label="Chart type" onChange={value => { setChartType(value as typeof chartType); setCreated(false); }} options={CHART_TYPES} value={chartType} />
    <Choice label="Values" onChange={setMeasure} options={[["all", "All value fields"], ...pivot.values.map((value, index) => [String(index), value.name] as const)]} value={measure} />
    <TextField label="Chart title" onChange={setTitle} value={title} />
    <Choice label="Chart worksheet" onChange={setSheet} options={viewer.sheets.map(sheet => [sheet.name, sheet.name])} value={sheet} />
    <TextField label="Starting cell (optional)" onChange={setCell} placeholder="Automatic placement" value={cell} />
    {["Pie", "Doughnut"].includes(chartType) ? <p className="text-xs text-muted-foreground">Pie and doughnut charts display the first series.</p> : null}
    <Button disabled={viewer.readOnly || panel.dirty || !pivot.id || !pivot.values.length} onClick={create} size="sm" type="button" variant="outline"><ChartColumn />Insert PivotChart</Button>
    {panel.dirty ? <p className="text-xs text-muted-foreground">Apply field changes before inserting a chart.</p> : null}
    {created ? <p aria-live="polite" className="text-xs text-brand">PivotChart inserted. Refreshing the report updates its data.</p> : null}
    {pivot.charts?.length ? <p className="text-xs text-muted-foreground">{pivot.charts.length} linked chart{pivot.charts.length === 1 ? "" : "s"} in this workbook.</p> : null}
  </div>;
}

function PivotSource({ pivot }: { pivot: XlsxPivotTable }) {
  const viewer = useXlsxViewer();
  const panel = usePivotPanel();
  const sheetOptions: Options = viewer.sheets.map(sheet => [sheet.name, sheet.name]);
  const newSheet = () => panel.run(() => {
    let name = "Pivot Report";
    for (let index = 2; viewer.sheets.some(sheet => sheet.name === name); index++) name = `Pivot Report ${index}`;
    viewer.addSheet(name);
    panel.edit(current => ({ ...current, target: { sheet: name, cell: "A3" } }));
  });
  return <>
    <Section open title="Report"><TextField label="PivotTable name" onChange={name => panel.edit(current => ({ ...current, name }))} value={pivot.name} /></Section>
    <Section open title="Data source">
      <Choice label="Source worksheet" onChange={sheet => panel.edit(current => ({ ...current, source: { ...current.source, sheet } }))} options={sheetOptions} value={pivot.source.sheet} />
      <TextField description="Include the column header row." label="Source range" onChange={range => panel.edit(current => ({ ...current, source: { ...current.source, range } }))} placeholder="A1:F100" value={pivot.source.range} />
      <TextField label="Excel table name (optional)" onChange={table => panel.edit(current => ({ ...current, source: { ...current.source, table: table || null } }))} value={pivot.source.table ?? ""} />
      <Button onClick={panel.readFields} size="sm" type="button" variant="outline"><ListFilter />Load source fields</Button>
      <p className="text-xs text-muted-foreground">Changing column names resets the field arrangement when you load fields.</p>
    </Section>
    <Section open title="Destination"><Choice label="Destination worksheet" onChange={sheet => panel.edit(current => ({ ...current, target: { ...current.target, sheet } }))} options={sheetOptions} value={pivot.target.sheet} /><TextField label="Starting cell" onChange={cell => panel.edit(current => ({ ...current, target: { ...current.target, cell } }))} placeholder="A3" value={pivot.target.cell} /><Button onClick={newSheet} size="sm" type="button" variant="outline"><Plus />New worksheet</Button><p className="text-xs text-muted-foreground">Choose an empty area. Refresh keeps other worksheet content intact.</p>{pivot.outputRange ? <p className="text-xs text-muted-foreground">Current report: {pivot.target.sheet}!{pivot.outputRange}</p> : null}</Section>
    <Section title="Refresh"><Toggle checked={pivot.refreshOnLoad} label="Refresh data when opening the workbook" onChange={refreshOnLoad => panel.edit(current => ({ ...current, refreshOnLoad }))} /><p className="text-xs text-muted-foreground">Refresh uses the current worksheet values.</p></Section>
  </>;
}

function DrillDownView({ detail, onBack }: { detail: DetailRows; onBack: () => void }) {
  const viewer = useXlsxViewer();
  const panel = usePivotPanel();
  const newSheet = () => panel.run(() => {
    const context = panel.detailContext;
    if (!context) return;
    viewer.showPivotDetails(context.id, context.row, context.col);
    panel.clearDetails();
  });
  const download = () => {
    const escape = (value: string | number | boolean | null) => {
      const text = String(value ?? "");
      return `"${text.replaceAll('"', '""')}"`;
    };
    const csv = [detail.fields, ...detail.rows].map(row => row.map(escape).join(",")).join("\r\n");
    const url = URL.createObjectURL(new Blob(["\ufeff", csv], { type: "text/csv;charset=utf-8" }));
    const link = document.createElement("a"); link.href = url; link.download = "pivot-details.csv"; link.click();
    window.setTimeout(() => URL.revokeObjectURL(url), 1000);
  };
  return <div className="flex min-h-0 flex-1 flex-col"><div className="flex flex-wrap items-center gap-2 border-b px-4 py-3"><Button onClick={onBack} size="xs" type="button" variant="ghost">Back to fields</Button><span className="ms-auto text-xs text-muted-foreground">{detail.rows.length.toLocaleString()} records</span><Button disabled={viewer.readOnly || !panel.detailContext} onClick={newSheet} size="xs" type="button" variant="outline"><Plus />New worksheet</Button><Button aria-label="Download detail records as CSV" onClick={download} size="icon-xs" type="button" variant="outline"><Download /></Button></div><div className="min-h-0 flex-1 overflow-auto"><table className="w-full border-collapse whitespace-nowrap text-xs"><thead className="sticky top-0 bg-chrome"><tr>{detail.fields.map((field, index) => <th className="border-b px-3 py-2 text-left font-medium" key={index} scope="col">{field}</th>)}</tr></thead><tbody>{detail.rows.slice(0, 200).map((row, index) => <tr className="even:bg-chrome/40" key={index}>{row.map((value, column) => <td className="border-b px-3 py-1.5" key={column}>{String(value ?? "")}</td>)}</tr>)}</tbody></table>{detail.rows.length === 0 ? <p className="p-4 text-xs text-muted-foreground">No source records match this value.</p> : null}</div>{detail.rows.length > 200 ? <p className="border-t p-3 text-xs text-muted-foreground">Showing the first 200 records. Download CSV for all {detail.rows.length.toLocaleString()} records.</p> : null}</div>;
}
