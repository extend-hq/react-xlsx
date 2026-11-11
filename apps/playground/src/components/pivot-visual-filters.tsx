import * as React from "react";
import { CalendarRange, FilterX, ListFilter } from "lucide-react";
import { useXlsxViewer, type XlsxPivotTable } from "@extend-ai/react-xlsx";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Select, SelectItem, SelectPopup, SelectTrigger, SelectValue } from "@/components/ui/select";
import { cn } from "@/lib/utils";

type Member = XlsxPivotTable["fields"][number]["items"][number];
type Filter = XlsxPivotTable["filters"][number];

function memberDate(member: Member, date1904: boolean): string | null {
  if (member.key.startsWith("n:")) {
    const serial = Number(member.key.slice(2));
    if (!Number.isFinite(serial) || serial < 0 || serial > 2_958_465 || (!date1904 && Math.floor(serial) === 60)) return null;
    const base = date1904 ? Date.UTC(1904, 0, 1) : Date.UTC(1899, 11, 31);
    const days = Math.floor(serial) - (!date1904 && serial >= 60 ? 1 : 0);
    const date = new Date(base + days * 86_400_000);
    return Number.isFinite(date.getTime()) ? date.toISOString().slice(0, 10) : null;
  }
  if (member.key.startsWith("s:") && /^\d{4}-\d{2}-\d{2}(?:$|T)/.test(member.key.slice(2))) {
    return member.key.slice(2, 12);
  }
  return null;
}

export function PivotVisualFilters({ pivot, onChange, disabled = false }: {
  pivot: XlsxPivotTable;
  onChange: (pivot: XlsxPivotTable) => void;
  disabled?: boolean;
}) {
  const viewer = useXlsxViewer();
  const [fieldIndex, setFieldIndex] = React.useState(pivot.fields[0]?.index ?? 0);
  const [query, setQuery] = React.useState("");
  const [mode, setMode] = React.useState<"items" | "dates">("items");
  const [start, setStart] = React.useState("");
  const [end, setEnd] = React.useState("");
  const [period, setPeriod] = React.useState<"months" | "quarters" | "years">("months");
  const field = pivot.fields.find(item => item.index === fieldIndex) ?? pivot.fields[0];
  const grouped = [...pivot.rows, ...pivot.columns].some(axis => axis.field === field?.index && axis.grouping);
  const date1904 = React.useMemo(() => {
    try { return viewer.workbook?.getSheetByName(pivot.source.sheet).date1904 ?? false; }
    catch { return false; }
  }, [pivot.source.sheet, viewer.workbook]);
  const active = pivot.filters.find(filter => filter.field === field?.index && filter.kind === "members" && !filter.groupingUnit);
  const selected = React.useMemo(() => new Set(active?.values ?? field?.items.map(item => item.key) ?? []), [active, field]);
  const items = field?.items.filter(item => item.caption.toLocaleLowerCase().includes(query.toLocaleLowerCase())) ?? [];
  const dates = React.useMemo(() => field?.items.flatMap(member => {
    const date = memberDate(member, date1904);
    return date ? [{ ...member, date }] : [];
  }) ?? [], [field, date1904]);
  const periods = React.useMemo(() => {
    const groups = new Map<string, { label: string; keys: string[] }>();
    for (const member of dates) {
      const year = member.date.slice(0, 4);
      const month = member.date.slice(5, 7);
      const key = period === "years" ? year : period === "quarters" ? `${year} Q${Math.ceil(Number(month) / 3)}` : `${year}-${month}`;
      const group = groups.get(key) ?? { label: key, keys: [] };
      group.keys.push(member.key); groups.set(key, group);
    }
    return [...groups].sort(([a], [b]) => a.localeCompare(b));
  }, [dates, period]);
  const replace = (keys: string[] | null) => {
    if (!field) return;
    const filters = pivot.filters.filter(filter => filter.field !== field.index || filter.kind !== "members" || Boolean(filter.groupingUnit));
    if (keys !== null) {
      const next: Filter = { field: field.index, kind: "members", values: keys, operator: "equal", value: "", value2: "", measure: 0, top: true, count: 10, percent: false };
      filters.push(next);
    }
    onChange({ ...pivot, filters });
  };
  const toggle = (keys: string[]) => {
    const next = new Set(selected);
    const remove = keys.every(key => next.has(key));
    for (const key of keys) { if (remove) next.delete(key); else next.add(key); }
    replace([...next]);
  };
  if (!field) return null;

  return (
    <details className="border-t pt-3">
      <summary className="flex cursor-pointer list-none items-center gap-2 text-xs font-medium [&::-webkit-details-marker]:hidden">
        <ListFilter aria-hidden className="size-3.5" /> Visual filters
        {active ? <span className="ms-auto text-[11px] font-normal text-muted-foreground">{selected.size} selected</span> : null}
      </summary>
      <div className="mt-3 space-y-3">
        <Select disabled={disabled} items={pivot.fields.map(item => ({ label: item.name, value: String(item.index) }))} onValueChange={value => { setFieldIndex(Number(value)); setQuery(""); setStart(""); setEnd(""); }} value={String(field.index)}>
          <SelectTrigger aria-label="Visual filter field" size="sm"><SelectValue /></SelectTrigger>
          <SelectPopup>{pivot.fields.map(item => <SelectItem key={item.index} value={String(item.index)}>{item.name}</SelectItem>)}</SelectPopup>
        </Select>
        {grouped ? <p className="text-xs text-muted-foreground">This field is grouped. Select its grouped items in the field settings, or remove grouping to use the timeline and source items.</p> : <>
        <div className="flex items-center gap-1" role="group" aria-label="Visual filter mode">
          <Button aria-pressed={mode === "items"} onClick={() => setMode("items")} size="xs" type="button" variant={mode === "items" ? "secondary" : "ghost"}><ListFilter />Items</Button>
          <Button aria-pressed={mode === "dates"} onClick={() => setMode("dates")} size="xs" type="button" variant={mode === "dates" ? "secondary" : "ghost"}><CalendarRange />Timeline</Button>
          <Button className="ms-auto" disabled={disabled || !active} onClick={() => replace(null)} size="icon-xs" type="button" variant="ghost" aria-label="Clear visual filter"><FilterX /></Button>
        </div>
        {mode === "items" ? <>
          <Input aria-label="Search filter items" onChange={event => setQuery(event.target.value)} placeholder="Search items" size="sm" type="search" value={query} />
          <div className="flex gap-2 text-[11px]">
            <button className="text-brand disabled:opacity-50" disabled={disabled} onClick={() => replace(null)} type="button">Select all</button>
            <button className="text-brand disabled:opacity-50" disabled={disabled} onClick={() => replace([])} type="button">Clear selection</button>
          </div>
          <div className="grid max-h-52 grid-cols-2 gap-1 overflow-y-auto" role="group" aria-label={`${field.name} filter items`}>
            {items.map(item => <Button aria-pressed={selected.has(item.key)} className={cn("min-w-0 justify-start truncate text-xs", selected.has(item.key) && "border-brand/30 bg-brand/10 text-brand hover:bg-brand/15")} disabled={disabled} key={item.key} onClick={() => toggle([item.key])} size="sm" title={item.caption} type="button" variant="outline"><span className="truncate">{item.caption || "(empty text)"}</span></Button>)}
          </div>
          {!items.length ? <p className="text-xs text-muted-foreground">No matching items.</p> : null}
        </> : <>
          <p className="text-[11px] text-muted-foreground">Select a date field. Numeric values are interpreted as spreadsheet dates.</p>
          <div className="grid grid-cols-2 gap-2">
            <label className="space-y-1 text-[11px]">From<Input aria-label="Timeline start date" disabled={disabled} onChange={event => setStart(event.target.value)} size="sm" type="date" value={start} /></label>
            <label className="space-y-1 text-[11px]">Through<Input aria-label="Timeline end date" disabled={disabled} onChange={event => setEnd(event.target.value)} size="sm" type="date" value={end} /></label>
          </div>
          <Button disabled={disabled || !dates.length || (!start && !end) || Boolean(start && end && start > end)} onClick={() => replace(dates.filter(member => (!start || member.date >= start) && (!end || member.date <= end)).map(member => member.key))} size="sm" type="button" variant="outline">Select date range</Button>
          <Select items={[{ label: "Months", value: "months" }, { label: "Quarters", value: "quarters" }, { label: "Years", value: "years" }]} onValueChange={value => setPeriod(value as typeof period)} value={period}>
            <SelectTrigger aria-label="Timeline periods" size="sm"><SelectValue /></SelectTrigger>
            <SelectPopup><SelectItem value="months">Months</SelectItem><SelectItem value="quarters">Quarters</SelectItem><SelectItem value="years">Years</SelectItem></SelectPopup>
          </Select>
          <div className="flex max-h-40 flex-wrap gap-1 overflow-y-auto" role="group" aria-label="Timeline periods">
            {periods.map(([key, group]) => {
              const count = group.keys.filter(member => selected.has(member)).length;
              return <Button aria-pressed={count === group.keys.length ? true : count ? "mixed" : false} className={cn("text-[11px]", count > 0 && "border-brand/30 bg-brand/10 text-brand")} disabled={disabled} key={key} onClick={() => toggle(group.keys)} size="xs" title={`${count} of ${group.keys.length} dates selected`} type="button" variant="outline">{group.label}</Button>;
            })}
          </div>
          {!dates.length ? <p className="text-xs text-muted-foreground">This field has no date values.</p> : null}
        </>}
        </>}
      </div>
    </details>
  );
}
