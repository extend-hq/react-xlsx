import * as React from "react";
import { parseXlsxRange, useXlsxViewer } from "@extend-ai/react-xlsx";
import { Button } from "./ui/button";
import { Field, FieldDescription, FieldError, FieldLabel } from "./ui/field";
import { Label } from "./ui/label";
import { Switch } from "./ui/switch";
import { Textarea } from "./ui/textarea";

export const HIGHLIGHT_DEMO_RANGES = ["Overview!B24:D27", "Overview!F24:H27", "Details!B20:E24"];
export const HIGHLIGHT_DEMO_URL = "/examples/range-highlights.xlsx";

export function RangeHighlightDemo({ customStyle, onCustomStyleChange, onLoadSample }: {
  customStyle: boolean;
  onCustomStyleChange: (enabled: boolean) => void;
  onLoadSample: (autoScroll: boolean) => void;
}) {
  const { clearHighlightedRanges, highlightRanges, highlightedRanges, isLoading, sheets } = useXlsxViewer();
  const [draft, setDraft] = React.useState("B3:D6\nF3:H6");
  const [autoScroll, setAutoScroll] = React.useState(true);
  const [error, setError] = React.useState<string | null>(null);
  const id = React.useId();
  const canHighlight = sheets.length > 0 && !isLoading;

  function applyRanges(event: React.FormEvent) {
    event.preventDefault();
    const lines = draft.split(/\r?\n/);
    const invalidLine = lines.findIndex((line) => line.trim() && !parseXlsxRange(line));
    if (invalidLine >= 0) {
      setError(`Line ${invalidLine + 1}: enter a valid Excel range.`);
      return;
    }
    const references = lines.map((line) => line.trim()).filter(Boolean);
    setError(highlightRanges(references, { autoScroll })
      ? null
      : "Check the worksheet names, then apply the ranges again.");
  }

  return (
    <form className="flex w-full min-w-0 flex-wrap items-start gap-x-5 gap-y-3" onSubmit={applyRanges}>
      <Field className="min-w-[180px] max-w-sm flex-1 gap-1" data-invalid={Boolean(error)}>
        <FieldLabel htmlFor={id}>Ranges to highlight</FieldLabel>
        <Textarea
          aria-describedby={`${id}-help${error ? ` ${id}-error` : ""}`}
          aria-invalid={Boolean(error)}
          className="min-h-12 max-h-24 py-1"
          id={id}
          onChange={(event) => { setDraft(event.target.value); setError(null); }}
          rows={2}
          spellCheck={false}
          value={draft}
        />
        <FieldDescription className="text-[11px]" id={`${id}-help`}>One range per line. Sheet names are optional.</FieldDescription>
        {error ? <FieldError id={`${id}-error`}>{error}</FieldError> : null}
      </Field>
      <div className="flex flex-col gap-3 pt-1">
        <Label><Switch checked={autoScroll} onCheckedChange={setAutoScroll} size="sm" />Scroll to first range</Label>
        <Label><Switch checked={customStyle} onCheckedChange={onCustomStyleChange} size="sm" />Custom CSS</Label>
        <span aria-live="polite" className="text-muted-foreground text-[11px]">
          {highlightedRanges.length > 0 ? `${highlightedRanges.length} ranges highlighted` : "No highlights applied"}
        </span>
      </div>
      <div className="flex flex-col gap-2 pt-1">
        <div className="flex gap-1.5">
          <Button disabled={!canHighlight} size="sm" type="submit">Apply ranges</Button>
          <Button disabled={highlightedRanges.length === 0} onClick={() => { clearHighlightedRanges(); setError(null); }} size="sm" type="button" variant="outline">Clear highlights</Button>
        </div>
        <Button
          disabled={isLoading}
          onClick={() => { setDraft(HIGHLIGHT_DEMO_RANGES.join("\n")); setError(null); onLoadSample(autoScroll); }}
          size="sm"
          type="button"
          variant="outline"
        >
          Load highlight demo
        </Button>
        <span className="text-muted-foreground text-[11px]">Three ranges across two sheets.</span>
      </div>
    </form>
  );
}
