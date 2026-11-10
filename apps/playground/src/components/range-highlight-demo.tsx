import * as React from "react";
import { Eraser, Highlighter, Sparkles } from "lucide-react";
import { parseXlsxRange, useXlsxViewer } from "@extend-ai/react-xlsx";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { Switch } from "@/components/ui/switch";
import { RibbonButton, RibbonGroup } from "@/components/ribbon";

export const HIGHLIGHT_DEMO_RANGES = ["'Scatter Matrix'!B7:B17", "'Scatter Matrix'!C7:C17", "'Scatter Matrix'!D7:D17"];
export const HIGHLIGHT_DEMO_URL = "/examples/crazy-chart-zoo-subtypes.xlsx";

/** Ribbon contents for the "Review" tab: highlight arbitrary ranges in the workbook. */
export function RangeHighlightDemo({ customStyle, onCustomStyleChange, onLoadSample }: {
  customStyle: boolean;
  onCustomStyleChange: (enabled: boolean) => void;
  onLoadSample: (autoScroll: boolean) => void;
}) {
  const { clearHighlightedRanges, highlightRanges, highlightedRanges, isLoading, sheets } = useXlsxViewer();
  const [draft, setDraft] = React.useState("B3:D6, F3:H6");
  const [autoScroll, setAutoScroll] = React.useState(true);
  const [error, setError] = React.useState<string | null>(null);
  const canHighlight = sheets.length > 0 && !isLoading;

  function applyRanges(event: React.FormEvent) {
    event.preventDefault();
    const parts = draft.split(/[,\n;]/).map((part) => part.trim()).filter(Boolean);
    const invalid = parts.find((part) => !parseXlsxRange(part));
    if (invalid) {
      setError(`"${invalid}" isn't a valid range`);
      return;
    }
    setError(highlightRanges(parts, { autoScroll }) ? null : "Check the sheet names and try again");
  }

  return (
    <>
      <RibbonGroup>
        <form className="flex items-center gap-1.5" onSubmit={applyRanges}>
          <Input
            aria-invalid={Boolean(error)}
            aria-label="Ranges to highlight"
            className="w-64 font-mono text-xs"
            onChange={(event) => {
              setDraft(event.target.value);
              setError(null);
            }}
            placeholder="A1:C4, Sheet2!B2:D8"
            size="sm"
            spellCheck={false}
            value={draft}
          />
          <RibbonButton disabled={!canHighlight} label="Highlight" showLabel type="submit">
            <Highlighter />
          </RibbonButton>
        </form>
        <RibbonButton
          disabled={highlightedRanges.length === 0}
          label="Clear"
          onClick={() => {
            clearHighlightedRanges();
            setError(null);
          }}
          showLabel
        >
          <Eraser />
        </RibbonButton>
      </RibbonGroup>
      <RibbonGroup className="gap-3 px-3">
        <Label className="gap-2 text-xs font-normal">
          <Switch checked={autoScroll} className="data-checked:bg-brand" onCheckedChange={setAutoScroll} />
          Scroll to range
        </Label>
        <Label className="gap-2 text-xs font-normal">
          <Switch checked={customStyle} className="data-checked:bg-brand" onCheckedChange={onCustomStyleChange} />
          Custom CSS
        </Label>
      </RibbonGroup>
      <RibbonGroup>
        <RibbonButton
          disabled={isLoading}
          label="Load demo workbook"
          onClick={() => {
            setDraft(HIGHLIGHT_DEMO_RANGES.join(", "));
            setError(null);
            onLoadSample(autoScroll);
          }}
          showLabel
        >
          <Sparkles />
        </RibbonButton>
      </RibbonGroup>
      <span aria-live="polite" className={error ? "px-2 text-xs text-destructive-foreground" : "px-2 text-xs text-muted-foreground"}>
        {error ?? (highlightedRanges.length > 0 ? `${highlightedRanges.length} ranges highlighted` : "")}
      </span>
    </>
  );
}
