import * as React from "react";
import { Check, Sparkles, X } from "lucide-react";
import type { XlsxSelectionPopoverRenderProps, XlsxViewerController } from "@extend-ai/react-xlsx";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { FormulaDemoPopover, formulaActionClassName, getDemoFormula, FormulaPreviewText } from "./formula-demo-ui";
import { createHeatmap, formulaRangePreview, type DemoHeatmap, type DemoStylePreview } from "./formula-demo-actions";

function cellLabel(row: number, col: number) {
  let column = "";
  for (let index = col + 1; index > 0; index = Math.floor((index - 1) / 26)) {
    column = String.fromCharCode(65 + (index - 1) % 26) + column;
  }
  return `${column}${row + 1}`;
}

export function FormulaSelectionDemo({ selection, isDark, controller, onPreview, inline = false }: {
  selection: Pick<XlsxSelectionPopoverRenderProps, "context" | "signal" | "applyFormula" | "commit" | "close">;
  inline?: boolean;
  isDark: boolean;
  controller: XlsxViewerController;
  onPreview: (preview: DemoStylePreview | null) => void;
}) {
  const [instructions, setInstructions] = React.useState("");
  const [formula, setFormula] = React.useState<string | null>(null);
  const [heatmap, setHeatmap] = React.useState<DemoHeatmap | null>(null);
  const [error, setError] = React.useState("");
  React.useEffect(() => {
    const preview = heatmap?.preview ?? (formula ? formulaRangePreview(formula, selection.context.sheetName) : null);
    onPreview(preview);
    return () => onPreview(null);
  }, [formula, heatmap, onPreview, selection.context.sheetName]);
  React.useEffect(() => { setInstructions(""); setFormula(null); setHeatmap(null); setError(""); }, [selection.signal]);
  const { cell, selection: range, readOnly } = selection.context;
  if (readOnly) return null;
  const target = cellLabel(cell.row, cell.col);
  const rangeLabel = `${cellLabel(range.start.row, range.start.col)}:${cellLabel(range.end.row, range.end.col)}`;

  function generate() {
    if (!instructions.trim()) return;
    setError("");
    try {
      if (/heat\s*map|colou?r\s*scale/i.test(instructions)) setHeatmap(createHeatmap(selection.context, controller));
      else setFormula(getDemoFormula(cell, selection.context.sheetName));
    } catch (cause) { setError(cause instanceof Error ? cause.message : "Could not generate the preview."); }
  }

  const content = (
    <>
      {formula || heatmap ? (
        <>
          <div className="px-2 pt-1 pb-2 text-xs text-muted-foreground">{heatmap ? `Heatmap · ${heatmap.updates.length} numeric cells` : `Formula for ${target}`}</div>
          {heatmap ? (
            <div className="px-2 pb-2">
              <div className="h-2 rounded-full bg-linear-to-r from-blue-100 via-yellow-100 to-red-400" />
              <div className="mt-1 flex justify-between text-xs text-muted-foreground"><span>{heatmap.min.toFixed(1)}</span><span>{heatmap.max.toFixed(1)}</span></div>
              <p className="mt-2 text-xs text-muted-foreground">Preview on the sheet. Keep to save the colors.</p>
            </div>
          ) : <FormulaPreviewText formula={formula ?? ""} />}
          <Button variant="ghost" className={formulaActionClassName} onClick={() => heatmap ? selection.commit(() => controller.setCellStyles(heatmap.updates)) : formula && selection.applyFormula(formula)}>
            <Check className="size-4" />Keep
          </Button>
          <Button variant="ghost" className={formulaActionClassName} onClick={selection.close}>
            <X className="size-4" />Discard
          </Button>
        </>
      ) : (
        <form onSubmit={event => { event.preventDefault(); generate(); }}>
          <div className="px-2 pt-1 pb-2 text-xs text-muted-foreground">{range.start.row === range.end.row && range.start.col === range.end.col ? target : `${rangeLabel} · Apply to ${target}`}</div>
          <div className="px-1 pb-1">
            <Input
              aria-label="Formula instructions"
              placeholder="Try “heatmap” or “average”…"
              value={instructions}
              onChange={event => setInstructions(event.target.value)}
              onKeyDown={event => { if (event.key !== "Escape" && event.key !== "Tab") event.stopPropagation(); }}
              onKeyUp={event => { if (event.key !== "Escape" && event.key !== "Tab") event.stopPropagation(); }}
              className="h-8 text-sm"
            />
          </div>
          <Button type="submit" variant="ghost" className={formulaActionClassName} disabled={!instructions.trim()}>
            <Sparkles className="size-4" /><span>{/heat\s*map|colou?r\s*scale/i.test(instructions) ? "Generate heatmap" : "Generate formula"} <span className="text-muted-foreground">(Demo)</span></span>
          </Button>
          {error && <p role="alert" className="px-2 py-1 text-xs text-destructive-foreground">{error}</p>}
        </form>
      )}
    </>
  );
  return inline ? content : <FormulaDemoPopover open isDark={isDark} onEscape={selection.close} label="Selection formula actions">{content}</FormulaDemoPopover>;
}
