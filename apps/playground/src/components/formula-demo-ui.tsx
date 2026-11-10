import * as React from "react";
import type { XlsxCellAddress } from "@extend-ai/react-xlsx";
import { Popover, PopoverPopup } from "@/components/ui/popover";
import "./formula-editor-demo.css";

export const formulaActionClassName = "h-auto min-h-8 w-full cursor-default justify-start gap-2 rounded-sm border-0 px-2 py-1 text-base font-normal sm:h-auto sm:min-h-7 sm:text-sm";

export function getDemoFormula(cell: XlsxCellAddress, sheetName = "") {
  if (sheetName === "Scatter Matrix") {
    const ranges = ["B", "C", "D"].flatMap((column, index) => {
      if (cell.col !== index + 1 || cell.row < 6 || cell.row > 16) return [`${column}7:${column}17`];
      return [
        ...(cell.row > 6 ? [`${column}7:${column}${cell.row}`] : []),
        ...(cell.row < 16 ? [`${column}${cell.row + 2}:${column}17`] : []),
      ];
    });
    return `=ROUND(AVERAGE(${ranges.join(",")}),2)`;
  }
  const columns = ["A", "C", "E", "G"].filter((_, index) => index * 2 !== cell.col);
  return `=SUM(${columns[0]}1:${columns[0]}3,${columns[1]}5:${columns[1]}7,${columns[2]}9:${columns[2]}11)`;
}

export function FormulaDemoPopover({ open, isDark, onEscape, label, children }: {
  open: boolean;
  isDark: boolean;
  onEscape: () => void;
  label: string;
  children: React.ReactNode;
}) {
  const [host, setHost] = React.useState<HTMLDivElement | null>(null);
  return (
    <div ref={setHost} className={`playground-formula-overlay${isDark ? " dark" : ""}`} style={{ position: "absolute", inset: 0 }}>
      <Popover open={open && host !== null} modal={false} onOpenChange={(next, details) => {
        if (!next && details.reason === "escape-key") onEscape();
        else details.cancel();
      }}>
        <PopoverPopup
          anchor={host}
          positionMethod="fixed"
          portalProps={{ container: host }}
          align="start"
          side="bottom"
          sideOffset={8}
          initialFocus={false}
          finalFocus={false}
          aria-label={label}
          className="playground-editor-actions pointer-events-auto w-64 max-w-[calc(100vw-16px)] transition-[scale,opacity] duration-150 ease-out data-ending-style:scale-98 data-ending-style:opacity-0 motion-reduce:transition-none [&_[data-slot=popover-viewport]]:p-1"
        >{children}</PopoverPopup>
      </Popover>
    </div>
  );
}

export function FormulaPreviewText({ formula }: { formula: string }) {
  const colors = ["#3472c9", "#c53645", "#9b4dbb", "#258447"];
  const ranges = [...new Set(formula.match(/[A-Z]+\d+:[A-Z]+\d+/g) ?? [])];
  const parts = formula.split(/([A-Z]+\d+:[A-Z]+\d+)/g);
  return (
    <div className="playground-formula-shimmer break-all px-2 pb-2 font-mono text-xs" aria-label="Generated formula">
      {parts.map((part, index) => ranges.includes(part)
        ? <span key={index} data-formula-range={part} style={{ color: colors[ranges.indexOf(part) % colors.length] }}>{part}</span>
        : part)}
    </div>
  );
}
