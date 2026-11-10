import type * as React from "react";
import { parseXlsxRange } from "@extend-ai/react-xlsx";
import type { XlsxCellRange, XlsxCellStyleInput, XlsxSelectionPopoverContext, XlsxViewerController } from "@extend-ai/react-xlsx";

export type DemoStylePreview = { sheetName: string; cells: Map<string, React.CSSProperties> };
export type DemoHeatmap = {
  kind: "heatmap";
  range: XlsxCellRange;
  updates: Array<{ cell: { row: number; col: number }; style: XlsxCellStyleInput }>;
  preview: DemoStylePreview;
  min: number;
  max: number;
};

export function createHeatmap(context: XlsxSelectionPopoverContext, controller: XlsxViewerController): DemoHeatmap {
  const selected = context.selection;
  const single = selected.start.row === selected.end.row && selected.start.col === selected.end.col;
  const range = single && context.sheetName === "Scatter Matrix"
    ? { start: { row: 6, col: 1 }, end: { row: 16, col: 3 } }
    : selected;
  if ((range.end.row - range.start.row + 1) * (range.end.col - range.start.col + 1) > 10000) {
    throw new Error("Select up to 10,000 cells for the demo heatmap.");
  }
  const values: Array<{ row: number; col: number; value: number }> = [];
  for (let row = range.start.row; row <= range.end.row; row++) {
    for (let col = range.start.col; col <= range.end.col; col++) {
      const text = controller.getCellDisplayValue({ row, col }).trim();
      const value = Number(text.replaceAll(",", ""));
      if (text && Number.isFinite(value)) values.push({ row, col, value });
    }
  }
  if (!values.length) throw new Error("Select numeric cells to create a heatmap.");
  const min = Math.min(...values.map(cell => cell.value));
  const max = Math.max(...values.map(cell => cell.value));
  const cells = new Map<string, React.CSSProperties>();
  const updates = values.map(({ row, col, value }) => {
    const t = max === min ? 0.5 : (value - min) / (max - min);
    const from = t < 0.5 ? [219, 234, 254] : [254, 249, 195];
    const to = t < 0.5 ? [254, 249, 195] : [248, 113, 113];
    const amount = t < 0.5 ? t * 2 : (t - 0.5) * 2;
    const hex = from.map((channel, index) => Math.round(channel + (to[index] - channel) * amount).toString(16).padStart(2, "0")).join("");
    cells.set(`${row}:${col}`, { backgroundColor: `#${hex}`, color: "#172554" });
    return {
      cell: { row, col },
      style: {
        fill: { fillType: "solid", color: { colorType: "rgb", hex } },
        font: { color: { colorType: "rgb", hex: "172554" } },
      } satisfies XlsxCellStyleInput,
    };
  });
  return { kind: "heatmap", range, updates, preview: { sheetName: context.sheetName, cells }, min, max };
}

export function formulaRangePreview(formula: string, sheetName: string): DemoStylePreview {
  const colors = ["#3472c9", "#c53645", "#9b4dbb", "#258447"];
  const cells = new Map<string, React.CSSProperties>();
  const ranges = [...new Set(formula.match(/[A-Z]+\d+:[A-Z]+\d+/g) ?? [])];
  ranges.forEach((reference, index) => {
    const range = parseXlsxRange(reference)?.range;
    if (!range) return;
    const color = colors[index % colors.length];
    for (let row = range.start.row; row <= range.end.row; row++) {
      for (let col = range.start.col; col <= range.end.col; col++) {
        cells.set(`${row}:${col}`, { backgroundColor: `${color}20`, boxShadow: `inset 0 0 0 1px ${color}` });
      }
    }
  });
  return { sheetName, cells };
}
