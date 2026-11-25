import type { XlsxCellAddress, XlsxCellRange } from "./types";

export function resolveFreezePanesCell(activeCell: XlsxCellAddress | null, selection: XlsxCellRange | null) {
  // A header selection can begin inside a merged cell whose anchor lies outside the selection.
  if (activeCell && selection && (activeCell.row < selection.start.row || activeCell.col < selection.start.col)) {
    return selection.start;
  }
  return activeCell;
}

export function subtractRange(area: XlsxCellRange, cut: XlsxCellRange): XlsxCellRange[] {
  const top = Math.max(area.start.row, cut.start.row);
  const bottom = Math.min(area.end.row, cut.end.row);
  const left = Math.max(area.start.col, cut.start.col);
  const right = Math.min(area.end.col, cut.end.col);
  if (top > bottom || left > right) return [area];

  const remaining: XlsxCellRange[] = [];
  if (area.start.row < top) remaining.push({ start: area.start, end: { row: top - 1, col: area.end.col } });
  if (bottom < area.end.row) remaining.push({ start: { row: bottom + 1, col: area.start.col }, end: area.end });
  if (area.start.col < left) remaining.push({ start: { row: top, col: area.start.col }, end: { row: bottom, col: left - 1 } });
  if (right < area.end.col) remaining.push({ start: { row: top, col: right + 1 }, end: { row: bottom, col: area.end.col } });
  return remaining;
}

export function addSelectionRange(areas: XlsxCellRange[], range: XlsxCellRange): XlsxCellRange[] {
  return [...areas.flatMap(area => subtractRange(area, range)), range];
}

export function selectionCoversRange(areas: readonly XlsxCellRange[], range: XlsxCellRange) {
  let remaining = [range];
  for (const area of areas) {
    remaining = remaining.flatMap(piece => subtractRange(piece, area));
    if (!remaining.length) return true;
  }
  return false;
}

export function clipboardBands(areas: readonly XlsxCellRange[]) {
  const first = areas[0];
  if (!first) return null;
  const sameExtent = (axis: "row" | "col") => areas.every(area =>
    area.start[axis] === first.start[axis] && area.end[axis] === first.end[axis]);
  const bands = (axis: "row" | "col", ranges: readonly XlsxCellRange[]) => ranges
    .map(range => ({ start: range.start[axis], end: range.end[axis] }))
    .sort((a, b) => a.start - b.start);
  if (sameExtent("col")) return { rows: bands("row", areas), cols: bands("col", [first]) };
  if (sameExtent("row")) return { rows: bands("row", [first]), cols: bands("col", areas) };
  return null;
}

export function* bandIndices(bands: readonly { start: number; end: number }[]) {
  let offset = 0;
  for (const band of bands) {
    for (let index = band.start; index <= band.end; index++) {
      yield { index, offset: offset++, remaining: band.end - index + 1 };
    }
  }
}
