import type { XlsxCellAddress, XlsxCellRange } from "./types";

export type MergedRangeLookup = (cell: XlsxCellAddress) => XlsxCellRange | null;

/**
 * Builds a lookup that returns the merged range containing a cell, or null.
 * Ranges are bucketed per row, so memory grows with the rows a merge spans
 * rather than with its area, and a lookup only scans the merges on that row.
 */
export function createMergedRangeLookup(ranges: readonly XlsxCellRange[]): MergedRangeLookup {
  const rangesByRow = new Map<number, XlsxCellRange[]>();

  for (const range of ranges) {
    const startRow = Math.min(range.start.row, range.end.row);
    const endRow = Math.max(range.start.row, range.end.row);
    const startCol = Math.min(range.start.col, range.end.col);
    const endCol = Math.max(range.start.col, range.end.col);
    if (startRow === endRow && startCol === endCol) {
      continue;
    }

    const normalized: XlsxCellRange = {
      end: { col: endCol, row: endRow },
      start: { col: startCol, row: startRow }
    };
    for (let row = startRow; row <= endRow; row += 1) {
      const rowRanges = rangesByRow.get(row);
      if (rowRanges) {
        rowRanges.push(normalized);
      } else {
        rangesByRow.set(row, [normalized]);
      }
    }
  }

  return (cell) => {
    for (const range of rangesByRow.get(cell.row) ?? []) {
      if (cell.col >= range.start.col && cell.col <= range.end.col) {
        return range;
      }
    }
    return null;
  };
}
