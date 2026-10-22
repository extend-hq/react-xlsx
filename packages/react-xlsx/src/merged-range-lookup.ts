import type { XlsxCellAddress, XlsxCellRange } from "./types";

export type MergedRangeLookup = (cell: XlsxCellAddress) => XlsxCellRange | null;

type RangeNode = {
  range: XlsxCellRange;
  bounds: XlsxCellRange;
  left: RangeNode | null;
  right: RangeNode | null;
};

/**
 * Builds a lookup that returns the merged range containing a cell, or null.
 * Subtree bounds prune the search without expanding ranges into rows or cells.
 * Storage grows with the number of merges, independently of their dimensions.
 */
export function createMergedRangeLookup(ranges: readonly XlsxCellRange[]): MergedRangeLookup {
  const normalizedRanges: XlsxCellRange[] = [];

  for (const range of ranges) {
    const startRow = Math.min(range.start.row, range.end.row);
    const endRow = Math.max(range.start.row, range.end.row);
    const startCol = Math.min(range.start.col, range.end.col);
    const endCol = Math.max(range.start.col, range.end.col);
    if (![startRow, endRow, startCol, endCol].every((value) => Number.isInteger(value) && value >= 0)
      || (startRow === endRow && startCol === endCol)) {
      continue;
    }

    normalizedRanges.push({
      end: { col: endCol, row: endRow },
      start: { col: startCol, row: startRow }
    });
  }

  normalizedRanges.sort((a, b) => a.start.row - b.start.row || a.start.col - b.start.col);

  const buildNode = (start: number, end: number): RangeNode | null => {
    if (start >= end) return null;
    const middle = (start + end) >>> 1;
    const range = normalizedRanges[middle];
    const left = buildNode(start, middle);
    const right = buildNode(middle + 1, end);
    return {
      range,
      left,
      right,
      bounds: {
        start: {
          row: normalizedRanges[start].start.row,
          col: Math.min(range.start.col, left?.bounds.start.col ?? Infinity, right?.bounds.start.col ?? Infinity)
        },
        end: {
          row: Math.max(range.end.row, left?.bounds.end.row ?? -1, right?.bounds.end.row ?? -1),
          col: Math.max(range.end.col, left?.bounds.end.col ?? -1, right?.bounds.end.col ?? -1)
        }
      }
    };
  };
  const root = buildNode(0, normalizedRanges.length);
  const contains = (range: XlsxCellRange, cell: XlsxCellAddress) => (
    cell.row >= range.start.row && cell.row <= range.end.row
    && cell.col >= range.start.col && cell.col <= range.end.col
  );
  const find = (node: RangeNode | null, cell: XlsxCellAddress): XlsxCellRange | null => {
    if (!node || !contains(node.bounds, cell)) return null;
    if (contains(node.range, cell)) {
      return node.range;
    }
    return find(node.left, cell) ?? find(node.right, cell);
  };
  return (cell) => find(root, cell);
}
