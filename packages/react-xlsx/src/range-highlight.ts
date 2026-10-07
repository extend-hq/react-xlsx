import type { XlsxCellRange } from "./types";

export function expandHighlightedRange(range: XlsxCellRange, mergedRanges: XlsxCellRange[]): XlsxCellRange {
  const expanded = { start: { ...range.start }, end: { ...range.end } };
  let changed: boolean;
  do {
    changed = false;
    for (const merged of mergedRanges) {
      if (merged.end.row < expanded.start.row || merged.start.row > expanded.end.row
        || merged.end.col < expanded.start.col || merged.start.col > expanded.end.col) {
        continue;
      }
      const startRow = Math.min(expanded.start.row, merged.start.row);
      const startCol = Math.min(expanded.start.col, merged.start.col);
      const endRow = Math.max(expanded.end.row, merged.end.row);
      const endCol = Math.max(expanded.end.col, merged.end.col);
      if (startRow !== expanded.start.row || startCol !== expanded.start.col
        || endRow !== expanded.end.row || endCol !== expanded.end.col) {
        expanded.start = { row: startRow, col: startCol };
        expanded.end = { row: endRow, col: endCol };
        changed = true;
      }
    }
  } while (changed);
  return expanded;
}

function lowerBound(indices: number[], value: number): number {
  let low = 0;
  let high = indices.length;
  while (low < high) {
    const middle = (low + high) >>> 1;
    if (indices[middle] < value) {
      low = middle + 1;
    } else {
      high = middle;
    }
  }
  return low;
}

export function getHighlightAxisExtent(indices: number[], prefixSums: number[], start: number, end: number) {
  const first = lowerBound(indices, start);
  const afterLast = lowerBound(indices, end + 1);
  if (first >= afterLast) {
    return null;
  }
  return { start: prefixSums[first] ?? 0, end: prefixSums[afterLast] ?? 0 };
}

export function getHighlightScrollOffset(
  start: number,
  end: number,
  frozenExtent: number,
  viewportExtent: number,
  currentOffset: number
): number {
  if (end <= frozenExtent || viewportExtent <= frozenExtent) {
    return currentOffset;
  }
  const scrollableStart = Math.max(start, frozenExtent);
  const availableExtent = viewportExtent - frozenExtent;
  if (scrollableStart < currentOffset + frozenExtent || end - scrollableStart > availableExtent) {
    return Math.max(0, scrollableStart - frozenExtent);
  }
  return end > currentOffset + viewportExtent ? end - viewportExtent : currentOffset;
}
