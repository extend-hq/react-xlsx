import type { XlsxFreezePanes, XlsxSheetData } from "./types";

export interface ReadOnlySheetLayout {
  freezePanes?: XlsxFreezePanes | null;
  hiddenRows?: Set<number>;
  hiddenCols?: Set<number>;
}

export function changedVisibilityRanges(
  existing: Set<number>, ranges: { start: number; end: number }[], hidden: boolean, fromSelection: boolean
) {
  if (!hidden && !existing.size) return [];
  let firstVisible = 0;
  if (!hidden && fromSelection) {
    while (existing.has(firstVisible)) firstVisible++;
  }
  const changed: { start: number; end: number }[] = [];
  for (const range of ranges) {
    const start = firstVisible > 0 && range.start === firstVisible ? 0 : range.start;
    for (let index = start; index <= range.end; index++) {
      if (existing.has(index) !== hidden) {
        changed.push({ start, end: range.end });
        break;
      }
    }
  }
  return changed;
}

export function nextVisibleIndex(index: number, limit: number, isHidden: (index: number) => boolean) {
  let next = index;
  while (next < limit && isHidden(next)) next++;
  if (next < limit) return next;
  next = index - 1;
  while (next >= 0 && isHidden(next)) next--;
  return next;
}

function visibleAxis(
  hidden: Set<number>, max: number, originalIndices: number[], originalSizes: number[],
  overrides: Record<number, number>, defaultSize: number
) {
  const indices: number[] = [];
  const sizes: number[] = [];
  let offset = 0;
  for (let index = 0; index <= max; index++) {
    const originalSize = originalIndices[offset] === index ? originalSizes[offset++] : undefined;
    if (hidden.has(index)) continue;
    indices.push(index);
    sizes.push(originalSize ?? overrides[index] ?? defaultSize);
  }
  return { indices, sizes };
}

export function applyReadOnlyLayout(sheet: XlsxSheetData, layout?: ReadOnlySheetLayout): XlsxSheetData {
  if (!layout) return sheet;
  const view = { ...sheet };
  if (layout.freezePanes !== undefined) view.freezePanes = layout.freezePanes;
  if (layout.hiddenRows) {
    const rows = visibleAxis(layout.hiddenRows, sheet.maxUsedRow, sheet.visibleRows, sheet.rowHeights,
      sheet.rowHeightOverridesPx, sheet.defaultRowHeightPx);
    view.hiddenRows = Array.from(layout.hiddenRows).sort((a, b) => a - b);
    view.visibleRows = rows.indices;
    view.rowHeights = rows.sizes;
    view.rowCount = rows.indices.length;
  }
  if (layout.hiddenCols) {
    const cols = visibleAxis(layout.hiddenCols, sheet.maxUsedCol, sheet.visibleCols, sheet.colWidths,
      sheet.colWidthOverridesPx, sheet.defaultColWidthPx);
    view.hiddenCols = Array.from(layout.hiddenCols).sort((a, b) => a - b);
    view.visibleCols = cols.indices;
    view.colWidths = cols.sizes;
    view.colCount = cols.indices.length;
  }
  return view;
}
