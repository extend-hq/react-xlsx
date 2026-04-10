import type { Workbook } from "@dukelib/sheets-wasm";

const CONTENT_ANCHOR_BATCH_ROW_COUNT = 256;

type WorksheetWithRowsBatch = ReturnType<Workbook["getSheet"]> & {
  getRowsBatch?: (startRow: number, maxRows: number, options?: unknown) => unknown;
};

function asRecord(value: unknown): Record<string, unknown> | null {
  return typeof value === "object" && value !== null ? value as Record<string, unknown> : null;
}

function asNonNegativeInteger(value: unknown) {
  return typeof value === "number" && Number.isInteger(value) && value >= 0 ? value : null;
}

function getBatchRowEntries(batch: unknown) {
  if (Array.isArray(batch)) {
    return batch;
  }

  const record = asRecord(batch);
  return Array.isArray(record?.rows) ? record.rows : [];
}

function resolveRowContentStartCol(rowEntry: unknown) {
  const rowRecord = asRecord(rowEntry);
  if (!rowRecord || !Array.isArray(rowRecord.cells)) {
    return null;
  }

  let firstContentCol: number | null = null;
  for (const cellEntry of rowRecord.cells) {
    const cellRecord = asRecord(cellEntry);
    const col = asNonNegativeInteger(cellRecord?.col);
    if (col === null) {
      continue;
    }

    const value = cellRecord?.value;
    const hasValue = typeof value === "string" ? value.length > 0 : value !== null && value !== undefined;
    const hasFormula = typeof cellRecord?.formula === "string" && cellRecord.formula.length > 0;
    const hasHyperlink = typeof cellRecord?.hyperlink === "string" && cellRecord.hyperlink.length > 0;
    const hasImage = cellRecord?.image !== null && cellRecord?.image !== undefined;
    const hasMergeSpan = cellRecord?.mergeSpan !== null && cellRecord?.mergeSpan !== undefined;

    if (!hasValue && !hasFormula && !hasHyperlink && !hasImage && !hasMergeSpan) {
      continue;
    }

    firstContentCol = firstContentCol === null ? col : Math.min(firstContentCol, col);
  }

  return firstContentCol;
}

export function resolveUsedRangeContentAnchor(
  worksheet: ReturnType<Workbook["getSheet"]>,
  usedRange: [number, number, number, number]
) {
  const [minUsedRow, minUsedCol, maxUsedRow] = usedRange;
  const worksheetWithRowsBatch = worksheet as WorksheetWithRowsBatch;

  if (typeof worksheetWithRowsBatch.getRowsBatch !== "function" || (minUsedRow <= 0 && minUsedCol <= 0)) {
    return {
      minUsedCol,
      minUsedRow
    };
  }

  for (let startRow = 0; startRow <= maxUsedRow; startRow += CONTENT_ANCHOR_BATCH_ROW_COUNT) {
    let rowsBatch: unknown;
    try {
      rowsBatch = worksheetWithRowsBatch.getRowsBatch(
        startRow,
        Math.min(CONTENT_ANCHOR_BATCH_ROW_COUNT, maxUsedRow - startRow + 1),
        {
          includeFormulas: true,
          includeHyperlinks: true,
          includeImages: true,
          includeMergeInfo: true,
          useFormattedValues: true
        }
      );
    } catch {
      break;
    }

    for (const rowEntry of getBatchRowEntries(rowsBatch)) {
      const rowRecord = asRecord(rowEntry);
      const row = asNonNegativeInteger(rowRecord?.index);
      const firstContentCol = resolveRowContentStartCol(rowEntry);
      if (row === null || firstContentCol === null) {
        continue;
      }

      return {
        minUsedCol: Math.min(minUsedCol, firstContentCol),
        minUsedRow: Math.min(minUsedRow, row)
      };
    }
  }

  return {
    minUsedCol,
    minUsedRow
  };
}
