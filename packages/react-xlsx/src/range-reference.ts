import type { XlsxCellAddress, XlsxParsedRange } from "./types";

export const XLSX_MAX_ROWS = 1_048_576;
export const XLSX_MAX_COLUMNS = 16_384;

function parseColumn(value: string): number | null {
  let column = 0;
  for (const character of value.toUpperCase()) {
    column = column * 26 + character.charCodeAt(0) - 64;
  }
  return column > 0 && column <= XLSX_MAX_COLUMNS ? column - 1 : null;
}

function parseRow(value: string): number | null {
  const row = Number(value);
  return row > 0 && row <= XLSX_MAX_ROWS ? row - 1 : null;
}

function parseCell(value: string): XlsxCellAddress | null {
  const match = /^\$?([a-z]{1,3})\$?([1-9]\d{0,6})$/i.exec(value);
  if (!match) {
    return null;
  }
  const col = parseColumn(match[1]);
  const row = parseRow(match[2]);
  return col === null || row === null ? null : { col, row };
}

/** Parses one rectangular A1 reference into normalized, zero-based coordinates. */
export function parseXlsxRange(reference: string): XlsxParsedRange | null {
  if (typeof reference !== "string") {
    return null;
  }
  let address = reference.trim().replace(/^=\s*/, "");
  let sheetName: string | null = null;
  if (address.startsWith("'")) {
    const match = /^'((?:[^']|'')+)'!(.+)$/.exec(address);
    if (!match) {
      return null;
    }
    sheetName = match[1].replace(/''/g, "'");
    address = match[2];
  } else if (address.includes("!")) {
    const match = /^([^\s'!\[\]:*?/\\+,;(){}=<>"&^%\-]+)!(.+)$/.exec(address);
    if (!match) {
      return null;
    }
    sheetName = match[1];
    address = match[2];
  }
  if (sheetName !== null && (
    sheetName.length > 31 || /[\[\]:*?/\\\u0000-\u001f]/.test(sheetName)
    || sheetName.startsWith("'") || sheetName.endsWith("'")
  )) {
    return null;
  }

  const parts = address.split(":");
  if (parts.length > 2) {
    return null;
  }
  const start = parseCell(parts[0]);
  const end = parseCell(parts[1] ?? parts[0]);
  if (start && end) {
    return {
      kind: "cells",
      range: {
        start: { col: Math.min(start.col, end.col), row: Math.min(start.row, end.row) },
        end: { col: Math.max(start.col, end.col), row: Math.max(start.row, end.row) }
      },
      sheetName
    };
  }

  const columns = /^\$?([a-z]{1,3}):\$?([a-z]{1,3})$/i.exec(address);
  if (columns) {
    const first = parseColumn(columns[1]);
    const last = parseColumn(columns[2]);
    if (first !== null && last !== null) {
      return {
        kind: "columns",
        range: {
          start: { col: Math.min(first, last), row: 0 },
          end: { col: Math.max(first, last), row: XLSX_MAX_ROWS - 1 }
        },
        sheetName
      };
    }
  }

  const rows = /^\$?([1-9]\d{0,6}):\$?([1-9]\d{0,6})$/.exec(address);
  if (rows) {
    const first = parseRow(rows[1]);
    const last = parseRow(rows[2]);
    if (first !== null && last !== null) {
      return {
        kind: "rows",
        range: {
          start: { col: 0, row: Math.min(first, last) },
          end: { col: XLSX_MAX_COLUMNS - 1, row: Math.max(first, last) }
        },
        sheetName
      };
    }
  }
  return null;
}
