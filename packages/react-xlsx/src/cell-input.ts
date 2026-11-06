import type { Worksheet } from "@dukelib/sheets-wasm";
import { parseLocaleNumber } from "./number-locale";
import { getFormulaReferences } from "./formula-references";
import type { XlsxCellAddress } from "./types";

export function isTextCell(worksheet: Worksheet, cell: XlsxCellAddress): boolean {
  const style = worksheet.getCellStyleAt(cell.row, cell.col);
  return style?.numberFormat?.formatString === "@" || style?.numberFormat?.id === 49;
}

export function coerceUserEnteredValue(value: string, locale = "en-US", source: "edit" | "paste" = "edit"): unknown {
  if (source === "edit" && value.startsWith("'")) return value.slice(1);
  const trimmed = value.trim();
  if (!trimmed) return value;
  if (/^(true|false)$/i.test(trimmed)) return trimmed.toLowerCase() === "true";
  const localized = parseLocaleNumber(trimmed, locale);
  if (localized !== undefined) return localized;
  if (/^[+-]?(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?$/.test(trimmed)) {
    const number = Number(trimmed);
    if (Number.isFinite(number)) return number;
  }
  return value;
}

export function getUserEnteredFormula(value: string, locale = "en-US"): string | null {
  if (value.startsWith("=") && value.length > 1) return value;
  if (/^[+-].+/.test(value) && typeof coerceUserEnteredValue(value, locale) !== "number") return `=${value}`;
  return null;
}

export function applyInputNumberFormat(worksheet: Worksheet, cell: XlsxCellAddress, input: string): void {
  const style = worksheet.getCellStyleAt(cell.row, cell.col);
  const format = style?.numberFormat;
  if (format?.formatString && format.formatString.toLowerCase() !== "general") return;
  const formatString = /^[+-]?(?:\d+\.?\d*|\.\d+)[eE][+-]?\d+$/.test(input.trim()) ? "0.00E+00"
    : /^=\s*(?:DATE|TODAY)\s*\(/i.test(input) ? "m/d/yy" : undefined;
  if (formatString) worksheet.setCellStyleAt(cell.row, cell.col, { ...style, numberFormat: { formatString } });
}

export function normalizeFormulaReferences(formula: string): string {
  for (const reference of getFormulaReferences(formula, "").reverse()) {
    const token = formula.slice(reference.start, reference.end);
    const separator = token.lastIndexOf("!") + 1;
    formula = formula.slice(0, reference.start) + token.slice(0, separator)
      + token.slice(separator).toUpperCase().replace(/\s*:\s*/g, ":") + formula.slice(reference.end);
  }
  return formula;
}

export function getCellEditingValue(worksheet: Worksheet, cell: XlsxCellAddress, locale: string): string {
  const value = worksheet.getCellAt(cell.row, cell.col);
  try {
    if (value.is_empty) return "";
    if (value.is_number) {
      const decimal = new Intl.NumberFormat(locale).formatToParts(1.1).find(part => part.type === "decimal")?.value ?? ".";
      return String(value.asNumber()).replace(".", decimal);
    }
    const text = value.toString();
    if (value.is_text && (text.startsWith("'") || (!isTextCell(worksheet, cell)
      && (getUserEnteredFormula(text, locale) || coerceUserEnteredValue(text, locale) !== text)))) return `'${text}`;
    return text;
  } finally {
    value.free();
  }
}

export function parseClipboardText(text: string): string[][] {
  const normalized = text.replace(/\r\n?/g, "\n");
  const rows: string[][] = [];
  let row: string[] = [];
  let offset = 0;
  while (offset < normalized.length) {
    let value = "";
    if (normalized[offset] === '"') {
      const quoted = /^"((?:[^"]|"")*)"(?=\t|\n|$)/.exec(normalized.slice(offset));
      if (quoted) {
        value = quoted[1].replace(/""/g, '"');
        offset += quoted[0].length;
      }
    }
    while (offset < normalized.length && normalized[offset] !== "\t" && normalized[offset] !== "\n") value += normalized[offset++];
    row.push(value);
    const delimiter = normalized[offset++];
    if (delimiter === "\n" || delimiter === undefined) {
      rows.push(row);
      row = [];
    } else if (offset === normalized.length) row.push("");
  }
  if (row.length) rows.push(row);
  return rows;
}

export function quoteClipboardField(value: string): string {
  return /[\t\r\n"]/.test(value) ? `"${value.replace(/"/g, '""')}"` : value;
}
