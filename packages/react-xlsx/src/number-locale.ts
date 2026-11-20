import type { Worksheet } from "./duke-sheets/generated/duke_sheets_wasm.js";

const separators = new Map<string, { decimal: string; group: string }>();
type DisplayNumberFormat = { isDateFormat?: boolean; formatString?: string };

export function parseLocaleNumber(text: string, locale: string): number | undefined {
  if (locale === "en-US") return undefined;
  const parts = new Intl.NumberFormat(locale).formatToParts(12345.6);
  const decimal = parts.find((part) => part.type === "decimal")?.value ?? ".";
  const group = parts.find((part) => part.type === "group")?.value ?? ",";
  let normalized = text.trim();
  const percent = normalized.endsWith("%");
  if (percent) normalized = normalized.slice(0, -1).trim();
  normalized = group.trim() === "" ? normalized.replace(/[ \u00a0\u202f]/g, "") : normalized.split(group).join("");
  normalized = normalized.replace(decimal, ".");
  if (!/^[+-]?(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?$/.test(normalized)) return undefined;
  const value = Number(normalized) / (percent ? 100 : 1);
  return Number.isFinite(value) ? value : undefined;
}

export function localizeCellNumber(worksheet: Worksheet, row: number, col: number, text: string, locale = "en-US"): string {
  if (locale === "en-US" || !text) return text;
  const value = worksheet.getCalculatedValueAt(row, col);
  try {
    if (!value.is_number) return text;
  } finally {
    value.free();
  }
  const style = worksheet.getCellStyleAt(row, col) as {
    numberFormat?: DisplayNumberFormat;
  } | null;
  return localizeFormattedNumber(text, style?.numberFormat, locale);
}

export function localizeFormattedNumber(text: string, format: DisplayNumberFormat | undefined, locale: string) {
  if (locale === "en-US" || !text || format?.isDateFormat) return text;

  let symbols = separators.get(locale);
  if (!symbols) {
    const parts = new Intl.NumberFormat(locale).formatToParts(12345.6);
    symbols = {
      decimal: parts.find((part) => part.type === "decimal")?.value ?? ".",
      group: parts.find((part) => part.type === "group")?.value ?? ","
    };
    if (separators.size >= 100) separators.clear();
    separators.set(locale, symbols);
  }
  if (symbols.decimal === "." && symbols.group === ",") return text;

  const protectedPositions = new Set<number>();
  const code = format?.formatString ?? "";
  const literalPattern = /"[^"]*"|\\.|\[[^\]]*\]/g;
  const firstPlaceholder = code.replace(literalPattern, (token) => " ".repeat(token.length)).search(/[0#?]/);
  const literals = [...code.matchAll(literalPattern)];
  let suffixEnd = text.length;
  let prefixStart = 0;
  const protect = (token: string, suffix: boolean) => {
    const literal = token.startsWith('"') ? token.slice(1, -1)
      : token.startsWith("\\") ? token.slice(1)
      : token.startsWith("[$") ? token.slice(2, -1).split("-")[0] : "";
    if (!literal) return;
    const offset = suffix ? text.lastIndexOf(literal, suffixEnd - literal.length) : text.indexOf(literal, prefixStart);
    if (offset >= 0) {
      for (let i = offset; i < offset + literal.length; i++) protectedPositions.add(i);
      if (suffix) suffixEnd = offset;
      else prefixStart = offset + literal.length;
    }
  };
  for (const token of literals.filter((token) => token.index < firstPlaceholder)) protect(token[0], false);
  for (const token of literals.filter((token) => token.index >= firstPlaceholder).reverse()) protect(token[0], true);
  return text.replace(/\d[\d,]*(?:\.\d*)?|\.\d+/g, (number, offset: number) =>
    number.replace(/[.,]/g, (separator, index: number) => protectedPositions.has(offset + index)
      ? separator
      : separator === "." ? symbols.decimal : symbols.group)
  );
}

export function localizeRows(rows: unknown, locale = "en-US"): void {
  if (locale === "en-US" || !Array.isArray(rows)) return;
  for (const row of rows) {
    for (const cell of row.cells ?? []) {
      if (cell.isMergedSecondary || typeof cell.value !== "string") continue;
      if (typeof cell.numericValue === "number") {
        cell.value = localizeFormattedNumber(cell.value, cell.style?.numberFormat, locale);
      } else {
        cell.numericValue = null;
      }
    }
  }
}
