import { parseXlsxRange } from "./range-reference";
import type { XlsxParsedRange } from "./types";

const REFERENCE_COLORS = ["#3472c9", "#c53645", "#9b4dbb", "#258447", "#b97812", "#16858e"];
const CELL = String.raw`\$?[A-Z]{1,3}\$?[1-9]\d{0,6}`;
const COLUMN = String.raw`\$?[A-Z]{1,3}`;
const ROW = String.raw`\$?[1-9]\d{0,6}`;
const SHEET = String.raw`(?:'(?:[^']|'')+'|[\p{L}\p{N}_\\.]+)!`;
const REFERENCE = new RegExp(
  `(?:${SHEET})?(?:${CELL}(?:\\s*:\\s*${CELL})?|${COLUMN}\\s*:\\s*${COLUMN}|${ROW}\\s*:\\s*${ROW})`,
  "iyu"
);
const NAME_CHARACTER = /[\p{L}\p{N}_.$\\]/u;

export type FormulaReference = XlsxParsedRange & {
  start: number;
  end: number;
  color: string;
};

export function getFormulaReferences(value: string, sheetName: string): FormulaReference[] {
  if (!/^[=+-]/.test(value)) return [];
  const references: FormulaReference[] = [];
  const colors = new Map<string, string>();
  let index = value.startsWith("=") ? 1 : 0;
  while (index < value.length) {
    if (value[index] === '"') {
      index += 1;
      while (index < value.length) {
        if (value[index++] === '"') {
          if (value[index] !== '"') break;
          index += 1;
        }
      }
      continue;
    }
    if (value[index] === "[") {
      let depth = 1;
      while (++index < value.length && depth) {
        if (value[index] === "[") depth += 1;
        if (value[index] === "]") depth -= 1;
      }
      while (index < value.length && /[^\s+\-*/^&=<>(),;]/.test(value[index])) index += 1;
      continue;
    }
    if (NAME_CHARACTER.test(value[index - 1])) {
      index += 1;
      continue;
    }
    REFERENCE.lastIndex = index;
    const match = REFERENCE.exec(value);
    if (match) {
      const end = index + match[0].length;
      const previous = value[index - 1];
      const next = value[end];
      const parsed = parseXlsxRange(match[0].replace(/\s*:\s*/g, ":"));
      if (parsed && !NAME_CHARACTER.test(previous) && previous !== "]" && previous !== "!"
        && previous !== ":" && (!next || (!NAME_CHARACTER.test(next) && !/[!\[]/.test(next)))
        && !/^\s*\(/.test(value.slice(end))) {
        const { range } = parsed;
        const key = `${(parsed.sheetName ?? sheetName).toLowerCase()}:${range.start.row},${range.start.col}:${range.end.row},${range.end.col}`;
        let color = colors.get(key);
        if (!color) {
          color = REFERENCE_COLORS[colors.size % REFERENCE_COLORS.length];
          colors.set(key, color);
        }
        references.push({ ...parsed, start: index, end, color });
      }
      index = end;
    } else if (value[index] === "'") {
      index += 1;
      while (index < value.length) {
        if (value[index++] === "'") {
          if (value[index] !== "'") break;
          index += 1;
        }
      }
    } else {
      index += 1;
    }
  }
  return references;
}
