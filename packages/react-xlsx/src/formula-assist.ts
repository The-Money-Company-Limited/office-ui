import type { XlsxCellAddress, XlsxCellRange } from "./types";
import { parseXlsxRange } from "./range-reference.ts";
import { FORMULA_FUNCTIONS, type FormulaFunction } from "./formula-functions.ts";

const MAX_FUNCTION_MATCHES = 8;
const FUNCTIONS_BY_NAME = new Map(FORMULA_FUNCTIONS.map((fn) => [fn.name, fn]));

/**
 * What the formula editor offers at the caret, as in Google Sheets: function
 * names while a name is being typed, otherwise the signature of the function
 * whose arguments the caret is in.
 */
export type FormulaAssist =
  | {
      kind: "functions";
      /** The typed name the chosen function replaces: text[start, end). */
      start: number;
      end: number;
      matches: FormulaFunction[];
    }
  | {
      kind: "signature";
      fn: FormulaFunction;
      /** Index into the call's arguments; past a repeating tail it stays on the tail. */
      argumentIndex: number;
    };

export function resolveFormulaAssist(text: string, caret: number): FormulaAssist | null {
  if (!text.startsWith("=")) {
    return null;
  }
  const before = text.slice(0, caret);
  const calls: Array<{ name: string | null; argumentIndex: number }> = [];
  let inString = false;
  let nameStart = -1;
  for (let index = 1; index < before.length; index += 1) {
    const char = before[index];
    if (char === "\"") {
      inString = !inString;
      nameStart = -1;
      continue;
    }
    if (inString) {
      continue;
    }
    if (/[A-Za-z0-9._]/.test(char)) {
      if (nameStart < 0) {
        nameStart = index;
      }
      continue;
    }
    if (char === "(") {
      const name = nameStart >= 0 ? before.slice(nameStart, index).toUpperCase() : null;
      calls.push({ argumentIndex: 0, name });
    } else if (char === ")") {
      calls.pop();
    } else if (char === "," && calls.length > 0) {
      calls[calls.length - 1].argumentIndex += 1;
    }
    nameStart = -1;
  }
  if (inString) {
    return null;
  }

  // A name being typed right before the caret, and not already called.
  if (nameStart >= 0 && /^[A-Za-z]/.test(before[nameStart]) && text[caret] !== "(") {
    const query = before.slice(nameStart).toUpperCase();
    // Shortest names first, so SUM comes before SUBSTITUTE, as in Sheets.
    const matches = FORMULA_FUNCTIONS.filter((fn) => fn.name.startsWith(query))
      .sort((left, right) => left.name.length - right.name.length || left.name.localeCompare(right.name))
      .slice(0, MAX_FUNCTION_MATCHES);
    if (matches.length > 0 && !(matches.length === 1 && matches[0].name === query)) {
      return { end: caret, kind: "functions", matches, start: nameStart };
    }
  }

  for (let index = calls.length - 1; index >= 0; index -= 1) {
    const fn = calls[index].name ? FUNCTIONS_BY_NAME.get(calls[index].name as string) : undefined;
    if (fn) {
      return {
        argumentIndex: Math.min(calls[index].argumentIndex, Math.max(0, fn.arguments.length - 1)),
        fn,
        kind: "signature"
      };
    }
    if (calls[index].name) {
      return null;
    }
  }
  return null;
}

/** Replaces the typed name with the chosen function and an open bracket. */
export function applyFormulaFunction(
  text: string,
  assist: Extract<FormulaAssist, { kind: "functions" }>,
  fn: FormulaFunction
): { text: string; caret: number } {
  const inserted = `${fn.name}(`;
  return {
    caret: assist.start + inserted.length,
    text: text.slice(0, assist.start) + inserted + text.slice(assist.end)
  };
}

/** The signature as Sheets shows it: SUM(number1, [number2, ...]). */
export function formatFormulaArgument(argument: FormulaFunction["arguments"][number]) {
  const name = argument.repeating ? `${argument.name}, ...` : argument.name;
  return argument.optional ? `[${name}]` : name;
}

export interface FormulaPoint {
  text: string;
  start: number;
  end: number;
  anchor: XlsxCellAddress;
  focus: XlsxCellAddress;
}

export function canPointFormula(text: string, start: number, end: number, point: FormulaPoint | null): boolean {
  if (!text.startsWith("=") || start !== end) return false;
  if (point && point.text === text && start === point.end) return true;
  const before = text.slice(0, start);
  // Quoted sheet names and string literals may contain operators and commas.
  let quote = "";
  for (let index = 1; index < before.length; index += 1) {
    const char = before[index];
    if (quote) {
      if (char === quote) {
        if (before[index + 1] === quote) index += 1;
        else quote = "";
      }
    } else if (char === '"' || char === "'") quote = char;
  }
  return !quote && /[=(,+\-*/^&<>]\s*$/.test(before);
}

export function applyFormulaPoint(
  text: string,
  start: number,
  end: number,
  previous: FormulaPoint | null,
  cell: XlsxCellAddress,
  extend = false
): FormulaPoint | null {
  if (!canPointFormula(text, start, end, previous)) return null;
  const point = previous?.text === text && previous.end === start ? previous : null;
  const anchor = extend && point ? point.anchor : cell;
  const first = { row: Math.min(anchor.row, cell.row), col: Math.min(anchor.col, cell.col) };
  const last = { row: Math.max(anchor.row, cell.row), col: Math.max(anchor.col, cell.col) };
  const reference = cellAddressToA1(first) + (first.row === last.row && first.col === last.col ? "" : `:${cellAddressToA1(last)}`);
  const insertAt = point?.start ?? start;
  return {
    anchor,
    focus: cell,
    start: insertAt,
    end: insertAt + reference.length,
    text: text.slice(0, insertAt) + reference + text.slice(point?.end ?? end)
  };
}

const REFERENCE_COLORS = ["#4285f4", "#d96535", "#a855c7", "#168f76", "#bd8513", "#dc5293"];

export function formulaReferences(text: string) {
  const references: Array<{ start: number; end: number; text: string; color: string; range: XlsxCellRange; sheetName: string | null }> = [];
  if (!text.startsWith("=")) return references;
  const colors = new Map<string, string>();
  const tokens = /"(?:[^"]|"")*(?:"|$)|(?:'(?:[^']|'')*'!|[A-Za-z_][\w.]*!)?\$?[A-Za-z]{1,3}\$?[1-9]\d*(?::\$?[A-Za-z]{1,3}\$?[1-9]\d*)?/g;
  for (const match of text.matchAll(tokens)) {
    const start = match.index;
    const end = start + match[0].length;
    if (match[0].startsWith('"') || /[\w.!$\[\]]/.test(text[start - 1] ?? "") || /[\w.(\[\]]/.test(text[end] ?? "")) continue;
    const parsed = parseXlsxRange(match[0]);
    if (!parsed) continue;
    const key = match[0].replaceAll("$", "").toUpperCase();
    const color = colors.get(key) ?? REFERENCE_COLORS[colors.size % REFERENCE_COLORS.length];
    colors.set(key, color);
    references.push({ start, end, text: match[0], color, range: parsed.range, sheetName: parsed.sheetName });
  }
  return references;
}

function cellAddressToA1(cell: XlsxCellAddress): string {
  let column = cell.col + 1;
  let name = "";
  while (column > 0) {
    column -= 1;
    name = String.fromCharCode(65 + column % 26) + name;
    column = Math.floor(column / 26);
  }
  return `${name}${cell.row + 1}`;
}
