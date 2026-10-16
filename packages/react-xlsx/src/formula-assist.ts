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
