import type { CellNumberFormat, Worksheet } from "../vendor/sheets-wasm/duke_sheets_wasm.js";

/** What a typed entry stores: a number, boolean or text, and the format it gives its cell. */
export type UserEntry = { value: unknown; numberFormat?: CellNumberFormat };

/**
 * The formats typing gives a cell (the engine's `coerce::typed_number`): dates, times and
 * percents. A cell showing one of these takes the format of whatever is typed next; any other
 * format is one the user chose.
 */
const TYPED_FORMAT_IDS = new Set([14, 15, 17, 18, 19, 20, 21, 22, 46]);
const TYPED_PERCENT_FORMAT = "0.##########%";

const formatKind = (format: CellNumberFormat) =>
  format.isDateFormat ? "date" : format.formatString.includes("%") ? "percent" : "number";

/**
 * Reads text typed into a cell as Excel does, in the workbook's regional settings: `5/11/2026` is
 * 5 November in en-GB and `1,5` is 1.5 in de-DE (the engine's `readTypedNumber`). A leading
 * apostrophe, or a Text (`@`) format on the cell, keeps the text as it is. A typed date, time or
 * percent gets its format unless the cell has one of the same kind that the user chose
 * (`dd mmm yyyy` stays for a typed date).
 */
export function readUserEntry(text: string, worksheet: Worksheet, current?: CellNumberFormat | null): UserEntry {
  if (text.startsWith("'")) return { value: text.slice(1) };
  const trimmed = text.trim();
  if (!trimmed || current?.id === 49 || current?.formatString === "@") return { value: text };
  if (/^(true|false)$/i.test(trimmed)) return { value: trimmed.toLowerCase() === "true" };
  const typed = worksheet.readTypedNumber(trimmed);
  if (!typed) return { value: text };
  const { value, numberFormat } = typed;
  if (!numberFormat || !current || current.formatType === "general") return { value, numberFormat };
  const chosen = !TYPED_FORMAT_IDS.has(current.id ?? -1) && current.formatString !== TYPED_PERCENT_FORMAT;
  return chosen && formatKind(current) === formatKind(numberFormat) ? { value } : { value, numberFormat };
}

/** Stores text typed or pasted into a cell (an A1 address) as Excel reads it. */
export function enterUserText(worksheet: Worksheet, address: string, text: string): void {
  const entry = readUserEntry(text, worksheet, worksheet.getCellStyle(address)?.numberFormat);
  worksheet.setCell(address, entry.value);
  if (entry.numberFormat) {
    worksheet.setCellStyle(address, { numberFormat: entry.numberFormat });
  }
}

/**
 * The cell's value as text to edit, which committed unchanged gives back the same value: text
 * that would read as something else (a number, a date, a formula) starts with an apostrophe.
 */
export function cellInputText(worksheet: Worksheet, row: number, col: number): string {
  const value = worksheet.getCellAt(row, col);
  if (value.is_empty) return "";
  const text = worksheet.getInputTextAt(row, col);
  if (!value.is_text) return text;
  const current = worksheet.getCellStyleAt(row, col)?.numberFormat;
  const readsAsText = typeof readUserEntry(text, worksheet, current).value === "string";
  return readsAsText && !text.startsWith("=") && !text.startsWith("'") ? text : `'${text}`;
}
