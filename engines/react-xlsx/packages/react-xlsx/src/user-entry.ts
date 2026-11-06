import type { Worksheet } from "../vendor/sheets-wasm/duke_sheets_wasm.js";
import type { XlsxCellNumberFormatInput } from "./types";

/** A number read from typed text, with the built-in format Excel gives a typed date or time. */
export type TypedNumber = { value: number; dateFormat?: number };

/** What a typed entry stores: a number, boolean or text, and the format it gives its cell. */
export type UserEntry = { value: unknown; numberFormat?: XlsxCellNumberFormatInput };

/**
 * Reads text typed into a cell as Excel does. `readNumber` reads numbers, dates and times in the
 * user's regional settings (the engine's `readTypedNumber`): `5/11/2026` is 5 November in en-GB
 * and `1,5` is 1.5 in de-DE. A leading apostrophe keeps the rest as text. A typed date or time
 * gets a date format unless the cell already has one (`dd mmm yyyy`), and a typed percent a
 * percent format.
 */
export function readUserEntry(
  text: string,
  readNumber: (text: string) => TypedNumber | undefined,
  cellHasDateFormat = false
): UserEntry {
  if (text.startsWith("'")) return { value: text.slice(1) };
  const trimmed = text.trim();
  if (!trimmed) return { value: text };
  if (/^(true|false)$/i.test(trimmed)) return { value: trimmed.toLowerCase() === "true" };
  // The engine reads $1,234 as a number too, but there's no currency format to give it here,
  // so it stays text.
  const number = trimmed.includes("$") ? undefined : readNumber(trimmed);
  if (!number) return { value: text };
  if (number.dateFormat !== undefined) {
    return cellHasDateFormat
      ? { value: number.value }
      : { value: number.value, numberFormat: { formatType: "builtin", id: number.dateFormat } };
  }
  return trimmed.endsWith("%")
    ? { value: number.value, numberFormat: { formatType: "custom", formatString: "0.##########%" } }
    : { value: number.value };
}

/** Stores text typed or pasted into a cell (an A1 address) as Excel reads it, in the workbook's
 * regional settings. */
export function enterUserText(worksheet: Worksheet, address: string, text: string): void {
  const style = worksheet.getCellStyle(address);
  const entry = readUserEntry(text, (typed) => worksheet.readTypedNumber(typed), Boolean(style?.numberFormat?.isDateFormat));
  worksheet.setCell(address, entry.value);
  if (entry.numberFormat) {
    worksheet.setCellStyle(address, { ...style, numberFormat: entry.numberFormat });
  }
}
