import type { CellStyle, NumberFormatInput, StyleInput, Worksheet } from "../vendor/sheets-wasm/duke_sheets_wasm.js";
import { cellAddressToA1 } from "./range-reference.ts";
import type { XlsxCellAddress } from "./types";

/** What a typed entry stores: a number, boolean or text, and the format it gives its cell. */
export type UserEntry = { value: unknown; numberFormat?: NumberFormatInput };

/** Whether the cell has a text format (`@`, `"ID: "@`), which keeps what is typed as text. */
const hasTextFormat = (style: CellStyle | null) => style?.numberFormat.kind === "text";

/**
 * Reads text typed into a cell as Excel does. A leading apostrophe keeps the rest as text, and
 * TRUE and FALSE are booleans. The engine reads numbers, percents, dates and times in the
 * workbook's regional settings (`5/11/2026` is 5 November in en-GB, `1,5` is 1.5 in de-DE) and
 * decides whether the cell's format changes: it stays for an entry of the same kind.
 */
export function readUserEntry(text: string, worksheet: Worksheet, cell: XlsxCellAddress): UserEntry {
  if (text.startsWith("'")) return { value: text.slice(1) };
  const trimmed = text.trim();
  if (!trimmed) return { value: text };
  if (/^(true|false)$/i.test(trimmed)) return { value: trimmed.toLowerCase() === "true" };
  return worksheet.readTypedNumberAt(cell.row, cell.col, trimmed) ?? { value: text };
}

/**
 * Stores text typed or pasted into a cell, as Excel does. A cell with a text format keeps the
 * text as it is (`=A1` and TRUE included, a leading apostrophe dropped); otherwise an entry
 * starting with `=` is a formula, and anything else is read by readUserEntry. A format the
 * entry gives the cell goes on top of `baseStyle`, the row or column style a cell without a
 * style of its own shows.
 */
export function enterUserText(worksheet: Worksheet, cell: XlsxCellAddress, text: string, baseStyle: StyleInput = {}): void {
  const address = cellAddressToA1(cell);
  const entry = text.trim();
  if (hasTextFormat(worksheet.getCellStyleAt(cell.row, cell.col))) {
    worksheet.setCell(address, text.startsWith("'") ? text.slice(1) : text);
  } else if (entry.length > 1 && entry.startsWith("=")) {
    worksheet.setFormula(address, entry);
  } else {
    const { value, numberFormat } = readUserEntry(text, worksheet, cell);
    worksheet.setCell(address, value);
    if (numberFormat) {
      worksheet.setCellStyle(address, { ...baseStyle, numberFormat });
    }
  }
}

/**
 * The 1900-system serials a date written as three numbers (`13/11/2026`, `5.11.26`) stands for,
 * in each order of day, month and year it is valid in.
 */
function numericDateSerials(text: string): number[] {
  const match = /^(\d{1,4})([./-])(\d{1,2})\2(\d{1,4})\.?$/.exec(text.trim());
  if (!match) return [];
  const [a, b, c] = [match[1], match[3], match[4]].map(Number);
  // Two-digit years are 1930-2029, as Excel reads them.
  const fullYear = (year: number, digits: number) => (digits > 2 ? year : year < 30 ? 2000 + year : 1900 + year);
  const orders = [
    [fullYear(c, match[4].length), a, b],
    [fullYear(c, match[4].length), b, a],
    [fullYear(a, match[1].length), b, c]
  ];
  return orders.flatMap(([year, month, day]) => {
    const time = Date.UTC(year, month - 1, day);
    const date = new Date(time);
    const valid = date.getUTCFullYear() === year && date.getUTCMonth() === month - 1 && date.getUTCDate() === day;
    return valid ? [(time - Date.UTC(1899, 11, 30)) / 86_400_000] : [];
  });
}

/**
 * Stores a cell pasted from desktop Excel, whose HTML carried its value (`number`, see
 * excelClipboardNumbers) beside the text it shows. The value wins over the text, which Excel
 * wrote in its computer's regional settings. The text gives the cell its format as typing it
 * would (a currency, a percent, a date) when it reads as that value, and the short date when it
 * is that date written in another order of day and month than the editor's. A Text-formatted
 * cell keeps the text.
 */
export function enterPastedNumber(worksheet: Worksheet, cell: XlsxCellAddress, text: string, number: number, baseStyle: StyleInput = {}): void {
  if (hasTextFormat(worksheet.getCellStyleAt(cell.row, cell.col))) {
    enterUserText(worksheet, cell, text, baseStyle);
    return;
  }
  const address = cellAddressToA1(cell);
  const entry = readUserEntry(text, worksheet, cell);
  const readsAsNumber = typeof entry.value === "number" && Math.abs(entry.value - number) <= 1e-9 * Math.max(1, Math.abs(number));
  const numberFormat: NumberFormatInput | undefined = readsAsNumber
    ? entry.numberFormat
    : numericDateSerials(text).includes(number) ? { formatType: "builtin", id: 14 } : undefined;
  worksheet.setCell(address, number);
  if (numberFormat) {
    worksheet.setCellStyle(address, { ...baseStyle, numberFormat });
  }
}

/**
 * The cell's value as text to edit, which committed unchanged gives back the same value: text
 * that would read as something else (a number, a date, a formula) starts with an apostrophe,
 * except in a Text-formatted cell, which keeps whatever is typed.
 */
export function cellInputText(worksheet: Worksheet, cell: XlsxCellAddress): string {
  const value = worksheet.getCellAt(cell.row, cell.col);
  if (value.is_empty) return "";
  const text = worksheet.getInputTextAt(cell.row, cell.col);
  if (!value.is_text || hasTextFormat(worksheet.getCellStyleAt(cell.row, cell.col))) return text;
  const readsAsText = !text.startsWith("=") && !text.startsWith("'") && typeof readUserEntry(text, worksheet, cell).value === "string";
  return readsAsText ? text : `'${text}`;
}
