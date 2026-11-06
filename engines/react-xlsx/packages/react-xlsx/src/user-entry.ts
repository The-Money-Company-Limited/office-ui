import type { CellStyle, NumberFormatInput, Worksheet } from "../vendor/sheets-wasm/duke_sheets_wasm.js";
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
 * text as it is (`=A1` and TRUE included); otherwise an entry starting with `=` is a formula, and
 * anything else is read by readUserEntry.
 */
export function enterUserText(worksheet: Worksheet, cell: XlsxCellAddress, text: string): void {
  const address = cellAddressToA1(cell);
  const entry = text.trim();
  if (hasTextFormat(worksheet.getCellStyleAt(cell.row, cell.col))) {
    worksheet.setCell(address, text);
  } else if (entry.length > 1 && entry.startsWith("=")) {
    worksheet.setFormula(address, entry);
  } else {
    const { value, numberFormat } = readUserEntry(text, worksheet, cell);
    worksheet.setCell(address, value);
    if (numberFormat) {
      worksheet.setCellStyle(address, { numberFormat });
    }
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
