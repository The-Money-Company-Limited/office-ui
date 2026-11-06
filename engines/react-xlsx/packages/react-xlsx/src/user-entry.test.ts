import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { before, describe, it } from "node:test";
import initSheetsWasm, { setLocale, Workbook, type Worksheet } from "../vendor/sheets-wasm/duke_sheets_wasm.js";
import { cellInputText, enterUserText, readUserEntry } from "./user-entry.ts";

/** A sheet in a new workbook with the regional settings of `tag`. */
function sheetIn(tag: string) {
  setLocale(tag);
  const sheet = new Workbook().getSheet(0);
  setLocale("en-US");
  return sheet;
}

const A1 = { row: 0, col: 0 };

/** The value a typed entry stores in an unformatted cell, and the built-in format it gives the cell. */
function read(tag: string, text: string) {
  const { value, numberFormat } = readUserEntry(text, sheetIn(tag), A1);
  return numberFormat ? [value, numberFormat.id] : [value];
}

/** Types `text` into A1 and gives back what the cell shows. */
function type(sheet: Worksheet, text: string) {
  enterUserText(sheet, A1, text, sheet.getCellStyleAt(0, 0));
  return sheet.getFormattedValueAt(0, 0);
}

describe("typed entries", () => {
  before(async () => {
    await initSheetsWasm({ module_or_path: readFileSync(new URL(import.meta.resolve("../vendor/sheets-wasm/duke_sheets_wasm_bg.wasm"))) });
  });

  it("percentages are fractions shown with as many decimals as Excel gives them", () => {
    assert.deepEqual(read("en-US", "12%"), [0.12, 9]);
    assert.deepEqual(read("en-US", "12.5%"), [0.125, 10]);
    assert.deepEqual(read("en-US", "(12%)"), [-0.12, 9]);
    assert.equal(type(sheetIn("en-US"), "12%"), "12%");
    assert.equal(type(sheetIn("en-US"), "12.5%"), "12.50%");
  });

  it("a plain number of 1 or more typed into a percent cell is a percent", () => {
    const sheet = sheetIn("en-US");
    type(sheet, "10%");
    assert.equal(type(sheet, "15"), "15%");
    assert.equal(type(sheet, "0.5"), "50%");
    assert.equal(cellInputText(sheet, A1), "50%");
  });

  it("grouped and scientific numbers are numeric; nonnumbers remain text", () => {
    for (const [input, expected] of [["1,234.5", 1234.5], ["-2e3", -2000], ["+.5", 0.5], ["(5)", -5], ["1,2", "1,2"], ["1e999", "1e999"], ["$1,234", "$1,234"], ["TRUE", true], ["false", false], ["  text  ", "  text  "]]) {
      assert.deepEqual(read("en-US", input as string), [expected]);
    }
  });

  it("lists and fractions stay text", () => {
    for (const input of ["3 1/4", "5 3/8", "1,2,3", "10, 20, 30", "5 10 15"]) {
      assert.deepEqual(read("en-US", input), [input]);
    }
  });

  it("an apostrophe or a Text format keeps the text", () => {
    for (const [input, expected] of [["'12.5%", "12.5%"], ["'TRUE", "TRUE"], ["'001 ", "001 "], [" '001 ", " '001 "], ["  ", "  "], ["'=SUM(A1:A3)", "=SUM(A1:A3)"], ["'5/11/2026", "5/11/2026"]]) {
      assert.deepEqual(read("en-US", input), [expected]);
    }
    const sheet = sheetIn("en-US");
    sheet.setCellStyle("A1", { numberFormat: { formatType: "builtin", id: 49 } });
    type(sheet, "5/11/2026");
    assert.equal(sheet.getCellAt(0, 0).asText(), "5/11/2026");
    type(sheet, "=A2+1");
    assert.ok(!sheet.getFormulaAt(0, 0));
    assert.equal(cellInputText(sheet, A1), "=A2+1");
  });

  it("numbers are read with the regional separators", () => {
    assert.deepEqual(read("de-DE", "1,5"), [1.5]);
    assert.deepEqual(read("de-DE", "1.234,5"), [1234.5]);
    assert.deepEqual(read("de-DE", "12,5%"), [0.125, 10]);
    assert.deepEqual(read("en-US", "1,5"), ["1,5"]);
  });

  it("dates and times are read in the regional date order and get Excel's format", () => {
    // 46331 is 5 November 2026; 46153 is 11 May 2026.
    assert.deepEqual(read("en-GB", "5/11/2026"), [46331, 14]);
    assert.deepEqual(read("en-US", "5/11/2026"), [46153, 14]);
    assert.deepEqual(read("cs-CZ", "5. 11. 2026"), [46331, 14]);
    assert.deepEqual(read("en-GB", "5/11/2026 13:30"), [46331.5625, 22]);
    assert.deepEqual(read("en-US", "15-Mar-2024"), [45366, 15]);
    assert.deepEqual(read("en-US", "May 2026"), [46143, 17]);
    assert.deepEqual(read("en-US", "13:30"), [0.5625, 20]);
    assert.deepEqual(read("en-US", "1:30 PM"), [0.5625, 18]);
    assert.deepEqual(read("en-US", "25:30"), [25.5 / 24, 46]);
  });

  it("a typed date shows in the user's order, and edits back to the same date", () => {
    const sheet = sheetIn("en-GB");
    assert.equal(type(sheet, "5/11/2026"), "05/11/2026");
    assert.equal(cellInputText(sheet, A1), "05/11/2026");
    type(sheet, cellInputText(sheet, A1));
    assert.equal(sheet.getCellAt(0, 0).asNumber(), 46331);
  });

  it("the cell keeps a format of the same kind and takes the entry's for another kind", () => {
    const sheet = sheetIn("en-GB");
    sheet.setCellStyle("A1", { numberFormat: { formatType: "custom", formatString: "d mmm yyyy" } });
    assert.equal(type(sheet, "5/11/2026"), "5 Nov 2026");
    assert.equal(type(sheet, "13:30"), "13:30");
    // The toolbar's Date (14) and Time (21) stay for entries of their kind.
    sheet.setCellStyle("A1", { numberFormat: { formatType: "builtin", id: 21 } });
    assert.equal(type(sheet, "9:30"), "9:30:00");
    sheet.setCellStyle("A1", { numberFormat: { formatType: "builtin", id: 14 } });
    assert.equal(type(sheet, "Mar 15, 2024"), "15/03/2024");
    assert.equal(type(sheet, "5/11/2026 9:00"), "05/11/2026 09:00");
    sheet.clearCellStyleAt(0, 0);
    assert.equal(sheet.getCellStyleAt(0, 0), null);
  });

  it("numbers show and edit with the regional separators", () => {
    const sheet = sheetIn("de-DE");
    assert.equal(type(sheet, "1,5"), "1,5");
    assert.equal(cellInputText(sheet, A1), "1,5");
    sheet.setCellStyle("A1", { numberFormat: { formatType: "custom", formatString: "#,##0.00" } });
    assert.equal(type(sheet, "1234,5"), "1.234,50");
  });

  it("text that would read as something else edits with an apostrophe", () => {
    const sheet = sheetIn("en-US");
    sheet.setCell("A1", "5/11/2026");
    sheet.setCell("A2", "plain");
    assert.equal(cellInputText(sheet, A1), "'5/11/2026");
    assert.equal(cellInputText(sheet, { row: 1, col: 0 }), "plain");
  });

  it("US English shows as before, except Excel's short date", () => {
    const sheet = sheetIn("en-US");
    assert.equal(type(sheet, "1234.5"), "1234.5");
    assert.equal(cellInputText(sheet, A1), "1234.5");
    sheet.clearCellStyleAt(0, 0);
    assert.equal(type(sheet, "11/5/2026"), "11/5/2026");
    sheet.clearCellStyleAt(0, 0);
    type(sheet, "1:30 PM");
    assert.equal(cellInputText(sheet, A1), "1:30:00 PM");
  });
});
