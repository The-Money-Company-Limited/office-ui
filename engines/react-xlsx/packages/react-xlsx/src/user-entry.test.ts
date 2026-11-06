import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { before, describe, it } from "node:test";
import initSheetsWasm, { setLocale, Workbook } from "../vendor/sheets-wasm/duke_sheets_wasm.js";
import { enterUserText, readUserEntry } from "./user-entry.ts";

/** A sheet in a new workbook with the regional settings of `tag`. */
function sheetIn(tag: string) {
  setLocale(tag);
  const sheet = new Workbook().getSheet(0);
  setLocale("en-US");
  return sheet;
}

describe("typed entries", () => {
  before(async () => {
    await initSheetsWasm({ module_or_path: readFileSync(new URL(import.meta.resolve("../vendor/sheets-wasm/duke_sheets_wasm_bg.wasm"))) });
  });

  const read = (tag: string, text: string) => {
    const sheet = sheetIn(tag);
    return readUserEntry(text, (typed) => sheet.readTypedNumber(typed));
  };

  it("percentages are numeric fractions with a percent format", () => {
    assert.deepEqual(read("en-US", "12.5%"), { value: 0.125, numberFormat: { formatType: "custom", formatString: "0.##########%" } });
    assert.equal(read("en-US", " -2% ").value, -0.02);
    assert.equal(read("en-US", "1e2%").value, 1);
    assert.equal(read("en-US", "0%").value, 0);
  });

  it("grouped and scientific numbers are numeric; nonnumbers remain text", () => {
    for (const [input, expected] of [["1,234.5", 1234.5], ["-2e3", -2000], ["+.5", 0.5], ["1,2", "1,2"], ["1e999", "1e999"], ["$1,234", "$1,234"], ["TRUE", true], ["false", false], ["  text  ", "  text  "]]) {
      assert.deepEqual(read("en-US", input as string), { value: expected });
    }
  });

  it("an apostrophe forces text and preserves its whitespace", () => {
    for (const [input, expected] of [["'12.5%", "12.5%"], ["'TRUE", "TRUE"], ["'001 ", "001 "], [" '001 ", " '001 "], ["  ", "  "], ["'=SUM(A1:A3)", "=SUM(A1:A3)"], ["'5/11/2026", "5/11/2026"]]) {
      assert.deepEqual(read("en-US", input), { value: expected });
    }
  });

  it("numbers are read with the regional separators", () => {
    assert.deepEqual(read("de-DE", "1,5"), { value: 1.5 });
    assert.deepEqual(read("de-DE", "1.234,5"), { value: 1234.5 });
    assert.equal(read("de-DE", "12,5%").value, 0.125);
    assert.deepEqual(read("en-US", "1,5"), { value: "1,5" });
  });

  it("dates and times are read in the regional date order and get Excel's format", () => {
    // 46331 is 5 November 2026; 46153 is 11 May 2026.
    assert.deepEqual(read("en-GB", "5/11/2026"), { value: 46331, numberFormat: { formatType: "builtin", id: 14 } });
    assert.deepEqual(read("en-US", "5/11/2026"), { value: 46153, numberFormat: { formatType: "builtin", id: 14 } });
    assert.deepEqual(read("en-GB", "5/11/2026 13:30"), { value: 46331.5625, numberFormat: { formatType: "builtin", id: 22 } });
    assert.deepEqual(read("en-US", "13:30"), { value: 0.5625, numberFormat: { formatType: "builtin", id: 20 } });
    assert.deepEqual(read("en-US", "1:30 PM"), { value: 0.5625, numberFormat: { formatType: "builtin", id: 18 } });
  });

  it("a typed date shows in the user's order, and edits back to the same date", () => {
    const sheet = sheetIn("en-GB");
    enterUserText(sheet, "A1", "5/11/2026");
    assert.equal(sheet.getCellAt(0, 0).asNumber(), 46331);
    assert.equal(sheet.getFormattedValueAt(0, 0), "05/11/2026");
    assert.equal(sheet.getInputTextAt(0, 0), "05/11/2026");
    enterUserText(sheet, "A1", sheet.getInputTextAt(0, 0));
    assert.equal(sheet.getCellAt(0, 0).asNumber(), 46331);
  });

  it("a typed date keeps a date format the cell already has", () => {
    const sheet = sheetIn("en-GB");
    sheet.setCellStyle("A1", { numberFormat: { formatType: "custom", formatString: "d mmm yyyy" } });
    enterUserText(sheet, "A1", "5/11/2026");
    assert.equal(sheet.getFormattedValueAt(0, 0), "5 Nov 2026");
  });

  it("numbers show and edit with the regional separators", () => {
    const sheet = sheetIn("de-DE");
    enterUserText(sheet, "A1", "1,5");
    assert.equal(sheet.getFormattedValueAt(0, 0), "1,5");
    assert.equal(sheet.getInputTextAt(0, 0), "1,5");
    sheet.setCellStyle("A2", { numberFormat: { formatType: "custom", formatString: "#,##0.00" } });
    enterUserText(sheet, "A2", "1234,5");
    assert.equal(sheet.getFormattedValueAt(1, 0), "1.234,50");
  });

  it("US English shows as before, except Excel's short date", () => {
    const sheet = sheetIn("en-US");
    enterUserText(sheet, "A1", "1234.5");
    assert.equal(sheet.getFormattedValueAt(0, 0), "1234.5");
    assert.equal(sheet.getInputTextAt(0, 0), "1234.5");
    enterUserText(sheet, "A2", "11/5/2026");
    assert.equal(sheet.getFormattedValueAt(1, 0), "11/5/2026");
  });
});
