import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { before, test } from "node:test";
import initSheetsWasm, { Workbook } from "../vendor/sheets-wasm/duke_sheets_wasm.js";
import { getWorksheetRowsBatch } from "./worksheet-batch.ts";

before(async () => {
  await initSheetsWasm({ module_or_path: readFileSync(new URL(import.meta.resolve("../vendor/sheets-wasm/duke_sheets_wasm_bg.wasm"))) });
});

test("formatted row batches retain type-based general alignment across the worker boundary", () => {
  const workbook = new Workbook();
  const sheet = workbook.getSheet(0);
  try {
    sheet.setCell("A1", 7319);
    sheet.setCell("B1", 1250.5);
    sheet.setCellStyle("B1", { numberFormat: { formatString: "$#,##0.00" } });
    sheet.setFormula("C1", "=A1*2");
    sheet.setCell("D1", "0042");
    sheet.setCell("E1", true);
    sheet.setFormula("F1", "=1/0");
    sheet.setCell("G1", 42);
    sheet.setCellStyle("G1", { alignment: { horizontal: "left" } });
    workbook.calculate();

    const rows = structuredClone(getWorksheetRowsBatch(sheet, 0, 1));
    assert.ok(rows);
    assert.deepEqual(rows[0].cells.map((cell) => [cell.value, cell.generalAlignment]), [
      ["7319", "right"],
      ["$1,250.50", "right"],
      ["14638", "right"],
      ["0042", "left"],
      ["TRUE", "center"],
      ["#DIV/0!", "left"],
      ["42", "right"]
    ]);
    assert.equal((rows[0].cells[6].style as { alignment: { horizontal: string } }).alignment.horizontal, "left");
  } finally {
    sheet.free();
    workbook.free();
  }
});
