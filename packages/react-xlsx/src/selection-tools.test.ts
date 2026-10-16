import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { before, test } from "node:test";
import initSheetsWasm, { Workbook } from "@dukelib/sheets-wasm";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import * as tools from "./selection-tools.ts";

before(async () => {
  await initSheetsWasm({ module_or_path: readFileSync(new URL(import.meta.resolve("@dukelib/sheets-wasm/duke_sheets_wasm_bg.wasm"))) });
});

function workbookWithSelection(selection: string) {
  const workbook = new Workbook();
  const archive = unzipSync(workbook.saveXlsxBytes());
  workbook.free();
  const path = "xl/worksheets/sheet1.xml";
  archive[path] = strToU8(strFromU8(archive[path]).replace(
    "<sheetFormatPr",
    `<sheetViews><sheetView workbookViewId="0">${selection}</sheetView></sheetViews><sheetFormatPr`
  ));
  return Workbook.fromBytes(zipSync(archive));
}

test("initial selection defaults to A1 when a worksheet has no saved active cell", () => {
  const workbook = workbookWithSelection("");
  assert.deepEqual(tools.getXlsxInitialActiveCell(workbook.getSheet(0)), { row: 0, col: 0 });
  workbook.free();
});

test("initial selection uses the saved active cell rather than the selection range corner", () => {
  const workbook = workbookWithSelection('<selection activeCell="D7" sqref="C4:F9"/>');
  assert.deepEqual(tools.getXlsxInitialActiveCell(workbook.getSheet(0)), { row: 6, col: 3 });
  workbook.free();
});

test("initial selection ignores frozen-pane placeholders without an active cell", () => {
  const workbook = workbookWithSelection(
    '<pane xSplit="1" ySplit="1" activePane="bottomRight" state="frozen"/>'
    + '<selection pane="topRight"/><selection pane="bottomLeft"/>'
    + '<selection pane="bottomRight" activeCell="AA12" sqref="AA12"/>'
  );
  assert.deepEqual(tools.getXlsxInitialActiveCell(workbook.getSheet(0)), { row: 11, col: 26 });
  workbook.free();
});

test("initial selection uses the active pane when each frozen pane has a saved cell", () => {
  const workbook = workbookWithSelection(
    '<pane xSplit="1" ySplit="1" activePane="bottomLeft" state="frozen"/>'
    + '<selection pane="topRight" activeCell="B1" sqref="B1"/>'
    + '<selection pane="bottomLeft" activeCell="A8" sqref="A8"/>'
    + '<selection pane="bottomRight" activeCell="D7" sqref="D7"/>'
  );
  assert.deepEqual(tools.getXlsxInitialActiveCell(workbook.getSheet(0), "bottomLeft"), { row: 7, col: 0 });
  workbook.free();
});

test("selection statistics use calculated numbers, count populated cells, and ignore numeric text and errors in the sum", () => {
  const workbook = new Workbook();
  const sheet = workbook.getSheet(0);
  sheet.setCell("A1", 1250.5);
  sheet.setCellStyle("A1", { numberFormat: { formatString: "$#,##0.00" } });
  sheet.setFormula("B1", "=A1*2");
  sheet.setCell("C1", "123");
  sheet.setCell("D1", true);
  sheet.setFormula("E1", "=1/0");
  workbook.calculate();
  assert.deepEqual(tools.getXlsxSelectionStatistics(sheet, { start: { row: 0, col: 5 }, end: { row: 0, col: 0 } }), {
    average: 1875.75, count: 5, numericCount: 2, sum: 3751.5
  });
  sheet.setCell("A1", 10);
  workbook.calculate();
  assert.equal(tools.getXlsxSelectionStatistics(sheet, { start: { row: 0, col: 0 }, end: { row: 1048575, col: 1 } })?.sum, 30);
  assert.equal(tools.getXlsxSelectionStatistics(sheet, { start: { row: 10, col: 0 }, end: { row: 20, col: 1 } })?.numericCount, 0);
});

test("find searches displayed values in the active sheet, case-insensitively and in row order", () => {
  const workbook = new Workbook();
  const sheet = workbook.getSheet(0);
  sheet.setCell("A1", 0.5);
  sheet.setCellStyle("A1", { numberFormat: { formatString: "0%" } });
  sheet.setFormula("B2", "=A1");
  sheet.setCellStyle("B2", { numberFormat: { formatString: "0%" } });
  sheet.setCell("A3", "Fifty PERCENT");
  workbook.addSheet("Other");
  workbook.getSheet(1).setCell("A1", "50%");
  workbook.calculate();
  const bounds = { visibleRows: [0, 1, 2, 3], visibleCols: [0, 1, 2], minUsedRow: 0, maxUsedRow: 2, minUsedCol: 0, maxUsedCol: 1 };
  const display = ({ row, col }: { row: number; col: number }) => sheet.getFormattedValueAt(row, col);
  assert.deepEqual(tools.findXlsxCells(bounds, display, "50%"), [{ row: 0, col: 0 }, { row: 1, col: 1 }]);
  assert.deepEqual(tools.findXlsxCells(bounds, display, "percent"), [{ row: 2, col: 0 }]);
  assert.deepEqual(tools.findXlsxCells(bounds, display, "=A1"), []);
  assert.deepEqual(tools.findXlsxCells(bounds, display, ""), []);
  assert.deepEqual(tools.findXlsxCells({ ...bounds, visibleRows: [0, 2] }, display, "50%"), [{ row: 0, col: 0 }]);
});

test("text-only and mixed selections expose Count with numeric statistics only when numbers exist", () => {
  const workbook = new Workbook();
  const sheet = workbook.getSheet(0);
  sheet.setCell("A1", "one");
  sheet.setCell("A2", "two");
  const range = { start: { row: 0, col: 0 }, end: { row: 1, col: 0 } };
  assert.deepEqual(tools.getXlsxSelectionStatistics(sheet, range), { count: 2, numericCount: 0, sum: null, average: null });
  sheet.setCell("A2", 7);
  assert.deepEqual(tools.getXlsxSelectionStatistics(sheet, range), { count: 2, numericCount: 1, sum: 7, average: 7 });
});

test("statistics clip whole-column selections to the used range and reject oversized intersections before reading", () => {
  const workbook = new Workbook();
  const sheet = workbook.getSheet(0);
  sheet.setCell("B3", 4);
  sheet.setCell("B4", 6);
  const reads: string[] = [];
  const original = sheet.getCalculatedValueAt.bind(sheet);
  sheet.getCalculatedValueAt = (row, col) => { reads.push(`${row}:${col}`); return original(row, col); };
  const wholeColumn = { start: { row: 1048575, col: 1 }, end: { row: 0, col: 1 } };
  assert.equal(tools.getXlsxSelectionStatistics(sheet, wholeColumn)?.sum, 10);
  assert.deepEqual(reads, ["2:1", "3:1"]);
  reads.length = 0;
  sheet.setCell("B100000", 1);
  assert.equal(tools.getXlsxSelectionStatistics(sheet, wholeColumn), null);
  assert.deepEqual(reads, []);
  assert.equal(tools.getXlsxSelectionStatistics(sheet, { start: { row: 0, col: 0 }, end: { row: 1048575, col: 0 } })?.count, 0);
  assert.deepEqual(reads, []);
});

test("find clips reads to used visible cells and caps large sheets before reading", () => {
  const bounds = { visibleRows: [0, 1, 2, 3], visibleCols: [0, 1, 2], minUsedRow: 1, maxUsedRow: 2, minUsedCol: 1, maxUsedCol: 1 };
  const reads: string[] = [];
  const display = ({ row, col }: { row: number; col: number }) => { reads.push(`${row}:${col}`); return "match"; };
  assert.deepEqual(tools.findXlsxCells(bounds, display, "match"), [{ row: 1, col: 1 }, { row: 2, col: 1 }]);
  assert.deepEqual(reads, ["1:1", "2:1"]);
  reads.length = 0;
  assert.ok(tools.findXlsxCells({ ...bounds, visibleRows: Array.from({ length: 100000 }, (_, i) => i), maxUsedRow: 99999 }, display, "match") === null);
  assert.deepEqual(reads, []);
});

test("find keeps its index on refresh, clamps when matches shrink, and resets only for a different query", () => {
  const current = { query: "match", index: 2 };
  assert.equal(tools.resolveXlsxFindIndex(current, "match", 5), 2);
  assert.equal(tools.resolveXlsxFindIndex(current, "match", 2), 1);
  assert.equal(tools.resolveXlsxFindIndex(current, "match", 0), 0);
  assert.equal(tools.resolveXlsxFindIndex(current, "other", 5), 0);
});

test("cut keeps source cells until paste and moves overlapping values without rewriting dependent formulas", () => {
  const workbook = new Workbook();
  const sheet = workbook.getSheet(0);
  sheet.setCell("A1", 10);
  sheet.setCell("A2", true);
  sheet.setFormula("B1", "=A1");
  const cut = tools.captureXlsxCut(sheet, { start: { row: 0, col: 0 }, end: { row: 1, col: 0 } });
  assert.equal(sheet.getCell("A1").asNumber(), 10);
  tools.applyXlsxCut(cut, sheet, sheet, { row: 1, col: 0 });
  assert.equal(sheet.getCell("A1").is_empty, true);
  assert.equal(sheet.getCell("A2").asNumber(), 10);
  assert.equal(sheet.getCell("A3").asBoolean(), true);
  assert.equal(sheet.getFormulaAt(0, 1), "=A1");
});

test("cut moves formulas and merged cells between sheets", () => {
  const workbook = new Workbook();
  const source = workbook.getSheet(0);
  workbook.addSheet("Target");
  const target = workbook.getSheet(1);
  source.setFormula("A1", "=1+2");
  source.mergeCells("A1:B1");
  const cut = tools.captureXlsxCut(source, { start: { row: 0, col: 0 }, end: { row: 0, col: 1 } });
  tools.applyXlsxCut(cut, source, target, { row: 2, col: 2 });
  assert.equal(source.getFormulaAt(0, 0), undefined);
  assert.equal(source.isMergedSecondary(0, 1), false);
  assert.equal(target.getFormulaAt(2, 2), "=1+2");
  assert.equal(target.isMergedSecondary(2, 3), true);
});
