import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { before, describe, it } from "node:test";
import initSheetsWasm, { Workbook } from "../vendor/sheets-wasm/duke_sheets_wasm.js";
import { applyCellMutationState, applyCellMutationStates, captureCellMutationState } from "./cell-history.ts";

before(async () => {
  await initSheetsWasm({
    module_or_path: readFileSync(new URL(import.meta.resolve("../vendor/sheets-wasm/duke_sheets_wasm_bg.wasm")))
  });
});

/** A1:A3 holding 1, 2, 3 and a Ctrl+Shift+Enter block {=A1:A3*2} over C1:C3. */
function workbookWithBlock() {
  const workbook = new Workbook();
  const sheet = workbook.getSheet(0);
  sheet.setCell("A1", 1);
  sheet.setCell("A2", 2);
  sheet.setCell("A3", 3);
  sheet.setArrayFormula("C1", "=A1:A3*2", 3, 1);
  workbook.calculate();
  return { workbook, sheet };
}

const value = (sheet: ReturnType<Workbook["getSheet"]>, row: number, col: number) =>
  sheet.getCalculatedValueAt(row, col).toJs();

describe("undo and redo of a Ctrl+Shift+Enter array formula", () => {
  it("rebuilds the block after one of its cells was changed", () => {
    const { workbook, sheet } = workbookWithBlock();
    assert.equal(value(sheet, 1, 2), 4);
    const cell = { row: 1, col: 2 };
    const before = captureCellMutationState(sheet, cell);
    assert.deepEqual(before.arrayFormula, { row: 0, col: 2, rows: 3, cols: 1, formula: "=A1:A3*2" });

    sheet.setCell("C2", 9);
    const after = captureCellMutationState(sheet, cell);
    workbook.calculate();
    assert.equal(sheet.arrayFormulaAt(0, 2), null, "changing part of the block turns it into values");
    assert.equal(value(sheet, 0, 2), 2);

    applyCellMutationState(sheet, cell, before);
    workbook.calculate();
    assert.equal(sheet.arrayFormulaAt(2, 2)?.formula, "=A1:A3*2");
    assert.deepEqual([value(sheet, 0, 2), value(sheet, 1, 2), value(sheet, 2, 2)], [2, 4, 6]);

    applyCellMutationState(sheet, cell, after);
    workbook.calculate();
    assert.equal(sheet.arrayFormulaAt(0, 2), null);
    assert.deepEqual([value(sheet, 0, 2), value(sheet, 1, 2), value(sheet, 2, 2)], [2, 9, 6]);
  });

  it("rebuilds it after a range edit over its cells", () => {
    const { workbook, sheet } = workbookWithBlock();
    const cells = [{ row: 0, col: 2 }, { row: 1, col: 2 }];
    const mutations = cells.map((cell) => {
      const before = captureCellMutationState(sheet, cell);
      sheet.setCell(cell.row === 0 ? "C1" : "C2", "x");
      return { cell, before, after: captureCellMutationState(sheet, cell) };
    });
    workbook.calculate();
    applyCellMutationStates(sheet, mutations.map(({ cell, before }) => ({ cell, state: before })), "undo");
    workbook.calculate();
    assert.deepEqual([value(sheet, 0, 2), value(sheet, 1, 2), value(sheet, 2, 2)], [2, 4, 6]);
    assert.equal(sheet.arrayFormulaAt(1, 2)?.rows, 3);
  });
});
