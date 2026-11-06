import type { CellStyle, Worksheet } from "../vendor/sheets-wasm/duke_sheets_wasm.js";
import type { XlsxCellAddress } from "./types";

/** A Ctrl+Shift+Enter array formula: its top-left cell, extent and formula. */
export type ArrayFormulaBlock = {
  col: number;
  cols: number;
  formula: string;
  row: number;
  rows: number;
};

/**
 * A cell as undo and redo put it back. A cell of a Ctrl+Shift+Enter array formula records the
 * whole block: changing one of its cells turns the block into values, so undo rebuilds it.
 */
export type CellMutationState = {
  arrayFormula?: ArrayFormulaBlock;
  formula: string | null;
  style: CellStyle | null;
  value: unknown;
};

function a1(row: number, col: number): string {
  let label = "";
  for (let n = col; n >= 0; n = Math.floor(n / 26) - 1) {
    label = String.fromCharCode(65 + (n % 26)) + label;
  }
  return `${label}${row + 1}`;
}

export function captureCellMutationState(worksheet: Worksheet, cell: XlsxCellAddress): CellMutationState {
  const arrayFormula = (worksheet.arrayFormulaAt(cell.row, cell.col) as ArrayFormulaBlock | null) ?? undefined;
  return {
    ...(arrayFormula ? { arrayFormula } : {}),
    formula: worksheet.getFormulaAt(cell.row, cell.col) ?? null,
    style: worksheet.getCellStyleAt(cell.row, cell.col),
    value: worksheet.getCellAt(cell.row, cell.col).toJs()
  };
}

export function applyCellMutationState(worksheet: Worksheet, cell: XlsxCellAddress, state: CellMutationState) {
  const block = state.arrayFormula;
  if (block) {
    // The block's other cells are filled when the workbook is calculated.
    worksheet.setArrayFormula(a1(block.row, block.col), block.formula, block.rows, block.cols);
  } else if (state.formula) {
    worksheet.setFormula(a1(cell.row, cell.col), state.formula);
  } else {
    worksheet.setCell(a1(cell.row, cell.col), state.value ?? "");
  }

  // A cell that had no style gets the default one back, dropping a format typing gave it.
  if (state.style) {
    worksheet.setCellStyleAt(cell.row, cell.col, state.style);
  } else {
    worksheet.clearCellStyleAt(cell.row, cell.col);
  }
}

/**
 * Puts a range's cells back. Undo goes in reverse, so a block whose cells were changed one by one
 * is rebuilt after the cells changed later are put back.
 */
export function applyCellMutationStates(
  worksheet: Worksheet,
  mutations: Array<{ cell: XlsxCellAddress; state: CellMutationState }>,
  direction: "undo" | "redo"
) {
  const ordered = direction === "undo" ? [...mutations].reverse() : mutations;
  for (const { cell, state } of ordered) {
    applyCellMutationState(worksheet, cell, state);
  }
}
