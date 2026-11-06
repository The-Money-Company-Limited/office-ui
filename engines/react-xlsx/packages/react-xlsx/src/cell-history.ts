import type { CellStyle, Worksheet } from "../vendor/sheets-wasm/duke_sheets_wasm.js";
import { cellAddressToA1 } from "./range-reference.ts";
import type { XlsxCellAddress } from "./types";

/**
 * A Ctrl+Shift+Enter array formula: its top-left cell, extent and formula, and the values its
 * cells showed, which undo puts back as they were (`setArrayFormula` takes them).
 */
export type ArrayFormulaBlock = {
  col: number;
  cols: number;
  formula: string;
  row: number;
  rows: number;
  values: unknown[][];
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
    // Its cells show the values they had, also when the workbook isn't calculated after undo.
    worksheet.setArrayFormula(cellAddressToA1(block), block.formula, block.values);
  } else if (state.formula) {
    worksheet.setFormula(cellAddressToA1(cell), state.formula);
  } else {
    worksheet.setCell(cellAddressToA1(cell), state.value ?? null);
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
