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
 * The blocks a history entry's cells belong to, by top-left cell. A block is read once per entry
 * and its cells' states share it, so an edit over a large block stays cheap.
 */
export type ArrayFormulaBlocks = Map<string, ArrayFormulaBlock>;

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

/** Reads a cell's state; pass one `blocks` for every cell of a history entry. */
export function captureCellMutationState(
  worksheet: Worksheet,
  cell: XlsxCellAddress,
  blocks: ArrayFormulaBlocks = new Map()
): CellMutationState {
  const found = worksheet.arrayFormulaAt(cell.row, cell.col) as Omit<ArrayFormulaBlock, "values"> | null;
  let arrayFormula: ArrayFormulaBlock | undefined;
  if (found) {
    const key = cellAddressToA1(found);
    const known = blocks.get(key);
    const same = known && known.formula === found.formula && known.rows === found.rows && known.cols === found.cols;
    arrayFormula = same ? known : { ...found, values: worksheet.arrayFormulaValuesAt(found.row, found.col) };
    blocks.set(key, arrayFormula);
  }
  return {
    ...(arrayFormula ? { arrayFormula } : {}),
    formula: worksheet.getFormulaAt(cell.row, cell.col) ?? null,
    style: worksheet.getCellStyleAt(cell.row, cell.col),
    value: worksheet.getCellAt(cell.row, cell.col).toJs()
  };
}

/** Puts a cell back; a block in `rebuilt` was already put back by another cell's state. */
export function applyCellMutationState(
  worksheet: Worksheet,
  cell: XlsxCellAddress,
  state: CellMutationState,
  rebuilt: Set<ArrayFormulaBlock> = new Set()
) {
  const block = state.arrayFormula;
  if (block) {
    if (!rebuilt.has(block)) {
      rebuilt.add(block);
      // Its cells show the values they had, also when the workbook isn't calculated after undo.
      worksheet.setArrayFormula(cellAddressToA1(block), block.formula, block.values);
    }
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
  const rebuilt = new Set<ArrayFormulaBlock>();
  for (const { cell, state } of ordered) {
    applyCellMutationState(worksheet, cell, state, rebuilt);
  }
}
