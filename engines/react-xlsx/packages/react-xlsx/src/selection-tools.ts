import type { Worksheet } from "../vendor/sheets-wasm/duke_sheets_wasm.js";
import type { XlsxCellAddress, XlsxCellRange, XlsxSheetData } from "./types";
import { parseXlsxRange } from "./range-reference.ts";

export function getXlsxInitialActiveCell(worksheet: Worksheet, activePane?: string): XlsxCellAddress {
  const selections = worksheet.selections as Array<{ activeCell?: string; pane?: string }>;
  for (const selection of selections) {
    if (activePane && selection.pane !== activePane) continue;
    if (!selection.activeCell) continue;
    const parsed = parseXlsxRange(selection.activeCell);
    if (parsed?.kind === "cells") return parsed.range.start;
  }
  return { row: 0, col: 0 };
}

// Bound synchronous WASM reads; null means the result is too large to calculate completely.
const MAX_SCAN_CELLS = 50_000;

/** Uses numeric results, including formulas, without parsing formatted currency or dates. */
export function getXlsxSelectionStatistics(worksheet: Worksheet, range: XlsxCellRange) {
  const used = worksheet.usedRange() as [number, number, number, number] | null;
  let count = 0;
  let numericCount = 0;
  let sum = 0;
  if (used) {
    const startRow = Math.max(used[0], Math.min(range.start.row, range.end.row));
    const endRow = Math.min(used[2], Math.max(range.start.row, range.end.row));
    const startCol = Math.max(used[1], Math.min(range.start.col, range.end.col));
    const endCol = Math.min(used[3], Math.max(range.start.col, range.end.col));
    if (Math.max(0, endRow - startRow + 1) * Math.max(0, endCol - startCol + 1) > MAX_SCAN_CELLS) return null;
    for (let row = startRow; row <= endRow; row += 1) {
      for (let col = startCol; col <= endCol; col += 1) {
        const value = worksheet.getCalculatedValueAt(row, col);
        if (!value.is_empty) count += 1;
        if (value.is_number) {
          sum += value.asNumber()!;
          numericCount += 1;
        }
        value.free();
      }
    }
  }
  return { average: numericCount ? sum / numericCount : null, count, numericCount, sum: numericCount ? sum : null };
}

/** Searches the same displayed values the grid exposes, in row-major order. */
export function findXlsxCells(
  sheet: Pick<XlsxSheetData, "visibleRows" | "visibleCols" | "minUsedRow" | "maxUsedRow" | "minUsedCol" | "maxUsedCol">,
  getDisplayValue: (cell: XlsxCellAddress) => string,
  query: string
): XlsxCellAddress[] | null {
  if (!query) return [];
  const needle = query.toLocaleLowerCase();
  const columns = sheet.visibleCols.filter((col) => col >= sheet.minUsedCol && col <= sheet.maxUsedCol);
  const rows = sheet.visibleRows.filter((row) => row >= sheet.minUsedRow && row <= sheet.maxUsedRow);
  if (rows.length * columns.length > MAX_SCAN_CELLS) return null;
  const matches: XlsxCellAddress[] = [];
  for (const row of rows) {
    for (const col of columns) {
      const cell = { row, col };
      if (getDisplayValue(cell).toLocaleLowerCase().includes(needle)) matches.push(cell);
    }
  }
  return matches;
}

export function resolveXlsxFindIndex(current: { query: string; index: number }, query: string, matchCount: number) {
  return current.query === query ? Math.min(current.index, Math.max(0, matchCount - 1)) : 0;
}

function cellReference(row: number, col: number) {
  let column = "";
  for (let value = col + 1; value > 0; value = Math.floor((value - 1) / 26)) {
    column = String.fromCharCode(65 + (value - 1) % 26) + column;
  }
  return `${column}${row + 1}`;
}

export function captureXlsxCut(worksheet: Worksheet, range: XlsxCellRange) {
  const start = { row: Math.min(range.start.row, range.end.row), col: Math.min(range.start.col, range.end.col) };
  const end = { row: Math.max(range.start.row, range.end.row), col: Math.max(range.start.col, range.end.col) };
  const cells = [];
  for (let row = start.row; row <= end.row; row += 1) {
    for (let col = start.col; col <= end.col; col += 1) {
      const value = worksheet.getCellAt(row, col);
      cells.push({
        row, col,
        value: value.toJs() as unknown,
        formula: worksheet.getFormulaAt(row, col),
        style: worksheet.getCellStyleAt(row, col),
        merge: worksheet.getMergeSpan(row, col) as { rowSpan: number; colSpan: number } | null
      });
      value.free();
    }
  }
  return { range: { start, end }, cells };
}

export function applyXlsxCut(
  cut: ReturnType<typeof captureXlsxCut>,
  source: Worksheet,
  target: Worksheet,
  destination: XlsxCellAddress
) {
  // The engine cannot rewrite dependent formula references when cells move.
  // Snapshot all cells before clearing so overlapping moves preserve their values.
  for (const cell of cut.cells) {
    if (cell.merge && (cell.merge.rowSpan > 1 || cell.merge.colSpan > 1)) {
      source.unmergeCells(`${cellReference(cell.row, cell.col)}:${cellReference(cell.row + cell.merge.rowSpan - 1, cell.col + cell.merge.colSpan - 1)}`);
    }
    source.setCell(cellReference(cell.row, cell.col), null);
  }
  for (const cell of cut.cells) {
    const row = destination.row + cell.row - cut.range.start.row;
    const col = destination.col + cell.col - cut.range.start.col;
    const address = cellReference(row, col);
    if (cell.formula) target.setFormula(address, cell.formula);
    else target.setCell(address, cell.value ?? null);
    if (cell.style) target.setCellStyle(address, cell.style);
  }
  for (const cell of cut.cells) {
    if (!cell.merge || (cell.merge.rowSpan <= 1 && cell.merge.colSpan <= 1)) continue;
    const row = destination.row + cell.row - cut.range.start.row;
    const col = destination.col + cell.col - cut.range.start.col;
    target.mergeCells(`${cellReference(row, col)}:${cellReference(row + cell.merge.rowSpan - 1, col + cell.merge.colSpan - 1)}`);
  }
}
