import type { Worksheet } from "../vendor/sheets-wasm/duke_sheets_wasm.js";

type WorksheetBatchRow = {
  index: number;
  cells: Array<{
    col: number;
    generalAlignment?: "left" | "center" | "right";
    [key: string]: unknown;
  }>;
};

export function getCellGeneralAlignment(worksheet: Worksheet, row: number, col: number) {
  const value = worksheet.getCalculatedValueAt(row, col);
  try {
    return value.is_number ? "right" : value.is_boolean ? "center" : "left";
  } finally {
    value.free();
  }
}

export function getWorksheetRowsBatch(worksheet: Worksheet, startRow: number, rowCount: number) {
  const worksheetWithRowsBatch = worksheet as Worksheet & {
    getRowsBatch?: (startRow: number, maxRows: number, options?: unknown) => WorksheetBatchRow[] | null;
  };
  if (typeof worksheetWithRowsBatch.getRowsBatch !== "function") {
    return null;
  }

  const rows = worksheetWithRowsBatch.getRowsBatch(startRow, rowCount, {
    includeFormulas: true,
    includeHyperlinks: true,
    includeMergeInfo: true,
    includeStyles: true,
    useFormattedValues: true
  });
  for (const row of rows ?? []) {
    for (const cell of row.cells) {
      // Formatted strings cannot distinguish numbers from numeric text in the worker response.
      cell.generalAlignment = getCellGeneralAlignment(worksheet, row.index, cell.col);
    }
  }
  return rows;
}
