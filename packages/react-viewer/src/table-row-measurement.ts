import type { TableNode } from "@extend-ai/react-docx-doc-model";

export function tableRowHasWholeMergeGeometry(
  table: TableNode,
  rowIndex: number,
  rowElement: HTMLTableRowElement
): boolean {
  const row = table.rows[rowIndex];
  if (!row || rowElement.cells.length === 0) return false;

  return Array.from(rowElement.cells).every((cellElement) => {
    const cellIndex = Number.parseInt(
      cellElement.getAttribute("data-docx-cell-index") ?? "",
      10
    );
    const cell = row.cells[cellIndex];
    if (!cell) return true;

    // A clipped or resumed merge changes the physical row's height. Its
    // fragment rectangle cannot replace the full row's pagination metric.
    return (
      cell.style?.vMergeContinuation !== true &&
      Math.max(1, cell.style?.rowSpan ?? 1) === cellElement.rowSpan
    );
  });
}
