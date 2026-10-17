import { describe, expect, it } from "vitest";
import type { TableCellNode, TableNode } from "@extend-ai/react-docx-doc-model";
import { tableRowHasWholeMergeGeometry } from "../../packages/react-viewer/src/table-row-measurement";

function table(styles: Array<TableCellNode["style"]>): TableNode {
  return {
    type: "table",
    rows: [{ type: "table-row", cells: styles.map((style) => ({ type: "table-cell", style, nodes: [] })) }],
  };
}

function renderedRow(cells: Array<{ index: number; rowSpan?: number }>): HTMLTableRowElement {
  return {
    cells: cells.map(({ index, rowSpan = 1 }) => ({
      rowSpan,
      getAttribute: (name: string) =>
        name === "data-docx-cell-index" ? String(index) : null,
    })),
  } as unknown as HTMLTableRowElement;
}

describe("whole table row measurement eligibility", () => {
  it("rejects a covered row whose height is distributed by ancestor rowspans", () => {
    const value = table([{ vMergeContinuation: true }, { vMergeContinuation: true }]);
    expect(tableRowHasWholeMergeGeometry(value, 0, renderedRow([]))).toBe(false);
  });

  it("accepts nonempty ordinary rows", () => {
    expect(tableRowHasWholeMergeGeometry(table([{}, {}]), 0, renderedRow([
      { index: 0 }, { index: 1 },
    ]))).toBe(true);
  });

  it("accepts a nonempty row with its entire merge extent", () => {
    expect(tableRowHasWholeMergeGeometry(table([{ rowSpan: 3 }, {}]), 0, renderedRow([
      { index: 0, rowSpan: 3 }, { index: 1 },
    ]))).toBe(true);
  });

  it("rejects a merge shortened to a page fragment", () => {
    expect(tableRowHasWholeMergeGeometry(table([{ rowSpan: 3 }, {}]), 0, renderedRow([
      { index: 0, rowSpan: 2 }, { index: 1 },
    ]))).toBe(false);
  });

  it("rejects resumed merge content at a page start", () => {
    expect(tableRowHasWholeMergeGeometry(table([{ vMergeContinuation: true }, {}]), 0, renderedRow([
      { index: 0 }, { index: 1 },
    ]))).toBe(false);
  });

  it("keeps ordinary cells beside omitted continuation slots eligible", () => {
    expect(tableRowHasWholeMergeGeometry(table([{ vMergeContinuation: true }, {}]), 0, renderedRow([
      { index: 1 },
    ]))).toBe(true);
  });

  it("rejects an absent source row", () => {
    expect(tableRowHasWholeMergeGeometry(table([{}]), 2, renderedRow([{ index: 0 }]))).toBe(false);
  });
});
