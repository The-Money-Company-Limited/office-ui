import { afterAll, beforeAll, describe, expect, it, vi } from "vitest";
import type { ParagraphNode, TableNode } from "../../packages/doc-model/src";
class MockCanvas {
  getContext() {
    return { font: "8px sans-serif", measureText: (text: string) => ({
      width: Array.from(text).length * 4,
      fontBoundingBoxAscent: 7, fontBoundingBoxDescent: 2,
      actualBoundingBoxAscent: 7, actualBoundingBoxDescent: 2,
    }) };
  }
}
beforeAll(() => { vi.stubGlobal("OffscreenCanvas", MockCanvas); vi.stubGlobal("document", {}); });
afterAll(() => vi.unstubAllGlobals());
async function height(spacing: NonNullable<NonNullable<ParagraphNode["style"]>["spacing"]>, options?: { grid?: number; tableGrid?: boolean; snap?: boolean; rowExact?: boolean }) {
  const api = await import("../../packages/react-viewer/src/editor");
  const p: ParagraphNode = { type: "paragraph", style: { spacing: { beforeTwips: 0, afterTwips: 0, ...spacing }, snapToGrid: options?.snap }, children: [{ type: "text", text: "Alpha\nBeta", style: { fontFamily: "Arial", fontSizePt: 6 } }] };
  const table: TableNode = { type: "table", style: { layout: "fixed", widthTwips: 1800, columnWidthsTwips: [1800], cellMarginTwips: { topTwips: 0, bottomTwips: 0, leftTwips: 0, rightTwips: 0 } }, rows: [{ type: "table-row", style: options?.rowExact ? { heightRule: "exact", heightTwips: 450 } : undefined, cells: [{ type: "table-cell", nodes: [p] }] }] };
  return api.estimateTableRowHeightsPx(table, 120, undefined, options?.grid, 200, 120, { adjustLineHeightInTable: options?.tableGrid })[0];
}
describe("table paragraph natural and declared line extents", () => {
  it("keeps two small-font lines at measured natural advance", async () => {
    expect(await height({ lineRule: "auto", lineTwips: 240 })).toBeCloseTo(18, 5);
  });
  it("honors an explicit at-least line minimum below the pagination reserve", async () => {
    expect(await height({ lineRule: "atLeast", lineTwips: 180 })).toBeCloseTo(24, 5);
  });
  it("honors active table grid and paragraph opt-out", async () => {
    expect(await height({ lineRule: "auto", lineTwips: 240 }, { grid: 10, tableGrid: true })).toBeCloseTo(20, 5);
    expect(await height({ lineRule: "auto", lineTwips: 240 }, { grid: 10, tableGrid: true, snap: false })).toBeCloseTo(18, 5);
    expect(await height({ lineRule: "auto", lineTwips: 240 }, { grid: 10, tableGrid: false })).toBeCloseTo(18, 5);
  });
  it("preserves exact paragraph and row heights", async () => {
    expect(await height({ lineRule: "exact", lineTwips: 120 })).toBeCloseTo(16, 5);
    expect(await height({ lineRule: "auto", lineTwips: 240 }, { rowExact: true })).toBeCloseTo(30, 5);
  });
});
