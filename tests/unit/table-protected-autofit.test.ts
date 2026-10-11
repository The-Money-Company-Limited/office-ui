import { afterAll, beforeAll, describe, expect, it, vi } from "vitest";
import type { TableNode } from "../../packages/doc-model/src";
import { invalidateFontMetrics } from "../../packages/react-viewer/src/font-metrics";
import { estimateFloatingTableWidthPx } from "../../packages/react-viewer/src/editor";

let advance = 8;
class MockCanvas {
  getContext() {
    return { font: "16px serif", measureText: (text: string) => ({ width: text.length * advance }) };
  }
}
beforeAll(() => { vi.stubGlobal("OffscreenCanvas", MockCanvas); vi.stubGlobal("document", {}); invalidateFontMetrics(); });
afterAll(() => { vi.unstubAllGlobals(); invalidateFontMetrics(); });

function table(text: string, options: { fixed?: boolean; noWrap?: boolean; split?: boolean } = {}): TableNode {
  const cell = (width: number, text: string) => ({
    type: "table-cell" as const,
    style: { preferredWidth: { type: "dxa" as const, value: width * 15 }, noWrap: options.noWrap ?? true },
    nodes: [{ type: "paragraph" as const, children: (options.split ? [text.slice(0, 6), text.slice(6)] : [text]).map(value => ({
      type: "text" as const, text: value, style: { fontFamily: "Arial", fontSizePt: 12 },
    })) }],
  });
  return {
    type: "table", style: {
      layout: options.fixed ? "fixed" : "autofit", preferredWidth: { type: "dxa", value: 6000 },
      columnWidthsTwips: [2400, 4800], cellMarginTwips: { leftTwips: 75, rightTwips: 75 },
    }, rows: [{ type: "table-row", cells: [cell(100, text), cell(300, "Beta")] }],
  };
}

describe("automatic table preferred minima", () => {
  it("expands for intact tokens without shrinking a protected neighboring column", () => {
    expect(estimateFloatingTableWidthPx(table("ABCDEFGHIJ/KLMNOPQRST"), 400)).toBe(478);
    expect(estimateFloatingTableWidthPx(table("ABCDEFGHIJ/KLMNOPQRST", { split: true }), 400)).toBe(478);
  });
  it("wraps ordinary spaces and hyphens without trusting an oversized stored grid", () => {
    expect(estimateFloatingTableWidthPx(table("Alpha Beta Gamma Delta"), 400)).toBe(400);
    expect(estimateFloatingTableWidthPx(table("ABCDEFGHIJ-KLMNOPQRST"), 400)).toBe(400);
    expect(estimateFloatingTableWidthPx(table("ABCDEFGHIJ\u2003KLMNOPQRST"), 400)).toBe(400);
    expect(estimateFloatingTableWidthPx(table("ABCDEFGHIJ\u00a0KLMNOPQRST"), 400)).toBe(478);
  });
  it("keeps fixed and unprotected tables on their existing width path", () => {
    expect(estimateFloatingTableWidthPx(table("ABCDEFGHIJ/KLMNOPQRST", { fixed: true }), 400)).toBe(400);
    expect(estimateFloatingTableWidthPx(table("ABCDEFGHIJ/KLMNOPQRST", { noWrap: false }), 400)).toBe(400);
  });
  it("retains column protection when other rows omit the setting", () => {
    const model = table("Alpha");
    model.rows.push(table("ABCDEFGHIJ/KLMNOPQRST", { noWrap: false }).rows[0]);
    expect(estimateFloatingTableWidthPx(model, 400)).toBe(478);
  });
  it("leaves inline geometry requiring other layout rules on the existing path", () => {
    expect(estimateFloatingTableWidthPx(table("ABCDEFGHIJ/KLMNOPQRST\tBeta"), 400)).toBe(400);
    const model = table("ABCDEFGHIJ/KLMNOPQRST");
    model.sourceXml = "<w:tbl><w:tcFitText/></w:tbl>";
    expect(estimateFloatingTableWidthPx(model, 400)).toBe(400);
  });
  it("keeps tracking continuous across equivalent formatting runs", () => {
    const model = table("ABCDEFGHIJ/KLMNOPQRST", { split: true });
    const paragraph = model.rows[0].cells[0].nodes[0];
    if (paragraph.type !== "paragraph") throw new Error("Expected paragraph");
    for (const run of paragraph.children) {
      if (run.type === "text") run.style = { ...run.style, characterSpacingTwips: 30 };
    }
    expect(estimateFloatingTableWidthPx(model, 400)).toBe(518);
  });
  it("refreshes intrinsic widths when resolved fonts change", () => {
    const model = table("ABCDEFGHIJ/KLMNOPQRST");
    expect(estimateFloatingTableWidthPx(model, 400)).toBe(478);
    advance = 9;
    invalidateFontMetrics();
    expect(estimateFloatingTableWidthPx(model, 400)).toBe(499);
    advance = 8;
    invalidateFontMetrics();
  });
});
