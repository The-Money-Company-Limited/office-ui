import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import type { ParagraphNode } from "@extend-ai/react-docx-doc-model";

class MockCanvas {
  getContext() {
    return {
      font: "16px Alpha",
      measureText(text: string) {
        const reverse = this.font.includes("Beta");
        return {
          width: text.length * 8,
          fontBoundingBoxAscent: reverse ? 4 : 18,
          fontBoundingBoxDescent: reverse ? 18 : 4,
        };
      },
    };
  }
}

beforeEach(() => {
  vi.resetModules();
  vi.stubGlobal("OffscreenCanvas", MockCanvas);
  vi.stubGlobal("document", {});
});
afterEach(() => vi.unstubAllGlobals());

describe("paragraph line grid", () => {
  it("rounds natural extent up while retaining a fractional auto-grid minimum", async () => {
    const { resolveLineGridHeightPx } = await import("../../packages/react-viewer/src/pretext-layout");
    expect(resolveLineGridHeightPx(19, 20)).toBe(20);
    expect(resolveLineGridHeightPx(20, 20)).toBe(20);
    expect(resolveLineGridHeightPx(22, 20)).toBe(40);
    expect(resolveLineGridHeightPx(22, 20, 50)).toBe(50);
    expect(resolveLineGridHeightPx(22, 0)).toBe(22);
  });

  it("uses the natural ascent/descent union of mixed fonts before rounding", async () => {
    const { layoutItemsWithPretextAroundExclusions } = await import("../../packages/react-viewer/src/pretext-layout");
    const items = [
      { text: "A", font: "16px Alpha", startOffset: 0, endOffset: 1, lineHeightPx: 22 },
      { text: "B", font: "16px Beta", startOffset: 1, endOffset: 2, lineHeightPx: 22 },
    ];
    const natural = layoutItemsWithPretextAroundExclusions("AB", items, 100, 22)!;
    const grid = layoutItemsWithPretextAroundExclusions("AB", items, 100, 22, [], undefined,
      { lineGrid: { pitchPx: 20 } })!;
    expect(natural.lines[0].height).toBe(36);
    expect(grid.lines[0].height).toBe(40);
    expect(grid.lines[0].ascent).toBe(20);
    expect(grid.lines[0].descent).toBe(20);
    expect(grid.lines[0].fragments.map(fragment => fragment.text).join("")).toBe("AB");
    expect(layoutItemsWithPretextAroundExclusions("AB", items, 100, 22)!.lines[0].height).toBe(36);
  });

  it("applies grid advances to hard breaks and empty ending marks but preserves exact lines", async () => {
    const { layoutTextWithPretextAroundExclusions } = await import("../../packages/react-viewer/src/pretext-layout");
    const options = { lineGrid: { pitchPx: 20 }, endingMark: { font: "16px Alpha", sourceOffset: 2 } };
    const grid = layoutTextWithPretextAroundExclusions("A\n", "16px Alpha", 100, 22, [], options)!;
    expect(grid.lines.map(line => [line.y, line.height])).toEqual([[0, 40], [40, 40]]);
    const exact = layoutTextWithPretextAroundExclusions("A\n", "16px Alpha", 100, 18, [],
      { ...options, exactLineHeight: true })!;
    expect(exact.lines.map(line => [line.y, line.height])).toEqual([[0, 18], [18, 18]]);
    const empty = layoutTextWithPretextAroundExclusions("", "16px Alpha", 100, 22, [],
      { lineGrid: { pitchPx: 20 }, endingMark: { font: "16px Alpha", sourceOffset: 0 } })!;
    expect(empty.lines[0].height).toBe(40);
  });

  it("checks the full rounded line height against wrap exclusions", async () => {
    const { layoutTextWithPretextAroundExclusions } = await import("../../packages/react-viewer/src/pretext-layout");
    const layout = layoutTextWithPretextAroundExclusions("A", "16px Alpha", 100, 22,
      [{ left: 0, right: 100, top: 30, bottom: 60 }], { lineGrid: { pitchPx: 20 } })!;
    expect(layout.lines[0].y).toBeGreaterThanOrEqual(60);
  });

  it("preserves paragraph opt-outs, inactive grids, table gates, and explicit spacing", async () => {
    const { estimateParagraphLineHeightPx, resolveParagraphForeignOnlyWrappedTextLayout } = await import("../../packages/react-viewer/src/editor");
    const paragraph: ParagraphNode = {
      type: "paragraph",
      style: { spacing: { lineRule: "auto", lineTwips: 240 } },
      children: [{ type: "text", text: "AB", style: { fontFamily: "Alpha", fontSizePt: 12 } }],
    };
    expect(estimateParagraphLineHeightPx(paragraph, 20)).toBe(40);
    expect(estimateParagraphLineHeightPx(paragraph)).toBe(22);
    expect(estimateParagraphLineHeightPx(paragraph, 20, true)).toBe(22);
    expect(estimateParagraphLineHeightPx({ ...paragraph, style: { ...paragraph.style, snapToGrid: false } }, 20)).toBe(22);
    expect(estimateParagraphLineHeightPx({ ...paragraph, style: { ...paragraph.style, sourceInheritedSnapToGrid: false } }, 20)).toBe(22);
    expect(estimateParagraphLineHeightPx({ ...paragraph, style: { spacing: { lineRule: "auto", lineTwips: 600 } } }, 20)).toBe(50);
    expect(estimateParagraphLineHeightPx({ ...paragraph, style: { spacing: { lineRule: "exact", lineTwips: 270 } } }, 20)).toBe(18);
    expect(estimateParagraphLineHeightPx({ ...paragraph, style: { spacing: { lineRule: "atLeast", lineTwips: 270 } } }, 20)).toBe(22);
    const rich: ParagraphNode = { ...paragraph, children: [
      { type: "text", text: "A", style: { fontFamily: "Alpha", fontSizePt: 12 } },
      { type: "text", text: "B", style: { fontFamily: "Beta", fontSizePt: 12 } },
    ] };
    const layout = resolveParagraphForeignOnlyWrappedTextLayout(rich, 100, 40,
      [{ left: 90, right: 100, top: 0, bottom: 1 }], { minimumLineHeightPx: 20, gridLinePitchPx: 20 })!;
    expect(layout.layout.lines[0].height).toBe(40);
  });
});
