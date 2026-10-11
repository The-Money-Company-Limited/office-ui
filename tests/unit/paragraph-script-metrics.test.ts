import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import type { ParagraphNode, TextStyle } from "../../packages/doc-model/src";

let tallExtent = 32;
class MetricCanvas {
  getContext() {
    return {
      font: "16px Latin Metric Face",
      measureText(text: string) {
        const extent = this.font.toLowerCase().includes("tall") ? tallExtent : 16;
        return { width: text.length * 8, fontBoundingBoxAscent: extent * 0.75, fontBoundingBoxDescent: extent * 0.25 };
      },
    };
  }
}

const slots: TextStyle = {
  fontFamily: "Latin Metric Face",
  sourceFontFamily: "Latin Metric Face",
  fontFamilyAscii: "Latin Metric Face",
  fontFamilyHAnsi: "Latin Metric Face",
  fontFamilyEastAsia: "East Metric Face",
  resolvedFontFamilyEastAsia: "Tall East Metric Face",
  fontFamilyCs: "Tall Complex Metric Face",
  fontSizePt: 12,
};
const paragraph = (text: string, style: TextStyle = slots): ParagraphNode => ({
  type: "paragraph",
  style: { spacing: { lineRule: "auto", lineTwips: 240 } },
  children: [{ type: "text", text, style }],
});

beforeEach(() => {
  vi.resetModules();
  tallExtent = 32;
  vi.stubGlobal("document", {});
  vi.stubGlobal("OffscreenCanvas", MetricCanvas);
});
afterEach(() => vi.unstubAllGlobals());

describe("paragraph metric family stability", () => {
  it("keeps the paragraph strut family separate from script font requests", async () => {
    const { estimateParagraphLineHeightPx } = await import("../../packages/react-viewer/src/editor");
    const p = paragraph("A漢字仮名", { ...slots, fontFamily: "Tall Base Metric Face", sourceFontFamily: "Tall Base Metric Face" });
    expect(estimateParagraphLineHeightPx(p, 20)).toBe(40);
    expect(estimateParagraphLineHeightPx(p)).toBe(32);
    expect(estimateParagraphLineHeightPx(paragraph("العربية"), 20)).toBe(20);
  });

  it("does not use an unused script slot for Latin text", async () => {
    const { estimateParagraphLineHeightPx } = await import("../../packages/react-viewer/src/editor");
    expect(estimateParagraphLineHeightPx(paragraph("Latin text"), 20)).toBe(20);
  });

  it("preserves a later explicit legacy-family edit", async () => {
    const { estimateParagraphLineHeightPx } = await import("../../packages/react-viewer/src/editor");
    expect(estimateParagraphLineHeightPx(paragraph("漢字仮名", { ...slots, fontFamily: "Replacement Metric Face" }), 20)).toBe(20);
  });

  it("keeps exact lines exact and preserves grid opt-outs", async () => {
    const { estimateParagraphLineHeightPx } = await import("../../packages/react-viewer/src/editor");
    const p = paragraph("漢字仮名", { ...slots, fontFamily: "Tall Base Metric Face" });
    expect(estimateParagraphLineHeightPx({ ...p, style: { spacing: { lineRule: "exact", lineTwips: 270 } } }, 20)).toBe(18);
    expect(estimateParagraphLineHeightPx({ ...p, style: { ...p.style, snapToGrid: false } }, 20)).toBe(32);
    expect(estimateParagraphLineHeightPx(p, 20, true)).toBe(32);
    expect(estimateParagraphLineHeightPx({ ...p, style: { spacing: { lineRule: "atLeast", lineTwips: 270 } } }, 20)).toBe(32);
  });

  it("retains the empty paragraph mark family", async () => {
    const { estimateParagraphLineHeightPx } = await import("../../packages/react-viewer/src/editor");
    const p: ParagraphNode = { ...paragraph(""), paragraphMarkStyle: { fontFamily: "Tall Mark Metric Face", fontSizePt: 12 } };
    expect(estimateParagraphLineHeightPx(p, 20)).toBe(40);
  });

  it("remeasures after the public font-metric refresh invalidation", async () => {
    const { estimateParagraphLineHeightPx } = await import("../../packages/react-viewer/src/editor");
    const { invalidateFontMetrics } = await import("../../packages/react-viewer/src/font-metrics");
    const p = paragraph("漢字仮名", { ...slots, fontFamily: "Tall Base Metric Face" });
    tallExtent = 18;
    expect(estimateParagraphLineHeightPx(p, 20)).toBe(20);
    tallExtent = 32;
    invalidateFontMetrics();
    expect(estimateParagraphLineHeightPx(p, 20)).toBe(40);
  });
});
