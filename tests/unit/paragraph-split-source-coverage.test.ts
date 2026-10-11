import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import type { DocModel, ParagraphNode } from "@extend-ai/react-docx-doc-model";

class MockCanvas {
  getContext() {
    return {
      font: "16px sans-serif",
      measureText: (text: string) => ({ width: Array.from(text).length * 10 }),
    };
  }
}

beforeEach(() => {
  vi.resetModules();
  vi.stubGlobal("OffscreenCanvas", MockCanvas);
  vi.stubGlobal("document", {});
});
afterEach(() => vi.unstubAllGlobals());

describe("paragraph split source coverage", () => {
  it.each([false, true])("covers every positioned line when the wrap estimate is shorter and widowControl=%s", async (widowControl) => {
    const { buildDocumentPageNodeSegments, paragraphLineCountWithinWidth } = await import(
      "../../packages/react-viewer/src/editor"
    );
    const { layoutTextWithPretextAroundExclusions } = await import(
      "../../packages/react-viewer/src/pretext-layout"
    );
    const text = "甲乙、丙丁戊己庚辛壬癸《子丑寅卯辰巳午未申酉戌亥甲乙（2023丙丁）》，戊己庚辛壬癸子丑、寅卯辰巳午未申酉戌亥甲、乙丙丁戊己庚辛壬癸，子丑寅卯辰巳午未申酉戌亥，甲乙丙丁戊己庚辛壬癸子丑寅卯辰巳午，未申酉戌亥甲乙丙丁戊己庚辛壬癸子丑寅卯辰巳午。";
    const paragraph: ParagraphNode = {
      type: "paragraph",
      style: {
        indent: { firstLineTwips: 300 },
        spacing: { lineRule: "exact", lineTwips: 300 },
        widowControl,
      },
      children: [{ type: "text", text, style: { fontFamily: "Arial", fontSizePt: 12 } }],
    };
    const model: DocModel = {
      nodes: [paragraph],
      metadata: { sourceParts: 1, warnings: [], headerSections: [], footerSections: [], paragraphStyles: [] },
    };
    const layout = layoutTextWithPretextAroundExclusions(
      text, "16px Arial", 240, 20,
      [{ left: 0, right: 20, top: 0, bottom: 1 }],
      { wordBreak: "keep-all", exactLineHeight: true }
    )!;
    const pages = buildDocumentPageNodeSegments(model, 60, 240, undefined, [
      { startNodeIndex: 0, pageContentHeightPx: 60, pageContentWidthPx: 240 },
    ]);
    const ranges = pages.flat().map(segment => segment.paragraphLineRange!);
    expect(paragraphLineCountWithinWidth(paragraph, 240)).toBeLessThan(layout.lineCount);
    expect(ranges[0]?.startLineIndex).toBe(0);
    expect(ranges.every((range, index) => range.totalLineCount === layout.lineCount &&
      (index === 0 || range.startLineIndex === ranges[index - 1].endLineIndex))).toBe(true);
    expect(ranges.at(-1)?.endLineIndex).toBe(layout.lineCount);
    const renderedText = ranges.flatMap(range => layout.lines.slice(range.startLineIndex, range.endLineIndex))
      .flatMap(line => line.fragments).map(fragment => fragment.text).join("");
    expect(renderedText).toBe(text);
  });
});
