import * as React from "react";
import { renderToStaticMarkup } from "react-dom/server";
import { afterAll, beforeAll, describe, expect, it, vi } from "vitest";
import type { DocModel, ParagraphNode } from "../../packages/doc-model/src";
import { DocxEditorViewer, useDocxEditor } from "../../packages/react-viewer/src/editor";
import {
  layoutItemsWithPretextAroundExclusions,
  resolveCaretRectAtOffset,
  resolveHorizontalTextScale,
  resolveOffsetAtPoint,
  resolveSelectionRects,
  sliceLayoutToLineRange,
  withPretextHorizontalAlignment,
  type PretextLayoutItem,
  type PretextVariableWidthLayout,
} from "../../packages/react-viewer/src/pretext-layout";

const graphemes = new Intl.Segmenter("en", { granularity: "grapheme" });
class CanvasContext {
  font = "16px Arial";
  letterSpacing = "0px";
  measureText(text: string) {
    const count = [...graphemes.segment(text)].filter(({ segment }) => !/[\r\n]/.test(segment)).length;
    return {
      width: count * (10 + Number.parseFloat(this.letterSpacing)),
      fontBoundingBoxAscent: 12, fontBoundingBoxDescent: 4,
      actualBoundingBoxAscent: 12, actualBoundingBoxDescent: 4,
    };
  }
}
class Canvas {
  getContext() { return new CanvasContext(); }
}
beforeAll(() => {
  vi.stubGlobal("OffscreenCanvas", Canvas);
  vi.stubGlobal("document", {});
});
afterAll(() => vi.unstubAllGlobals());

function item(text: string, scale = 1, start = 0, spacing = 0): PretextLayoutItem {
  return { text, font: "16px Arial", startOffset: start, endOffset: start + text.length, horizontalScale: scale, letterSpacingPx: spacing };
}
function layout(text: string, scale = 1, width = 1000, spacing = 0) {
  return layoutItemsWithPretextAroundExclusions(text, [item(text, scale, 0, spacing)], width, 20)!;
}
const fragments = (value: PretextVariableWidthLayout) => value.lines.flatMap(line => line.fragments);
function expectSourceCoverage(value: PretextVariableWidthLayout, text: string) {
  expect(fragments(value).map(fragment => fragment.text).join("")).toBe(text);
  let offset = 0;
  for (const fragment of fragments(value)) {
    expect(fragment.startOffset).toBe(offset);
    expect(fragment.text).toBe(text.slice(fragment.startOffset, fragment.endOffset));
    offset = fragment.endOffset;
  }
  expect(offset).toBe(text.length);
}
function Viewer({ paragraph }: { paragraph: ParagraphNode }) {
  const model: DocModel = {
    nodes: [paragraph],
    metadata: {
      sourceParts: 1, warnings: [], headerSections: [], footerSections: [], paragraphStyles: [],
      sectionPropertiesXml: '<w:sectPr><w:pgSz w:w="3000" w:h="6000"/><w:pgMar w:top="300" w:right="300" w:bottom="300" w:left="300"/></w:sectPr>',
    },
  };
  const editor = useDocxEditor({ starterModel: model });
  return React.createElement(DocxEditorViewer, { editor, mode: "read-only", deferInitialPaginationPaint: false, pageVirtualization: { enabled: false } });
}
function render(align: "left" | "center" | "right" | "justify", percent: number, spacing = 0) {
  return renderToStaticMarkup(React.createElement(Viewer, { paragraph: {
    type: "paragraph", blockId: "body0",
    style: { align, spacing: { beforeTwips: 0, afterTwips: 0, lineTwips: 240, lineRule: "auto" } },
    children: [{ type: "text", text: "Alpha beta gamma delta", style: { fontFamily: "Arial", fontSizePt: 12, characterScalePercent: percent, characterSpacingTwips: spacing } }],
  } }));
}

describe("horizontal character scaling", () => {
  it("keeps scaled runs breakable and preserves UTF16 source ranges", () => {
    const text = "a😀b e\u0301f ABCDEFGHIJKLMNOP";
    const value = layout(text, 1.5, 70);
    expect(value.lineCount).toBeGreaterThan(1);
    expectSourceCoverage(value, text);
    expect(fragments(value).every(fragment => fragment.width <= 70)).toBe(true);
    expect(resolveSelectionRects(value, 1, 3)[0]?.width).toBe(15);
  });

  it("separates scale caches and mixed shaping items without changing vertical metrics", () => {
    const values = [1, 1.5, 0.75, 1].map(scale => layout("ABCDE", scale));
    expect(values.map(value => fragments(value)[0].width)).toEqual([50, 75, 37.5, 50]);
    expect(new Set(values.map(value => value.height)).size).toBe(1);
    const mixed = layoutItemsWithPretextAroundExclusions("AB", [item("A", 1.5), item("B", 0.5, 1)], 100, 20)!;
    expect(fragments(mixed).map(fragment => [fragment.x, fragment.width])).toEqual([[0, 15], [15, 5]]);
  });

  it.each([0.75, 1, 1.5])("adds tracking after glyph scaling at %s", (scale) => {
    const plain = layout("HHHHHHHHHH", scale);
    const tracked = layout("HHHHHHHHHH", scale, 1000, 3);
    for (const offset of [1, 2, 5, 9]) {
      const x = resolveCaretRectAtOffset(tracked, offset)!.left;
      expect(x - resolveCaretRectAtOffset(plain, offset)!.left).toBe(offset * 3);
      expect(resolveOffsetAtPoint(tracked, x, 10)).toBe(offset);
    }
  });

  it.each(["center", "right"] as const)("aligns wrapped %s fragments before slicing and selection", (align) => {
    const source = layout("ABCD EFGH IJKL", 1.5, 100);
    const aligned = withPretextHorizontalAlignment(source, align);
    const x = align === "center" ? 20 : 40;
    expectSourceCoverage(aligned, "ABCD EFGH IJKL");
    expect(aligned.lines[0].fragments[0].x).toBe(x);
    expect(resolveCaretRectAtOffset(aligned, 2)?.left).toBe(x + 30);
    expect(resolveSelectionRects(aligned, 1, 3)[0]?.left).toBe(x + 15);
    const sliced = sliceLayoutToLineRange(aligned, 1, 2);
    expect(sliced.lines[0].fragments[0].startOffset).toBe(5);
    expect(resolveCaretRectAtOffset(sliced, 6, { affinity: "downstream" })?.left).toBe(x + 15);
  });

  it("keeps independent exclusion intervals and their source offsets", () => {
    const text = "AB CD EF GH";
    const source = layoutItemsWithPretextAroundExclusions(text, [item(text, 1.5)], 180, 20, [{ left: 60, right: 120, top: 0, bottom: 20 }])!;
    const aligned = withPretextHorizontalAlignment(source, "center");
    expectSourceCoverage(aligned, text);
    expect(aligned.lines[0].fragments.length).toBeGreaterThanOrEqual(2);
    for (const fragment of aligned.lines[0].fragments) {
      expect(fragment.x).toBeGreaterThanOrEqual(fragment.intervalX);
      expect(fragment.x).toBeLessThanOrEqual(fragment.intervalX + fragment.intervalWidth);
    }
  });

  it.each(["left", "center", "right"] as const)("routes ordinary scaled %s text through source-preserving fragments", (align) => {
    const html = render(align, 150);
    expect(html).toContain("transform:scaleX(1.5)");
    const spans = [...html.matchAll(/data-docx-text-start="(\d+)" data-docx-text-end="(\d+)"[^>]*>([^<]*)/g)];
    expect(spans.length).toBeGreaterThan(1);
    expect(spans.map(span => span[3]).join("")).toBe("Alpha beta gamma delta");
    expect(spans[0][1]).toBe("0");
    expect(spans.at(-1)?.[2]).toBe("22");
  });

  it.each([[150, 2], [75, 4]])("counter-scales inner tracking before painting %s percent glyphs", (percent, spacing) => {
    const html = render("center", percent, 45);
    const spans = [...html.matchAll(/data-docx-text-start="(\d+)" data-docx-text-end="(\d+)" style="([^"]*)"/g)];
    expect(spans.length).toBeGreaterThan(1);
    expect(spans.every(span => span[3].includes(`letter-spacing:${spacing}px`))).toBe(true);
    expect(html).toContain(`transform:scaleX(${percent / 100})`);
  });

  it("retains the ordinary route for unscaled centered and unsupported justified content", () => {
    expect(render("center", 100)).not.toContain("data-docx-wrapped-paragraph-root");
    expect(render("justify", 150)).not.toContain("transform:scaleX");
    expect(resolveHorizontalTextScale(0)).toBe(1);
    expect(resolveHorizontalTextScale(150)).toBe(1.5);
  });
});
