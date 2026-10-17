import { afterAll, beforeAll, describe, expect, it, vi } from "vitest";
import type { ParagraphNode } from "../../packages/doc-model/src";
import {
  resolveParagraphDualWrappedTextLayout,
  resolveParagraphForeignOnlyWrappedTextLayout,
} from "../../packages/react-viewer/src/editor";
import {
  resolveCaretRectAtOffset,
  resolveOffsetAtPoint,
  resolveSelectionRects,
  sliceLayoutToLineRange,
  type PretextVariableWidthLayout,
} from "../../packages/react-viewer/src/pretext-layout";

class CanvasContext {
  font = "16px Arial";
  letterSpacing = "0px";
  measureText(text: string) {
    return {
      width: text.length * (10 + Number.parseFloat(this.letterSpacing)),
      fontBoundingBoxAscent: 12,
      fontBoundingBoxDescent: 4,
      actualBoundingBoxAscent: 12,
      actualBoundingBoxDescent: 4,
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

type Alignment = "left" | "center" | "right";
function paragraph(align: Alignment, scaled = true): ParagraphNode {
  return {
    type: "paragraph",
    style: { align },
    children: [
      { type: "text", text: "AB", style: { fontFamily: "Arial", fontSizePt: 12, characterScalePercent: scaled ? 150 : 100 } },
      { type: "text", text: "CD", style: { fontFamily: "Arial", fontSizePt: 12, characterScalePercent: scaled ? 75 : 100 } },
    ],
  };
}
function owned(paragraph: ParagraphNode, left = 0, width = 80, height = 100) {
  return resolveParagraphDualWrappedTextLayout({ ...paragraph, children: [{
    type: "image", widthPx: width, heightPx: height,
    floating: { xPx: left, yPx: 0, distLPx: 0, distRPx: 0, distTPx: 0, distBPx: 0, wrapType: "square", wrapText: "bothSides", behindDocument: false },
  }, ...paragraph.children] }, 320, 20)!.layout;
}
function foreign(paragraph: ParagraphNode, left = 0, right = 80, bottom = 100) {
  return resolveParagraphForeignOnlyWrappedTextLayout(paragraph, 320, 20, [{ left, right, top: 0, bottom }])!.layout;
}
function expectCoverage(value: PretextVariableWidthLayout, text: string) {
  const fragments = value.lines.flatMap(line => line.fragments);
  expect(fragments.map(fragment => fragment.text).join("")).toBe(text);
  let offset = 0;
  for (const fragment of fragments) {
    expect(fragment.startOffset).toBe(offset);
    expect(fragment.text).toBe(text.slice(fragment.startOffset, fragment.endOffset));
    offset = fragment.endOffset;
  }
  expect(offset).toBe(text.length);
}

describe("wrapped character scaling alignment", () => {
  it.each(["left", "center", "right"] as const)("anchors %s mixed-scale fragments once in an owned and foreign interval", (align) => {
    const expectedX = 80 + (240 - 45) * (align === "left" ? 0 : align === "center" ? 0.5 : 1);
    for (const value of [owned(paragraph(align)), foreign(paragraph(align))]) {
      expectCoverage(value, "ABCD");
      const line = value.lines[0];
      expect(line.fragments.map(fragment => [fragment.x, fragment.width])).toEqual([[expectedX, 30], [expectedX + 30, 15]]);
      expect(resolveCaretRectAtOffset(value, 1)?.left).toBe(expectedX + 15);
      expect(resolveOffsetAtPoint(value, expectedX + 15, 10)).toBe(1);
      expect(resolveSelectionRects(value, 0, 2)[0]?.left).toBe(expectedX);
      expect(resolveSelectionRects(value, 0, 2)[0]?.width).toBe(30);
    }
  });

  it.each(["left", "center", "right"] as const)("keeps %s anchors through two intervals, wrapping, and line slices", (align) => {
    const text = "AB CD EF GH IJ KL MN OP QR ST UV WX YZ";
    const node = paragraph(align);
    node.children = [{ type: "text", text, style: { fontFamily: "Arial", fontSizePt: 12, characterScalePercent: 150 } }];
    for (const value of [owned(node, 80, 80, 20), foreign(node, 80, 160, 20)]) {
      expectCoverage(value, text);
      expect(value.lineCount).toBeGreaterThan(1);
      expect(new Set(value.lines[0].fragments.map(fragment => fragment.intervalX)).size).toBe(2);
      for (const line of value.lines) {
        const groups = new Map<string, typeof line.fragments>();
        for (const fragment of line.fragments) {
          const key = `${fragment.intervalX}:${fragment.intervalWidth}`;
          groups.set(key, [...(groups.get(key) ?? []), fragment]);
        }
        for (const group of groups.values()) {
          const first = group[0];
          const visibleText = group.map(fragment => fragment.text).join("").trimEnd();
          const free = Math.max(0, first.intervalWidth - visibleText.length * 15);
          const expectedX = first.intervalX + free * (align === "left" ? 0 : align === "center" ? 0.5 : 1);
          expect(first.x).toBe(expectedX);
          expect(resolveCaretRectAtOffset(value, first.startOffset + 1)?.left).toBe(expectedX + 15);
        }
      }
      const sliced = sliceLayoutToLineRange(value, 1, value.lineCount);
      expect(sliced.lines.map(line => line.fragments)).toEqual(value.lines.slice(1).map(line => line.fragments));
    }
  });

  it.each(["left", "center", "right"] as const)("preserves the existing unscaled %s placement", (align) => {
    const expectedX = 80 + (240 - 40) * (align === "left" ? 0 : align === "center" ? 0.5 : 1);
    for (const value of [owned(paragraph(align, false)), foreign(paragraph(align, false))]) {
      expectCoverage(value, "ABCD");
      expect(value.lines[0].fragments[0].x).toBe(expectedX);
      expect(value.lines[0].fragments.reduce((width, fragment) => width + fragment.width, 0)).toBe(40);
    }
  });
});
