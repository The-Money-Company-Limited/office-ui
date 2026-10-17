import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import type { ParagraphNode } from "@extend-ai/react-docx-doc-model";
import type { PretextExclusionRect } from "../../packages/react-viewer/src/pretext-layout";

class MockCanvas {
  getContext() {
    return {
      font: "16px sans-serif",
      measureText: (text: string) => ({ width: Array.from(text).length * 7 }),
    };
  }
}

beforeEach(() => {
  vi.resetModules();
  vi.stubGlobal("OffscreenCanvas", MockCanvas);
  vi.stubGlobal("document", {});
});
afterEach(() => vi.unstubAllGlobals());

function paragraph(text: string, lineHeightPx = 24): ParagraphNode {
  return {
    type: "paragraph",
    style: { spacing: { lineRule: "exact", lineTwips: lineHeightPx * 15 } },
    children: [{ type: "text", text, style: { fontFamily: "Arial", fontSizePt: 12 } }],
  };
}

describe("static foreign wrap flow", () => {
  it("keeps consecutive one-line paragraphs beside one continuing obstacle", async () => {
    const api = await import("../../packages/react-viewer/src/editor");
    const obstacle = { sourceNodeIndex: 0, left: 0, right: 120, top: 8, bottom: 490 };
    let flowTopPx = 0;
    const rendered: string[] = [];
    for (let index = 1; index <= 12; index++) {
      const text = `Label ${index}`;
      const exclusions = api.resolveForeignWrapExclusionsForFlowRange(
        [obstacle], index, flowTopPx, flowTopPx + 24
      );
      const wrapped = api.resolveParagraphForeignOnlyWrappedTextLayout(
        paragraph(text), 400, 24, exclusions
      )!;
      expect(wrapped).toBeDefined();
      expect(wrapped.layout.lines).toHaveLength(1);
      expect(wrapped.layout.lines[0].fragments[0].x).toBe(120);
      const height = api.resolveDragPreviewAwareWrapBlockHeightPx(wrapped.layout, [], exclusions);
      expect(height).toBe(24);
      rendered.push(wrapped.layout.lines.flatMap(line => line.fragments).map(fragment => fragment.text).join(""));
      flowTopPx += height;
    }
    expect(flowTopPx).toBe(288);
    expect(rendered).toEqual(Array.from({ length: 12 }, (_, index) => `Label ${index + 1}`));
  });

  it("advances by multiline text while retaining the exclusion for the next paragraph", async () => {
    const api = await import("../../packages/react-viewer/src/editor");
    const obstacle = { sourceNodeIndex: 0, left: 0, right: 120, top: 0, bottom: 480 };
    const text = "Alpha beta gamma delta epsilon zeta eta theta iota kappa";
    const firstExclusions = api.resolveForeignWrapExclusionsForFlowRange([obstacle], 1, 0, 24);
    const first = api.resolveParagraphForeignOnlyWrappedTextLayout(paragraph(text), 240, 24, firstExclusions)!;
    expect(first.layout.lines.length).toBeGreaterThan(1);
    expect(first.layout.lines.every(line => line.fragments.every(fragment => fragment.x >= 120))).toBe(true);
    const height = api.resolveDragPreviewAwareWrapBlockHeightPx(first.layout, [], firstExclusions);
    expect(height).toBe(first.layout.lines.length * 24);
    expect(height).toBeLessThan(480);
    const nextExclusions = api.resolveForeignWrapExclusionsForFlowRange([obstacle], 2, height, height + 24);
    const next = api.resolveParagraphForeignOnlyWrappedTextLayout(paragraph("Next"), 240, 24, nextExclusions)!;
    expect(next.layout.lines[0].fragments[0].x).toBe(120);
    expect(api.resolveDragPreviewAwareWrapBlockHeightPx(next.layout, [], nextExclusions)).toBe(24);
    expect(first.source.text).toBe(text);
    expect(next.source.text).toBe("Next");
  });

  it("retains vertical clearance when the full column is blocked", async () => {
    const api = await import("../../packages/react-viewer/src/editor");
    const exclusions = [{ left: 0, right: 240, top: 0, bottom: 120 }];
    const wrapped = api.resolveParagraphForeignOnlyWrappedTextLayout(paragraph("After"), 240, 24, exclusions)!;
    expect(wrapped.layout.lines[0].y).toBe(120);
    expect(api.resolveDragPreviewAwareWrapBlockHeightPx(wrapped.layout, [], exclusions)).toBe(144);
  });

  it.each([true, false])("keeps foreign-only flow at text extent with dragPreview=%s", async fromDragPreview => {
    const api = await import("../../packages/react-viewer/src/editor");
    const exclusions = [{ left: 0, right: 120, top: 0, bottom: 300, fromDragPreview }];
    const wrapped = api.resolveParagraphForeignOnlyWrappedTextLayout(paragraph("Moving"), 300, 24, exclusions)!;
    expect(api.resolveDragPreviewAwareWrapBlockHeightPx(wrapped.layout, [], exclusions)).toBe(24);
  });

  it("preserves owned-image clearance and its existing drag calculation", async () => {
    const api = await import("../../packages/react-viewer/src/editor");
    const { layoutTextWithPretextAroundExclusions } = await import("../../packages/react-viewer/src/pretext-layout");
    const geometry = api.resolveDualWrappedFloatingImageGeometry({
      type: "image", widthPx: 60, heightPx: 20,
      floating: { xPx: 0, yPx: 0, wrapType: "square", wrapText: "bothSides" },
    }, 300)!;
    expect(geometry).toBeDefined();
    const staticExclusion = { left: 180, right: 240, top: 0, bottom: 120 };
    const dragExclusion = { left: 240, right: 300, top: 0, bottom: 240, fromDragPreview: true };
    const layout = layoutTextWithPretextAroundExclusions("Text", "16px Arial", 300, 24,
      [geometry.exclusion, staticExclusion, dragExclusion], { exactLineHeight: true })!;
    expect(layout.lines).toHaveLength(1);
    expect(api.resolveDragPreviewAwareWrapBlockHeightPx(layout, [geometry], [staticExclusion])).toBe(240);
    expect(api.resolveDragPreviewAwareWrapBlockHeightPx(layout, [geometry], [staticExclusion, dragExclusion])).toBe(120);
    expect(api.resolveDragPreviewAwareWrapBlockHeightPx(layout, [geometry], [dragExclusion])).toBe(24);
  });

  it("preserves exact short line heights without imposing the obstacle height", async () => {
    const api = await import("../../packages/react-viewer/src/editor");
    const exclusions = [{ left: 0, right: 100, top: 0, bottom: 200 }];
    const wrapped = api.resolveParagraphForeignOnlyWrappedTextLayout(paragraph("Short", 12), 300, 12, exclusions)!;
    expect(api.resolveDragPreviewAwareWrapBlockHeightPx(wrapped.layout, [], exclusions)).toBe(12);
  });

  it("keeps sliced fragments at their text height and preserves every source span", async () => {
    const api = await import("../../packages/react-viewer/src/editor");
    const { sliceLayoutToLineRange } = await import("../../packages/react-viewer/src/pretext-layout");
    const text = "Alpha beta gamma delta epsilon zeta eta theta iota kappa";
    const exclusions: PretextExclusionRect[] = [{ left: 0, right: 120, top: 0, bottom: 480 }];
    const wrapped = api.resolveParagraphForeignOnlyWrappedTextLayout(paragraph(text), 240, 24, exclusions)!;
    const middle = Math.floor(wrapped.layout.lineCount / 2);
    const first = sliceLayoutToLineRange(wrapped.layout, 0, middle);
    const last = sliceLayoutToLineRange(wrapped.layout, middle, wrapped.layout.lineCount);
    expect(first.sourceRange?.startOffset).toBe(0);
    expect(first.sourceRange?.endOffset).toBe(last.sourceRange?.startOffset);
    expect(last.sourceRange?.endOffset).toBe(text.length);
    expect(last.sourceRange?.includesEnd).toBe(true);
    for (const fragment of [first, last]) {
      expect(fragment.height).toBe(fragment.lineCount * 24);
      expect(api.resolveDragPreviewAwareWrapBlockHeightPx(fragment, [], exclusions)).toBe(fragment.height);
    }
    expect(first.sourceLineRange?.endLineIndex).toBe(last.sourceLineRange?.startLineIndex);
    expect(text.slice(first.sourceRange!.startOffset, first.sourceRange!.endOffset) +
      text.slice(last.sourceRange!.startOffset, last.sourceRange!.endOffset)).toBe(text);
  });
});
