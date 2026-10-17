import { describe, expect, it } from "vitest";
import type { DocModel, ParagraphNode } from "../../packages/doc-model/src";
import { isSectionParityBlankPage } from "../../packages/layout-core/src";
import { buildDocumentPageNodeSegments } from "../../packages/react-viewer/src/editor";

function paragraph(text: string, content: string, properties = ""): ParagraphNode {
  return {
    type: "paragraph",
    blockId: `paragraph-${text.length}`,
    children: [{ type: "text", text, style: { fontFamily: "Arial", fontSizePt: 11 } }],
    sourceXml: `<w:p><w:pPr>${properties}</w:pPr><w:r>${content}</w:r></w:p>`,
  };
}

const pageBreak = '<w:br w:type="page"/>';
const visible = () => paragraph("Visible text", "<w:t>Visible text</w:t>");

function document(nodes: ParagraphNode[]): DocModel {
  return { nodes, metadata: { sourceParts: 1, warnings: [] } };
}

describe("terminal manual page breaks", () => {
  it.each([
    { name: "standalone", nodes: () => [visible(), paragraph("\n", pageBreak)], count: 2, anchor: { nodeIndex: 1, sourceOffset: 1 } },
    { name: "inline", nodes: () => [paragraph("Visible text\n", `<w:t>Visible text</w:t>${pageBreak}`)], count: 2, anchor: { nodeIndex: 0, sourceOffset: 13 } },
    { name: "double", nodes: () => [visible(), paragraph("\n\n", pageBreak + pageBreak)], count: 3, anchor: { nodeIndex: 1, sourceOffset: 2 } },
    { name: "break-only", nodes: () => [paragraph("\n", pageBreak)], count: 2, anchor: { nodeIndex: 0, sourceOffset: 1 } },
  ])("preserves the empty final page for a $name break", ({ nodes, count, anchor }) => {
    const model = document(nodes());
    const pages = buildDocumentPageNodeSegments(model, 864, 624);
    expect(pages).toHaveLength(count);
    expect(pages.at(-1)).toHaveLength(0);
    expect(pages.at(-1)).toMatchObject({ terminalPageBreakAnchor: anchor });
    expect(pages.some(isSectionParityBlankPage)).toBe(false);
  });

  it.each([
    { name: "soft line break", nodes: () => [visible(), paragraph("\n", "<w:br/>")], count: 1 },
    { name: "rendered break hint", nodes: () => [visible(), paragraph("", "<w:lastRenderedPageBreak/>")], count: 1 },
    { name: "empty paragraph after a break", nodes: () => [visible(), paragraph("\n", pageBreak), paragraph("", "")], count: 2 },
    { name: "page break before an empty paragraph", nodes: () => [visible(), paragraph("", "", "<w:pageBreakBefore/>")], count: 2 },
  ])("does not append an extra page for $name", ({ nodes, count }) => {
    const pages = buildDocumentPageNodeSegments(document(nodes()), 864, 624);
    expect(pages).toHaveLength(count);
    expect(pages.at(-1)).not.toHaveProperty("terminalPageBreakAnchor");
  });

  it("does not repeat paragraph content on the empty page", () => {
    const model = document([paragraph("Visible text\n", `<w:t>Visible text</w:t>${pageBreak}`)]);
    const pages = buildDocumentPageNodeSegments(model, 864, 624);
    expect(pages.flat().map(segment => segment.nodeIndex)).toEqual([0]);
    expect(model.nodes).toHaveLength(1);
  });

  it("does not apply a retained break after the source text was edited", () => {
    const model = document([paragraph("Changed text", `<w:t>Visible text</w:t>${pageBreak}`)]);
    const pages = buildDocumentPageNodeSegments(model, 864, 624);
    expect(pages).toHaveLength(1);
  });
});
