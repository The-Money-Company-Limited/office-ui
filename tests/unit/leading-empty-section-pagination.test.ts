import { describe, expect, it } from "vitest";
import type { DocModel, ParagraphNode } from "../../packages/doc-model/src";
import {
  buildDocumentPageNodeSegments,
  buildDocumentPageNodeSegmentsFromLastRenderedPageBreakHints,
} from "../../packages/react-viewer/src/editor";
import { isSectionParityBlankPage } from "../../packages/layout-core/src";

function sectionProperties(type?: string): string {
  return `<w:sectPr>${type ? `<w:type w:val="${type}"/>` : ""}<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>`;
}

function paragraph(text: string, properties = ""): ParagraphNode {
  return {
    type: "paragraph",
    children: [{ type: "text", text, style: { fontFamily: "Arial", fontSizePt: 11 } }],
    sourceXml: `<w:p><w:pPr>${properties}</w:pPr>${text ? `<w:r><w:t>${text}</w:t></w:r>` : ""}</w:p>`,
  };
}

function document(nodes: ParagraphNode[], sections: [number, string][]): DocModel {
  return {
    nodes,
    metadata: {
      sourceParts: 1,
      warnings: [],
      sections: sections.map(([startNodeIndex, sectionPropertiesXml]) => ({
        startNodeIndex,
        sectionPropertiesXml,
        headerSections: [],
        footerSections: [],
      })),
    },
  };
}

describe("leading empty section pagination", () => {
  it.each([
    { type: undefined, indexes: [[0], [1]] },
    { type: "nextPage", indexes: [[0], [1]] },
    { type: "continuous", indexes: [[1]] },
    { type: "evenPage", indexes: [[0], [1]] },
    { type: "oddPage", indexes: [[0], [], [1]] },
  ])("preserves the initial section for $type", ({ type, indexes }) => {
    const model = document(
      [paragraph("", sectionProperties()), paragraph("Visible text")],
      [[0, sectionProperties()], [1, sectionProperties(type)]],
    );
    const pages = buildDocumentPageNodeSegments(model, 864, 624);
    expect(pages.map(page => page.map(segment => segment.nodeIndex))).toEqual(indexes);
    expect(pages.filter(isSectionParityBlankPage)).toHaveLength(type === "oddPage" ? 1 : 0);
  });

  it.each(["nextPage", "evenPage", "oddPage"])(
    "agrees with section-aware rendered hints for %s",
    type => {
      const model = document(
        [paragraph("", sectionProperties()), paragraph("Visible text")],
        [[0, sectionProperties()], [1, sectionProperties(type)]],
      );
      expect(buildDocumentPageNodeSegments(model, 864, 624)).toEqual(
        buildDocumentPageNodeSegmentsFromLastRenderedPageBreakHints(model),
      );
    },
  );

  it("does not add a page for an ordinary trailing empty paragraph", () => {
    const model = document(
      [paragraph("Visible text"), paragraph("")],
      [[0, sectionProperties()]],
    );
    expect(buildDocumentPageNodeSegments(model, 864, 624)).toHaveLength(1);
  });

  it("still suppresses a final structural paragraph after visible content", () => {
    const model = document(
      [paragraph("Visible text"), paragraph("", sectionProperties())],
      [[0, sectionProperties()]],
    );
    expect(buildDocumentPageNodeSegments(model, 864, 624)).toEqual([[{ nodeIndex: 0 }]]);
  });

  it.each([
    { type: "nextPage", indexes: [[0], [2]] },
    { type: "continuous", indexes: [[0, 2]] },
  ])("keeps a populated section's $type boundary", ({ type, indexes }) => {
    const model = document(
      [paragraph("First section"), paragraph("", sectionProperties()), paragraph("Second section")],
      [[0, sectionProperties()], [2, sectionProperties(type)]],
    );
    const pages = buildDocumentPageNodeSegments(model, 864, 624);
    expect(pages.map(page => page.map(segment => segment.nodeIndex))).toEqual(indexes);
  });
});
