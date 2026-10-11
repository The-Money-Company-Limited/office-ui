import { describe, expect, it } from "vitest";
import type { DocModel, ParagraphNode, TableNode } from "@extend-ai/react-docx-doc-model";
import {
  buildDocumentPageNodeSegments,
  estimateRenderedPageSegmentHeightPx,
  resolveColumnBlockingFloatingTableFlowSpacing,
} from "../../packages/react-viewer/src/editor";

function paragraph(text: string): ParagraphNode {
  return {
    type: "paragraph",
    style: { spacing: { beforeTwips: 0, afterTwips: 0, lineTwips: 360, lineRule: "exact" }, widowControl: false },
    children: [{ type: "text", text, style: { fontFamily: "Arial", fontSizePt: 11 } }],
  };
}

function table(rowCount = 2, widthTwips = 9360): TableNode {
  return {
    type: "table",
    style: {
      widthTwips,
      columnWidthsTwips: [widthTwips],
      layout: "fixed",
      cellMarginTwips: { topTwips: 0, bottomTwips: 0, leftTwips: 0, rightTwips: 0 },
      floating: {
        verticalAnchor: "text", horizontalAnchor: "margin", xTwips: 0, yTwips: 121,
        leftFromTextTwips: 180, rightFromTextTwips: 180,
        topFromTextTwips: 0, bottomFromTextTwips: 0,
      },
    },
    rows: Array.from({ length: rowCount }, (_, index) => ({
      type: "table-row",
      style: { heightTwips: 3600, heightRule: "atLeast" },
      cells: [{ type: "table-cell", nodes: [paragraph(`Row ${index + 1}`)] }],
    })),
  };
}

function document(floatingTable: TableNode): DocModel {
  return {
    nodes: [
      ...Array.from({ length: 8 }, (_, index) => paragraph(`Lead ${index + 1}`)),
      floatingTable,
      ...Array.from({ length: 12 }, (_, index) => paragraph(`Following ${index + 1}`)),
    ],
    metadata: { sourceParts: 1, warnings: [], headerSections: [], footerSections: [], paragraphStyles: [] },
  };
}

describe("full-column floating table pagination", () => {
  it("reserves the full table and initial offset before following text", () => {
    const model = document(table());
    const pages = buildDocumentPageNodeSegments(model, 864, 624);
    expect(pages).toHaveLength(2);
    expect(pages[0].map(segment => segment.nodeIndex)).toEqual(Array.from({ length: 16 }, (_, index) => index));
    expect(pages[1].map(segment => segment.nodeIndex)).toEqual([16, 17, 18, 19, 20]);
    expect(pages[0][8].tableRowRange).toEqual({
      startRowIndex: 0, endRowIndex: 2, flowBeforeSpacingPx: 8, flowAfterSpacingPx: 0,
    });
    expect(estimateRenderedPageSegmentHeightPx(model.nodes[8], pages[0][8], model, 624)).toBeGreaterThanOrEqual(488);
  });

  it("moves a minimum-height row intact and applies the offset only to the first fragment", () => {
    const model = document(table(3));
    const pages = buildDocumentPageNodeSegments(model, 864, 624);
    expect(pages).toHaveLength(2);
    expect(pages[0].map(segment => segment.nodeIndex)).toEqual(Array.from({ length: 9 }, (_, index) => index));
    expect(pages[0][8].tableRowRange).toMatchObject({ startRowIndex: 0, endRowIndex: 2, flowBeforeSpacingPx: 8 });
    expect(pages[1][0].tableRowRange).toMatchObject({ startRowIndex: 2, endRowIndex: 3, flowBeforeSpacingPx: 0 });
    expect(pages.flat().some(segment => segment.tableRowSlice)).toBe(false);
    expect(pages[1].slice(1).map(segment => segment.nodeIndex)).toEqual(Array.from({ length: 12 }, (_, index) => index + 9));
  });

  it("keeps side-wrapped tables out of flow", () => {
    const model = document(table(2, 2880));
    const pages = buildDocumentPageNodeSegments(model, 864, 624);
    expect(pages).toHaveLength(1);
    expect(pages[0][8]).toEqual({ nodeIndex: 8 });
    expect(estimateRenderedPageSegmentHeightPx(model.nodes[8], pages[0][8], model, 624)).toBe(0);
  });

  it.each(["page", "margin"] as const)("preserves absolute %s vertical anchors", verticalAnchor => {
    const floatingTable = table();
    floatingTable.style!.floating!.verticalAnchor = verticalAnchor;
    const pages = buildDocumentPageNodeSegments(document(floatingTable), 864, 624);
    expect(pages).toHaveLength(1);
    expect(pages[0][8]).toEqual({ nodeIndex: 8 });
  });

  it("uses the wrap exclusion to decide whether a column is blocked", () => {
    const floatingTable = table(2, 7800);
    floatingTable.style!.floating!.rightFromTextTwips = 1560;
    expect(resolveColumnBlockingFloatingTableFlowSpacing(floatingTable, 624)).toEqual({ beforePx: 8, afterPx: 0 });
    floatingTable.style!.floating!.rightFromTextTwips = 0;
    expect(resolveColumnBlockingFloatingTableFlowSpacing(floatingTable, 624)).toBeUndefined();
  });

  it("preserves backward overlap and page-relative horizontal placement", () => {
    const floatingTable = table();
    floatingTable.style!.floating!.yTwips = -120;
    expect(resolveColumnBlockingFloatingTableFlowSpacing(floatingTable, 624)).toBeUndefined();
    floatingTable.style!.floating!.yTwips = 121;
    floatingTable.style!.floating!.topFromTextTwips = 300;
    expect(resolveColumnBlockingFloatingTableFlowSpacing(floatingTable, 624)).toBeUndefined();
    floatingTable.style!.floating!.topFromTextTwips = 0;
    floatingTable.style!.floating!.horizontalAnchor = "page";
    expect(resolveColumnBlockingFloatingTableFlowSpacing(floatingTable, 624)).toBeUndefined();
  });

  it("reserves bottom wrap distance only on the final fragment", () => {
    const floatingTable = table(3);
    floatingTable.style!.floating!.bottomFromTextTwips = 300;
    const pages = buildDocumentPageNodeSegments(document(floatingTable), 864, 624);
    expect(pages[0][8].tableRowRange?.flowAfterSpacingPx).toBe(0);
    expect(pages[1][0].tableRowRange?.flowAfterSpacingPx).toBe(20);
  });

  it("does not repeat the initial offset with table headers", () => {
    const floatingTable = table(4);
    floatingTable.rows[0].style!.isHeader = true;
    const pages = buildDocumentPageNodeSegments(document(floatingTable), 864, 624);
    const tableSegments = pages.flat().filter(segment => segment.nodeIndex === 8);
    expect(tableSegments.reduce((sum, segment) => sum + (segment.tableRowRange?.flowBeforeSpacingPx ?? 0), 0)).toBe(8);
    expect(pages[1][0].tableRowRange).toEqual({ startRowIndex: 0, endRowIndex: 1 });
  });

  it("keeps the initial offset when the first row moves to a fresh page", () => {
    const model = document(table());
    model.nodes.splice(0, 0, ...Array.from({ length: 26 }, (_, index) => paragraph(`Prefix ${index + 1}`)));
    const pages = buildDocumentPageNodeSegments(model, 864, 624);
    expect(pages[0].some(segment => segment.nodeIndex === 34)).toBe(false);
    expect(pages[1][0]).toMatchObject({ nodeIndex: 34, tableRowRange: { flowBeforeSpacingPx: 8 } });
  });

  it("still slices rows taller than a fresh page", () => {
    const floatingTable = table(1);
    floatingTable.rows[0].style!.heightTwips = 15000;
    const pages = buildDocumentPageNodeSegments(document(floatingTable), 864, 624);
    expect(pages.length).toBeGreaterThan(1);
    expect(pages.flat().some(segment => segment.tableRowSlice)).toBe(true);
    expect(pages.flat().reduce((sum, segment) => sum + (segment.tableRowRange?.flowBeforeSpacingPx ?? 0), 0)).toBe(8);
  });
});
