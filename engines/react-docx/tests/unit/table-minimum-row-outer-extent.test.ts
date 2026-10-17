import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import type { DocModel, ParagraphNode, TableNode } from "../../packages/doc-model/src";
import { buildDocumentPageNodeSegments, estimateRenderedPageSegmentHeightPx, estimateTableRowHeightsPx } from "../../packages/react-viewer/src/editor";

beforeEach(() => {
  vi.stubGlobal("document", {
    defaultView: { devicePixelRatio: 1, getComputedStyle: () => ({ borderLeftWidth: "1px" }) },
    documentElement: { appendChild() {} },
    createElement: () => ({ style: {}, remove() {} }),
  });
});
afterEach(() => vi.unstubAllGlobals());

function paragraph(text: string): ParagraphNode {
  return { type: "paragraph", style: { spacing: { beforeTwips: 0, afterTwips: 0, lineTwips: 360, lineRule: "exact" }, widowControl: false }, children: [{ type: "text", text, style: { fontFamily: "Arial", fontSizePt: 11 } }] };
}
function table(): TableNode {
  return { type: "table", style: { widthTwips: 9360, columnWidthsTwips: [9360], layout: "fixed", cellMarginTwips: { topTwips: 0, rightTwips: 0, bottomTwips: 0, leftTwips: 0 }, borders: Object.fromEntries(["top", "left", "bottom", "right", "insideH", "insideV"].map(side => [side, { type: "single", sizeEighthPt: 4, color: "#000000" }])) }, rows: Array.from({ length: 2 }, (_, index) => ({ type: "table-row", style: { heightTwips: 3600, heightRule: "atLeast" }, cells: [{ type: "table-cell", nodes: [paragraph(`Row ${index + 1}`)] }] })) };
}
function model(node = table()): DocModel {
  return { nodes: [...Array.from({ length: 8 }, (_, index) => paragraph(`Lead ${index + 1}`)), node, ...Array.from({ length: 12 }, (_, index) => paragraph(`Following ${index + 1}`))], metadata: { sourceParts: 1, warnings: [] } };
}

describe("collapsed minimum-row outer extent", () => {
  it("moves the next paragraph when row centers fit but outer table edges do not", () => {
    const document = model();
    const pages = buildDocumentPageNodeSegments(document, 864, 624);
    expect(pages).toHaveLength(2);
    expect(pages[0].map(segment => segment.nodeIndex)).toEqual(Array.from({ length: 16 }, (_, index) => index));
    expect(pages[1].map(segment => segment.nodeIndex)).toEqual([16, 17, 18, 19, 20]);
  });
  it("includes one outer-edge pair without changing individual row heights", () => {
    const document = model();
    expect(estimateRenderedPageSegmentHeightPx(document.nodes[8], { nodeIndex: 8 }, document, 624)).toBeCloseTo(482.3333333333333, 8);
  });
  it("preserves a borderless table that fits exactly", () => {
    const node = table();
    delete node.style!.borders;
    const pages = buildDocumentPageNodeSegments(model(node), 864, 624);
    expect(pages).toHaveLength(2);
    expect(pages[0].map(segment => segment.nodeIndex)).toEqual(Array.from({ length: 17 }, (_, index) => index));
    expect(pages[1].map(segment => segment.nodeIndex)).toEqual([17, 18, 19, 20]);
  });
});

describe("collapsed outer-edge accounting boundaries", () => {
  it("adds outer edges once when pagination receives measured row rectangles", () => {
    const document = model();
    const pages = buildDocumentPageNodeSegments(document, 864, 624, undefined, undefined, {
      measuredTableRowHeightsByNodeIndex: { 8: [240.656, 240.656] },
    });
    expect(pages).toHaveLength(2);
    expect(pages[0].at(-1)?.nodeIndex).toBe(15);
    expect(pages[1][0].nodeIndex).toBe(16);
  });
  it("reserves one edge pair for each whole-row fragment", () => {
    const node = table();
    node.rows.push(structuredClone(node.rows[1]));
    node.rows.forEach(row => { row.style!.cantSplit = true; });
    const document: DocModel = { nodes: [node], metadata: { sourceParts: 1, warnings: [] } };
    const pages = buildDocumentPageNodeSegments(document, 300, 624);
    expect(pages).toHaveLength(3);
    for (const page of pages) {
      expect(page).toHaveLength(1);
      expect(page[0].tableRowSlice).toBeUndefined();
      expect(estimateRenderedPageSegmentHeightPx(node, page[0], document, 624)).toBeCloseTo(241.66666666666666, 8);
    }
  });
  it("preserves the explicit height of a clipped row slice", () => {
    const node = table();
    const document: DocModel = { nodes: [node], metadata: { sourceParts: 1, warnings: [] } };
    expect(estimateRenderedPageSegmentHeightPx(node, {
      nodeIndex: 0,
      tableRowRange: { startRowIndex: 0, endRowIndex: 1 },
      tableRowSlice: { rowIndex: 0, startOffsetPx: 24, sliceHeightPx: 120, totalRowHeightPx: 240.66666666666666 },
    }, document, 624)).toBe(120);
  });
  it("keeps explicit cell margins within the existing minimum-row metric", () => {
    const node = table();
    node.style!.cellMarginTwips = { topTwips: 90, rightTwips: 108, bottomTwips: 90, leftTwips: 108 };
    const before = structuredClone(node);
    const document = model(node);
    expect(estimateTableRowHeightsPx(node, 624)).toEqual([246.66666666666666, 246.66666666666666]);
    expect(estimateRenderedPageSegmentHeightPx(node, { nodeIndex: 8 }, document, 624)).toBeCloseTo(494.3333333333333, 8);
    expect(node).toEqual(before);
  });
  it("leaves a containing nested table on its previous outer-edge path", () => {
    const nested = table();
    delete nested.style!.borders;
    nested.rows = nested.rows.slice(0, 1);
    nested.rows[0].style!.heightTwips = 360;
    const node = table();
    node.rows[0].cells[0].nodes = [nested];
    const document = model(node);
    expect(estimateRenderedPageSegmentHeightPx(node, { nodeIndex: 8 }, document, 624)).toBeCloseTo(481.3333333333333, 8);
  });
  it.each(["header", "separate", "span", "grid-span", "mixed-border", "automatic", "exact", "empty"] as const)("preserves the previous edge path for %s tables", mode => {
    const node = table();
    if (mode === "header") node.rows[0].style!.isHeader = true;
    if (mode === "separate") node.style!.cellSpacingTwips = 30;
    if (mode === "span") node.rows[0].cells[0].style = { rowSpan: 2 };
    if (mode === "grid-span") node.rows[0].cells[0].style = { gridSpan: 2 };
    if (mode === "mixed-border") node.style!.borders!.insideH!.sizeEighthPt = 8;
    if (mode === "automatic") node.rows.forEach(row => { row.style!.heightRule = "auto"; });
    if (mode === "exact") node.rows.forEach(row => { row.style!.heightRule = "exact"; });
    if (mode === "empty") node.rows[0].cells = [];
    const document = model(node);
    const rows = estimateTableRowHeightsPx(node, 624);
    expect(estimateRenderedPageSegmentHeightPx(node, { nodeIndex: 8 }, document, 624)).toBeCloseTo(rows.reduce((sum, height) => sum + height, 0), 8);
  });

  it.each([4, 8])("preserves minimum row pitch and outer edges for equivalent %i eighth-point cell borders", sizeEighthPt => {
    const uniform = table();
    for (const border of Object.values(uniform.style!.borders!)) border!.sizeEighthPt = sizeEighthPt;
    const complementary = structuredClone(uniform);
    delete complementary.style!.borders;
    complementary.rows.forEach((row, index) => {
      row.cells[0].style = { borders: {
        top: index === 0 ? { type: "single", sizeEighthPt, color: "#000000" } : { type: "nil" },
        left: { type: "single", sizeEighthPt, color: "#000000" },
        right: { type: "single", sizeEighthPt, color: "#000000" },
        bottom: { type: "single", sizeEighthPt, color: "#000000" },
      } };
    });
    expect(estimateTableRowHeightsPx(complementary, 624)).toEqual(
      Array(2).fill(240 + sizeEighthPt / 6)
    );
    expect(estimateRenderedPageSegmentHeightPx(complementary, { nodeIndex: 8 }, model(complementary), 624))
      .toBeCloseTo(estimateRenderedPageSegmentHeightPx(uniform, { nodeIndex: 8 }, model(uniform), 624), 8);
    expect(buildDocumentPageNodeSegments(model(complementary), 864, 624))
      .toEqual(buildDocumentPageNodeSegments(model(uniform), 864, 624));
    complementary.rows.forEach(row => { row.style!.heightRule = "exact"; });
    expect(estimateTableRowHeightsPx(complementary, 624)).toEqual([240, 240]);
  });

  it("includes shared border centers in minimum rows with repeated headers", () => {
    const node = table();
    node.rows.forEach((row, index) => {
      row.style = { heightTwips: 270, heightRule: "atLeast", isHeader: index === 0 };
      row.cells[0].nodes = [paragraph("Row")];
      (row.cells[0].nodes[0] as ParagraphNode).style!.spacing!.lineTwips = 120;
      row.cells[0].style = { borders: {
        top: index === 0 ? { type: "single", sizeEighthPt: 4 } : { type: "nil" },
        left: { type: "single", sizeEighthPt: 4 },
        right: { type: "single", sizeEighthPt: 4 },
        bottom: { type: "single", sizeEighthPt: 4 },
      } };
    });
    delete node.style!.borders;
    expect(estimateTableRowHeightsPx(node, 624)).toEqual([18 + 2 / 3, 18 + 2 / 3]);
    node.rows.forEach(row => { delete row.cells[0].style!.borders; });
    expect(estimateTableRowHeightsPx(structuredClone(node), 624)).toEqual([18, 18]);
  });


});
