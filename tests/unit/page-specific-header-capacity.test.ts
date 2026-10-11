import { describe, expect, it } from "vitest";
import type { DocModel, HeaderSection, ParagraphNode } from "@extend-ai/react-docx-doc-model";
import { resolveDocumentSectionsFromMetadata } from "../../packages/layout-core/src";
import {
  buildDocumentPageNodeSegments,
  buildPaginationSectionMetrics,
  resolveRenderPageContentHeightPxForPageSegments,
} from "../../packages/react-viewer/src/editor";

const layout = {
  pageWidthPx: 816, pageHeightPx: 1056,
  marginsPx: { left: 96, right: 96, top: 96, bottom: 96 },
  headerDistancePx: 48, footerDistancePx: 48,
};
const geometryXml = '<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/>';

function paragraph(text: string): ParagraphNode {
  return {
    type: "paragraph",
    style: {
      spacing: { beforeTwips: 0, afterTwips: 0, lineTwips: 300, lineRule: "exact" },
      widowControl: false,
    },
    children: [{ type: "text", text, style: { fontFamily: "Arial", fontSizePt: 11 } }],
  };
}

function modelWithHeaders(types: string[], titlePage: boolean, evenAndOddHeaders = false): DocModel {
  return {
    nodes: Array.from({ length: 150 }, (_, index) => paragraph(`Paragraph ${index + 1}.`)),
    metadata: {
      sourceParts: 0, warnings: [], paragraphStyles: [], footerSections: [],
      sectionPropertiesXml: `<w:sectPr>${geometryXml}${titlePage ? "<w:titlePg/>" : ""}</w:sectPr>`,
      compatibility: { evenAndOddHeaders },
      headerSections: types.map((referenceType) => ({
        partName: `word/header-${referenceType}.xml`, referenceType,
        nodes: Array.from({ length: 10 }, (_, index) => paragraph(`Header ${index + 1}.`)),
      })),
    },
  };
}

function metricsFor(model: DocModel) {
  return buildPaginationSectionMetrics(
    resolveDocumentSectionsFromMetadata(model.metadata), layout,
    model.metadata.compatibility?.evenAndOddHeaders, model.metadata.compatibility, model
  );
}

function pagesFor(model: DocModel, measuredHeights?: number[]) {
  return buildDocumentPageNodeSegments(model, 864, 624, undefined, metricsFor(model), {
    measuredPageContentHeightsPxByPageIndex: measuredHeights,
  });
}

describe("page-specific header capacity", () => {
  it.each([
    { types: ["first"], titlePage: true, even: false, counts: [35, 43, 43, 29] },
    { types: ["first", "default"], titlePage: true, even: false, counts: [35, 35, 35, 35, 10] },
    { types: ["default"], titlePage: false, even: true, counts: [35, 43, 35, 37] },
    { types: ["default"], titlePage: true, even: false, counts: [43, 35, 35, 35, 2] },
  ])("uses only the header selected for each page", ({ types, titlePage, even, counts }) => {
    const model = modelWithHeaders(types, titlePage, even);
    expect(pagesFor(model).map((page) => page.length)).toEqual(counts);
    expect(pagesFor(model, [712, 864, 864, 864, 864]).map((page) => page.length)).toEqual(
      types.includes("first") || !titlePage ? counts : [35, 35, 35, 35, 10]
    );
  });

  it("restarts the first-page reserve at a hard section boundary and inherits its reference", () => {
    const model = modelWithHeaders(["first"], true);
    model.metadata.sections = [
      { startNodeIndex: 0, sectionPropertiesXml: model.metadata.sectionPropertiesXml, headerSections: model.metadata.headerSections, footerSections: [] },
      { startNodeIndex: 70, sectionPropertiesXml: model.metadata.sectionPropertiesXml, headerSections: [], footerSections: [] },
    ];
    expect(pagesFor(model).map((page) => page.length)).toEqual([35, 35, 35, 43, 2]);
    const metrics = metricsFor(model);
    expect(resolveRenderPageContentHeightPxForPageSegments({
      pageSegments: [{ nodeIndex: 70 }], pageIndex: 2, pageIndexWithinSection: 0,
      defaultPageContentHeightPx: 864, metricsBySection: metrics,
    })).toBe(712);
    expect(resolveRenderPageContentHeightPxForPageSegments({
      pageSegments: [{ nodeIndex: 105 }], pageIndex: 3, pageIndexWithinSection: 1,
      defaultPageContentHeightPx: 864, metricsBySection: metrics,
    })).toBe(864);
  });

  it("keeps a paragraph together when it fits the larger next page", () => {
    const model = modelWithHeaders(["first"], true);
    model.nodes = [paragraph("Lead"), {
      ...paragraph(Array.from({ length: 40 }, () => "Line").join("\n")),
      style: { ...paragraph("").style, keepLines: true },
    }];
    const pages = pagesFor(model);
    expect(pages).toEqual([[{ nodeIndex: 0 }], [{ nodeIndex: 1 }]]);
  });

  it.each(["continuous", "nextColumn"])("retains the conservative section reserve for %s section flow", (startType) => {
    const model = modelWithHeaders(["first"], true);
    model.metadata.sections = [
      { startNodeIndex: 0, sectionPropertiesXml: model.metadata.sectionPropertiesXml, headerSections: model.metadata.headerSections, footerSections: [] },
      { startNodeIndex: 70, sectionPropertiesXml: `<w:sectPr>${geometryXml}<w:type w:val="${startType}"/><w:cols w:num="2"/></w:sectPr>`, headerSections: [], footerSections: [] },
    ];
    expect(metricsFor(model).every((metrics) => metrics.pageContentHeightsPxByPageType === undefined)).toBe(true);
  });

  it.each(["header", "footer", "nested"])("retains the conservative reserve for legacy frames in %s stories", (story) => {
    const model = modelWithHeaders(["first"], true);
    const framed = {
      ...paragraph("Framed text"),
      sourceXml: '<w:p><w:pPr><w:framePr w:hAnchor="text" w:vAnchor="margin"/></w:pPr><w:r><w:t>Framed text</w:t></w:r></w:p>',
    };
    const reference: HeaderSection = {
      partName: "word/story.xml", referenceType: "first",
      nodes: [
        ...model.metadata.headerSections[0].nodes,
        story === "nested" ? {
          type: "table", rows: [{ type: "table-row", cells: [{ type: "table-cell", nodes: [framed] }] }],
        } : framed,
      ],
    };
    if (story === "footer") model.metadata.footerSections = [reference];
    else model.metadata.headerSections = [reference];
    expect(metricsFor(model)[0].pageContentHeightsPxByPageType).toBeUndefined();
  });
});
