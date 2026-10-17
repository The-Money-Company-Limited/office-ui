import * as React from "react";
import { renderToStaticMarkup } from "react-dom/server";
import { afterAll, beforeAll, describe, expect, it, vi } from "vitest";
import type { DocModel, ParagraphNode, TableNode } from "../../packages/doc-model/src";

class MockCanvas {
  getContext() {
    return {
      font: "16px serif",
      measureText: (text: string) => ({
        width: Array.from(text).length * 10,
        fontBoundingBoxAscent: 12,
        fontBoundingBoxDescent: 4,
        actualBoundingBoxAscent: 12,
        actualBoundingBoxDescent: 4,
      }),
    };
  }
}

beforeAll(() => {
  vi.stubGlobal("OffscreenCanvas", MockCanvas);
  vi.stubGlobal("document", {});
});

afterAll(() => {
  vi.unstubAllGlobals();
});

function paragraph(text: string, line = 300): ParagraphNode {
  return {
    type: "paragraph",
    style: {
      spacing: { beforeTwips: 0, afterTwips: 0, lineTwips: line, lineRule: "exact" },
      widowControl: false,
    },
    children: [{ type: "text", text, style: { fontFamily: "Arial", fontSizePt: 12 } }],
  };
}

const text = "甲乙丙丁戊己庚辛壬癸子丑".repeat(7);

function table(rule: "atLeast" | "exact" | undefined, columns = 9): TableNode {
  return {
    type: "table",
    blockId: "table",
    style: {
      layout: "fixed",
      widthTwips: columns * 1800,
      columnWidthsTwips: Array(columns).fill(1800),
      cellMarginTwips: { topTwips: 0, bottomTwips: 0, leftTwips: 0, rightTwips: 0 },
    },
    rows: [{
      type: "table-row",
      style: rule ? { heightTwips: 960, heightRule: rule } : undefined,
      cells: Array.from({ length: columns }, (_, index) => ({
        type: "table-cell",
        nodes: [paragraph(index ? `Cell${index}` : text)],
      })),
    }],
  };
}

describe("row minimum height preserves cell content", () => {
  it.each([1, 9])("retains all positioned text in a %i-cell minimum-height row", async (columns) => {
    const api = await import("../../packages/react-viewer/src/editor");
    const value = table("atLeast", columns);
    const intrinsic = api.estimateTableRowHeightsPx(table(undefined, columns), 1080, undefined, undefined, 200)[0];
    const minimum = api.estimateTableRowHeightsPx(value, 1080, undefined, undefined, 200)[0];
    const model: DocModel = {
      nodes: [paragraph("Lead", 2250), value],
      metadata: {
        sourceParts: 1,
        warnings: [],
        headerSections: [],
        footerSections: [],
        paragraphStyles: [],
        sectionPropertiesXml: '<w:sectPr><w:pgSz w:w="17400" w:h="4200"/><w:pgMar w:top="600" w:right="600" w:bottom="600" w:left="600"/></w:sectPr>',
      },
    };
    function Viewer(): React.JSX.Element {
      const editor = api.useDocxEditor({ starterModel: model });
      return React.createElement(api.DocxEditorViewer, {
        editor,
        mode: "read-only",
        deferInitialPaginationPaint: false,
        pageVirtualization: { enabled: false },
      });
    }
    const html = renderToStaticMarkup(React.createElement(Viewer));
    const cells = [...html.matchAll(/<td\b([^>]*)>([\s\S]*?)<\/td>/g)]
      .filter(match => match[1].includes('data-docx-cell-index="0"'))
      .map(match => ({
        text: match[2].replace(/<[^>]*>/g, ""),
        ranges: [...match[2].matchAll(/data-docx-text-start="(\d+)" data-docx-text-end="(\d+)"[^>]*>([^<]*)/g)]
          .map(fragment => ({ start: Number(fragment[1]), end: Number(fragment[2]), text: fragment[3] })),
      }));

    expect(cells.length).toBeGreaterThan(1);
    expect(cells.map(cell => cell.text).join("")).toBe(text);
    expect(minimum).toBeGreaterThanOrEqual(intrinsic);
    const ranges = cells.flatMap(cell => cell.ranges);
    expect(ranges[0].start).toBe(0);
    for (let index = 0; index < ranges.length; index++) {
      expect(ranges[index].start).toBe(index ? ranges[index - 1].end : 0);
      expect(ranges[index].text).toBe(text.slice(ranges[index].start, ranges[index].end));
    }
    expect(ranges.at(-1)?.end).toBe(text.length);
  });

  it("retains an exact row height independently of content", async () => {
    const api = await import("../../packages/react-viewer/src/editor");
    expect(api.estimateTableRowHeightsPx(table("exact"), 1080, undefined, undefined, 200)).toEqual([64]);
  });

  it.each([false, true])("retains the text height when a measured row splits with horizontalSpan=%s", async (horizontalSpan) => {
    const api = await import("../../packages/react-viewer/src/editor");
    const value = table("atLeast", 2);
    value.style!.columnWidthsTwips = horizontalSpan ? [900, 900, 1800] : [1800, 1800];
    value.rows[0].cells[0].style = { gridSpan: horizontalSpan ? 2 : 1 };
    const model: DocModel = {
      nodes: [paragraph("Lead", 2250), value],
      metadata: {
        sourceParts: 1,
        warnings: [],
        headerSections: [],
        footerSections: [],
        paragraphStyles: [],
      },
    };
    const intrinsic = api.estimateTableRowHeightsPx(value, 240, undefined, undefined, 200)[0];
    const pages = api.buildDocumentPageNodeSegments(
      model,
      200,
      240,
      undefined,
      [{ startNodeIndex: 0, pageContentHeightPx: 200, pageContentWidthPx: 240 }],
      { measuredTableRowHeightsByNodeIndex: { 1: [80] } }
    );
    const slices = pages.flat()
      .filter(segment => segment.nodeIndex === 1)
      .map(segment => segment.tableRowSlice);

    expect(slices.length).toBeGreaterThan(1);
    expect(slices.every(slice => slice && slice.totalRowHeightPx >= intrinsic)).toBe(true);
    expect(slices.reduce((sum, slice) => sum + slice!.sliceHeightPx, 0)).toBeGreaterThanOrEqual(intrinsic);
  });
});
