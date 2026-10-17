import * as React from "react";
import { renderToStaticMarkup } from "react-dom/server";
import { describe, expect, it } from "vitest";
import { buildDocModel } from "../../packages/doc-model/src";
import { parseDocx } from "../../packages/ooxml-core/src";
import { DocxEditorViewer, useDocxEditor } from "../../packages/react-viewer/src/editor";
import { createZip } from "./helpers/zip";

function border(side: string, value = "single", size = 4): string {
  return `<w:${side} w:val="${value}" w:sz="${size}" w:color="000000"/>`;
}

function documentXml(mode: "cell" | "table", options: {
  emptyBorder?: "nil" | "none";
  spacing?: number;
  conflict?: boolean;
  merged?: boolean;
} = {}): string {
  const tableBorders = mode === "table"
    ? `<w:tblBorders>${["top", "right", "bottom", "left", "insideH", "insideV"].map((side) => border(side)).join("")}</w:tblBorders>`
    : "";
  const rows = [0, 1].map((rowIndex) => `<w:tr>${[0, 1].map((columnIndex) => {
    const empty = options.emptyBorder ?? "nil";
    const cellBorders = mode === "cell"
      ? `<w:tcBorders>${border("top", rowIndex === 0 ? "single" : empty)}${border("left", columnIndex === 0 ? "single" : empty)}${border("bottom")}${border("right", "single", options.conflict && columnIndex === 0 ? 8 : 4)}</w:tcBorders>`
      : "";
    const span = options.merged && rowIndex === 0 && columnIndex === 0
      ? '<w:gridSpan w:val="2"/>' : "";
    return `<w:tc><w:tcPr><w:tcW w:type="dxa" w:w="1800"/>${span}${cellBorders}</w:tcPr><w:p><w:r><w:t>Shared edge</w:t></w:r></w:p></w:tc>`;
  }).join("")}</w:tr>`).join("");
  return `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:tbl><w:tblPr><w:tblW w:type="dxa" w:w="3600"/>${tableBorders}${options.spacing ? `<w:tblCellSpacing w:type="dxa" w:w="${options.spacing}"/>` : ""}<w:tblCellMar><w:left w:w="108" w:type="dxa"/><w:right w:w="108" w:type="dxa"/><w:top w:w="0" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/></w:tblCellMar></w:tblPr><w:tblGrid><w:gridCol w:w="1800"/><w:gridCol w:w="1800"/></w:tblGrid>${rows}</w:tbl><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>`;
}

function Viewer({ model }: { model: Awaited<ReturnType<typeof buildDocModel>> }): React.JSX.Element {
  const editor = useDocxEditor({ starterModel: model });
  return React.createElement(DocxEditorViewer, {
    editor, mode: "read-only", deferInitialPaginationPaint: false,
  });
}

function css(style: string): Record<string, string> {
  return Object.fromEntries(style.split(";").filter(Boolean).map((part) => {
    const colon = part.indexOf(":");
    return [part.slice(0, colon), part.slice(colon + 1)];
  }));
}

async function render(xml: string): Promise<{ table: Record<string, string>; cells: Record<string, string>[] }> {
  const zip = createZip([
    { name: "[Content_Types].xml", content: '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>', deflate: true },
    { name: "_rels/.rels", content: '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="doc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>', deflate: true },
    { name: "word/document.xml", content: xml, deflate: true },
  ]);
  const model = await buildDocModel(await parseDocx(zip));
  const html = renderToStaticMarkup(React.createElement(Viewer, { model }));
  const table = html.match(/<table\b[^>]*style="([^"]*)"/);
  expect(table).not.toBeNull();
  return {
    table: css(table![1]),
    cells: [...html.matchAll(/<td\b[^>]*style="([^"]*)"/g)].map((match) => css(match[1])),
  };
}

const horizontalPadding = (cell: Record<string, string>) => [cell["padding-left"], cell["padding-right"]];

describe("shared cell border geometry", () => {
  it.each(["nil", "none"] as const)("keeps equivalent table and complementary cell borders at equal margin origins with %s edges", async (emptyBorder) => {
    const table = await render(documentXml("table"));
    const cells = await render(documentXml("cell", { emptyBorder }));
    expect(cells.table["border-collapse"]).toBe("collapse");
    expect(cells.table.width).toBe(table.table.width);
    expect(cells.cells).toHaveLength(4);
    expect(cells.cells.map(horizontalPadding)).toEqual(table.cells.map(horizontalPadding));
    expect(cells.cells[1]["border-left"]).toBe("none");
    for (const cell of cells.cells) {
      expect(parseFloat(cell["padding-left"])).toBeLessThan(7);
      expect(cell["padding-left"]).toBe(cell["padding-right"]);
    }
  });

  it("retains separate borders when explicit cell spacing separates the edges", async () => {
    const result = await render(documentXml("cell", { spacing: 30 }));
    expect(result.table["border-collapse"]).toBe("separate");
    expect(result.table["border-spacing"]).toBe("2px");
    expect(parseFloat(result.cells[0]["padding-left"])).toBeGreaterThan(7);
  });

  it("leaves conflicting stroke widths on the existing border path", async () => {
    const result = await render(documentXml("cell", { conflict: true }));
    expect(result.table["border-collapse"]).toBe("separate");
    expect(parseFloat(result.cells[1]["padding-left"])).toBeGreaterThan(7);
  });

  it("leaves spanning cells on the existing border path", async () => {
    const result = await render(documentXml("cell", { merged: true }));
    expect(result.table["border-collapse"]).toBe("separate");
    expect(parseFloat(result.cells[0]["padding-left"])).toBeGreaterThan(7);
  });
});
