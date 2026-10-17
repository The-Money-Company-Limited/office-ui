import { describe, expect, it } from "vitest";
import {
  createMinimalDocxPackage,
  packageToArrayBuffer,
  parseDocx,
} from "@extend-ai/react-docx-ooxml-core";
import { buildDocModel } from "@extend-ai/react-docx-doc-model";
import { buildDocumentPageNodeSegments } from "../../packages/react-viewer/src/editor";

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

describe("blank paragraph before a page-break-before paragraph", () => {
  it("stays on the earlier page, so it can be drawn and edited", async () => {
    const seed = createMinimalDocxPackage(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="${W_NS}"><w:body>` +
        `<w:p><w:r><w:t>One</w:t></w:r></w:p>` +
        `<w:p/>` +
        `<w:p><w:pPr><w:pageBreakBefore/></w:pPr><w:r><w:t>Two</w:t></w:r></w:p>` +
        `</w:body></w:document>`
    );
    const model = await buildDocModel(await parseDocx(await packageToArrayBuffer(seed)));

    const pages = buildDocumentPageNodeSegments(model, 900, 600);

    expect(pages.map((page) => page.map((segment) => segment.nodeIndex))).toEqual([[0, 1], [2]]);
  });
});
