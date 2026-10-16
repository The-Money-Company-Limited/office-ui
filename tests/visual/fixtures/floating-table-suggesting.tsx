import * as React from "react";
import { createRoot } from "react-dom/client";
import { buildDocModel } from "@extend-ai/react-docx-doc-model";
import {
  createMinimalDocxPackage,
  packageToArrayBuffer,
  parseDocx,
} from "@extend-ai/react-docx-ooxml-core";
import {
  DocxEditorViewer,
  useDocxEditor,
  type DocxEditorViewerMode,
} from "../../../packages/react-viewer/src/editor";

// The calibre demo's Items/Needed table and adjacent paragraph, reduced to
// their layout properties so the regression does not need an external DOCX.
const rows = [
  ["ITEM", "NEEDED"], ["Books", "1"], ["Pens", "3"],
  ["Pencils", "2"], ["Highlighter", "2 colors"], ["Scissors", "1 pair"],
];
const text = "Tables in Word can vary from the extremely simple to the extremely complex. " +
  "calibre tries to do its best when converting tables. While you may run into trouble " +
  "with the occasional table, the vast majority of common cases should be converted " +
  "very well, as demonstrated in this section. Note that for optimum results, when " +
  "creating tables in Word, you should set their widths using percentages, rather " +
  "than absolute units. To the left of this paragraph is a floating two column table " +
  "with a nice green border and header row.";

export async function mount(initialMode: DocxEditorViewerMode) {
  const seed = createMinimalDocxPackage(`
    <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
      <w:body>
        <w:p><w:r><w:t>Cover</w:t></w:r></w:p>
        <w:p><w:r><w:br w:type="page"/></w:r></w:p>
        <w:p><w:r><w:t>Tables</w:t></w:r></w:p>
        <w:tbl>
          <w:tblPr>
            <w:tblpPr w:rightFromText="187" w:bottomFromText="72" w:vertAnchor="text" w:tblpY="1"/>
            <w:tblOverlap w:val="never"/><w:tblW w:w="0" w:type="auto"/>
          </w:tblPr>
          <w:tblGrid><w:gridCol w:w="1818"/><w:gridCol w:w="1620"/></w:tblGrid>
          ${rows.map(row => `<w:tr>${row.map((value, i) => `<w:tc>
            <w:tcPr><w:tcW w:w="${i ? 1620 : 1818}" w:type="dxa"/></w:tcPr>
            <w:p><w:r><w:t>${value}</w:t></w:r></w:p>
          </w:tc>`).join("")}</w:tr>`).join("")}
        </w:tbl>
        <w:p><w:r><w:t>${text}</w:t></w:r></w:p>
        <w:p><w:r><w:t>After the floating table.</w:t></w:r></w:p>
        <w:p><w:ins w:id="7" w:author="Ada"><w:r><w:t>Existing suggestion</w:t></w:r></w:ins></w:p>
        <w:sectPr>
          <w:pgSz w:w="12240" w:h="15840"/>
          <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>
        </w:sectPr>
      </w:body>
    </w:document>`);
  const model = await buildDocModel(await parseDocx(await packageToArrayBuffer(seed)));
  const host = document.createElement("div");
  Object.assign(host.style, { position: "fixed", inset: "0", background: "white", zIndex: "10000" });
  document.body.append(host);
  function Fixture() {
    const editor = useDocxEditor({ starterModel: model });
    const [mode, setMode] = React.useState(initialMode);
    return <>
      <nav>{(["read-only", "suggest", "edit"] as const).map(value =>
        <button key={value} onClick={() => setMode(value)}>{value}</button>
      )}</nav>
      <div data-testid="float-scroll" style={{ height: 600, overflow: "auto" }}>
        <DocxEditorViewer editor={editor} mode={mode} showTrackedChanges={false} />
      </div>
    </>;
  }
  createRoot(host).render(<Fixture />);
}
