import * as React from "react";
import { Input as BaseInput } from "@base-ui/react/input";
import { createRoot } from "react-dom/client";
import { zipSync, strToU8 } from "fflate";
import {
  useXlsxCellEditor, setWasmSource, useXlsxViewerController, useXlsxFormulaAssist,
  XlsxFormulaInput, XlsxFormulaAssistPopup, XlsxViewer, XlsxViewerProvider, useXlsxViewerPalette
} from "../packages/react-xlsx/src/index";

setWasmSource("/packages/react-xlsx/dist/duke_sheets_wasm_bg.wasm");
const xml = {
  "[Content_Types].xml": '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>',
  "_rels/.rels": '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>',
  "xl/workbook.xml": '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>',
  "xl/_rels/workbook.xml.rels": '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>',
  "xl/worksheets/sheet1.xml": '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><v>2</v></c><c r="B1"><v>3</v></c></row></sheetData></worksheet>'
};
const file = zipSync(Object.fromEntries(Object.entries(xml).map(([name, text]) => [name, strToU8(text)]))).buffer;
const params = new URLSearchParams(location.search);
const isDark = params.has("dark");
function FormulaBar({ controller }: { controller: ReturnType<typeof useXlsxViewerController> }) {
  const editor = useXlsxCellEditor();
  const ref = React.useRef<HTMLInputElement>(null);
  const raw = controller.getCellInputValue(controller.activeCell);
  const draft = editor.draft?.value ?? raw;
  const assist = useXlsxFormulaAssist(draft, editor.change, ref);
  const palette = useXlsxViewerPalette();
  return <div style={{ position: "relative" }}>
    <BaseInput render={<XlsxFormulaInput />} aria-label="Formula bar" ref={ref} value={draft}
      style={{ height: 32, font: "14px monospace", color: isDark ? "white" : "black", background: isDark ? "#222" : "white" }}
      onFocus={() => { if (!editor.draft && controller.activeCell) editor.begin(controller.activeCell, raw); }}
      onChange={(event) => editor.change(event.target.value)} onSelect={assist.onSelect}
      onBlur={(event) => { assist.onBlur(); editor.onBlur(event); }}
      onKeyDown={(event) => {
        if (assist.handleKeyDown(event)) return;
        if (event.key === "Enter") { event.preventDefault(); editor.commit(); }
        if (event.key === "Escape") { event.preventDefault(); editor.cancel(); }
      }} />
    <XlsxFormulaAssistPopup state={assist} palette={palette} />
  </div>;
}
function App() {
  const controller = useXlsxViewerController({ file, readOnly: false });
  (window as any).controller = controller;
  return <XlsxViewerProvider controller={controller} isDark={isDark}>
    <FormulaBar controller={controller} />
    <XlsxViewer cellAnchorIdPrefix="test-" controller={controller} isDark={isDark} experimentalCanvas={params.has("canvas")} height={600} showDefaultToolbar={false} />
  </XlsxViewerProvider>;
}
createRoot(document.getElementById("root")!).render(<App />);
