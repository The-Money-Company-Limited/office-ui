import * as React from "react";
import { createRoot } from "react-dom/client";
import { setWasmSource, useXlsxViewerController, XlsxViewer } from "../packages/react-xlsx/src/index";

setWasmSource("/packages/react-xlsx/vendor/sheets-wasm/duke_sheets_wasm_bg.wasm");
const params = new URLSearchParams(location.search);
const initial = await (await fetch(params.has("real") ? "/tests/fixtures/tmp-agent-demo.xlsx" : `/tmp/drawings-${params.has("frozen") ? "frozen" : "normal"}.xlsx`)).arrayBuffer();
function App() {
  const [file, setFile] = React.useState(initial);
  const controller = useXlsxViewerController({ file, readOnly: params.has("readonly") });
  React.useEffect(() => {
    if (params.has("real") && controller.workbook) controller.setActiveSheetByName("9. Charts");
  }, [controller.workbook]);
  (window as any).controller = controller;
  (window as any).reloadWorkbook = (bytes: number[]) => {
    (window as any).previousWorkbook = controller.workbook;
    setFile(new Uint8Array(bytes).buffer);
  };
  return <>
    <div style={{ height: params.has("window") ? 150 : 0 }} />
    <XlsxViewer controller={controller} isDark={params.has("dark")} experimentalCanvas={!params.has("dom")}
      scrollTarget={params.has("window") ? "window" : "element"} height={params.has("window") ? "auto" : 650}
      showDefaultToolbar={false} showImages />
  </>;
}
createRoot(document.getElementById("root")!).render(<App />);
