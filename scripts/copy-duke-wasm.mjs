import { copyFileSync, mkdirSync } from "node:fs";
const distDir = new URL("../packages/react-xlsx/dist/", import.meta.url);
mkdirSync(distDir, { recursive: true });
copyFileSync(new URL("../packages/react-xlsx/vendor/sheets-wasm/duke_sheets_wasm_bg.wasm", import.meta.url), new URL("duke_sheets_wasm_bg.wasm", distDir));
