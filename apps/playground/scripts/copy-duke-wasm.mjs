import { copyFileSync, mkdirSync } from "node:fs";
const publicDir = new URL("../public/", import.meta.url);
mkdirSync(publicDir, { recursive: true });
copyFileSync(new URL("../../../packages/react-xlsx/vendor/sheets-wasm/duke_sheets_wasm_bg.wasm", import.meta.url), new URL("duke_sheets_wasm_bg.wasm", publicDir));
