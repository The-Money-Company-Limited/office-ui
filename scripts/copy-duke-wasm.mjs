import { copyFileSync, mkdirSync } from "node:fs";
import { resolve } from "node:path";

const destination = resolve(process.argv[2]);
mkdirSync(destination, { recursive: true });
copyFileSync(new URL("../packages/react-xlsx/vendor/sheets-wasm/duke_sheets_wasm_bg.wasm", import.meta.url), resolve(destination, "duke_sheets_wasm_bg.wasm"));
