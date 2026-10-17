// On Modal: sh tests/formula-point.setup.sh && node --test tests/formula-point.browser.mjs
// Set EVERYDAY_BASELINE=1 to exercise HEAD without the fixes.
import assert from "node:assert/strict";
import { test } from "node:test";
import path from "node:path";
import { readFileSync } from "node:fs";
import { unzipSync, strFromU8 } from "../packages/react-xlsx/node_modules/fflate/esm/index.mjs";
import { execFileSync } from "node:child_process";
import { createServer } from "../apps/playground/node_modules/vite/dist/node/index.js";

const root = process.cwd();
process.env.PLAYWRIGHT_BROWSERS_PATH ??= path.join(root, "tests/.browser-cache");
const { chromium } = await import("./.browser-tools/node_modules/playwright/index.mjs");
const baselineFiles = [
  ...["XlsxViewer.tsx", "FormulaAssist.tsx", "FormulaEditing.tsx", "FormulaInput.tsx", "controller.tsx", "index.ts"].map(file => `packages/react-xlsx/src/${file}`),
  "tests/formula-point.fixture.tsx"
];

const server = await createServer({
  configFile: false, root,
  plugins: process.env.EVERYDAY_BASELINE ? [{
    name: "everyday-baseline",
    load(id) {
      const file = path.relative(root, id);
      if (baselineFiles.includes(file)) return execFileSync("git", ["show", `HEAD:${file}`], { encoding: "utf8" });
    }
  }] : [],
  resolve: { alias: {
    react: path.join(root, "apps/playground/node_modules/react"),
    "react-dom": path.join(root, "apps/playground/node_modules/react-dom"),
    "@base-ui/react": path.join(root, "apps/playground/node_modules/@base-ui/react"),
    fflate: path.join(root, "packages/react-xlsx/node_modules/fflate")
  } },
  optimizeDeps: { entries: ["tests/formula-point.html"] },
  server: { host: "127.0.0.1", port: 0, watch: { ignored: ["**/.browser*/**"], followSymlinks: false } }
});
await server.listen();
const browser = await chromium.launch({
  headless: true, args: ["--no-sandbox"],
  env: { ...process.env, LD_LIBRARY_PATH: [
    path.join(root, "tests/.browser-libs/root/usr/lib/x86_64-linux-gnu"),
    path.join(root, "tests/.browser-libs/root/lib/x86_64-linux-gnu"),
    process.env.LD_LIBRARY_PATH
  ].filter(Boolean).join(":") }
});
try {
  for (const canvas of [false, true]) for (const scenario of ["click target", "live draft", "grow and wrap", "typed values", "sheet names"]) {
    await test(`${canvas ? "canvas" : "DOM"}: ${scenario}`, async () => {
      const page = await browser.newPage({ viewport: { width: 900, height: 800 } });
      page.setDefaultTimeout(8000);
      try {
        await page.goto(`${server.resolvedUrls.local[0]}tests/formula-point.html?${canvas ? "canvas" : ""}`);
        await page.waitForFunction(() => window.controller?.activeCell && !window.controller.isLoading);
        const cell = (row, col) => page.locator(canvas ? `#test-c-${String.fromCharCode(65 + col)}${row + 1}` : `[data-xlsx-cell="${row}:${col}"]`).first();
        const editor = () => page.locator('.react-xlsx-viewer input, .react-xlsx-viewer textarea[data-xlsx-editor]').first();
        const bar = page.getByRole("textbox", { name: "Formula bar" });
        if (scenario === "click target") {
          await cell(0, 0).click();
          await page.keyboard.type("1");
          await page.keyboard.press("Enter");
          await page.keyboard.type("2");
          await page.keyboard.press("Enter");
          await page.waitForFunction(() => window.controller.activeCell.row === 2);
          await cell(0, 1).click();
          const target = await cell(0, 1).boundingBox();
          await page.keyboard.type("x");
          const box = await editor().boundingBox();
          assert.ok(Math.abs(box.x - target.x) < 3 && Math.abs(box.y - target.y) < 3, "editor must anchor to B1");
          await page.keyboard.press("Enter");
          assert.equal(await page.evaluate(() => window.controller.getCellDisplayValue({ row: 0, col: 1 })), "x");
          assert.equal(await page.evaluate(() => window.controller.getCellDisplayValue({ row: 2, col: 0 })), "");
        } else if (scenario === "live draft") {
          await bar.focus();
          await bar.press("Enter");
          assert.equal(await page.evaluate(() => window.controller.canUndo), false);
          await cell(3, 0).dblclick();
          await editor().pressSequentially("=SUM(");
          assert.equal(await bar.inputValue(), "=SUM(");
          await editor().press("Home");
          await editor().press("Shift+End");
          await bar.focus();
          assert.deepEqual(await bar.evaluate(el => [el.selectionStart, el.selectionEnd]), [0, 5]);
          assert.equal(await page.evaluate(() => window.controller.getCellFormula({ row: 3, col: 0 })), "");
          await bar.press("End");
          await bar.pressSequentially("A1:B1)");
          assert.equal(await editor().inputValue(), "=SUM(A1:B1)");
          await editor().focus();
          assert.equal(await editor().evaluate(el => el.selectionStart), 11);
          await editor().press("Enter");
          await cell(3, 0).click();
          await page.waitForFunction(() => window.controller.activeCell.row === 3);
          assert.equal(await bar.inputValue(), "=SUM(A1:B1)");
          await bar.focus();
          await bar.fill("=123");
          await editor().focus();
          await editor().press("Escape");
          assert.equal(await page.evaluate(() => window.controller.getCellDisplayValue({ row: 3, col: 0 })), "5");
        } else if (scenario === "grow and wrap") {
          await cell(3, 0).dblclick();
          const initial = await editor().boundingBox();
          await editor().fill("=SUM(A1:B1,A1:B1,A1:B1)");
          assert.ok((await editor().boundingBox()).width > initial.width + 40, "editor grows over adjacent cells");
          await editor().fill("long text ".repeat(60));
          const box = await editor().boundingBox();
          assert.ok(box.x + box.width <= 900, "editor stops at viewport edge");
          assert.ok(box.height > initial.height * 2, "editor wraps at viewport edge");
        } else if (scenario === "typed values") {
          await cell(2, 0).click();
          await page.keyboard.type("3");
          await page.keyboard.press("Enter");
          assert.equal(await page.evaluate(() => window.controller.getActiveWorksheet().getCell("A3").toJs()), 3);
          if (!canvas) assert.equal(await cell(2, 0).evaluate(el => getComputedStyle(el).textAlign), "right");
          await page.evaluate(() => window.controller.setCellStyle({ row: 3, col: 0 }, { font: { bold: true } }));
          await cell(3, 0).dblclick();
          await editor().fill("12.5%");
          await editor().press("Enter");
          assert.equal(await page.evaluate(() => window.controller.getActiveWorksheet().getCell("A4").toJs()), 0.125);
          assert.equal(await page.evaluate(() => window.controller.getCellDisplayValue({ row: 3, col: 0 })), "12.5%");
          assert.equal(await page.evaluate(() => window.controller.getActiveWorksheet().getCellStyleAt(3, 0).font.bold), true);
          await cell(3, 0).click();
          assert.equal(await bar.inputValue(), "0.125");
          await cell(4, 0).dblclick();
          await editor().fill("'003");
          await editor().press("Enter");
          await cell(4, 0).dblclick();
          assert.equal(await editor().inputValue(), "'003");
          await editor().press("Enter");
          await cell(5, 0).dblclick();
          await editor().fill("TRUE");
          await editor().press("Enter");
          const downloadPromise = page.waitForEvent("download");
          await page.evaluate(() => window.controller.exportXlsx());
          const download = await downloadPromise;
          const zip = unzipSync(new Uint8Array(readFileSync(await download.path())));
          const xml = strFromU8(zip["xl/worksheets/sheet1.xml"]);
          assert.match(xml, /<c[^>]*r="A3"[^>]*>[^<]*<v>3<\/v>/);
          assert.match(xml, /<c[^>]*r="A4"[^>]*>[^<]*<v>0.125<\/v>/);
          assert.match(xml, /<c[^>]*r="A6"[^>]*t="b"[^>]*>[^<]*<v>1<\/v>/);
          assert.equal(await page.evaluate(() => window.controller.getActiveWorksheet().getCell("A5").toJs()), "003");
        } else {
          await page.evaluate(() => window.controller.addSheet());
          await page.waitForFunction(() => window.controller.sheets.length === 2);
          await page.evaluate(() => window.controller.addSheet());
          await page.waitForFunction(() => window.controller.sheets.length === 3);
          assert.deepEqual(await page.evaluate(() => window.controller.sheets.map(s => s.name)), ["Sheet1", "Sheet2", "Sheet3"]);
        }
      } catch (error) {
        console.error(await page.evaluate(() => ({ active: window.controller.activeCell, inputs: [...document.querySelectorAll("input, textarea")].map(el => ({ value: el.value, selection: [el.selectionStart, el.selectionEnd], rect: el.getBoundingClientRect().toJSON(), active: document.activeElement === el })) })));
        throw error;
      } finally { await page.close(); }
    });
  }
} finally { await browser.close(); await server.close(); }
