// On Modal: sh tests/formula-point.setup.sh && node --test tests/formula-point.browser.mjs
// Set FORMULA_POINT_BASELINE=1 to exercise the original grid and editor hook.
import assert from "node:assert/strict";
import { test } from "node:test";
import path from "node:path";
import { execFileSync } from "node:child_process";
import { createServer } from "../apps/playground/node_modules/vite/dist/node/index.js";

const root = process.cwd();
process.env.PLAYWRIGHT_BROWSERS_PATH ??= path.join(root, "tests/.browser-cache");
const { chromium } = await import("./.browser-tools/node_modules/playwright/index.mjs");
const baselineFiles = ["packages/react-xlsx/src/XlsxViewer.tsx", "packages/react-xlsx/src/FormulaAssist.tsx"];

const server = await createServer({
  configFile: false, root,
  plugins: process.env.FORMULA_POINT_BASELINE ? [{
    name: "formula-point-baseline",
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
  for (const canvas of [false, true]) for (const dark of [false, true]) for (const bar of [false, true]) {
    await test(`${canvas ? "canvas" : "DOM"} ${dark ? "dark" : "light"} ${bar ? "formula bar" : "cell"}: point, drag, colour, commit and cancel`, async () => {
      const page = await browser.newPage({ viewport: { width: 1100, height: 800 } });
      page.setDefaultTimeout(8000);
      const errors = [];
      page.on("pageerror", (error) => errors.push(error.message));
      try {
        await page.goto(`${server.resolvedUrls.local[0]}tests/formula-point.html?${canvas ? "canvas&" : ""}${dark ? "dark" : ""}`);
        await page.waitForFunction(() => window.controller?.activeCell && !window.controller.isLoading);
        const cell = (row, col) => page.locator(canvas ? `#test-c-${String.fromCharCode(65 + col)}${row + 1}` : `[data-xlsx-cell="${row}:${col}"]`).first();
        await cell(2, 2).click();
        await page.waitForFunction(() => window.controller.activeCell?.row === 2 && window.controller.activeCell?.col === 2);
        let input;
        if (bar) {
          input = page.getByRole("textbox", { name: "Formula bar" });
          await input.click();
        } else {
          await cell(2, 2).dblclick();
          input = page.locator(".react-xlsx-viewer textarea[data-xlsx-editor]").first();
        }
        await input.pressSequentially("=SU");
        await page.getByRole("option", { name: "SUM", exact: true }).click();
        assert.equal(await input.inputValue(), "=SUM(");
        await input.press("ArrowLeft");
        assert.equal(await input.inputValue(), "=SUM(B3");
        await input.press("Shift+ArrowUp");
        assert.equal(await input.inputValue(), "=SUM(B2:B3");
        await input.press("Shift+ArrowDown");
        assert.equal(await input.inputValue(), "=SUM(B3");
        await cell(0, 0).click();
        assert.equal(await input.inputValue(), "=SUM(A1");
        const a = await cell(0, 0).boundingBox();
        const b = await cell(0, 1).boundingBox();
        await page.mouse.move(a.x + a.width / 2, a.y + a.height / 2);
        await page.mouse.down();
        await page.mouse.move(b.x + b.width / 2, b.y + b.height / 2, { steps: 5 });
        await page.mouse.up();
        assert.equal(await input.inputValue(), "=SUM(A1:B1");
        assert.equal(await input.evaluate((element) => document.activeElement === element), true);
        const outline = page.locator('[data-xlsx-formula-reference="A1:B1"]');
        await outline.waitFor();
        const box = await outline.boundingBox();
        assert.ok(Math.abs(box.x - a.x) < 2 && Math.abs(box.y - a.y) < 2);
        assert.ok(Math.abs(box.width - (b.x + b.width - a.x)) < 2);
        if (canvas && dark && bar) await page.screenshot({ path: "tests/.formula-point-dark.png" });
        const tint = input.locator("..").locator('span[aria-hidden="true"] span span').first();
        assert.equal(await tint.evaluate((element) => getComputedStyle(element).color), await outline.evaluate((element) => getComputedStyle(element).borderTopColor));
        await input.press(")");
        await input.press("ArrowLeft");
        assert.equal(await input.inputValue(), "=SUM(A1:B1)");
        assert.equal(await input.evaluate((element) => element.selectionStart), 10);
        await input.press("Enter");
        await page.waitForFunction(() => window.controller.getCellDisplayValue({ row: 2, col: 2 }) === "5");
        assert.equal(await page.evaluate(() => window.controller.getCellFormula({ row: 2, col: 2 }).replace(/^=/, "")), "SUM(A1:B1)");
        assert.equal(await page.locator("[data-xlsx-formula-reference]").count(), 0);
        await cell(2, 2).click();
        await page.waitForFunction(() => window.controller.activeCell?.row === 2 && window.controller.activeCell?.col === 2);
        if (bar) await input.click();
        else await cell(2, 2).dblclick();
        const editor = bar ? input : page.locator(".react-xlsx-viewer textarea[data-xlsx-editor]").first();
        assert.equal(await editor.inputValue(), "=SUM(A1:B1)");
        await editor.fill("=");
        await cell(0, 0).click();
        assert.equal(await editor.inputValue(), "=A1");
        await editor.press("Escape");
        assert.equal(await page.evaluate(() => window.controller.getCellDisplayValue({ row: 2, col: 2 })), "5");
        assert.deepEqual(errors, []);
      } catch (error) {
        console.error(errors, await page.evaluate(() => ({ error: String(window.controller?.error), active: window.controller?.activeCell, inputs: [...document.querySelectorAll("input")].map((input) => ({ value: input.value, caret: input.selectionStart, focused: input === document.activeElement })) })));
        throw error;
      } finally { await page.close(); }
    });
  }
} finally {
  await browser.close();
  await server.close();
}
