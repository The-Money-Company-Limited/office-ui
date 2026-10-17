// Run on Modal: node --test tests/drawings.browser.mjs (DRAWINGS_BASELINE=1 for HEAD).
import assert from "node:assert/strict";
import { test } from "node:test";
import path from "node:path";
import { execFileSync } from "node:child_process";
import { readFile, writeFile, mkdir } from "node:fs/promises";
import { createRequire } from "node:module";
import { createServer } from "../apps/playground/node_modules/vite/dist/node/index.js";
import init, { Workbook } from "../packages/react-xlsx/vendor/sheets-wasm/duke_sheets_wasm.js";
import { addPreservationChart, assertPreservedErrorBars } from "./structural-preservation.mjs";
const require = createRequire(new URL("../packages/react-xlsx/package.json", import.meta.url));
const { unzipSync, zipSync, strFromU8, strToU8 } = require("fflate");
await init({ module_or_path: await readFile(new URL("../packages/react-xlsx/vendor/sheets-wasm/duke_sheets_wasm_bg.wasm", import.meta.url)) });
await mkdir("tmp", { recursive: true });
const fixture = new Workbook();
const sheet = fixture.getSheet(0);
sheet.setCell("A1", "Keep cell data");
sheet.setCell("B3", 42);
sheet.setCell("B4", 21);
sheet.setCell("A100", "Scroll extent");
const png = [...Buffer.from("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+aWQAAAABJRU5ErkJggg==", "base64")];
sheet.addDrawing({ kind: "image", name: "One cell image", anchor: { type: "oneCell", from: { col: 1, row: 3, colOffsetEmu: 3 * 9525, rowOffsetEmu: 2 * 9525 }, widthEmu: 160 * 9525, heightEmu: 80 * 9525 }, image: { data: png, format: "png", widthEmu: 160 * 9525, heightEmu: 80 * 9525 } });
sheet.addDrawing({ kind: "image", name: "Absolute image", anchor: { type: "absolute", xEmu: 70 * 9525, yEmu: 280 * 9525, widthEmu: 160 * 9525, heightEmu: 80 * 9525 }, image: { data: png, format: "png", widthEmu: 160 * 9525, heightEmu: 80 * 9525 } });
const archive = unzipSync(addPreservationChart(fixture));
const drawingPath = Object.keys(archive).find((name) => /^xl\/drawings\/drawing\d+\.xml$/.test(name));
archive[drawingPath] = strToU8(strFromU8(archive[drawingPath]).replace('<xdr:twoCellAnchor', '<xdr:twoCellAnchor editAs="oneCell"'));
await writeFile("tmp/drawings-normal.xlsx", zipSync(archive));
archive["xl/worksheets/sheet1.xml"] = strToU8(strFromU8(archive["xl/worksheets/sheet1.xml"]).replace('<sheetFormatPr',
  '<sheetViews><sheetView workbookViewId="0"><pane xSplit="5" ySplit="12" topLeftCell="F13" activePane="bottomRight" state="frozen"/></sheetView></sheetViews><sheetFormatPr'));
await writeFile("tmp/drawings-frozen.xlsx", zipSync(archive));
fixture.free();
const root = process.cwd();
process.env.PLAYWRIGHT_BROWSERS_PATH ??= path.join(root, "tests/.browser-cache");
const { chromium } = await import("./.browser-tools/node_modules/playwright/index.mjs");
const baselineFiles = ["packages/react-xlsx/src/XlsxViewer.tsx", "packages/react-xlsx/src/controller.tsx", "packages/react-xlsx/src/images.ts", "packages/react-xlsx/src/charts.ts"];

const server = await createServer({
  configFile: false, root,
  plugins: process.env.DRAWINGS_BASELINE ? [{
    name: "drawings-baseline",
    load(id) {
      const file = path.relative(root, id);
      if (baselineFiles.includes(file)) {
        // Add selectors only: baseline interactions and persistence remain unchanged.
        return execFileSync("git", ["show", `HEAD:${file}`], { encoding: "utf8" })
          .replace('onClick={() => handleImageClick(image)}', 'data-xlsx-drawing="image" onClick={() => handleImageClick(image)}')
          .replace('onClick={() => handleChartClick(chart)}', 'data-xlsx-drawing="chart" onClick={() => handleChartClick(chart)}')
          .replaceAll('key={position}', 'key={position} data-xlsx-resize={position}');
      }
    }
  }] : [],
  resolve: { alias: {
    react: path.join(root, "apps/playground/node_modules/react"),
    "react-dom": path.join(root, "apps/playground/node_modules/react-dom"),
    "@base-ui/react": path.join(root, "apps/playground/node_modules/@base-ui/react"),
    fflate: path.join(root, "packages/react-xlsx/node_modules/fflate")
  } },
  optimizeDeps: { entries: ["tests/drawings.html"] },
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

async function drawingState(page) {
  return page.evaluate(() => window.controller.workbook.getSheet(0).drawings.map(({ name, kind, anchor }) => ({ name, kind, anchor })));
}
async function historyRoundTrip(page, before, after) {
  await page.keyboard.press("Control+z");
  await page.waitForFunction((state) => JSON.stringify(window.controller.workbook.getSheet(0).drawings.map(({ name, kind, anchor }) => ({ name, kind, anchor }))) === state, JSON.stringify(before));
  await page.keyboard.press("Control+Shift+z");
  await page.waitForFunction((state) => JSON.stringify(window.controller.workbook.getSheet(0).drawings.map(({ name, kind, anchor }) => ({ name, kind, anchor }))) === state, JSON.stringify(after));
}
async function exported(page) {
  const pending = page.waitForEvent("download");
  await page.evaluate(() => window.controller.exportXlsx());
  return new Uint8Array(await readFile(await (await pending).path()));
}
async function drag(page, locator, dx, dy, modifiers = [], scrollY = 0) {
  const box = await locator.boundingBox();
  assert.ok(box);
  for (const key of modifiers) await page.keyboard.down(key);
  await page.mouse.move(box.x + box.width / 2, box.y + box.height / 2);
  await page.mouse.down();
  if (scrollY) await page.evaluate((dy) => window.scrollBy(0, dy), scrollY);
  await page.mouse.move(box.x + box.width / 2 + dx, box.y + box.height / 2 + dy, { steps: 6 });
  await page.mouse.up();
  for (const key of modifiers) await page.keyboard.up(key);
}
try {
  for (const query of ["", "dark&window&frozen", "dom&frozen", "dom&dark&window"]) {
    await test(`drawings: ${query || "canvas light"}: move, ratio resize, nudge, delete, history and XLSX reload`, async () => {
      const page = await browser.newPage({ viewport: { width: 1200, height: 900 } });
      page.setDefaultTimeout(10000);
      const errors = [];
      page.on("pageerror", (error) => errors.push(error.message));
      try {
        await page.goto(`${server.resolvedUrls.local[0]}tests/drawings.html?${query}`);
        await page.waitForFunction(() => window.controller?.charts.length === 1 && window.controller.images.length === 2 && !window.controller.isLoading && !window.controller.isChartsLoading);
        if (query.includes("frozen")) assert.deepEqual(await page.evaluate(() => window.controller.activeSheet.freezePanes), { col: 5, row: 12 });
        if (query.includes("window")) await page.evaluate(() => window.scrollTo(0, 100));
        const original = await drawingState(page);
        for (const [kind, index] of [["chart", 0], ["image", 0], ["image", 1]]) {
          const drawing = kind === "chart" ? page.locator('[data-xlsx-drawing="chart"]')
            : page.locator(`[data-xlsx-drawing="image"][aria-label="${index === 0 ? "One cell image" : "Absolute image"}"]`);
          await drawing.click({ position: { x: 12, y: 12 } });
          assert.equal(await page.evaluate(() => window.controller.selection), null, `${kind} ${index} selection clears cells`);
          assert.equal(await page.locator("[data-xlsx-resize]").count(), 8);
          const before = await drawingState(page);
          if (kind === "chart") {
            await drag(page, drawing, 4, 4);
            assert.deepEqual(await drawingState(page), before, "drag snaps back to nearby cell edges");
          }
          const scrollY = kind === "image" && index === 1 && query.includes("window") ? 80 : 0;
          await drag(page, drawing, 37, 27, ["Alt"], scrollY);
          await page.waitForFunction((previous) => JSON.stringify(window.controller.workbook.getSheet(0).drawings.map(({ name, kind, anchor }) => ({ name, kind, anchor }))) !== previous, JSON.stringify(before));
          const moved = await drawingState(page);
          if (kind === "image" && index === 1) {
            const previousAnchor = before.find(({ anchor }) => anchor.type === "absolute").anchor;
            const nextAnchor = moved.find(({ anchor }) => anchor.type === "absolute").anchor;
            assert.equal(nextAnchor.yEmu - previousAnchor.yEmu, (27 + scrollY) * 9525, "drag follows window scrolling");
          }
          await historyRoundTrip(page, before, moved);
          await drawing.click({ position: { x: 12, y: 12 } });
          const box = await drawing.boundingBox();
          await drag(page, page.locator('[data-xlsx-resize="se"]'), 44, 7, kind === "chart" ? ["Shift"] : []);
          const resizedBox = await drawing.boundingBox();
          assert.ok(Math.abs(resizedBox.width / resizedBox.height - box.width / box.height) < .02, "resize keeps ratio");
          assert.ok(resizedBox.width > box.width + 30);
          const resized = await drawingState(page);
          assert.notDeepEqual(resized, moved);
          await historyRoundTrip(page, moved, resized);
          await drawing.click({ position: { x: 12, y: 12 } });
          await page.keyboard.press("ArrowRight");
          const nudged = await drawingState(page);
          assert.notDeepEqual(nudged, resized, "arrow nudges the drawing");
          await page.keyboard.press("Control+z");
          await page.waitForFunction((state) => JSON.stringify(window.controller.workbook.getSheet(0).drawings.map(({ name, kind, anchor }) => ({ name, kind, anchor }))) === state, JSON.stringify(resized));
          await page.keyboard.press("Control+Shift+z");
          await page.waitForFunction((state) => JSON.stringify(window.controller.workbook.getSheet(0).drawings.map(({ name, kind, anchor }) => ({ name, kind, anchor }))) === state, JSON.stringify(nudged));
          await drawing.click({ position: { x: 12, y: 12 } });
          await page.keyboard.press("Escape");
          assert.equal(await page.locator("[data-xlsx-resize]").count(), 0);
        }
        const changed = await drawingState(page);
        assert.notDeepEqual(changed, original);
        assert.deepEqual(changed.map(({ anchor }) => anchor.type), original.map(({ anchor }) => anchor.type));
        assert.equal(changed.find(({ kind }) => kind === "chart").anchor.editAs, "oneCell");
        const bytes = await exported(page);
        assertPreservedErrorBars(bytes);
        const saved = Workbook.fromBytes(bytes);
        assert.deepEqual(saved.getSheet(0).drawings.map(({ name, kind, anchor }) => ({ name, kind, anchor })), changed, "saved anchors match live engine");
        assert.equal(saved.getSheet(0).getCell("A1").toJs(), "Keep cell data");
        saved.free();
        await page.evaluate((bytes) => window.reloadWorkbook(bytes), [...bytes]);
        await page.waitForFunction(() => window.controller.workbook !== window.previousWorkbook && !window.controller.isLoading && !window.controller.isChartsLoading && window.controller.images.length === 2);
        assert.deepEqual(await drawingState(page), changed, "reloaded anchors match");
        for (const kind of ["chart", "image"]) {
          await page.locator(`[data-xlsx-drawing="${kind}"]`).first().click({ position: { x: 12, y: 12 } });
          const before = await drawingState(page);
          await page.keyboard.press("Delete");
          await page.waitForFunction((length) => window.controller.workbook.getSheet(0).drawings.length === length, before.length - 1);
          const deletedBytes = await exported(page);
          const deleted = Workbook.fromBytes(deletedBytes);
          assert.equal(deleted.getSheet(0).drawings.length, before.length - 1);
          deleted.free();
          await page.keyboard.press("Control+z");
          await page.waitForFunction((length) => window.controller.workbook.getSheet(0).drawings.length === length, before.length);
          assert.deepEqual(await drawingState(page), before, "undo restores deleted drawing");
          await page.keyboard.press("Control+Shift+z");
          await page.waitForFunction((length) => window.controller.workbook.getSheet(0).drawings.length === length, before.length - 1);
        }
        assert.deepEqual(errors, []);
        await page.screenshot({ path: `tmp/drawings-${query.replaceAll("&", "-") || "canvas"}.png` });
      } catch (error) {
        await page.screenshot({ path: `tmp/drawings-failed-${query.replaceAll("&", "-") || "canvas"}.png` });
        console.error(await page.evaluate(() => ({ selectedChart: window.controller.selectedChartId, selectedImage: window.controller.selectedImageId, selection: window.controller.selection })));
        throw error;
      } finally { await page.close(); }
    });
  }
} finally {
  await browser.close();
  await server.close();
}
