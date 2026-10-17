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
        return execFileSync("git", ["show", `HEAD:${file}`], { encoding: "utf8" });
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
  return page.evaluate(() => window.controller.workbook.getSheet(window.controller.activeSheet.workbookSheetIndex).drawings.map(({ name, kind, anchor }) => ({ name, kind, anchor })));
}
async function historyRoundTrip(page, before, after) {
  await page.keyboard.press("Control+z");
  await page.waitForFunction((state) => JSON.stringify(window.controller.workbook.getSheet(window.controller.activeSheet.workbookSheetIndex).drawings.map(({ name, kind, anchor }) => ({ name, kind, anchor }))) === state, JSON.stringify(before));
  await page.keyboard.press("Control+Shift+z");
  await page.waitForFunction((state) => JSON.stringify(window.controller.workbook.getSheet(window.controller.activeSheet.workbookSheetIndex).drawings.map(({ name, kind, anchor }) => ({ name, kind, anchor }))) === state, JSON.stringify(after));
}
async function exported(page) {
  const pending = page.waitForEvent("download");
  await page.evaluate(() => window.controller.exportXlsx());
  return new Uint8Array(await readFile(await (await pending).path()));
}
async function assertChartPixels(page, chart, name) {
  const rect = await chart.boundingBox();
  assert.ok(rect);
  const png = await page.screenshot({ path: `tmp/real-chart-${name}.png` });
  const pixels = await page.evaluate(async ({ bytes, rect }) => {
    const bitmap = await createImageBitmap(new Blob([new Uint8Array(bytes)], { type: "image/png" }));
    const canvas = document.createElement("canvas");
    canvas.width = bitmap.width;
    canvas.height = bitmap.height;
    const ctx = canvas.getContext("2d");
    ctx.drawImage(bitmap, 0, 0);
    const scale = bitmap.width / innerWidth;
    let black = 0, samples = 0;
    for (let y = Math.max(0, rect.y) + 5; y < Math.min(innerHeight, rect.y + rect.height) - 5; y += 8) {
      for (let x = Math.max(0, rect.x) + 5; x < Math.min(innerWidth, rect.x + rect.width) - 5; x += 8) {
        const [r, g, b] = ctx.getImageData(Math.floor(x * scale), Math.floor(y * scale), 1, 1).data;
        if (r < 8 && g < 8 && b < 8) black++;
        samples++;
      }
    }
    bitmap.close();
    return { black, samples };
  }, { bytes: [...png], rect });
  console.log(`real chart ${name}`, pixels);
  assert.ok(pixels.samples > 100 && pixels.black / pixels.samples < .05, "visible chart pixels are not black");
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
  await test("real workbook: window-scroll pixels, chart move, resize, history and reload", async () => {
    const page = await browser.newPage({ viewport: { width: 1200, height: 900 }, deviceScaleFactor: 2 });
    page.setDefaultTimeout(10000);
    const errors = [];
    page.on("pageerror", (error) => errors.push(error.message));
    try {
      await page.goto(`${server.resolvedUrls.local[0]}tests/drawings.html?real&window`);
      await page.waitForFunction(() => window.controller?.activeSheet?.name === "9. Charts" && !window.controller.isLoading && !window.controller.isChartsLoading);
      const measurements = [];
      for (const x of [0, 400, 800, 400, 0]) {
        await page.evaluate((x) => window.scrollTo(x, 0), x);
        await page.waitForTimeout(500);
        const pixels = await page.evaluate(() => {
          const canvas = [...document.querySelectorAll("canvas")].sort((a, b) => b.width * b.height - a.width * a.height)[0];
          const rect = canvas.getBoundingClientRect();
          const ctx = canvas.getContext("2d");
          let black = 0, samples = 0;
          for (let y = 200; y < 850; y += 10) for (let x = 50; x < 1150; x += 10) {
            const px = Math.floor((x - rect.left) * canvas.width / rect.width);
            const py = Math.floor((y - rect.top) * canvas.height / rect.height);
            if (px < 0 || py < 0 || px >= canvas.width || py >= canvas.height) continue;
            const [r, g, b] = ctx.getImageData(px, py, 1, 1).data;
            if (r < 8 && g < 8 && b < 8) black++;
            samples++;
          }
          return { scrollX, black, samples, fraction: black / samples, canvas: { left: rect.left, top: rect.top, width: rect.width, height: rect.height } };
        });
        console.log("real pixels", pixels);
        await page.screenshot({ path: `tmp/real-${process.env.DRAWINGS_BASELINE ? "before" : "after"}-${x}.png` });
        measurements.push(pixels);
      }
      for (const pixels of measurements.sort((a, b) => b.scrollX - a.scrollX)) {
        assert.ok(pixels.samples > 6000 && pixels.fraction < .05, `scroll ${pixels.scrollX}: ${pixels.black}/${pixels.samples} grid samples are black`);
      }
      const chart = page.locator('[data-xlsx-drawing="chart"]').first();
      await page.evaluate(() => window.scrollTo(400, 800));
      const initialBox = await chart.boundingBox();
      assert.ok(initialBox);
      await page.evaluate((y) => window.scrollTo(400, window.scrollY + y - 200), initialBox.y);
      await chart.click({ position: { x: 12, y: 12 } });
      assert.equal(await page.evaluate(() => window.controller.selection), null);
      assert.equal(await page.locator("[data-xlsx-resize]").count(), 8);
      await assertChartPixels(page, chart, "selected");
      assert.equal(await page.evaluate(() => window.controller.activeSheet.showGridLines), false);
      const original = await drawingState(page);
      const beforeBox = await chart.boundingBox();
      await drag(page, chart, 37, 27, ["Alt"]);
      const moved = await drawingState(page);
      assert.notDeepEqual(moved, original, "real chart moved");
      assert.equal(await page.evaluate(() => window.controller.activeSheet.showGridLines), false, "drawing saves retain hidden gridlines");
      await page.waitForTimeout(200);
      const movedBox = await chart.boundingBox();
      assert.ok(Math.abs(movedBox.x - beforeBox.x - 37) < 1);
      assert.ok(Math.abs(movedBox.y - beforeBox.y - 27) < 1);
      await historyRoundTrip(page, original, moved);
      await chart.click({ position: { x: 12, y: 12 } });
      await drag(page, page.locator('[data-xlsx-resize="se"]'), 44, 7, ["Shift"]);
      const resized = await drawingState(page);
      const resizedBox = await chart.boundingBox();
      assert.notDeepEqual(resized, moved, "real chart resized");
      assert.ok(resizedBox.width > movedBox.width + 30);
      assert.ok(Math.abs(resizedBox.width / resizedBox.height - movedBox.width / movedBox.height) < .02);
      await historyRoundTrip(page, moved, resized);
      const bytes = await exported(page);
      await writeFile("tmp/real-chart-edited.xlsx", bytes);
      const saved = Workbook.fromBytes(bytes);
      assert.deepEqual(saved.getSheet(9).drawings.map(({ name, kind, anchor }) => ({ name, kind, anchor })), resized);
      saved.free();
      await page.evaluate((bytes) => window.reloadWorkbook(bytes), [...bytes]);
      await page.waitForFunction(() => window.controller.workbook !== window.previousWorkbook && window.controller.activeSheet?.name === "9. Charts" && !window.controller.isLoading && !window.controller.isChartsLoading);
      assert.deepEqual(await drawingState(page), resized, "real chart anchors survive save and reload");
      assert.equal(await page.evaluate(() => window.controller.activeSheet.showGridLines), false);
      await page.evaluate(() => window.scrollTo(400, 800));
      await chart.scrollIntoViewIfNeeded();
      await chart.click({ position: { x: 12, y: 12 } });
      assert.equal(await page.locator("[data-xlsx-resize]").count(), 8);
      await assertChartPixels(page, chart, "reloaded");
      assert.deepEqual(errors, []);
    } finally { await page.close(); }
  });
  for (const query of (process.env.REAL_DRAWINGS_ONLY ? [] : ["", "dark&window&frozen", "dom&frozen", "dom&dark&window"])) {
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
          await page.waitForFunction((previous) => JSON.stringify(window.controller.workbook.getSheet(window.controller.activeSheet.workbookSheetIndex).drawings.map(({ name, kind, anchor }) => ({ name, kind, anchor }))) !== previous, JSON.stringify(before));
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
          await page.waitForFunction((state) => JSON.stringify(window.controller.workbook.getSheet(window.controller.activeSheet.workbookSheetIndex).drawings.map(({ name, kind, anchor }) => ({ name, kind, anchor }))) === state, JSON.stringify(resized));
          await page.keyboard.press("Control+Shift+z");
          await page.waitForFunction((state) => JSON.stringify(window.controller.workbook.getSheet(window.controller.activeSheet.workbookSheetIndex).drawings.map(({ name, kind, anchor }) => ({ name, kind, anchor }))) === state, JSON.stringify(nudged));
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
          await page.waitForFunction((length) => window.controller.workbook.getSheet(window.controller.activeSheet.workbookSheetIndex).drawings.length === length, before.length - 1);
          const deletedBytes = await exported(page);
          const deleted = Workbook.fromBytes(deletedBytes);
          assert.equal(deleted.getSheet(0).drawings.length, before.length - 1);
          deleted.free();
          await page.keyboard.press("Control+z");
          await page.waitForFunction((length) => window.controller.workbook.getSheet(window.controller.activeSheet.workbookSheetIndex).drawings.length === length, before.length);
          assert.deepEqual(await drawingState(page), before, "undo restores deleted drawing");
          await page.keyboard.press("Control+Shift+z");
          await page.waitForFunction((length) => window.controller.workbook.getSheet(window.controller.activeSheet.workbookSheetIndex).drawings.length === length, before.length - 1);
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
