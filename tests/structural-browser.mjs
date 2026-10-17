import assert from "node:assert/strict";
import { createServer } from "node:http";
import { extname, resolve } from "node:path";
import { readFile, mkdir, writeFile } from "node:fs/promises";
import { chromium } from "@playwright/test";
import init, { Workbook } from "../packages/react-xlsx/vendor/sheets-wasm/duke_sheets_wasm.js";

await init({ module_or_path: await readFile(new URL("../packages/react-xlsx/vendor/sheets-wasm/duke_sheets_wasm_bg.wasm", import.meta.url)) });
const output = new URL("../.tmp-structural-site/", import.meta.url);
await mkdir(new URL("document/", output), { recursive: true });
const fixture = new Workbook();
fixture.getSheet(0).setCell("A1", "Structural edits");
fixture.getSheet(0).setCell("B3", 42);
fixture.getSheet(0).setCell("D8", "End");
fixture.addSheet("Other");
fixture.getSheet(1).setFormula("A1", "=Sheet1!B3");
await writeFile(new URL("document/structural.xlsx", output), fixture.saveXlsxBytes());
await writeFile(new URL("artifact.json", output), JSON.stringify({ kind: "xlsx", title: "Structural edits", file: "document/structural.xlsx" }));
fixture.free();

const server = createServer(async (request, response) => {
  const pathname = decodeURIComponent(new URL(request.url, "http://localhost").pathname);
  const root = output.pathname;
  const path = resolve(root, `.${pathname === "/" ? "/index.html" : pathname}`);
  if (!path.startsWith(root)) { response.writeHead(403).end(); return; }
  try {
    const bytes = await readFile(path);
    const mime = { ".html": "text/html", ".js": "text/javascript", ".css": "text/css", ".json": "application/json", ".wasm": "application/wasm" }[extname(path)] ?? "application/octet-stream";
    response.writeHead(200, { "Content-Type": mime }).end(bytes);
  } catch { response.writeHead(404).end(); }
});
await new Promise((resolve) => server.listen(0, "127.0.0.1", resolve));
const browser = await chromium.launch({ headless: true, args: ["--no-sandbox"] });
const page = await browser.newPage({ viewport: { width: 1440, height: 1000 } });
const errors = [];
page.on("pageerror", (error) => errors.push(error.message));
async function select(ref) {
  const box = page.getByRole("textbox", { name: "Name box", exact: true });
  await box.fill(ref);
  await box.press("Enter");
  await page.evaluate(async () => {
    window.scrollTo({ left: 0, top: 0 });
    document.querySelectorAll('[role="grid"]').forEach((grid) => { grid.scrollLeft = 0; grid.scrollTop = 0; });
    await new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve)));
  });
}
async function download() {
  await page.getByRole("button", { name: "Open workbook actions", exact: true }).click();
  const pending = page.waitForEvent("download");
  await page.getByRole("menuitem", { name: "Download", exact: true }).click();
  const result = await pending;
  return Workbook.fromBytes(new Uint8Array(await readFile(await result.path())));
}
async function expectCell(ref, expected) {
  const wb = await download();
  const actual = wb.getSheet(0).getCell(ref).toJs();
  if (actual !== expected) {
    const cells = [];
    for (let row = 0; row < 12; row++) for (let col = 0; col < 10; col++) {
      const value = wb.getSheet(0).getCellAt(row, col).toJs();
      if (value !== null) cells.push({ row, col, value });
    }
    console.error("Exported cells", cells);
  }
  assert.equal(actual, expected);
  wb.free();
}
try {
  await page.goto(process.env.STRUCTURAL_TEST_URL ?? `http://127.0.0.1:${server.address().port}`);
  await page.getByRole("tab", { name: "Sheet1", exact: true }).waitFor({ timeout: 60_000 });
  await select("B3");
  const grid = page.getByRole("grid", { name: "Sheet1 worksheet grid" });
  const rect = await grid.boundingBox();
  assert.ok(rect);
  // Header positions are measured from the rendered grid; the name box tests
  // the same selection path used for multi-row and multi-column operations.
  await page.mouse.click(rect.x + 15, rect.y + 24 + 2 * 20 + 10, { button: "right" });
  await page.getByRole("menuitem", { name: "Insert 1 row above", exact: true }).click();
  await expectCell("B4", 42);
  await page.getByRole("button", { name: "Undo", exact: true }).click();
  await expectCell("B3", 42);
  await page.getByRole("button", { name: "Redo", exact: true }).click();
  await expectCell("B4", 42);
  await select("B4:B5");
  await page.mouse.click(rect.x + 15, rect.y + 24 + 3 * 20 + 10, { button: "right" });
  await page.getByRole("menuitem", { name: "Insert 2 rows below", exact: true }).click();
  await expectCell("B4", 42);
  await select("B4");
  await page.mouse.click(rect.x + 40 + 64 + 20, rect.y + 12, { button: "right" });
  await page.getByRole("menuitem", { name: "Insert 1 column left", exact: true }).click();
  await expectCell("C4", 42);
  await select("C4:D4");
  await page.mouse.click(rect.x + 40 + 2 * 64 + 20, rect.y + 12, { button: "right" });
  await page.getByRole("menuitem", { name: "Insert 2 columns right", exact: true }).click();
  await expectCell("G11", "End");
  await select("E4:F4");
  await page.mouse.click(rect.x + 40 + 4 * 64 + 20, rect.y + 12, { button: "right" });
  await page.getByRole("menuitem", { name: "Delete 2 columns", exact: true }).click();
  await expectCell("E11", "End");
  await select("C4");
  await page.mouse.click(rect.x + 40 + 2 * 64 + 20, rect.y + 24 + 3 * 20 + 10, { button: "right" });
  await page.getByRole("menuitem", { name: "Delete row", exact: true }).click();
  const deleted = await download();
  assert.equal(deleted.getSheet(1).getFormulaAt(0, 0)?.replace(/^=/, ""), "#REF!");
  deleted.free();
  await page.getByRole("button", { name: "Undo", exact: true }).click();
  await page.getByRole("tab", { name: "Sheet1", exact: true }).dblclick();
  const rename = page.getByRole("textbox", { name: "Sheet name", exact: true });
  await rename.fill("Other");
  await rename.press("Enter");
  await page.getByRole("alert").filter({ hasText: /already exists|duplicate/i }).waitFor();
  await rename.fill("Renamed Sheet");
  await rename.press("Enter");
  await page.getByRole("tab", { name: "Renamed Sheet", exact: true }).waitFor();
  const renamed = await download();
  assert.equal(renamed.getSheet(1).getFormulaAt(0, 0)?.replace(/^=/, ""), "'Renamed Sheet'!C4");
  renamed.free();
  await page.getByRole("tab", { name: "Renamed Sheet", exact: true }).dragTo(page.getByRole("tab", { name: "Other", exact: true }));
  const reordered = await download();
  assert.equal(reordered.getSheet(1).name, "Renamed Sheet");
  reordered.free();
  await page.screenshot({ path: new URL("../.tmp-structural-site/verified.png", import.meta.url).pathname });
  assert.deepEqual(errors, []);
  console.log("PASS: header/cell menus, multi-row insertion, multi-column insertion/deletion, deletion/#REF!, undo/redo, inline rename, drag reorder, exported XLSX");
} catch (error) {
  console.error("Browser errors:", errors);
  console.error((await page.locator("body").innerText()).slice(0, 4000));
  await page.screenshot({ path: new URL("../.tmp-structural-site/failure.png", import.meta.url).pathname });
  throw error;
} finally {
  await browser.close();
  server.closeAllConnections();
  await new Promise((resolve) => server.close(resolve));
}
