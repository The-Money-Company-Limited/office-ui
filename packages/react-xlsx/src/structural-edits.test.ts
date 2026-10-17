import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { before, it } from "node:test";
import init, { Workbook } from "../vendor/sheets-wasm/duke_sheets_wasm.js";
import { unzipSync } from "fflate";
import { addCommentThread, applyCommentThreads, readCommentThreads } from "./comments.ts";
import { editWorkbookStructure } from "./structural-edits.ts";

before(async () => {
  await init({ module_or_path: readFileSync(new URL("../vendor/sheets-wasm/duke_sheets_wasm_bg.wasm", import.meta.url)) });
});

function fixture() {
  const workbook = new Workbook();
  workbook.getSheet(0).setCell("B3", 42);
  workbook.addSheet("Other");
  workbook.getSheet(1).setFormula("A1", "=Sheet1!B3");
  const { threads } = addCommentThread([], { sheet: "Sheet1", ref: "B3", text: "Review", author: "Ann" });
  const bytes = applyCommentThreads(workbook.saveXlsxBytes(), threads);
  workbook.free();
  return { bytes, threads };
}

it("inserts and deletes rows and columns through WASM, including comments and cross-sheet formulas", () => {
  const { bytes, threads } = fixture();
  const inserted = editWorkbookStructure(bytes, threads, { kind: "rows", sheetIndex: 0, index: 1, count: 2, delete: false });
  const insertedThreads = readCommentThreads(unzipSync(inserted));
  assert.equal(insertedThreads[0]?.ref, "B5");
  const shifted = Workbook.fromBytes(inserted);
  assert.equal(shifted.getSheet(0).getCell("B5").toJs(), 42);
  assert.equal(shifted.getSheet(1).getFormulaAt(0, 0)?.replace(/^=/, ""), "Sheet1!B5");
  shifted.free();
  const columns = editWorkbookStructure(inserted, insertedThreads, { kind: "columns", sheetIndex: 0, index: 1, count: 2, delete: false });
  const movedThreads = readCommentThreads(unzipSync(columns));
  assert.equal(movedThreads[0]?.ref, "D5");
  const removed = editWorkbookStructure(columns, movedThreads, { kind: "columns", sheetIndex: 0, index: 3, count: 1, delete: true });
  assert.deepEqual(readCommentThreads(unzipSync(removed)), []);
  const deleted = Workbook.fromBytes(removed);
  assert.equal(deleted.getSheet(1).getFormulaAt(0, 0)?.replace(/^=/, ""), "#REF!");
  deleted.free();
  const rowsDeleted = editWorkbookStructure(inserted, insertedThreads, { kind: "rows", sheetIndex: 0, index: 4, count: 1, delete: true });
  assert.deepEqual(readCommentThreads(unzipSync(rowsDeleted)), []);
});

it("renames and reorders loaded sheets and restores the before/after snapshots", () => {
  const { bytes, threads } = fixture();
  const renamed = editWorkbookStructure(bytes, threads, { kind: "rename", sheetIndex: 0, name: "O'Brien" });
  const renamedThreads = readCommentThreads(unzipSync(renamed));
  assert.equal(renamedThreads[0]?.sheet, "O'Brien");
  const moved = editWorkbookStructure(renamed, renamedThreads, { kind: "move", sheetIndex: 0, to: 1 });
  const after = Workbook.fromBytes(moved);
  assert.equal(after.getSheet(1).name, "O'Brien");
  assert.equal(after.getSheet(0).getFormulaAt(0, 0)?.replace(/^=/, ""), "'O''Brien'!B3");
  assert.equal(after.getSheet(1).getCell("B3").toJs(), 42);
  after.free();
  const before = Workbook.fromBytes(bytes);
  assert.equal(before.getSheet(0).name, "Sheet1");
  assert.equal(readCommentThreads(unzipSync(bytes))[0]?.sheet, "Sheet1");
  before.free();
  const deleted = editWorkbookStructure(moved, readCommentThreads(unzipSync(moved)), { kind: "deleteSheet", sheetIndex: 1 });
  const remaining = Workbook.fromBytes(deleted);
  assert.equal(remaining.sheetCount, 1);
  assert.equal(remaining.getSheet(0).getFormulaAt(0, 0)?.replace(/^=/, ""), "#REF!");
  remaining.free();
});

it("rejects invalid edits without changing snapshot bytes", () => {
  const { bytes, threads } = fixture();
  const original = bytes.slice();
  assert.throws(() => editWorkbookStructure(bytes, threads, { kind: "rows", sheetIndex: 0, index: -1, count: 1, delete: false }));
  assert.throws(() => editWorkbookStructure(bytes, threads, { kind: "rename", sheetIndex: 0, name: "Other" }));
  assert.deepEqual(bytes, original);
});
