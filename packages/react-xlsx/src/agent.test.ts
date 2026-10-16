import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { before, describe, it } from "node:test";
import initSheetsWasm, { Workbook } from "@dukelib/sheets-wasm";
import { strFromU8, unzipSync } from "fflate";
import { openXlsx } from "./agent.ts";

function sourceBytes(): Uint8Array {
  const workbook = new Workbook();
  const sheet = workbook.getSheet(0);
  sheet.setCell("A1", "Item");
  sheet.setCell("B1", "Amount");
  sheet.setCell("A2", "Rent");
  sheet.setCell("B2", 1000);
  sheet.setFormula("B3", "=B2*2");
  workbook.calculate();
  return workbook.saveXlsxBytes();
}

before(async () => {
  await initSheetsWasm({
    module_or_path: readFileSync(new URL(import.meta.resolve("@dukelib/sheets-wasm/duke_sheets_wasm_bg.wasm")))
  });
});

describe("headless agent API", () => {
  it("reads sheets and cells", () => {
    const workbook = openXlsx(sourceBytes());
    assert.deepEqual(workbook.sheets(), ["Sheet1"]);
    assert.deepEqual(workbook.cells("Sheet1", "A1:B2"), [
      { ref: "A1", value: "Item" },
      { ref: "B1", value: "Amount" },
      { ref: "A2", value: "Rent" },
      { ref: "B2", value: 1000 }
    ]);
    assert.deepEqual(workbook.cells("sheet1", "B3"), [{ formula: "B2*2", ref: "B3", value: 2000 }]);
    assert.throws(() => workbook.cells("Nope"), /No sheet named "Nope"\. Sheets: Sheet1\./);
  });

  it("comments, suggests and accepts, and the engine opens the result", () => {
    const workbook = openXlsx(sourceBytes());
    const question = workbook.comment({ author: "Agent", ref: "A2", sheet: "Sheet1", text: "Is rent monthly?" });
    const valueSuggestion = workbook.suggest({ author: "Agent", note: "Per the new lease.", ref: "B2", sheet: "Sheet1", value: 1200 });
    const newCell = workbook.suggest({ author: "Agent", ref: "D5", sheet: "Sheet1", value: "=B2+1" });
    assert.deepEqual(workbook.suggestions().map((entry) => [entry.ref, entry.previous, entry.status]), [["B2", 1000, "open"], ["D5", null, "open"]]);
    assert.deepEqual(workbook.accept("all", "Ann"), [valueSuggestion.id, newCell.id]);
    const reopened = openXlsx(workbook.save());
    assert.deepEqual(reopened.cells("Sheet1", "B2:D5"), [
      { ref: "B2", value: 1200 },
      { formula: "B2*2", ref: "B3", value: 2000 },
      { formula: "B2+1", ref: "D5", value: null }
    ]);
    assert.deepEqual(reopened.suggestions().map((entry) => entry.status), ["accepted", "accepted"]);
    assert.deepEqual(reopened.comments().find((thread) => thread.id === question.id)?.comments.map((comment) => comment.text), ["Is rent monthly?"]);
    const accepted = reopened.comments().find((thread) => thread.id === valueSuggestion.id)!;
    assert.equal(accepted.resolved, true);
    assert.deepEqual(accepted.comments.map((comment) => [comment.author, comment.text]), [
      ["Agent", "Per the new lease.\n[suggestion] {\"previous\":1000,\"value\":1200}"],
      ["Ann", "Accepted"]
    ]);

    const saved = workbook.save();
    assert.match(strFromU8(unzipSync(saved)["xl/workbook.xml"]!), /<calcPr fullCalcOnLoad="1"/);
    const engine = Workbook.fromBytes(saved);
    engine.calculate();
    assert.equal(engine.getSheet(0).getCalculatedValue("B3").toJs(), 2400);
    assert.equal(engine.getSheet(0).getCalculatedValue("D5").toJs(), 1201);
  });

  it("rejects without touching the cell and replies to threads", () => {
    const workbook = openXlsx(sourceBytes());
    const suggestion = workbook.suggest({ author: "Agent", ref: "A2", sheet: "Sheet1", value: "Lease" });
    assert.deepEqual(workbook.reject(suggestion.id), [suggestion.id]);
    assert.throws(() => workbook.accept(suggestion.id), /already rejected/);
    assert.equal(workbook.reply({ author: "Agent", id: suggestion.id, text: "Understood." }), suggestion.id);
    const reopened = openXlsx(workbook.save());
    assert.deepEqual(reopened.cells("Sheet1", "A2"), [{ ref: "A2", value: "Rent" }]);
    assert.deepEqual(reopened.comments()[0]!.comments.map((comment) => comment.text).slice(1), ["Rejected", "Understood."]);
    assert.equal(reopened.suggestions()[0]!.status, "rejected");
  });

  it("overwrites a formula cell with a value", () => {
    const bytes = sourceBytes();
    const archive = unzipSync(bytes);
    const sheet = strFromU8(archive["xl/worksheets/sheet1.xml"]!);
    assert.match(sheet, /<f>B2\*2<\/f>/);
    const workbook = openXlsx(bytes);
    workbook.suggest({ author: "Agent", ref: "B3", sheet: "Sheet1", value: 5 });
    assert.deepEqual(workbook.accept("all"), [workbook.suggestions()[0]!.threadId]);
    assert.deepEqual(openXlsx(workbook.save()).cells("Sheet1", "B3"), [{ ref: "B3", value: 5 }]);
  });

  it("locates the cell a page comment points at", () => {
    const workbook = new Workbook();
    workbook.getSheet(0).setCell("A1", "x");
    workbook.addSheet("Q1 Plan");
    workbook.getSheet(1).setCell("B2", 1200);
    const agent = openXlsx(workbook.saveXlsxBytes());
    const anchor = { page: "/book.html?sheet=Q1+Plan", selector: "main > div[id=\"wb-c-B2\"]" };
    assert.deepEqual(agent.locate(anchor), { ref: "B2", sheet: "Q1 Plan", value: 1200 });
    assert.deepEqual(agent.locate({ selector: "#wb-c-A1" }, { page: "/" }), { ref: "A1", sheet: "Sheet1", value: "x" });
    assert.deepEqual(agent.locate({ selector: "#wb-c-C9" }, { page: "/?sheet=Q1%20Plan" }), { ref: "C9", sheet: "Q1 Plan", value: null });
    assert.equal(agent.locate({ selector: "canvas" }, { page: "/?sheet=Q1+Plan" }), null);
    assert.equal(agent.locate(anchor, { page: "/?sheet=Nope" }), null);
  });
});
