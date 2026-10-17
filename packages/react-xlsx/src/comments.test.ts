import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { before, describe, it } from "node:test";
import initSheetsWasm, { Workbook } from "../vendor/sheets-wasm/duke_sheets_wasm.js";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import {
  addCommentThread,
  applyCommentThreads,
  closeSuggestion,
  formatSuggestionText,
  listSuggestions,
  parseSuggestionText,
  readCommentThreads,
  removeCommentThread,
  replyToCommentThread,
  setCommentThreadResolved,
  type XlsxCommentThread
} from "./comments.ts";

function engineWorkbookBytes(): Uint8Array {
  const workbook = new Workbook();
  workbook.getSheet(0).setCell("A1", "Revenue");
  workbook.getSheet(0).setCell("B2", 1000);
  workbook.addSheet("Data");
  workbook.getSheet(1).setCell("C3", "x");
  return workbook.saveXlsxBytes();
}

function text(bytes: Uint8Array, path: string): string {
  const entry = unzipSync(bytes)[path];
  assert.ok(entry, `missing ${path}`);
  return strFromU8(entry);
}

function threadedFixture(): { bytes: Uint8Array; threads: XlsxCommentThread[] } {
  let { thread, threads } = addCommentThread([], { author: "Ann", ref: "b2", sheet: "Sheet1", text: "Is this right?" });
  threads = replyToCommentThread(threads, thread.id, "Yes, per the invoice.", "Agent");
  threads = setCommentThreadResolved(threads, thread.id, true);
  threads = addCommentThread(threads, { author: "Ann", ref: "C3", sheet: "Data", text: "Line one\nline <two> & \"three\"" }).threads;
  return { bytes: applyCommentThreads(engineWorkbookBytes(), threads), threads };
}

before(async () => {
  await initSheetsWasm({
    module_or_path: readFileSync(new URL(import.meta.resolve("../vendor/sheets-wasm/duke_sheets_wasm_bg.wasm")))
  });
});

describe("threaded comments", () => {
  it("round-trips threads, replies, authors and the resolved state", () => {
    const { bytes, threads } = threadedFixture();
    assert.deepEqual(readCommentThreads(unzipSync(bytes)), threads);
    assert.equal(threads[0]!.ref, "B2");
    assert.equal(threads[0]!.comments[1]!.parentId, threads[0]!.id);
    assert.equal(threads[0]!.comments[0]!.personId, threads[1]!.comments[0]!.personId);
  });

  it("writes every part Excel pairs with a thread", () => {
    const { bytes, threads } = threadedFixture();
    const contentTypes = text(bytes, "[Content_Types].xml");
    assert.match(contentTypes, /PartName="\/xl\/threadedComments\/threadedComment1\.xml" ContentType="application\/vnd\.ms-excel\.threadedcomments\+xml"/);
    assert.match(contentTypes, /PartName="\/xl\/persons\/person\.xml" ContentType="application\/vnd\.ms-excel\.person\+xml"/);
    assert.match(contentTypes, /PartName="\/xl\/comments1\.xml"/);
    assert.match(contentTypes, /Default Extension="vml"/);
    assert.match(text(bytes, "xl/_rels/workbook.xml.rels"), /relationships\/person" Target="persons\/person\.xml"/);
    const sheetRels = text(bytes, "xl/worksheets/_rels/sheet1.xml.rels");
    assert.match(sheetRels, /threadedComment" Target="\.\.\/threadedComments\/threadedComment1\.xml"/);
    assert.match(sheetRels, /comments" Target="\.\.\/comments1\.xml"/);
    assert.match(sheetRels, /vmlDrawing" Target="\.\.\/drawings\/vmlDrawing1\.vml"/);
    assert.match(text(bytes, "xl/worksheets/sheet1.xml"), /<legacyDrawing r:id="rIdCmt\d"\/><\/worksheet>$/);
    const note = text(bytes, "xl/comments1.xml");
    assert.ok(note.includes(`<author>tc=${threads[0]!.id}</author>`));
    assert.match(note, /\[Threaded comment\][\s\S]*Comment:\n {4}Is this right\?\nReply:\n {4}Yes, per the invoice\./);
    assert.match(text(bytes, "xl/drawings/vmlDrawing1.vml"), /ObjectType="Note"[\s\S]*<x:Row>1<\/x:Row><x:Column>1<\/x:Column>/);
    assert.match(text(bytes, "xl/threadedComments/threadedComment1.xml"), /done="1"/);
    const vmlIds = [text(bytes, "xl/drawings/vmlDrawing1.vml"), text(bytes, "xl/drawings/vmlDrawing2.vml")]
      .map((xml) => /data="(\d+)"/.exec(xml)![1]);
    assert.notEqual(vmlIds[0], vmlIds[1]);
  });

  it("survives the engine: it reads the paired note and the post-processor restores the threads", () => {
    const { bytes, threads } = threadedFixture();
    const workbook = Workbook.fromBytes(bytes);
    assert.match(String(workbook.getSheet(0).getCommentAt(1, 1)?.text), /^\[Threaded comment\]/);
    const resaved = workbook.saveXlsxBytes();
    assert.equal(unzipSync(resaved)["xl/persons/person.xml"], undefined, "engine drops threaded parts");
    const restored = applyCommentThreads(resaved, threads);
    assert.deepEqual(readCommentThreads(unzipSync(restored)), threads);
    assert.equal(Object.keys(unzipSync(restored)).filter((path) => path.startsWith("xl/comments")).length, 2);
  });

  it("reads a legacy note as a thread and converts it on reply", () => {
    const note: XlsxCommentThread = {
      comments: [{ author: "Bob", date: "", id: "note:Sheet1!A1", personId: "", text: "Old note" }],
      id: "note:Sheet1!A1",
      legacy: true,
      ref: "A1",
      resolved: false,
      sheet: "Sheet1"
    };
    const bytes = applyCommentThreads(engineWorkbookBytes(), [note]);
    assert.equal(unzipSync(bytes)["xl/persons/person.xml"], undefined);
    const threads = readCommentThreads(unzipSync(Workbook.fromBytes(bytes).saveXlsxBytes()));
    assert.deepEqual(threads, [note]);

    const replied = replyToCommentThread(threads, note.id, "Answered", "Agent");
    assert.equal(replied[0]!.legacy, false);
    assert.match(replied[0]!.id, /^\{[0-9A-F-]{36}\}$/);
    assert.deepEqual(readCommentThreads(unzipSync(applyCommentThreads(bytes, replied))), replied);
  });

  it("removes comment parts when the last thread goes, keeping other VML shapes", () => {
    const { bytes, threads } = threadedFixture();
    const archive = unzipSync(bytes);
    archive["xl/drawings/vmlDrawing1.vml"] = strToU8(strFromU8(archive["xl/drawings/vmlDrawing1.vml"]!).replace(
      "</xml>",
      "<v:shape id=\"_x0000_s2000\" type=\"#_x0000_t201\"><x:ClientData ObjectType=\"Button\"/></v:shape></xml>"
    ));
    const withButton = zipSync(archive);
    const remaining = threads.slice(1);
    const sheetOneCleared = applyCommentThreads(withButton, removeCommentThread(threads, threads[0]!.id));
    assert.deepEqual(readCommentThreads(unzipSync(sheetOneCleared)), remaining);
    const vml = text(sheetOneCleared, "xl/drawings/vmlDrawing1.vml");
    assert.match(vml, /ObjectType="Button"/);
    assert.doesNotMatch(vml, /ObjectType="Note"/);
    assert.match(text(sheetOneCleared, "xl/worksheets/sheet1.xml"), /<legacyDrawing /);

    const cleared = unzipSync(applyCommentThreads(bytes, []));
    assert.deepEqual(readCommentThreads(cleared), []);
    assert.deepEqual(Object.keys(cleared).filter((path) => /comment|vml/i.test(path)), []);
    assert.doesNotMatch(strFromU8(cleared["xl/worksheets/sheet1.xml"]!), /legacyDrawing/);
    assert.doesNotMatch(strFromU8(cleared["[Content_Types].xml"]!), /comments|threadedcomments/);
  });
});

describe("suggestions", () => {
  it("keeps the proposal in a readable last line", () => {
    const text = formatSuggestionText(1200, 1000, "Matches the invoice.");
    assert.equal(text, "Matches the invoice.\n[suggestion] {\"previous\":1000,\"value\":1200}");
    assert.deepEqual(parseSuggestionText(text), { note: "Matches the invoice.", previous: 1000, value: 1200 });
    assert.equal(parseSuggestionText("[suggestion] not json"), null);
    assert.equal(parseSuggestionText("plain comment"), null);
  });

  it("tracks status through accept and reject", () => {
    let { thread: first, threads } = addCommentThread([], { author: "Agent", ref: "B2", sheet: "Sheet1", text: formatSuggestionText("=A1*2", null) });
    const added = addCommentThread(threads, { author: "Agent", ref: "B3", sheet: "Sheet1", text: formatSuggestionText("x", undefined) });
    threads = added.threads;
    assert.deepEqual(listSuggestions(threads).map((entry) => [entry.value, entry.status]), [["=A1*2", "open"], ["x", "open"]]);
    threads = closeSuggestion(threads, first.id, "accepted", "Ann");
    threads = closeSuggestion(threads, added.thread.id, "rejected", "Ann");
    assert.deepEqual(listSuggestions(threads).map((entry) => entry.status), ["accepted", "rejected"]);
    assert.throws(() => closeSuggestion(threads, first.id, "rejected", "Ann"), /already accepted/);
    assert.deepEqual(listSuggestions(readCommentThreads(unzipSync(applyCommentThreads(engineWorkbookBytes(), threads)))).map((entry) => entry.status), ["accepted", "rejected"]);
  });
});
