import assert from "node:assert/strict";
import { describe, it } from "node:test";
import { cellAnchorId, cellRefFromAnchor, cellRefFromSelector, listCellAnchors, sheetFromPage } from "./cell-anchors.ts";

describe("cell anchors", () => {
  it("lists one anchor per visible cell and one per merge, at its top-left cell", () => {
    const rows = [{ actualRow: 0, size: 20, start: 0 }, { actualRow: 2, size: 20, start: 20 }];
    const cols = [{ actualCol: 1, size: 64, start: 0 }, { actualCol: 2, size: 64, start: 64 }];
    // B3:C3 is merged and starts at B3; hidden row 2 is not listed.
    const mergeAnchorOf = (row: number, col: number) => (row === 2 && col >= 1 && col <= 2 ? { col: 1, row: 2 } : null);
    assert.deepEqual(listCellAnchors("wb-", rows, cols, mergeAnchorOf).map(({ id, rect }) => [id, rect]), [
      ["wb-c-B1", { height: 20, left: 0, top: 0, width: 64 }],
      ["wb-c-C1", { height: 20, left: 64, top: 0, width: 64 }],
      ["wb-c-B3", null]
    ]);
    assert.equal(cellAnchorId("x", "AA10"), "xc-AA10");
    assert.equal(cellAnchorId("", "B4"), "c-B4");
    assert.equal(cellRefFromSelector("#c-B4"), "B4");
  });

  it("reads the cell from either selector form, deepest first", () => {
    assert.equal(cellRefFromSelector("div#viewer > div[id=\"wb-c-B12\"]"), "B12");
    assert.equal(cellRefFromSelector("#wb-c-A1 span"), "A1");
    assert.equal(cellRefFromSelector("main > canvas:nth-of-type(2)"), null);
  });

  it("prefers the most-inked cell of a drawing", () => {
    assert.equal(cellRefFromAnchor({
      ink: { marks: [
        { on: [{ el: "#wb-c-A1", ink: 30 }, { el: "#wb-c-B2", ink: 25 }] },
        { on: [{ el: "#wb-c-B2", ink: 20 }, { el: "canvas", ink: 90 }] }
      ] },
      selector: "#wb-c-A1"
    }), "B2");
    assert.equal(cellRefFromAnchor({ ink: null, selector: "#wb-c-D4" }), "D4");
    assert.equal(cellRefFromAnchor({ selector: "canvas" }), null);
  });

  it("reads the sheet from the page query", () => {
    assert.equal(sheetFromPage("/book.html?sheet=Q1+Plan%20%26+more&x=1"), "Q1 Plan & more");
    assert.equal(sheetFromPage("/#/view?sheet=Data"), "Data");
    assert.equal(sheetFromPage("/book.html"), null);
    assert.equal(sheetFromPage("/?sheet=%E0%A4%A"), null);
  });
});
