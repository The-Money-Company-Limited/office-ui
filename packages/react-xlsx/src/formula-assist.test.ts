import assert from "node:assert/strict";
import { describe, it } from "node:test";
import { applyFormulaFunction, applyFormulaPoint, canPointFormula, formulaReferences, formatFormulaArgument, resolveFormulaAssist } from "./formula-assist.ts";

const at = (text: string) => resolveFormulaAssist(text, text.length);

describe("formula assist", () => {
  it("suggests functions while a name is typed, shortest first", () => {
    const assist = at("=A1+su");
    assert.equal(assist?.kind, "functions");
    if (assist?.kind !== "functions") return;
    assert.equal(assist.start, 4);
    assert.equal(assist.end, 6);
    assert.ok(assist.matches.every((fn) => fn.name.startsWith("SU")));
    assert.equal(assist.matches[0].name, "SUM");
  });

  it("inserts the chosen function with an open bracket", () => {
    const assist = at("=SU");
    assert.equal(assist?.kind, "functions");
    if (assist?.kind !== "functions") return;
    const sum = assist.matches.find((fn) => fn.name === "SUM");
    assert.ok(sum);
    assert.deepEqual(applyFormulaFunction("=SU", assist, sum), { caret: 5, text: "=SUM(" });
  });

  it("shows the signature of the innermost call and the argument being typed", () => {
    const sum = at("=IF(A1>0, SUM(B1:B4, ");
    assert.equal(sum?.kind, "signature");
    if (sum?.kind !== "signature") return;
    assert.equal(sum.fn.name, "SUM");
    assert.equal(sum.argumentIndex, 1);

    const outer = at("=IF(A1>0, SUM(B1:B4), ");
    assert.equal(outer?.kind === "signature" && outer.fn.name, "IF");
    assert.equal(outer?.kind === "signature" && outer.argumentIndex, 2);
  });

  it("offers nothing outside formulas, inside text, or after the call closes", () => {
    assert.equal(at("SU"), null);
    assert.equal(at('="SU'), null);
    assert.equal(at('=LEN("a,b'), null);
    assert.equal(at("=SUM(1,2)"), null);
    assert.equal(at("=A1"), null);
  });

  it("writes optional and repeating arguments as Sheets does", () => {
    assert.equal(formatFormulaArgument({ name: "number2", optional: true, repeating: true }), "[number2, ...]");
    assert.equal(formatFormulaArgument({ name: "number1" }), "number1");
  });
});


describe("formula pointing", () => {
  it("allows operand positions, but not strings, names, values or text selections", () => {
    for (const text of ["=", "=SUM(", "=SUM(A1, ", "=A1+", "=A1>=", "=A1&"]) {
      assert.equal(canPointFormula(text, text.length, text.length, null), true, text);
    }
    for (const text of ["plain", "=SUM", "=A1", "=12", '=IF("a,', '=IF("a""+', "='Sheet+", "=SUM(A1)"]) {
      assert.equal(canPointFormula(text, text.length, text.length, null), false, text);
    }
    assert.equal(canPointFormula("=SUM()", 5, 6, null), false);
  });

  it("inserts at the caret, replaces pointing, extends and shrinks a normalized range", () => {
    const first = applyFormulaPoint("=SUM()+7", 5, 5, null, { row: 2, col: 2 });
    assert.ok(first);
    assert.equal(first.text, "=SUM(C3)+7");
    const range = applyFormulaPoint(first.text, first.end, first.end, first, { row: 0, col: 0 }, true);
    assert.ok(range);
    assert.equal(range.text, "=SUM(A1:C3)+7");
    const shrunk = applyFormulaPoint(range.text, range.end, range.end, range, { row: 2, col: 2 }, true);
    assert.ok(shrunk);
    assert.equal(shrunk.text, "=SUM(C3)+7");
    const replaced = applyFormulaPoint(range.text, range.end, range.end, range, { row: 4, col: 26 });
    assert.equal(replaced?.text, "=SUM(AA5)+7");
  });

  it("ends replacement on typing or caret movement, and starts again after an operator", () => {
    const point = applyFormulaPoint("=", 1, 1, null, { row: 0, col: 0 });
    assert.ok(point);
    assert.equal(canPointFormula("=A12", 4, 4, point), false);
    assert.equal(canPointFormula(point.text, 2, 2, point), false);
    assert.equal(canPointFormula(point.text, point.end, point.end, null), false);
    assert.equal(applyFormulaPoint("=A1+", 4, 4, null, { row: 1, col: 1 })?.text, "=A1+B2");
  });

  it("colours actual references consistently, excluding strings, functions and named identifiers", () => {
    const refs = formulaReferences(`=SUM($A$1:B3,A1,"C4""D5",LOG10(2),fooA1,'Bob''s tab'!C2,Sheet2!D4,A1)`);
    assert.deepEqual(refs.map((ref) => ref.text), ["$A$1:B3", "A1", "'Bob''s tab'!C2", "Sheet2!D4", "A1"]);
    assert.equal(refs[1].color, refs[4].color);
    assert.equal(refs[2].sheetName, "Bob's tab");
    assert.deepEqual(refs[0].range, { start: { row: 0, col: 0 }, end: { row: 2, col: 1 } });
    assert.deepEqual(formulaReferences("=XFE1+A1048577"), []);
  });
});
