import assert from "node:assert/strict";
import { describe, it } from "node:test";
import { applyFormulaFunction, formatFormulaArgument, resolveFormulaAssist } from "./formula-assist.ts";

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
