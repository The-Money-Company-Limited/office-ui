import assert from "node:assert/strict";
import { test } from "node:test";
import { coerceUserEnteredValue } from "./user-entry.ts";

test("percentages are numeric fractions, including decimals and exponents", () => {
  assert.equal(coerceUserEnteredValue("12.5%"), 0.125);
  assert.equal(coerceUserEnteredValue(" -2% "), -0.02);
  assert.equal(coerceUserEnteredValue("1e2%"), 1);
  assert.equal(coerceUserEnteredValue("0%"), 0);
});

test("grouped and scientific numbers are numeric; nonnumbers remain text", () => {
  for (const [input, expected] of [["1,234.5", 1234.5], ["-2e3", -2000], ["+.5", 0.5], ["1,2", "1,2"], ["1e999", "1e999"], ["TRUE", true], ["false", false], ["  text  ", "  text  "]]) {
    assert.equal(coerceUserEnteredValue(input as string), expected);
  }
});

test("an apostrophe forces text and preserves its whitespace", () => {
  assert.equal(coerceUserEnteredValue("'12.5%"), "12.5%");
  assert.equal(coerceUserEnteredValue("'TRUE"), "TRUE");
  assert.equal(coerceUserEnteredValue("'001 "), "001 ");
  assert.equal(coerceUserEnteredValue(" '001 "), " '001 ");
  assert.equal(coerceUserEnteredValue("  "), "  ");
  assert.equal(coerceUserEnteredValue("'=SUM(A1:A3)"), "=SUM(A1:A3)");
});
