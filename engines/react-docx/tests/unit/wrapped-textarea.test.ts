import { describe, expect, it } from "vitest";
import { applyTextareaTextEdit } from "../../packages/react-viewer/src/wrapped-textarea";

describe("applyTextareaTextEdit undo classes", () => {
  it("classifies one typed character as an insert, with the caret before it", () => {
    const edit = applyTextareaTextEdit("abc", "abXc", 3, 3, { start: 2, end: 2, inputType: "insertText" });
    expect(edit.editClass).toBe("insert");
    expect(edit.sourceCaretBefore).toBe(2);
  });

  it("gives Backspace the caret before the keystroke, after the deleted character", () => {
    const edit = applyTextareaTextEdit("abc", "ac", 1, 1, { start: 2, end: 2, inputType: "deleteContentBackward" });
    expect(edit.replacement).toEqual({ sourceStart: 1, sourceEnd: 2, inserted: "" });
    expect(edit.editClass).toBe("delete");
    expect(edit.sourceCaretBefore).toBe(2);
  });

  it("gives forward Delete the caret before the deleted character", () => {
    const edit = applyTextareaTextEdit("abc", "ac", 1, 1, { start: 1, end: 1, inputType: "deleteContentForward" });
    expect(edit.editClass).toBe("delete");
    expect(edit.sourceCaretBefore).toBe(1);
  });

  it("maps the caret through CRLF source text", () => {
    const edit = applyTextareaTextEdit("a\r\nbc", "a\nbXc", 4, 4, { start: 3, end: 3, inputType: "insertText" });
    expect(edit.replacement).toEqual({ sourceStart: 4, sourceEnd: 4, inserted: "X" });
    expect(edit.sourceCaretBefore).toBe(4);
  });

  it("does not group pastes or selection replacements", () => {
    expect(applyTextareaTextEdit("abc", "abXYc", 4, 4, { start: 2, end: 2, inputType: "insertFromPaste" }).editClass).toBeUndefined();
    expect(applyTextareaTextEdit("abc", "aXc", 2, 2, { start: 1, end: 2, inputType: "insertText" }).editClass).toBeUndefined();
  });

  it("has no caret before without a usable hint", () => {
    expect(applyTextareaTextEdit("abc", "ac", 1, 1).sourceCaretBefore).toBeUndefined();
  });
});
