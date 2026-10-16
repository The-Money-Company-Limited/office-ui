import * as React from "react";
import { renderToStaticMarkup } from "react-dom/server";
import { describe, expect, it } from "vitest";
import { buildDocModel, type DocModel, type ParagraphNode } from "@extend-ai/react-docx-doc-model";
import { createMinimalDocxPackage, packageToArrayBuffer, parseDocx } from "@extend-ai/react-docx-ooxml-core";
import { useDocxEditor, type DocxEditorController } from "../../packages/react-viewer/src/editor";

async function load(body: string): Promise<DocModel> {
  return buildDocModel(await parseDocx(await packageToArrayBuffer(createMinimalDocxPackage(
    `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${body}</w:body></w:document>`
  ))));
}
const p = (text: string) => `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
function run(model: DocModel, steps: Array<(editor: DocxEditorController) => void> = []) {
  let index = 0;
  let last!: DocxEditorController;
  renderToStaticMarkup(React.createElement(function Probe() {
    const editor = useDocxEditor({ starterModel: model, author: "Ada" });
    last = editor;
    steps[index++]?.(editor);
    return null;
  }));
  expect(index).toBeGreaterThanOrEqual(steps.length);
  return last;
}
function text(editor: DocxEditorController, index = 0) {
  return (editor.model.nodes[index] as ParagraphNode).children.map(child => child.type === "text" ? child.text : "").join("");
}

describe("controller find and replace", () => {
  it.each([false, true])("deletes the exact match across repeated styled runs (suggesting: %s)", async suggesting => {
    const model = await load('<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>a</w:t></w:r><w:r><w:t>aa</w:t></w:r></w:p>');
    const editor = run(model, [editor => {
      if (suggesting) editor.registerSuggestingMode();
      expect(editor.replaceText([editor.findText("aa")[0]!], "")).toBe(1);
    }]);
    const paragraph = editor.model.nodes[0] as ParagraphNode;
    expect(text(editor)).toBe("a");
    expect(paragraph.children.filter(child => child.type === "text" && child.text).map(child => child.style?.bold)).toEqual([undefined]);
    if (suggesting) {
      expect(editor.trackedChanges.filter(change => change.kind === "deletion").map(change => change.text).join("")).toBe("aa");
      expect(paragraph.sourceXml).toMatch(/<w:del\b[\s\S]*<w:b\s*\//);
    }
  });

  it("suggests replacements in table cells and body paragraphs in one undo step", async () => {
    const model = await load(p("cat") + '<w:tbl><w:tr><w:tc>' + p("cat cat") + p("cat") + '</w:tc><w:tc>' + p("cat") + '</w:tc></w:tr></w:tbl>');
    const editor = run(model, [editor => {
      editor.registerSuggestingMode();
      expect(editor.replaceText(editor.findText("cat"), "dog")).toBe(5);
    }, editor => {
      expect(editor.findText("dog")).toHaveLength(5);
      expect(editor.findText("cat")).toHaveLength(0);
      expect(editor.trackedChanges.filter(change => change.kind === "insertion")).toHaveLength(5);
      expect(editor.trackedChanges.filter(change => change.kind === "deletion")).toHaveLength(5);
      expect(new Set(editor.trackedChanges.map(change => change.revisionId)).size).toBe(10);
      editor.undo();
    }, editor => {
      expect(editor.findText("cat")).toHaveLength(5);
      expect(editor.canUndo).toBe(false);
      editor.redo();
    }]);
    expect(editor.findText("dog")).toHaveLength(5);
  });

  it.each([false, true])("counts only changed matches beside protected fields (suggesting: %s)", async suggesting => {
    const model = await load(p("cat") + '<w:p><w:fldSimple w:instr="REF name"><w:r><w:t>cat</w:t></w:r></w:fldSimple></w:p>');
    const editor = run(model, [editor => {
      if (suggesting) editor.registerSuggestingMode();
      expect(editor.findText("cat")).toHaveLength(2);
      expect(editor.replaceText(editor.findText("cat"), "dog")).toBe(1);
    }, editor => {
      expect(editor.status).toBe("Replaced 1 match");
      expect(editor.replaceText(editor.findText("cat"), "dog")).toBe(0);
      editor.undo();
    }]);
    expect(editor.findText("cat")).toHaveLength(2);
    expect(editor.canUndo).toBe(false);
    expect((editor.model.nodes[1] as ParagraphNode).sourceXml).toBe((model.nodes[1] as ParagraphNode).sourceXml);
  });

  it("finds across styled runs and every model paragraph without a mounted viewer", async () => {
    const editor = run(await load('<w:p><w:r><w:t>He</w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>llo HELLO</w:t></w:r></w:p>' + p("hello")));
    expect(editor.findText("hello").map(range => [range.start.location, range.start.offset, range.end.offset])).toEqual([
      [{ kind: "paragraph", nodeIndex: 0 }, 0, 5],
      [{ kind: "paragraph", nodeIndex: 0 }, 6, 11],
      [{ kind: "paragraph", nodeIndex: 1 }, 0, 5],
    ]);
    expect(editor.findText("hello", { matchCase: true })).toHaveLength(1);
  });

  it("treats the query literally, keeps UTF-16 offsets, and returns nonoverlapping matches", async () => {
    const editor = run(await load(p("😀 İ a.a AAAA")));
    expect(editor.findText("a.a")[0]?.start.offset).toBe(5);
    expect(editor.findText("AA").map(range => range.start.offset)).toEqual([9, 11]);
    expect(editor.findText("")).toEqual([]);
    expect(editor.findText("missing")).toEqual([]);
    expect(editor.findText(" ")).toHaveLength(3);
  });

  it("finds table cell paragraphs in document order", async () => {
    const editor = run(await load(p("hit") + '<w:tbl><w:tr><w:tc>' + p("hit") + p("hit") + '</w:tc><w:tc>' + p("hit") + '</w:tc></w:tr></w:tbl>' + p("hit")));
    expect(editor.findText("hit").map(range => range.start.location)).toEqual([
      { kind: "paragraph", nodeIndex: 0 },
      { kind: "table-cell", tableIndex: 1, rowIndex: 0, cellIndex: 0, paragraphIndex: 0 },
      { kind: "table-cell", tableIndex: 1, rowIndex: 0, cellIndex: 0, paragraphIndex: 1 },
      { kind: "table-cell", tableIndex: 1, rowIndex: 0, cellIndex: 1, paragraphIndex: 0 },
      { kind: "paragraph", nodeIndex: 2 },
    ]);
  });

  it("replaces one match and preserves paragraph formatting even when it is the whole document", async () => {
    const editor = run(await load('<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t>hello</w:t></w:r></w:p>'), [editor => {
      editor.setActiveTextRange(editor.findText("hello")[0]);
      expect(editor.replaceText(editor.findText("hello"), "world")).toBe(1);
    }]);
    expect(text(editor)).toBe("world");
    expect(editor.historyRestoreRequest).toBeUndefined();
    expect((editor.model.nodes[0] as ParagraphNode).style?.styleId).toBe("Heading1");
    expect((editor.model.nodes[0] as ParagraphNode).children[0]).toMatchObject({ style: { bold: true } });
    expect(editor.canUndo).toBe(true);
  });

  it("replaces all original matches once, including tables, with one undo step and redo", async () => {
    const model = await load(p("cat cat") + '<w:tbl><w:tr><w:tc>' + p("cat") + '</w:tc></w:tr></w:tbl>');
    const editor = run(model, [editor => {
      expect(editor.replaceText(editor.findText("cat"), "catcat")).toBe(3);
    }, editor => {
      expect(editor.findText("cat")).toHaveLength(6);
      editor.undo();
    }, editor => {
      expect(editor.findText("cat")).toHaveLength(3);
      expect(editor.canUndo).toBe(false);
      editor.redo();
    }]);
    expect(editor.findText("cat")).toHaveLength(6);
  });

  it("allows an empty replacement and does not add history for an unchanged replacement", async () => {
    const editor = run(await load(p("cat cat")), [editor => {
      expect(editor.replaceText(editor.findText("cat"), "cat")).toBe(0);
      expect(editor.replaceText(editor.findText("cat"), "")).toBe(2);
    }, editor => {
      expect(text(editor)).toBe(" ");
      editor.undo();
    }]);
    expect(text(editor)).toBe("cat cat");
    expect(editor.canUndo).toBe(false);
  });

  it("tracks each replacement in suggesting and rejecting restores the original text", async () => {
    const editor = run(await load(p("cat and cat") + p("cat")), [editor => {
      editor.registerSuggestingMode();
      expect(editor.replaceText(editor.findText("cat"), "dog")).toBe(3);
    }, editor => {
      expect(text(editor)).toBe("dog and dog");
      expect(editor.trackedChanges.filter(change => change.kind === "insertion")).toHaveLength(3);
      expect(editor.trackedChanges.every(change => change.author === "Ada")).toBe(true);
      expect(editor.findText("cat")).toEqual([]);
      expect(editor.rejectAllRevisions().ok).toBe(true);
    }]);
    expect(text(editor)).toBe("cat and cat");
    expect(text(editor, 1)).toBe("cat");
  });

  it("rejects overlapping or invalid ranges without a partial edit", async () => {
    const editor = run(await load(p("cat cat")), [editor => {
      const matches = editor.findText("cat");
      expect(editor.replaceText([matches[0]!, matches[0]!], "dog")).toBe(0);
      expect(editor.replaceText([...matches, { start: { location: { kind: "paragraph", nodeIndex: 8 }, offset: 0 }, end: { location: { kind: "paragraph", nodeIndex: 8 }, offset: 3 } }], "dog")).toBe(0);
    }]);
    expect(text(editor)).toBe("cat cat");
    expect(editor.canUndo).toBe(false);
  });
});
