import * as React from "react";
import { renderToStaticMarkup } from "react-dom/server";
import { describe, expect, it, vi } from "vitest";
import {
  buildDocModel,
  type DocModel,
  type ParagraphNode,
  type TableNode,
} from "@extend-ai/react-docx-doc-model";
import {
  createMinimalDocxPackage,
  packageToArrayBuffer,
  parseDocx,
} from "@extend-ai/react-docx-ooxml-core";
import { serializeDocx } from "@extend-ai/react-docx-serializer";
import {
  useDocxEditor,
  DocxEditorViewer,
  buildParagraphNumberingLabels,
  type DocxEditorController,
  type DocxTextRangeLocation,
} from "../../packages/react-viewer/src/editor";
import {
  applyTextEditingIntent,
  editingIntentCoalesceClass,
  resolveEditingIntent,
} from "../../packages/react-viewer/src/editing-intents";

async function load(inCell = false, cellPrefix = ""): Promise<DocModel> {
  const paragraph = "<w:p/>";
  const body = inCell
    ? `<w:tbl><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc>${cellPrefix}${paragraph}</w:tc></w:tr></w:tbl>`
    : paragraph;
  const pkg = createMinimalDocxPackage(
    `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${body}</w:body></w:document>`
  );
  return buildDocModel(await parseDocx(await packageToArrayBuffer(pkg)));
}

function runControllerSteps(
  model: DocModel,
  steps: Array<(editor: DocxEditorController, view?: React.ReactNode) => void>,
  keyboard = false
): DocxEditorController {
  let next = 0;
  let last: DocxEditorController | undefined;
  renderToStaticMarkup(
    React.createElement(function Probe(): null {
      const editor = useDocxEditor({ starterModel: model });
      last = editor;
      const view = keyboard ? DocxEditorViewer({
        editor, deferInitialPaginationPaint: false, pageVirtualization: { enabled: false },
      }) : undefined;
      const step = steps[next];
      if (step) {
        next += 1;
        step(editor, view);
      }
      return null;
    })
  );
  expect(next).toBe(steps.length);
  return last!;
}

const bodyLocation: DocxTextRangeLocation = { kind: "paragraph", nodeIndex: 0 };
const cellLocation: DocxTextRangeLocation = {
  kind: "table-cell", tableIndex: 0, rowIndex: 0, cellIndex: 0, paragraphIndex: 0,
};
const paths = [
  { name: "body beforeinput", inCell: false, commit: (e: DocxEditorController, text: string) =>
    e.commitParagraphTextAtRange(bodyLocation, text, { start: text.length, end: text.length },
      { insertionOffset: text.length - 1, coalesceKey: "body:typing" }) },
  { name: "wrapped body", inCell: false, commit: (e: DocxEditorController, text: string) =>
    e.commitParagraphText(0, text) },
  { name: "cell beforeinput", inCell: true, commit: (e: DocxEditorController, text: string) =>
    e.commitParagraphTextAtRange(cellLocation, text, { start: text.length, end: text.length }) },
  { name: "wrapped cell", inCell: true, commit: (e: DocxEditorController, text: string) =>
    e.commitTableCellParagraphTextRecursive(0, 0, 0, 0, text) },
  { name: "cell draft", inCell: true, commit: (e: DocxEditorController, text: string) =>
    e.commitTableCellText(0, 0, 0, text) },
  { name: "cell range", inCell: true, commit: (e: DocxEditorController, text: string) =>
    e.commitTableCellTextAtRange(0, 0, 0, text, { paragraphIndex: 0, start: text.length, end: text.length }) },
];
function paragraph(editor: DocxEditorController, inCell: boolean, index = 0): ParagraphNode {
  return (inCell
    ? (editor.model.nodes[0] as TableNode).rows[0]!.cells[0]!.nodes[index]
    : editor.model.nodes[index]) as ParagraphNode;
}
function textOf(node: ParagraphNode): string {
  return node.children.map((child) => child.type === "text" ? child.text : "").join("");
}
function expectList(editor: DocxEditorController, inCell: boolean, format: string): void {
  const node = paragraph(editor, inCell);
  expect(textOf(node)).toBe("");
  expect(node.style?.numbering?.numId).toBeGreaterThan(0);
  const definitions = editor.model.metadata.numberingDefinitions!;
  const instance = definitions.instances.find((item) => item.numId === node.style?.numbering?.numId)!;
  expect(definitions.abstracts.find((item) => item.abstractNumId === instance.abstractNumId)?.levels[0]?.format).toBe(format);
  expect(editor.activeTextRange?.start.offset).toBe(0);
}

for (const path of paths) {
  describe(path.name, () => {
    it.each(["- ", "* ", "1. ", "1) "])("converts %j and undoes only the conversion", async (prefix) => {
      const model = await load(path.inCell);
      const editor = runControllerSteps(model, [
        ...Array.from(prefix, (_, index) => (e: DocxEditorController) => path.commit(e, prefix.slice(0, index + 1))),
        (e) => {
          expectList(e, path.inCell, prefix.startsWith("1") ? "decimal" : "bullet");
          e.undo();
        },
        (e) => {
          expect(textOf(paragraph(e, path.inCell))).toBe(prefix);
          expect(paragraph(e, path.inCell).style?.numbering).toBeUndefined();
          expect(e.activeTextRange?.start.offset).toBe(prefix.length);
          e.redo();
        },
      ]);
      expectList(editor, path.inCell, prefix.startsWith("1") ? "decimal" : "bullet");
      const reopened = await buildDocModel(await parseDocx(await serializeDocx(editor.model)));
      const saved = paragraph({ ...editor, model: reopened }, path.inCell);
      expect(textOf(saved)).toBe("");
      expect(saved.style?.numbering).toEqual(paragraph(editor, path.inCell).style?.numbering);
    });

    it("indents, outdents, continues and exits the list", async () => {
      const editor = runControllerSteps(await load(path.inCell), [
        (e) => path.commit(e, "-"),
        (e) => path.commit(e, "- "),
        (e) => { expect(e.adjustSelectedListDepth(1)).toBe(true); },
        (e) => {
          expect(paragraph(e, path.inCell).style?.numbering?.ilvl).toBe(1);
          expect(e.adjustSelectedListDepth(-1)).toBe(true);
        },
        (e) => {
          expect(paragraph(e, path.inCell).style?.numbering?.ilvl).toBe(0);
          path.commit(e, "Item");
        },
        (e) => e.undo(),
        (e) => {
          expectList(e, path.inCell, "bullet");
          e.redo();
        },
        (e) => {
          expect(e.insertListItemAfterSelection("Item", 4, 4, path.inCell ? cellLocation : bodyLocation))
            .toEqual({ paragraphIndex: 1, caretOffset: 0 });
        },
        (e) => {
          expect(paragraph(e, path.inCell, 1).style?.numbering?.numId).toBeGreaterThan(0);
          expect(e.insertListItemAfterSelection("", 0)).toEqual({ paragraphIndex: 1, caretOffset: 0 });
        },
      ]);
      expect(paragraph(editor, path.inCell, 1).style?.numbering).toBeUndefined();
      expect(textOf(paragraph(editor, path.inCell, 1))).toBe("");
      expect(path.inCell ? (editor.model.nodes[0] as TableNode).rows[0]!.cells[0]!.nodes.length : editor.model.nodes.length).toBe(2);
    });
  });
}

it.each(["unordered", "ordered"] as const)("Enter exits an empty toolbar %s list", async (type) => {
  const editor = runControllerSteps(await load(), [
    (e) => e.toggleList(type),
    (e) => {
      const text = textOf(paragraph(e, false));
      expect(e.insertListItemAfterSelection(text, text.length)).toEqual({ paragraphIndex: 0, caretOffset: 0 });
    },
  ]);
  expect(textOf(paragraph(editor, false))).toBe("");
  expect(editor.model.nodes).toHaveLength(1);
});

it.each([" - ", "hello - ", "2. ", "1. pasted"])("keeps ordinary text %j", async (text) => {
  const editor = runControllerSteps(await load(), [
    (e) => e.commitParagraphText(0, text.slice(0, -1)),
    (e) => e.commitParagraphText(0, text),
  ]);
  expect(textOf(paragraph(editor, false))).toBe(text);
  expect(paragraph(editor, false).style?.numbering).toBeUndefined();
});

function mixedCellModel(): Promise<DocModel> {
  return load(true, `<w:tbl><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid>
    <w:tr><w:tc><w:p><w:r><w:t>Nested text</w:t></w:r></w:p></w:tc></w:tr></w:tbl>`);
}

// Invoke the rendered keyboard handler; only DOM selection and deferred focus
// are stubbed, so the handler and controller execute together.
function pressKey(view: React.ReactNode, key: string, text: string, paragraphIndex = 0): void {
  function findHost(node: React.ReactNode): React.ReactElement<React.HTMLAttributes<HTMLElement>> | undefined {
    for (const child of React.Children.toArray(node)) {
      if (!React.isValidElement<React.HTMLAttributes<HTMLElement>>(child)) continue;
      if (child.props.contentEditable && child.props.onKeyDown) return child;
      const found = findHost(child.props.children);
      if (found) return found;
    }
    return undefined;
  }
  const host = findHost(view);
  expect(host).toBeDefined();
  const element = {
    textContent: text,
    cloneNode: () => element,
    querySelectorAll: () => [],
    closest: () => element,
    contains: () => true,
    getAttribute: () => String(paragraphIndex),
  };
  vi.stubGlobal("Element", class {});
  vi.stubGlobal("window", {
    getSelection: () => ({ rangeCount: 0, anchorNode: { parentElement: element }, focusNode: null }),
    requestAnimationFrame: vi.fn(), setTimeout: vi.fn(), clearTimeout: vi.fn(),
  });
  const preventDefault = vi.fn();
  try {
    host!.props.onKeyDown!({
      key, currentTarget: element, target: element, nativeEvent: {},
      preventDefault, stopPropagation: vi.fn(),
      shiftKey: false, ctrlKey: false, altKey: false, metaKey: false,
    } as unknown as React.KeyboardEvent<HTMLElement>);
    expect(preventDefault).toHaveBeenCalled();
  } finally {
    vi.unstubAllGlobals();
  }
}

it.each(["cell draft", "recursive paragraph"])("%s converts a direct paragraph after a nested table without changing nested text", async (path) => {
  const model = await mixedCellModel();
  const nested = (model.nodes[0] as TableNode).rows[0]!.cells[0]!.nodes[0];
  const editor = runControllerSteps(model, [
    (e) => path === "cell draft" ? e.commitTableCellText(0, 0, 0, "-") : e.commitTableCellParagraphTextRecursive(0, 0, 0, 1, "-"),
    (e) => path === "cell draft" ? e.commitTableCellText(0, 0, 0, "- ") : e.commitTableCellParagraphTextRecursive(0, 0, 0, 1, "- "),
  ]);
  expect(textOf(paragraph(editor, true, 1))).toBe("");
  expect(paragraph(editor, true, 1).style?.numbering?.numId).toBeGreaterThan(0);
  expect((editor.model.nodes[0] as TableNode).rows[0]!.cells[0]!.nodes[0]).toEqual(nested);
});

it.each(["continue", "exit"])("Enter key can %s a list beside a nested table", async (action) => {
  const model = await mixedCellModel();
  const nested = (model.nodes[0] as TableNode).rows[0]!.cells[0]!.nodes[0];
  const direct = (model.nodes[0] as TableNode).rows[0]!.cells[0]!.nodes[1] as ParagraphNode;
  direct.children = [{ type: "text", text: action === "continue" ? "Item" : "" }];
  direct.style = { numbering: { numId: 1, ilvl: 0 } };
  const editor = runControllerSteps(model, [
    (e) => e.selectTableCell(0, 0, 0),
    (e, view) => pressKey(view, "Enter", textOf(paragraph(e, true, 1))),
  ], true);
  const nodes = (editor.model.nodes[0] as TableNode).rows[0]!.cells[0]!.nodes;
  expect(nodes[0]).toEqual(nested);
  expect(nodes).toHaveLength(action === "continue" ? 3 : 2);
  if (action === "continue") {
    expect(textOf(nodes[1] as ParagraphNode)).toBe("Item");
    expect(textOf(nodes[2] as ParagraphNode)).toBe("");
    expect((nodes[2] as ParagraphNode).style?.numbering).toEqual(direct.style.numbering);
    expect(editor.activeTextRange?.start.location).toEqual({ ...cellLocation, paragraphIndex: 1 });
  } else {
    expect((nodes[1] as ParagraphNode).style?.numbering).toBeUndefined();
  }
});

it.each([
  { inCell: false, level: 1 }, { inCell: true, level: 1 },
  { inCell: false, level: 2 }, { inCell: true, level: 2 },
])("empty Enter outdents one level and survives save/reopen ($inCell, $level)", async ({ inCell, level }) => {
  let previousStyle: ParagraphNode["style"];
  const editor = runControllerSteps(await load(inCell), [
    (e) => e.commitParagraphTextAtRange(inCell ? cellLocation : bodyLocation, "-", { start: 1, end: 1 }),
    (e) => e.commitParagraphTextAtRange(inCell ? cellLocation : bodyLocation, "- ", { start: 2, end: 2 }),
    (e) => { expect(e.adjustSelectedListDepth(level)).toBe(true); },
    (e, view) => {
      previousStyle = paragraph(e, inCell).style;
      pressKey(view, "Enter", "");
    },
  ], true);
  const node = paragraph(editor, inCell);
  expect(node.style?.numbering).toEqual({ ...previousStyle!.numbering, ilvl: level - 1 });
  expect(node.style?.indent).toBeUndefined();
  expect(textOf(node)).toBe("");
  expect(editor.activeTextRange?.start).toEqual({ location: inCell ? cellLocation : bodyLocation, offset: 0 });
  expect(inCell ? (editor.model.nodes[0] as TableNode).rows[0]!.cells[0]!.nodes : editor.model.nodes).toHaveLength(1);
  const reopened = await buildDocModel(await parseDocx(await serializeDocx(editor.model)));
  const saved = paragraph({ ...editor, model: reopened }, inCell);
  expect(saved.style?.numbering).toEqual(node.style?.numbering);
  expect(saved.style?.indent).toEqual(node.style?.indent);
});

it.each([
  { inCell: false, styleIndent: 0 }, { inCell: true, styleIndent: 0 },
  { inCell: false, styleIndent: 360 }, { inCell: true, styleIndent: 360 },
])("level-0 empty Enter restores style indent and survives save/reopen ($inCell, $styleIndent)", async ({ inCell, styleIndent }) => {
  const item = `<w:p><w:pPr><w:pStyle w:val="Body"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:p>`;
  const body = inCell ? `<w:tbl><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc>${item}</w:tc></w:tr></w:tbl>` : item;
  const pkg = createMinimalDocxPackage(`<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${body}</w:body></w:document>`);
  pkg.parts.set("word/styles.xml", {
    name: "word/styles.xml",
    content: `<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:style w:type="paragraph" w:styleId="Body"><w:name w:val="Body"/><w:pPr><w:ind w:left="${styleIndent}" w:right="120" w:firstLine="0"/></w:pPr></w:style></w:styles>`,
  });
  const model = await buildDocModel(await parseDocx(await packageToArrayBuffer(pkg)));
  const expectedIndent = model.metadata.paragraphStyles.find((style) => style.id === "Body")!.indent;
  const editor = runControllerSteps(model, [
    (e) => {
      expect(e.insertListItemAfterSelection("", 0, 0, inCell ? cellLocation : bodyLocation))
        .toEqual({ paragraphIndex: 0, caretOffset: 0 });
    },
  ]);
  const check = (value: DocModel): void => {
    const node = paragraph({ ...editor, model: value }, inCell);
    expect(node.style?.numbering).toBeUndefined();
    expect(node.style?.indent).toEqual(expectedIndent);
    expect(textOf(node)).toBe("");
    expect(inCell ? (value.nodes[0] as TableNode).rows[0]!.cells[0]!.nodes : value.nodes).toHaveLength(1);
  };
  check(editor.model);
  check(await buildDocModel(await parseDocx(await serializeDocx(editor.model))));
});

it("shares abstracts by kind while restarting independent lists and continuing one numId", async () => {
  const model = await load();
  const prefixes = ["- ", "* ", "1. ", "1. ", "1) ", "1) "];
  model.nodes = prefixes.map(() => ({ type: "paragraph", children: [{ type: "text", text: "" }] }));
  const editor = runControllerSteps(model, prefixes.flatMap((prefix, index) => [
    (e: DocxEditorController) => e.commitParagraphText(index, prefix.trimEnd()),
    (e: DocxEditorController) => e.commitParagraphText(index, prefix),
    (e: DocxEditorController) => e.commitParagraphText(index, "Item"),
  ]).concat([(e) => { e.insertListItemAfterSelection("Item", 4, 4, { kind: "paragraph", nodeIndex: 5 }); }]));
  function check(value: DocModel): void {
    const definitions = value.metadata.numberingDefinitions!;
    expect(definitions.abstracts).toHaveLength(3);
    expect(definitions.instances).toHaveLength(6);
    expect(new Set(definitions.instances.map((instance) => instance.numId)).size).toBe(6);
    expect(definitions.instances.map((instance) => instance.abstractNumId)).toEqual([0, 0, 1, 1, 2, 2]);
    expect((value.nodes[6] as ParagraphNode).style?.numbering).toEqual((value.nodes[5] as ParagraphNode).style?.numbering);
    const labels = buildParagraphNumberingLabels(value);
    expect([2, 3, 4, 5, 6].map((index) => labels.get(`p:${index}`)?.text?.trim())).toEqual(["1.", "1.", "1)", "1)", "2)"]);
  }
  check(editor.model);
  const reloaded = JSON.parse(JSON.stringify(editor.model)) as DocModel;
  reloaded.nodes.push({ type: "paragraph", children: [{ type: "text", text: "1." }] });
  const restarted = runControllerSteps(reloaded, [(e) => e.commitParagraphText(7, "1. ")]);
  expect(restarted.model.metadata.numberingDefinitions!.abstracts).toHaveLength(3);
  expect(restarted.model.metadata.numberingDefinitions!.instances).toHaveLength(7);
  expect(buildParagraphNumberingLabels(restarted.model).get("p:7")?.text?.trim()).toBe("1.");
});

it("preserves shared numbering definitions when saving and reopening", async () => {
  const model = await load();
  model.nodes.push({ type: "paragraph", children: [{ type: "text", text: "" }] });
  const editor = runControllerSteps(model, [
    (e) => e.commitParagraphText(0, "1."),
    (e) => e.commitParagraphText(0, "1. "),
    (e) => e.commitParagraphText(1, "1."),
    (e) => e.commitParagraphText(1, "1. "),
  ]);
  const reopened = await buildDocModel(await parseDocx(await serializeDocx(editor.model)));
  expect(reopened.metadata.numberingDefinitions).toEqual(editor.model.metadata.numberingDefinitions);
  expect(buildParagraphNumberingLabels(reopened).get("p:1")?.text?.trim()).toBe("1.");
});


it.each(["- ", "* ", "1. ", "1) "])("keeps character-by-character %j typing and Enter in one input session", async (prefix) => {
  const model = await load();
  const preceding = 'Paragraph 3 ends with click "OK".';
  model.nodes[0] = { type: "paragraph", children: [{ type: "text", text: preceding }] };
  let location: DocxTextRangeLocation = bodyLocation;
  let continuation = { text: preceding, start: preceding.length, end: preceding.length };
  const type = (editor: DocxEditorController, text: string): void => {
    // A burst deliberately uses the same controller, without a React render
    // between characters, just as beforeinput uses its continuation ref.
    for (const character of text) {
      const intent = resolveEditingIntent({ inputType: "insertText", data: character })!;
      const applied = applyTextEditingIntent(continuation.text, continuation.start, continuation.end, intent)!;
      const committed = editor.commitParagraphTextAtRange(location, applied.text,
        { start: applied.caret, end: applied.caret }, {
          insertionOffset: continuation.start,
          coalesceKey: `p:${location.kind === "paragraph" ? location.nodeIndex : 0}:${editingIntentCoalesceClass(intent)}`,
        });
      expect(committed).toEqual({
        text: applied.text === prefix ? "" : applied.text,
        start: applied.text === prefix ? 0 : applied.caret,
        end: applied.text === prefix ? 0 : applied.caret,
      });
      continuation = committed!;
    }
  };
  const enter = (editor: DocxEditorController, list: boolean): void => {
    const inserted = list
      ? editor.insertListItemAfterSelection(continuation.text, continuation.start, continuation.end, location)
      : editor.splitParagraphAtSelection(continuation.text, continuation.start, continuation.end, location);
    expect(inserted).toBeDefined();
    location = { kind: "paragraph", nodeIndex: inserted!.paragraphIndex };
    continuation = { text: "", start: inserted!.caretOffset, end: inserted!.caretOffset };
  };
  const editor = runControllerSteps(model, [
    (e) => enter(e, false),
    (e) => type(e, `${prefix}apples`),
    (e) => {
      expect(textOf(paragraph(e, false, 1))).toBe("apples");
      expect(e.activeTextRange?.end.offset).toBe(6);
      expect(paragraph(e, false, 1).style?.numbering?.numId).toBeGreaterThan(0);
      enter(e, true);
    },
    (e) => type(e, "pears"),
    (e) => { expect(e.adjustSelectedListDepth(1, continuation.text)).toBe(true); },
    (e) => {
      expect(paragraph(e, false, 2).style?.numbering?.ilvl).toBe(1);
      enter(e, true);
    },
    (e) => enter(e, true),
    (e) => enter(e, true),
    (e) => type(e, "after list"),
    (e) => {
      expect(e.model.nodes.map((node) => textOf(node as ParagraphNode)))
        .toEqual([preceding, "apples", "pears", "after list"]);
      expect(paragraph(e, false, 3).style?.numbering).toBeUndefined();
      expect(paragraph(e, false, 3).style?.indent).toBeUndefined();
      e.undo();
    },
    (e) => {
      // Undo takes back one word at a time.
      expect(textOf(paragraph(e, false, 3))).toBe("after ");
      e.undo();
    },
    (e) => {
      expect(textOf(paragraph(e, false, 3))).toBe("");
      expect(textOf(paragraph(e, false, 1))).toBe("apples");
      expect(textOf(paragraph(e, false, 2))).toBe("pears");
      e.redo();
    },
    (e) => {
      expect(textOf(paragraph(e, false, 3))).toBe("after ");
      e.redo();
    },
  ]);
  expect(textOf(paragraph(editor, false, 3))).toBe("after list");
  const reopened = await buildDocModel(await parseDocx(await serializeDocx(editor.model)));
  expect(reopened.nodes.map((node) => textOf(node as ParagraphNode)))
    .toEqual([preceding, "apples", "pears", "after list"]);
  expect((reopened.nodes[1] as ParagraphNode).style?.numbering)
    .toEqual(paragraph(editor, false, 1).style?.numbering);
  expect((reopened.nodes[2] as ParagraphNode).style?.numbering)
    .toEqual(paragraph(editor, false, 2).style?.numbering);
});


it.each([false, true])("keeps character-by-character list typing and caret continuation (cell=%s)", async (inCell) => {
  let location = inCell ? cellLocation : bodyLocation;
  let continuation = { text: "", start: 0, end: 0 };
  const type = (editor: DocxEditorController, text: string): void => {
    // Several native beforeinput events can arrive before React renders again.
    for (const character of text) {
      const before = continuation.start;
      const applied = applyTextEditingIntent(continuation.text, continuation.start, continuation.end, {
        kind: "insertText", text: character,
      })!;
      const committed = editor.commitParagraphTextAtRange(location, applied.text,
        { start: applied.caret, end: applied.caret },
        { insertionOffset: before, coalesceKey: `${JSON.stringify(location)}:insertText` });
      expect(committed).toEqual({
        text: applied.text === "- " ? "" : applied.text,
        start: applied.text === "- " ? 0 : applied.caret,
        end: applied.text === "- " ? 0 : applied.caret,
      });
      if (!committed) throw new Error("Text commit was rejected");
      continuation = committed;
    }
  };
  const enter = (editor: DocxEditorController): void => {
    const inserted = editor.insertListItemAfterSelection(
      continuation.text, continuation.start, continuation.end, location
    )!;
    location = inCell
      ? { ...cellLocation, paragraphIndex: inserted.paragraphIndex }
      : { ...bodyLocation, nodeIndex: inserted.paragraphIndex };
    continuation = { text: "", start: inserted.caretOffset, end: inserted.caretOffset };
  };
  const editor = runControllerSteps(await load(inCell), [
    (e) => type(e, "- apples"),
    (e) => {
      expect(textOf(paragraph(e, inCell))).toBe("apples");
      expect(e.activeTextRange?.start.offset).toBe(6);
      enter(e);
      type(e, "pears");
      expect(e.adjustSelectedListDepth(1, continuation.text)).toBe(true);
      enter(e);
      enter(e);
      enter(e);
      type(e, "after list");
    },
  ]);
  const nodes = inCell ? (editor.model.nodes[0] as TableNode).rows[0]!.cells[0]!.nodes : editor.model.nodes;
  expect(nodes.map((node) => textOf(node as ParagraphNode))).toEqual(["apples", "pears", "after list"]);
  expect((nodes[0] as ParagraphNode).style?.numbering?.ilvl).toBe(0);
  expect((nodes[1] as ParagraphNode).style?.numbering).toEqual({
    numId: (nodes[0] as ParagraphNode).style?.numbering?.numId, ilvl: 1,
  });
  expect((nodes[2] as ParagraphNode).style?.numbering).toBeUndefined();
  expect((nodes[2] as ParagraphNode).style?.indent).toBeUndefined();
  expect(editor.activeTextRange?.start).toEqual({ location, offset: 10 });
});

function findElement<P>(view: React.ReactNode, matches: (props: P) => boolean): React.ReactElement<P> | undefined {
  for (const child of React.Children.toArray(view)) {
    if (!React.isValidElement<P & { children?: React.ReactNode }>(child)) continue;
    if (matches(child.props)) return child;
    const found = findElement(child.props.children, matches);
    if (found) return found;
  }
  return undefined;
}

it("keeps the rendered wrapped textarea synchronized after prefix conversion", async () => {
  const model = await load();
  model.nodes[0] = {
    type: "paragraph",
    style: { spacing: { lineRule: "exact", lineTwips: 240 } },
    children: [{ type: "text", text: "-tail" }],
  };
  vi.stubGlobal("OffscreenCanvas", class {
    getContext() {
      return { font: "16px Calibri", measureText: (text: string) => ({ width: text.length * 8 }) };
    }
  });
  vi.stubGlobal("window", {
    requestAnimationFrame: vi.fn(), setTimeout: vi.fn(), clearTimeout: vi.fn(),
  });
  type SurfaceProps = React.HTMLAttributes<HTMLSpanElement> & {
    "data-docx-wrapped-paragraph-root"?: string;
  };
  try {
    const editor = runControllerSteps(model, [
      (_e, view) => {
        const surface = findElement<SurfaceProps>(view, (props) => props["data-docx-wrapped-paragraph-root"] === "true");
        expect(surface).toBeDefined();
        surface!.props.onClick!({ detail: 3, preventDefault: vi.fn(), stopPropagation: vi.fn() } as unknown as React.MouseEvent<HTMLSpanElement>);
      },
      ...Array.from(" apples", (character, index) => (e: DocxEditorController, view?: React.ReactNode) => {
        const textarea = findElement<React.TextareaHTMLAttributes<HTMLTextAreaElement>>(view,
          (props) => props["aria-label"] === "Wrapped paragraph text");
        expect(textarea, `textarea before character ${index}`).toBeDefined();
        const value = textarea!.props.value as string;
        expect(value).toBe(index === 0 ? "-tail" : "apples".slice(0, index - 1) + "tail");
        const caret = index === 0 ? 1 : e.activeTextRange!.end.offset;
        const next = value.slice(0, caret) + character + value.slice(caret);
        textarea!.props.onChange!({
          currentTarget: { value: next, selectionStart: caret + 1, selectionEnd: caret + 1 },
          nativeEvent: { inputType: "insertText", data: character },
        } as unknown as React.ChangeEvent<HTMLTextAreaElement>);
      }),
    ], true);
    expect(textOf(paragraph(editor, false))).toBe("applestail");
    expect(paragraph(editor, false).style?.numbering?.numId).toBeGreaterThan(0);
    expect(editor.activeTextRange?.end.offset).toBe(6);
  } finally {
    vi.unstubAllGlobals();
  }
});


for (const saved of [false, true]) {
  it.each([false, true])(`empty Enter cancels style-defined numbering ${saved ? "after save/reopen" : "in the model"} (cell=%s)`, async (inCell) => {
    const item = '<w:p><w:pPr><w:pStyle w:val="ListPara"/></w:pPr></w:p>';
    const body = inCell
      ? `<w:tbl><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc>${item}</w:tc></w:tr></w:tbl>`
      : item;
    const pkg = createMinimalDocxPackage(`<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${body}</w:body></w:document>`);
    pkg.parts.set("word/styles.xml", {
      name: "word/styles.xml",
      content: `<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:style w:type="paragraph" w:styleId="ListBase"><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr></w:style><w:style w:type="paragraph" w:styleId="ListPara"><w:basedOn w:val="ListBase"/></w:style></w:styles>`,
    });
    const model = await buildDocModel(await parseDocx(await packageToArrayBuffer(pkg)));
    const editor = runControllerSteps(model, [
      (e) => {
        expect(paragraph(e, inCell).style?.numbering).toEqual({ numId: 1, ilvl: 0 });
        expect(e.insertListItemAfterSelection("", 0, 0, inCell ? cellLocation : bodyLocation))
          .toEqual({ paragraphIndex: 0, caretOffset: 0 });
      },
    ]);
    let result = editor.model;
    if (saved) {
      for (let round = 0; round < 2; round += 1) {
        const reopenedPackage = await parseDocx(await serializeDocx(result));
        expect(reopenedPackage.parts.get("word/document.xml")!.content).toContain('<w:numId w:val="0"/>');
        result = await buildDocModel(reopenedPackage);
        expect(paragraph({ ...editor, model: result }, inCell).style?.numbering).toBeUndefined();
      }
    } else {
      expect(paragraph(editor, inCell).style?.numbering).toEqual({ numId: 0, ilvl: 0 });
    }
    expect(paragraph({ ...editor, model: result }, inCell).style?.styleId).toBe("ListPara");
    expect(textOf(paragraph({ ...editor, model: result }, inCell))).toBe("");
  });

  it.each([
    { inCell: false, replace: false }, { inCell: true, replace: false },
    { inCell: false, replace: true }, { inCell: true, replace: true },
  ])(`wrapped typing keeps the pre-edit run formatting ${saved ? "after save/reopen" : "in the model"} ($inCell, $replace)`, async ({ inCell, replace }) => {
    const model = await load(inCell);
    const node = inCell ? (model.nodes[0] as TableNode).rows[0]!.cells[0]!.nodes[0] as ParagraphNode : model.nodes[0] as ParagraphNode;
    node.sourceXml = undefined;
    node.style = { spacing: { lineRule: "exact", lineTwips: 240 } };
    node.children = [
      { type: "text", text: "ab", style: { bold: true } },
      { type: "text", text: "cd", style: { italic: true } },
    ];
    const location = inCell ? cellLocation : bodyLocation;
    vi.stubGlobal("OffscreenCanvas", class {
      getContext() {
        return { font: "16px Calibri", measureText: (text: string) => ({ width: text.length * 8 }) };
      }
    });
    vi.stubGlobal("window", {
      requestAnimationFrame: vi.fn(), setTimeout: vi.fn(), clearTimeout: vi.fn(),
    });
    try {
      const editor = runControllerSteps(model, [
        (_e, view) => {
          const surface = findElement<React.HTMLAttributes<HTMLSpanElement> & { "data-docx-wrapped-paragraph-root"?: string }>(view,
            (props) => props["data-docx-wrapped-paragraph-root"] === "true");
          expect(surface).toBeDefined();
          surface!.props.onClick!({ detail: 3, preventDefault: vi.fn(), stopPropagation: vi.fn() } as unknown as React.MouseEvent<HTMLSpanElement>);
        },
        (e) => e.setActiveTextRange({ start: { location, offset: 1 }, end: { location, offset: replace ? 3 : 1 } }),
        (_e, view) => {
          const textarea = findElement<React.TextareaHTMLAttributes<HTMLTextAreaElement>>(view,
            (props) => props["aria-label"] === "Wrapped paragraph text");
          expect(textarea).toBeDefined();
          textarea!.props.onBeforeInput!({
            currentTarget: { selectionStart: 1, selectionEnd: replace ? 3 : 1 },
            nativeEvent: { inputType: "insertText" },
          } as unknown as React.InputEvent<HTMLTextAreaElement>);
          textarea!.props.onChange!({
            currentTarget: { value: replace ? "axd" : "axbcd", selectionStart: 2, selectionEnd: 2 },
            nativeEvent: { inputType: "insertText", data: "x" },
          } as unknown as React.ChangeEvent<HTMLTextAreaElement>);
        },
      ], true);
      const result = saved ? await buildDocModel(await parseDocx(await serializeDocx(editor.model))) : editor.model;
      const runs = paragraph({ ...editor, model: result }, inCell).children.filter((child) => child.type === "text");
      expect(runs.map((run) => [run.text, run.style?.bold === true, run.style?.italic === true]))
        .toEqual([[replace ? "ax" : "axb", true, false], [replace ? "d" : "cd", false, true]]);
      expect(editor.activeTextRange?.start.offset).toBe(2);
    } finally {
      vi.unstubAllGlobals();
    }
  });
}
