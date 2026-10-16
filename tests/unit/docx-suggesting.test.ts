import * as React from "react";
import { renderToStaticMarkup } from "react-dom/server";
import { describe, expect, it } from "vitest";

import {
  createMinimalDocxPackage,
  packageToArrayBuffer,
  parseDocx,
} from "@extend-ai/react-docx-ooxml-core";
import {
  buildDocModel,
  type DocModel,
  type ParagraphNode,
} from "@extend-ai/react-docx-doc-model";
import {
  acceptParagraphRevision,
  joinParagraphsTracked,
  readParagraphAnnotations,
  rejectAllRevisions,
  rejectParagraphRevision,
  splitParagraphTracked,
  textReplacementBetween,
  trackModelTextEdits,
  updateParagraphText,
} from "@extend-ai/react-docx-editor-ops";
import {
  useDocxEditor,
  type DocxEditorController,
} from "../../packages/react-viewer/src/editor";

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const DATE = "2026-10-16T09:00:00Z";
const AUTHOR = { author: "Grace", date: DATE };

async function load(body: string): Promise<DocModel> {
  const seed = createMinimalDocxPackage(
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="${W_NS}"><w:body>${body}</w:body></w:document>`
  );
  return buildDocModel(await parseDocx(await packageToArrayBuffer(seed)));
}

const PLAIN = '<w:p><w:r><w:t xml:space="preserve">Hello world</w:t></w:r></w:p>';
const TWO_PARAGRAPHS = `${PLAIN}<w:p><w:r><w:t>Second</w:t></w:r></w:p>`;

function sourceXml(model: DocModel, index = 0): string {
  return (model.nodes[index] as ParagraphNode).sourceXml ?? "";
}

function visible(model: DocModel, index = 0): string {
  return (model.nodes[index] as ParagraphNode).children
    .map((child) => (child.type === "text" ? child.text : ""))
    .join("");
}

describe("textReplacementBetween", () => {
  it("finds the minimal replacement", () => {
    expect(textReplacementBetween("Hello world", "Hello world")).toBeUndefined();
    expect(textReplacementBetween("Hello world", "Hello brave world")).toEqual({
      startOffset: 6,
      endOffset: 6,
      text: "brave ",
    });
    expect(textReplacementBetween("Hello world", "Hello")).toEqual({
      startOffset: 5,
      endOffset: 11,
      text: "",
    });
    expect(textReplacementBetween("Hello world", "Help world")).toEqual({
      startOffset: 3,
      endOffset: 5,
      text: "p",
    });
  });

  it("never splits a surrogate pair", () => {
    // U+1F600 and U+1F601 share their high surrogate.
    const replacement = textReplacementBetween("a\u{1F600}b", "a\u{1F601}b");
    expect(replacement).toEqual({ startOffset: 1, endOffset: 3, text: "\u{1F601}" });
  });
});

describe("trackModelTextEdits", () => {
  it("records typing as an insertion by the author", async () => {
    const model = await load(PLAIN);
    const result = trackModelTextEdits(
      model,
      updateParagraphText(model, 0, "Hello brave world"),
      AUTHOR
    );
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    expect(visible(result.model)).toBe("Hello brave world");
    expect(sourceXml(result.model)).toMatch(
      /<w:ins w:id="\d+" w:author="Grace" w:date="2026-10-16T09:00:00Z"><w:r><w:t xml:space="preserve">brave <\/w:t><\/w:r><\/w:ins>/
    );
  });

  it("records deleting as a deletion that stays in the source", async () => {
    const model = await load(PLAIN);
    const result = trackModelTextEdits(
      model,
      updateParagraphText(model, 0, "Hello"),
      AUTHOR
    );
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    expect(visible(result.model)).toBe("Hello");
    expect(sourceXml(result.model)).toMatch(
      /<w:del w:id="\d+" w:author="Grace"[^>]*><w:r><w:delText xml:space="preserve"> world<\/w:delText><\/w:r><\/w:del>/
    );
  });

  it("removes the author's own pending insertion outright on backspace", async () => {
    const model = await load(PLAIN);
    const typed = trackModelTextEdits(
      model,
      updateParagraphText(model, 0, "Hello world!"),
      AUTHOR
    );
    expect(typed.ok).toBe(true);
    if (!typed.ok) return;
    const erased = trackModelTextEdits(
      typed.model,
      updateParagraphText(typed.model, 0, "Hello world"),
      AUTHOR
    );
    expect(erased.ok).toBe(true);
    if (!erased.ok) return;
    expect(visible(erased.model)).toBe("Hello world");
    expect(sourceXml(erased.model)).not.toMatch(/<w:(ins|del) /);
  });

  it("keeps an earlier deletion while typing after it", async () => {
    const model = await load(PLAIN);
    const deleted = trackModelTextEdits(
      model,
      updateParagraphText(model, 0, "Hello"),
      AUTHOR
    );
    expect(deleted.ok).toBe(true);
    if (!deleted.ok) return;
    const typed = trackModelTextEdits(
      deleted.model,
      updateParagraphText(deleted.model, 0, "Hello there"),
      AUTHOR
    );
    expect(typed.ok).toBe(true);
    if (!typed.ok) return;
    expect(visible(typed.model)).toBe("Hello there");
    expect(sourceXml(typed.model)).toContain("<w:delText xml:space=\"preserve\"> world</w:delText>");
    expect(sourceXml(typed.model)).toMatch(/<w:ins [^>]*><w:r><w:t xml:space="preserve"> there<\/w:t>/);
  });

  it("refuses structure, formatting and metadata changes", async () => {
    const model = await load(TWO_PARAGRAPHS);
    const merged: DocModel = { ...model, nodes: [model.nodes[0]!] };
    expect(trackModelTextEdits(model, merged, AUTHOR)).toEqual({
      ok: false,
      reason: "unsupported",
    });

    const paragraph = model.nodes[0] as ParagraphNode;
    const bolded: DocModel = {
      ...model,
      nodes: [
        {
          ...paragraph,
          children: paragraph.children.map((child) =>
            child.type === "text" ? { ...child, style: { bold: true } } : child
          ),
        },
        model.nodes[1]!,
      ],
    };
    expect(trackModelTextEdits(model, bolded, AUTHOR).ok).toBe(false);

    const retitled: DocModel = { ...model, metadata: { ...model.metadata } };
    expect(trackModelTextEdits(model, retitled, AUTHOR).ok).toBe(false);
  });

  it("refuses a text change that also changes formatting, such as a styled paste", async () => {
    const model = await load(PLAIN);
    const paragraph = model.nodes[0] as ParagraphNode;
    const pasted: DocModel = {
      ...model,
      nodes: [
        {
          ...paragraph,
          sourceXml: undefined,
          children: [
            ...paragraph.children,
            { type: "text", text: " again", style: { bold: true } },
          ],
        },
      ],
    };
    expect(trackModelTextEdits(model, pasted, AUTHOR)).toEqual({
      ok: false,
      reason: "unsupported",
    });
  });

  it("tracks typing next to formatted text, which takes that formatting", async () => {
    const model = await load(
      '<w:p><w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">Bold </w:t></w:r><w:r><w:t>plain</w:t></w:r></w:p>'
    );
    for (const text of ["Oh Bold plain", "Bold, plain", "Bold plainly"]) {
      const result = trackModelTextEdits(model, updateParagraphText(model, 0, text), AUTHOR);
      expect(result.ok, text).toBe(true);
    }
  });

  it("passes an unchanged model through", async () => {
    const model = await load(PLAIN);
    expect(
      trackModelTextEdits(model, { ...model, nodes: [...model.nodes] }, AUTHOR)
    ).toEqual({ ok: true, model });
  });
});

/**
 * Drives the controller through server rendering: each step runs during a
 * render and its state update re-renders the probe with the result.
 */
function runControllerSteps(
  model: DocModel,
  steps: Array<(editor: DocxEditorController) => void>
): DocxEditorController {
  let next = 0;
  let last: DocxEditorController | undefined;
  renderToStaticMarkup(
    React.createElement(function Probe(): null {
      const editor = useDocxEditor({ starterModel: model, author: "Grace" });
      last = editor;
      const step = steps[next];
      if (step) {
        next += 1;
        step(editor);
      }
      return null;
    })
  );
  expect(next).toBe(steps.length);
  return last!;
}

describe("suggesting controller", () => {
  it("tracks commits while suggesting and accepts the change by revision id", async () => {
    const model = await load(PLAIN);
    let revisionId: string | undefined;
    const editor = runControllerSteps(model, [
      (editor) => {
        editor.registerSuggestingMode();
        editor.commitParagraphText(0, "Hello brave world");
      },
      (editor) => {
        const insertion = editor.trackedChanges.find(
          (change) => change.kind === "insertion"
        );
        expect(insertion).toMatchObject({ author: "Grace" });
        revisionId = insertion?.revisionId;
        expect(editor.acceptRevision(revisionId!)).toEqual({ ok: true });
      },
    ]);
    expect(revisionId).toBeDefined();
    expect(visible(editor.model)).toBe("Hello brave world");
    expect(editor.trackedChanges).toHaveLength(0);
    expect(sourceXml(editor.model)).not.toContain("<w:ins ");
  });

  it("rejects all suggestions in one step", async () => {
    const model = await load(PLAIN);
    const editor = runControllerSteps(model, [
      (editor) => {
        editor.registerSuggestingMode();
        editor.commitParagraphText(0, "Hello");
      },
      (editor) => {
        expect(editor.trackedChanges.map((change) => change.kind)).toEqual([
          "deletion",
        ]);
        expect(editor.rejectAllRevisions()).toEqual({
          ok: true,
          remainingNodeIndexes: [],
        });
      },
    ]);
    expect(visible(editor.model)).toBe("Hello world");
    expect(editor.trackedChanges).toHaveLength(0);
  });

  it("refuses edits it cannot track and leaves the model unchanged", async () => {
    const model = await load(TWO_PARAGRAPHS);
    const editor = runControllerSteps(model, [
      (editor) => {
        editor.registerSuggestingMode();
        editor.appendParagraph("Third");
      },
      (editor) => {
        expect(editor.suggestionRefusalNonce).toBe(1);
        editor.setAlignment("center");
      },
    ]);
    expect(editor.suggestionRefusalNonce).toBe(2);
    expect(editor.model.nodes).toHaveLength(2);
    expect((editor.model.nodes[0] as ParagraphNode).style?.align).toBeUndefined();
    expect(editor.canUndo).toBe(false);
  });

  it("lists a deleted space as a change the reviewer can see", async () => {
    const model = await load(PLAIN);
    const editor = runControllerSteps(model, [
      (editor) => {
        editor.registerSuggestingMode();
        editor.commitParagraphText(0, "Helloworld");
      },
    ]);
    expect(editor.trackedChanges.map(({ kind, text }) => [kind, text])).toEqual([
      ["deletion", "Space"],
    ]);
  });

  it("keeps tracking while another suggesting viewer is still registered", async () => {
    const model = await load(PLAIN);
    const editor = runControllerSteps(model, [
      (editor) => {
        editor.registerSuggestingMode();
        const unregisterSecond = editor.registerSuggestingMode();
        unregisterSecond();
        unregisterSecond();
        editor.commitParagraphText(0, "Hello brave world");
      },
    ]);
    expect(visible(editor.model)).toBe("Hello brave world");
    expect(sourceXml(editor.model)).toContain("<w:ins ");
  });

  it("leaves edit mode untracked", async () => {
    const model = await load(PLAIN);
    const editor = runControllerSteps(model, [
      (editor) => editor.commitParagraphText(0, "Hello brave world"),
    ]);
    expect(visible(editor.model)).toBe("Hello brave world");
    expect(sourceXml(editor.model)).not.toContain("<w:ins ");
  });
});

/** The model the editor makes when Enter splits paragraph 0 at `offset`. */
function splitByEditor(model: DocModel, offset: number): DocModel {
  const paragraph = model.nodes[0] as ParagraphNode;
  const text = visible(model);
  const half = (part: string, blockId?: string): ParagraphNode => ({
    type: "paragraph",
    blockId,
    style: paragraph.style,
    children: [{ type: "text", text: part }],
  });
  return {
    ...model,
    nodes: [
      half(text.slice(0, offset), paragraph.blockId),
      half(text.slice(offset), "new-block"),
      ...model.nodes.slice(1),
    ],
  };
}

describe("tracked paragraph breaks", () => {
  it("records Enter as an inserted paragraph mark on the first half", async () => {
    const model = await load(TWO_PARAGRAPHS);
    const tracked = trackModelTextEdits(model, splitByEditor(model, 6), AUTHOR);
    expect(tracked.ok).toBe(true);
    const next = (tracked as { model: DocModel }).model;
    expect(next.nodes).toHaveLength(3);
    expect([visible(next, 0), visible(next, 1), visible(next, 2)]).toEqual([
      "Hello ",
      "world",
      "Second",
    ]);
    expect(sourceXml(next, 0)).toMatch(
      /^<w:p\b[^>]*><w:pPr><w:rPr><w:ins w:id="\d+" w:author="Grace" w:date="2026-10-16T09:00:00Z"\/><\/w:rPr><\/w:pPr>/
    );
    expect(sourceXml(next, 1)).not.toContain("<w:ins");
    expect((next.nodes[1] as ParagraphNode).blockId).toBe("new-block");
    expect(next.nodes[2]).toBe(model.nodes[1]);
    expect(readParagraphAnnotations(next, 0).revisions).toEqual([
      expect.objectContaining({ kind: "paragraph-insertion", author: "Grace", text: "" }),
    ]);
  });

  it("joins the paragraphs again on reject and keeps the break on accept", async () => {
    const model = await load(PLAIN);
    const split = splitParagraphTracked(model, { nodeIndex: 0, offset: 6, ...AUTHOR });
    expect(split.ok).toBe(true);
    const next = (split as { model: DocModel }).model;
    const [revision] = readParagraphAnnotations(next, 0).revisions;
    const target = { nodeIndex: 0, revisionId: revision!.id, kind: "paragraph-insertion" as const };

    const rejected = rejectParagraphRevision(next, target);
    expect(rejected.ok).toBe(true);
    const joined = (rejected as { model: DocModel }).model;
    expect(joined.nodes).toHaveLength(1);
    expect(visible(joined)).toBe("Hello world");
    expect(sourceXml(joined)).not.toContain("<w:ins");

    const accepted = acceptParagraphRevision(next, target);
    expect(accepted.ok).toBe(true);
    const kept = (accepted as { model: DocModel }).model;
    expect([visible(kept, 0), visible(kept, 1)]).toEqual(["Hello ", "world"]);
    expect(sourceXml(kept, 0)).not.toContain("<w:ins");
    expect(sourceXml(kept, 0)).not.toContain("<w:rPr>");
  });

  it("rejects a break and the typing after it back to the original paragraph", async () => {
    const model = await load(PLAIN);
    let next = (splitParagraphTracked(model, { nodeIndex: 0, offset: 11, ...AUTHOR }) as {
      model: DocModel;
    }).model;
    const typed = updateParagraphText(next, 1, "New line");
    next = (trackModelTextEdits(next, typed, AUTHOR) as { model: DocModel }).model;
    expect(visible(next, 1)).toBe("New line");

    const result = rejectAllRevisions(next);
    expect(result.remainingNodeIndexes).toEqual([]);
    expect(result.model.nodes).toHaveLength(1);
    expect(visible(result.model)).toBe("Hello world");
  });

  it("gives a revision that spans the break a fresh id in the second half", async () => {
    const model = await load(
      '<w:p><w:ins w:id="3" w:author="Bob"><w:r><w:t>Hello world</w:t></w:r></w:ins></w:p>'
    );
    const next = (splitParagraphTracked(model, { nodeIndex: 0, offset: 6, ...AUTHOR }) as {
      model: DocModel;
    }).model;
    const ids = [0, 1].map((index) =>
      readParagraphAnnotations(next, index)
        .revisions.filter((revision) => revision.kind === "insertion")
        .map((revision) => revision.id)
    );
    expect(ids[0]).toEqual(["3"]);
    expect(ids[1]).toHaveLength(1);
    expect(ids[1]![0]).not.toBe("3");
  });

  it("keeps a comment that starts at the cut on the text it covers", async () => {
    const model = await load(
      '<w:p><w:r><w:t xml:space="preserve">Hello </w:t></w:r><w:commentRangeStart w:id="4"/>' +
        '<w:r><w:t>world</w:t></w:r><w:commentRangeEnd w:id="4"/></w:p>'
    );
    const next = (splitParagraphTracked(model, { nodeIndex: 0, offset: 6, ...AUTHOR }) as {
      model: DocModel;
    }).model;
    expect(readParagraphAnnotations(next, 0).comments).toEqual([]);
    expect(readParagraphAnnotations(next, 1).comments).toEqual([
      { commentId: 4, startOffset: 0, endOffset: 5, text: "world" },
    ]);
  });

  it("moves later sections by one node on a split and back on a join", async () => {
    const model = await load(
      `${TWO_PARAGRAPHS}<w:p><w:pPr><w:sectPr/></w:pPr><w:r><w:t>End of one</w:t></w:r></w:p>` +
        "<w:p><w:r><w:t>Next section</w:t></w:r></w:p>"
    );
    const starts = (next: DocModel) =>
      (next.metadata.sections ?? []).map((section) => section.startNodeIndex);
    const before = starts(model);
    expect(before.at(-1)).toBe(3);
    const split = (splitParagraphTracked(model, { nodeIndex: 0, offset: 6, ...AUTHOR }) as {
      model: DocModel;
    }).model;
    expect(starts(split).at(-1)).toBe(4);
    expect(rejectAllRevisions(split).model.metadata.sections?.map((s) => s.startNodeIndex)).toEqual(
      before
    );
    expect(joinParagraphsTracked(model, { nodeIndex: 2, ...AUTHOR })).toEqual({
      ok: false,
      reason: "unsupported",
    });
  });

  it("refuses a split that also changes text", async () => {
    const model = await load(PLAIN);
    const edited = splitByEditor(model, 6);
    (edited.nodes[1] as ParagraphNode).children = [{ type: "text", text: "there" }];
    expect(trackModelTextEdits(model, edited, AUTHOR)).toEqual({
      ok: false,
      reason: "unsupported",
    });
  });
});

describe("suggesting controller paragraph breaks", () => {
  it("tracks Enter and accepts the break by revision id", async () => {
    const model = await load(PLAIN);
    const editor = runControllerSteps(model, [
      (editor) => {
        editor.registerSuggestingMode();
        editor.splitParagraphAtSelection("Hello world", 6, 6, {
          kind: "paragraph",
          nodeIndex: 0,
        });
      },
      (editor) => {
        const change = editor.trackedChanges.find(
          (candidate) => candidate.kind === "paragraph-insertion"
        );
        expect(change).toMatchObject({ author: "Grace", nodeIndex: 0 });
        expect(editor.acceptRevision(change!.revisionId!)).toEqual({ ok: true });
      },
    ]);
    expect(editor.suggestionRefusalNonce).toBe(0);
    expect([visible(editor.model, 0), visible(editor.model, 1)]).toEqual([
      "Hello ",
      "world",
    ]);
    expect(editor.trackedChanges).toHaveLength(0);
  });
});

describe("tracked paragraph joins", () => {
  function joinedByEditor(model: DocModel): DocModel {
    const first = model.nodes[0] as ParagraphNode;
    return {
      ...model,
      nodes: [
        {
          ...first,
          children: [{ type: "text", text: `${visible(model, 0)}${visible(model, 1)}` }],
        },
        ...model.nodes.slice(2),
      ],
    };
  }

  it("records Backspace at a paragraph start as a deleted paragraph mark", async () => {
    const model = await load(TWO_PARAGRAPHS);
    const tracked = trackModelTextEdits(model, joinedByEditor(model), AUTHOR);
    expect(tracked.ok).toBe(true);
    const next = (tracked as { model: DocModel }).model;
    expect(next.nodes).toHaveLength(2);
    expect(sourceXml(next, 0)).toMatch(
      /<w:pPr><w:rPr><w:del w:id="\d+" w:author="Grace" w:date="2026-10-16T09:00:00Z"\/><\/w:rPr><\/w:pPr>/
    );
    expect(next.nodes[1]).toBe(model.nodes[1]);
    const [revision] = readParagraphAnnotations(next, 0).revisions;
    expect(revision).toMatchObject({ kind: "paragraph-deletion", author: "Grace" });

    const target = { nodeIndex: 0, revisionId: revision!.id, kind: "paragraph-deletion" as const };
    const accepted = (acceptParagraphRevision(next, target) as { model: DocModel }).model;
    expect(accepted.nodes).toHaveLength(1);
    expect(visible(accepted)).toBe("Hello worldSecond");
    const rejected = (rejectParagraphRevision(next, target) as { model: DocModel }).model;
    expect(rejected.nodes).toHaveLength(2);
    expect(sourceXml(rejected, 0)).not.toContain("<w:del");
  });

  it("removes the author's own inserted break outright", async () => {
    const model = await load(PLAIN);
    const split = (splitParagraphTracked(model, { nodeIndex: 0, offset: 6, ...AUTHOR }) as {
      model: DocModel;
    }).model;
    const joined = joinParagraphsTracked(split, { nodeIndex: 0, ...AUTHOR });
    expect(joined.ok).toBe(true);
    const next = (joined as { model: DocModel }).model;
    expect(next.nodes).toHaveLength(1);
    expect(visible(next)).toBe("Hello world");
    expect(sourceXml(next)).not.toMatch(/<w:(?:ins|del)\b/);
  });

  it("marks another author's inserted break deleted after its insertion", async () => {
    const model = await load(PLAIN);
    const split = (splitParagraphTracked(model, {
      nodeIndex: 0,
      offset: 6,
      author: "Bob",
      date: DATE,
    }) as { model: DocModel }).model;
    const next = (joinParagraphsTracked(split, { nodeIndex: 0, ...AUTHOR }) as {
      model: DocModel;
    }).model;
    expect(sourceXml(next, 0)).toMatch(/<w:rPr><w:ins [^>]*w:author="Bob"[^>]*\/><w:del [^>]*w:author="Grace"/);
    expect(rejectAllRevisions(next).model.nodes).toHaveLength(1);
  });

  it("refuses joining a paragraph with a table", async () => {
    const model = await load(
      `${PLAIN}<w:tbl><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl>`
    );
    expect(joinParagraphsTracked(model, { nodeIndex: 0, ...AUTHOR })).toEqual({
      ok: false,
      reason: "unsupported",
    });
  });
});

describe("suggesting controller paragraph joins", () => {
  it("tracks Backspace across paragraphs and joins them on accept", async () => {
    const model = await load(TWO_PARAGRAPHS);
    const editor = runControllerSteps(model, [
      (editor) => {
        editor.registerSuggestingMode();
        editor.deleteExpandedSelection({
          start: { location: { kind: "paragraph", nodeIndex: 0 }, offset: 11 },
          end: { location: { kind: "paragraph", nodeIndex: 1 }, offset: 0 },
        });
      },
      (editor) => {
        expect(editor.model.nodes).toHaveLength(2);
        const change = editor.trackedChanges.find(
          (candidate) => candidate.kind === "paragraph-deletion"
        );
        expect(change).toMatchObject({ author: "Grace", nodeIndex: 0 });
        expect(editor.acceptRevision(change!.revisionId!)).toEqual({ ok: true });
      },
    ]);
    expect(editor.suggestionRefusalNonce).toBe(0);
    expect(editor.model.nodes).toHaveLength(1);
    expect(visible(editor.model)).toBe("Hello worldSecond");
  });
});
