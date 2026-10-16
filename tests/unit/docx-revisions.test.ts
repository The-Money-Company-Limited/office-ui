import { describe, expect, it } from "vitest";
import {
  createMinimalDocxPackage,
  packageToArrayBuffer,
  parseDocx,
  withPart,
  type OoxmlPackage,
} from "@extend-ai/react-docx-ooxml-core";
import {
  buildDocModel,
  type DocModel,
  type ParagraphNode,
} from "@extend-ai/react-docx-doc-model";
import {
  acceptAllRevisions,
  acceptParagraphRevision,
  createParagraphComment,
  deleteComment,
  rejectAllRevisions,
  rejectParagraphRevision,
  replaceParagraphTextTracked,
  replyToComment,
  updateParagraphText,
  type AnnotationMutationResult,
} from "@extend-ai/react-docx-editor-ops";
import { serializeDocModel } from "@extend-ai/react-docx-serializer";

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const DATE = "2026-10-16T09:00:00Z";

function documentXml(body: string): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="${W_NS}" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body>${body}</w:body></w:document>`;
}

const COMMENTS_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:comments xmlns:w="${W_NS}" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:comment w:id="0" w:author="Carol" w:initials="C" w:date="${DATE}"><w:p w14:paraId="0A000001"><w:r><w:t>Is this right?</w:t></w:r></w:p></w:comment></w:comments>`;
const COMMENTS_EXTENDED_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w15:commentsEx xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"><w15:commentEx w15:paraId="0A000001" w15:done="0"/></w15:commentsEx>`;
const RELS_WITH_COMMENTS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/><Relationship Id="rId2" Type="http://schemas.microsoft.com/office/2011/relationships/commentsExtended" Target="commentsExtended.xml"/></Relationships>`;

/** "The quick brown fox": "quick" carries comment 0, " slow" is deleted, " brown" inserted by Bob. */
const ANNOTATED_PARAGRAPH =
  `<w:p><w:r><w:t xml:space="preserve">The </w:t></w:r>` +
  `<w:commentRangeStart w:id="0"/><w:r><w:rPr><w:i/></w:rPr><w:t>quick</w:t></w:r><w:commentRangeEnd w:id="0"/>` +
  `<w:r><w:commentReference w:id="0"/></w:r>` +
  `<w:del w:id="5" w:author="Bob" w:date="${DATE}"><w:r><w:delText xml:space="preserve"> slow</w:delText></w:r></w:del>` +
  `<w:ins w:id="6" w:author="Bob" w:date="${DATE}"><w:r><w:t xml:space="preserve"> brown</w:t></w:r></w:ins>` +
  `<w:r><w:t xml:space="preserve"> fox</w:t></w:r></w:p>`;

async function load(
  body: string,
  withComments = false
): Promise<{ pkg: OoxmlPackage; model: DocModel }> {
  let seed = createMinimalDocxPackage(documentXml(body));
  if (withComments) {
    seed = withPart(seed, { name: "word/comments.xml", content: COMMENTS_XML });
    seed = withPart(seed, {
      name: "word/commentsExtended.xml",
      content: COMMENTS_EXTENDED_XML,
    });
    seed = withPart(seed, {
      name: "word/_rels/document.xml.rels",
      content: RELS_WITH_COMMENTS,
    });
  }
  const pkg = await parseDocx(await packageToArrayBuffer(seed));
  return { pkg, model: await buildDocModel(pkg) };
}

async function exported(model: DocModel, pkg: OoxmlPackage) {
  const out = await serializeDocModel(model, pkg);
  return {
    out,
    xml: out.parts.get("word/document.xml")?.content ?? "",
    comments: out.parts.get("word/comments.xml")?.content ?? "",
    commentsExtended: out.parts.get("word/commentsExtended.xml")?.content ?? "",
    model: await buildDocModel(out),
  };
}

function paragraphText(model: DocModel, index = 0): string {
  const paragraph = model.nodes[index] as ParagraphNode;
  return paragraph.children
    .map((child) => (child.type === "text" ? child.text : ""))
    .join("");
}

function ok(result: AnnotationMutationResult): DocModel {
  if (!result.ok) {
    throw new Error(`mutation failed: ${result.reason}`);
  }
  return result.model;
}

function revisionIds(xml: string): string[] {
  return [...xml.matchAll(/<w:(?:ins|del) w:id="(\d+)"/g)].map((match) => match[1]!);
}

describe("editing paragraphs that carry revisions and comments", () => {
  it("keeps deleted text, insertions and comment anchors when text is typed", async () => {
    const { pkg, model } = await load(ANNOTATED_PARAGRAPH, true);
    expect(paragraphText(model)).toBe("The quick brown fox");

    // One keystroke-sized edit at a time, as the editor commits them.
    const edited = updateParagraphText(
      updateParagraphText(model, 0, "The very quick brown fox"),
      0,
      "The very quick brown fox jumps"
    );
    const result = await exported(edited, pkg);

    expect(result.xml).toContain(
      `<w:del w:id="5" w:author="Bob" w:date="${DATE}"><w:r><w:delText xml:space="preserve"> slow</w:delText></w:r></w:del>`
    );
    expect(result.xml).toContain(
      `<w:ins w:id="6" w:author="Bob" w:date="${DATE}"><w:r><w:t xml:space="preserve"> brown</w:t></w:r></w:ins>`
    );
    expect(result.xml).toContain('<w:t xml:space="preserve">The very </w:t>');
    expect(result.xml).toContain(
      '<w:commentRangeStart w:id="0"/><w:r><w:rPr><w:i/></w:rPr><w:t>quick</w:t></w:r><w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r>'
    );
    expect(result.xml).toContain('<w:t xml:space="preserve"> fox jumps</w:t>');
    expect(paragraphText(result.model)).toBe("The very quick brown fox jumps");
    expect(result.comments).toContain("Is this right?");
  });

  it("keeps text typed inside an insertion part of that insertion", async () => {
    const { pkg, model } = await load(ANNOTATED_PARAGRAPH);
    const edited = updateParagraphText(model, 0, "The quick brownish fox");
    const { xml } = await exported(edited, pkg);
    expect(xml).toContain(
      `<w:ins w:id="6" w:author="Bob" w:date="${DATE}"><w:r><w:t xml:space="preserve"> brownish</w:t></w:r></w:ins>`
    );
    expect(xml).toContain("<w:delText xml:space=\"preserve\"> slow</w:delText>");
  });

  it("collapses comment markers to the edit point when their text is removed", async () => {
    const { pkg, model } = await load(ANNOTATED_PARAGRAPH);
    const edited = updateParagraphText(model, 0, "The  brown fox");
    const result = await exported(edited, pkg);
    expect(result.xml).toContain(
      '<w:commentRangeStart w:id="0"/><w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r>'
    );
    expect(result.xml).toContain("> slow</w:delText>");
    expect(paragraphText(result.model)).toBe("The  brown fox");
  });

  it("refuses an edit it cannot apply without dropping revisions", () => {
    const paragraph: ParagraphNode = {
      type: "paragraph",
      sourceXml:
        '<w:p><w:ins w:id="1" w:author="A"><w:r><w:t>A</w:t></w:r></w:ins><w:r><w:fldChar w:fldCharType="begin"/></w:r></w:p>',
      children: [{ type: "text", text: "A" }],
    };
    const model: DocModel = {
      nodes: [paragraph],
      metadata: {
        sourceParts: 1,
        warnings: [],
        headerSections: [],
        footerSections: [],
        paragraphStyles: [],
      },
    };
    expect(updateParagraphText(model, 0, "AB")).toBe(model);
  });
});

describe("tracked text replacement", () => {
  const PLAIN =
    '<w:p><w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">Hello world</w:t></w:r></w:p>';

  it("writes a replacement as a deletion plus an insertion with copied formatting", async () => {
    const { pkg, model } = await load(PLAIN);
    const edited = ok(
      replaceParagraphTextTracked(model, {
        nodeIndex: 0,
        startOffset: 6,
        endOffset: 11,
        text: "there",
        author: "Ada",
        date: DATE,
      })
    );
    const result = await exported(edited, pkg);
    expect(result.xml).toContain(
      `<w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">Hello </w:t></w:r>` +
        `<w:del w:id="1" w:author="Ada" w:date="${DATE}"><w:r><w:rPr><w:b/></w:rPr><w:delText>world</w:delText></w:r></w:del>` +
        `<w:ins w:id="0" w:author="Ada" w:date="${DATE}"><w:r><w:rPr><w:b/></w:rPr><w:t>there</w:t></w:r></w:ins>`
    );
    expect(paragraphText(result.model)).toBe("Hello there");
    expect((result.model.nodes[0] as ParagraphNode).children[1]).toMatchObject({
      text: "there",
      style: { bold: true },
    });
  });

  it("removes the author's own pending insertion and nests other authors' deletions", async () => {
    const { pkg, model } = await load(PLAIN);
    const inserted = ok(
      replaceParagraphTextTracked(model, {
        nodeIndex: 0,
        startOffset: 5,
        endOffset: 5,
        text: " brave new",
        author: "Ada",
        date: DATE,
      })
    );
    expect(paragraphText(inserted)).toBe("Hello brave new world");

    const ownDelete = ok(
      replaceParagraphTextTracked(inserted, {
        nodeIndex: 0,
        startOffset: 5,
        endOffset: 11,
        text: "",
        author: "Ada",
      })
    );
    const own = await exported(ownDelete, pkg);
    expect(own.xml).not.toContain("<w:del ");
    expect(own.xml).toContain(
      `<w:ins w:id="0" w:author="Ada" w:date="${DATE}"><w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve"> new</w:t></w:r></w:ins>`
    );

    const otherDelete = ok(
      replaceParagraphTextTracked(ownDelete, {
        nodeIndex: 0,
        startOffset: 5,
        endOffset: 9,
        text: "",
        author: "Bob",
      })
    );
    const other = await exported(otherDelete, pkg);
    expect(other.xml).toContain(
      `<w:ins w:id="0" w:author="Ada" w:date="${DATE}"><w:del w:id="1" w:author="Bob"><w:r><w:rPr><w:b/></w:rPr><w:delText xml:space="preserve"> new</w:delText></w:r></w:del></w:ins>`
    );
    expect(paragraphText(other.model)).toBe("Hello world");
  });

  it("splits a revision it cuts through with fresh unique ids", async () => {
    const { pkg, model } = await load(ANNOTATED_PARAGRAPH);
    // Ada replaces "row" in Bob's " brown" insertion.
    const edited = ok(
      replaceParagraphTextTracked(model, {
        nodeIndex: 0,
        startOffset: 11,
        endOffset: 14,
        text: "ROW",
        author: "Ada",
      })
    );
    const result = await exported(edited, pkg);
    expect(paragraphText(result.model)).toBe("The quick bROWn fox");
    const ids = revisionIds(result.xml);
    expect(new Set(ids).size).toBe(ids.length);
    expect(ids).toContain("5");
    expect(result.xml).toContain(
      '<w:ins w:id="7" w:author="Ada"><w:r><w:t>ROW</w:t></w:r></w:ins>'
    );
    expect(result.xml).toContain(
      '<w:del w:id="8" w:author="Ada"><w:r><w:delText>row</w:delText></w:r></w:del>'
    );
    expect(result.xml).toContain("<w:commentRangeStart w:id=\"0\"/>");
  });

  it("accepts and rejects all revisions", async () => {
    const { pkg, model } = await load(PLAIN);
    const edited = ok(
      replaceParagraphTextTracked(model, {
        nodeIndex: 0,
        startOffset: 6,
        endOffset: 11,
        text: "there",
        author: "Ada",
      })
    );

    const accepted = acceptAllRevisions(edited);
    expect(accepted.remainingNodeIndexes).toEqual([]);
    const acceptedXml = (await exported(accepted.model, pkg)).xml;
    expect(acceptedXml).not.toMatch(/<w:(?:ins|del)\b/);
    expect(acceptedXml).toContain(
      '<w:r><w:rPr><w:b/></w:rPr><w:t>Hello there</w:t></w:r>'
    );

    const rejected = rejectAllRevisions(edited);
    const rejectedResult = await exported(rejected.model, pkg);
    expect(rejectedResult.xml).not.toMatch(/<w:(?:ins|del)\b/);
    expect(paragraphText(rejectedResult.model)).toBe("Hello world");
  });

  it("resolves one revision spanning several runs", async () => {
    const body =
      `<w:p><w:r><w:t xml:space="preserve">A </w:t></w:r>` +
      `<w:ins w:id="3" w:author="Bob"><w:r><w:rPr><w:b/></w:rPr><w:t>bold</w:t></w:r><w:r><w:rPr><w:i/></w:rPr><w:t xml:space="preserve"> italic</w:t></w:r></w:ins>` +
      `<w:del w:id="4" w:author="Bob"><w:r><w:rPr><w:u w:val="single"/></w:rPr><w:delText xml:space="preserve"> under</w:delText></w:r><w:r><w:delText xml:space="preserve"> plain</w:delText></w:r></w:del>` +
      `<w:r><w:t xml:space="preserve"> end</w:t></w:r></w:p>`;
    const { pkg, model } = await load(body);
    expect(paragraphText(model)).toBe("A bold italic end");

    const rejectedInsertion = ok(
      rejectParagraphRevision(model, { nodeIndex: 0, revisionId: 3, kind: "insertion" })
    );
    expect(paragraphText(rejectedInsertion)).toBe("A  end");

    const acceptedInsertion = ok(
      acceptParagraphRevision(model, { nodeIndex: 0, revisionId: 3, kind: "insertion" })
    );
    const rejectedDeletion = ok(
      rejectParagraphRevision(acceptedInsertion, {
        nodeIndex: 0,
        revisionId: "4",
        kind: "deletion",
      })
    );
    const result = await exported(rejectedDeletion, pkg);
    expect(result.xml).not.toMatch(/<w:(?:ins|del)\b/);
    expect(result.xml).toContain(
      '<w:r><w:rPr><w:u w:val="single"/></w:rPr><w:t xml:space="preserve"> under</w:t></w:r>'
    );
    expect(paragraphText(result.model)).toBe("A bold italic under plain end");
    expect(
      rejectParagraphRevision(rejectedDeletion, { nodeIndex: 0, revisionId: 4, kind: "deletion" })
    ).toEqual({ ok: false, reason: "stale" });
  });
});

describe("comment threads", () => {
  it("anchors, replies to and deletes comments in a paragraph with revisions", async () => {
    const { pkg, model } = await load(ANNOTATED_PARAGRAPH, true);

    const created = createParagraphComment(model, {
      nodeIndex: 0,
      startOffset: 9,
      endOffset: 19,
      text: "Colour?",
      author: "Ada",
      initials: "A",
      date: DATE,
    });
    expect(created.ok).toBe(true);
    if (!created.ok) return;
    expect(created.commentId).toBe(1);

    const replied = replyToComment(created.model, {
      commentId: 1,
      text: "Brown is fine.",
      author: "Bob",
      date: DATE,
    });
    expect(replied.ok).toBe(true);
    if (!replied.ok) return;
    expect(replied.commentId).toBe(2);

    const result = await exported(replied.model, pkg);
    expect(result.xml).toContain(
      `<w:del w:id="5" w:author="Bob" w:date="${DATE}"><w:r><w:delText xml:space="preserve"> slow</w:delText></w:r></w:del>` +
        `<w:commentRangeStart w:id="1"/><w:commentRangeStart w:id="2"/>` +
        `<w:ins w:id="6" w:author="Bob" w:date="${DATE}"><w:r><w:t xml:space="preserve"> brown</w:t></w:r></w:ins>` +
        `<w:r><w:t xml:space="preserve"> fox</w:t></w:r>` +
        `<w:commentRangeEnd w:id="1"/><w:commentRangeEnd w:id="2"/>` +
        `<w:r><w:commentReference w:id="1"/></w:r><w:r><w:commentReference w:id="2"/></w:r>`
    );
    expect(result.comments).toContain("Colour?");
    expect(result.comments).toContain("Brown is fine.");
    const comments = result.model.metadata.comments ?? [];
    expect(comments.map((comment) => [comment.id, comment.author, comment.parentId])).toEqual([
      [0, "Carol", undefined],
      [1, "Ada", undefined],
      [2, "Bob", 1],
    ]);

    const deleted = ok(deleteComment(replied.model, 1));
    const afterDelete = await exported(deleted, pkg);
    expect(afterDelete.xml).not.toMatch(/w:id="[12]"\/>/);
    expect(afterDelete.xml).toContain('<w:commentRangeStart w:id="0"/>');
    expect(afterDelete.comments).not.toContain("Colour?");
    expect(afterDelete.comments).not.toContain("Brown is fine.");
    expect(afterDelete.model.metadata.comments?.map((comment) => comment.id)).toEqual([0]);
  });

  it("deletes an imported comment from the body and from comments.xml", async () => {
    const { pkg, model } = await load(ANNOTATED_PARAGRAPH, true);
    const deleted = ok(deleteComment(model, 0));
    const result = await exported(deleted, pkg);
    expect(result.xml).not.toContain("w:commentRange");
    expect(result.xml).not.toContain("w:commentReference");
    expect(result.comments).not.toContain("Is this right?");
    expect(result.commentsExtended).not.toContain("0A000001");
    expect(paragraphText(result.model)).toBe("The quick brown fox");
    expect(result.xml).toContain("> slow</w:delText>");
  });
});
