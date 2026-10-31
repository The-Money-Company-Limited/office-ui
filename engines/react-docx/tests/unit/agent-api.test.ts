import { describe, expect, it } from "vitest";
import {
  createMinimalDocxPackage,
  packageToArrayBuffer,
  withPart,
} from "@extend-ai/react-docx-ooxml-core";
import { openDocx, type Doc } from "@extend-ai/react-docx/agent";

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const W14_NS = "http://schemas.microsoft.com/office/word/2010/wordml";
const DATE = "2026-10-16T09:00:00Z";

const BODY =
  `<w:p w14:paraId="1a2b3c4d"><w:r><w:t>Demonstration of the agent tools</w:t></w:r></w:p>` +
  `<w:p><w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">Hello world, hello again</w:t></w:r></w:p>` +
  `<w:p><w:r><w:t xml:space="preserve">The </w:t></w:r>` +
  `<w:commentRangeStart w:id="0"/><w:r><w:t>quick</w:t></w:r><w:commentRangeEnd w:id="0"/>` +
  `<w:r><w:commentReference w:id="0"/></w:r>` +
  `<w:del w:id="5" w:author="Bob" w:date="${DATE}"><w:r><w:delText xml:space="preserve"> slow</w:delText></w:r></w:del>` +
  `<w:r><w:t xml:space="preserve"> fox</w:t></w:r></w:p>`;

const COMMENTS_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:comments xmlns:w="${W_NS}" xmlns:w14="${W14_NS}"><w:comment w:id="0" w:author="Carol" w:initials="C" w:date="${DATE}"><w:p w14:paraId="0A000001"><w:r><w:t>Is this right?</w:t></w:r></w:p></w:comment></w:comments>`;
const COMMENTS_EXTENDED_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w15:commentsEx xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"><w15:commentEx w15:paraId="0A000001" w15:done="0"/></w15:commentsEx>`;
const RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/><Relationship Id="rId2" Type="http://schemas.microsoft.com/office/2011/relationships/commentsExtended" Target="commentsExtended.xml"/></Relationships>`;

async function fixture(): Promise<Doc> {
  let pkg = createMinimalDocxPackage(
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="${W_NS}" xmlns:w14="${W14_NS}"><w:body>${BODY}</w:body></w:document>`
  );
  pkg = withPart(pkg, { name: "word/comments.xml", content: COMMENTS_XML });
  pkg = withPart(pkg, { name: "word/commentsExtended.xml", content: COMMENTS_EXTENDED_XML });
  pkg = withPart(pkg, { name: "word/_rels/document.xml.rels", content: RELS });
  return openDocx(await packageToArrayBuffer(pkg));
}

async function reopened(doc: Doc): Promise<Doc> {
  return openDocx(await doc.save());
}

describe("agent document API", () => {
  it("lists paragraphs under the viewer's DOM ids and finds text", async () => {
    const doc = await fixture();
    expect(doc.paragraphs()).toEqual([
      { id: "p-1A2B3C4D", index: 0, text: "Demonstration of the agent tools" },
      { id: "p-n1", index: 1, text: "Hello world, hello again" },
      { id: "p-n2", index: 2, text: "The quick fox" },
    ]);
    expect(doc.find("ello")).toEqual([
      { paragraphId: "p-n1", start: 1, end: 5 },
      { paragraphId: "p-n1", start: 14, end: 18 },
    ]);
    expect(doc.find("ello", { occurrence: 2 })).toEqual([
      { paragraphId: "p-n1", start: 14, end: 18 },
    ]);
    expect(doc.find("the", { within: "p-1A2B3C4D" })).toEqual([
      { paragraphId: "p-1A2B3C4D", start: 17, end: 20 },
    ]);
  });

  it("locates click anchors by paraId and index ids, narrowing to their text", async () => {
    const doc = await fixture();
    expect(
      doc.locate({ kind: "element", selector: "#p-n1 > span", text: "hello again" })
    ).toEqual({ paragraphId: "p-n1", index: 1, text: "Hello world, hello again", start: 13, end: 24 });
    expect(
      doc.locate({ kind: "element", selector: "#p-1a2b3c4d-l1 > span:nth-child(2)", text: "not here" })
    ).toEqual({ paragraphId: "p-1A2B3C4D", index: 0, text: "Demonstration of the agent tools" });
    expect(
      doc.locate({ kind: "element", selector: "#doc_p-n2-l3 > span", text: "quick" })
    ).toEqual({ paragraphId: "p-n2", index: 2, text: "The quick fox", start: 4, end: 9 });
    expect(doc.locate({ kind: "element", selector: "#page-1", text: "" })).toBeNull();
    expect(doc.locate({ kind: "element", selector: "#doc_page-1", text: "" })).toBeNull();
    expect(doc.locate({ kind: "element", selector: "#p-n9", text: "" })).toBeNull();
  });

  it("locates ink anchors at the most-inked paragraph", async () => {
    const doc = await fixture();
    expect(
      doc.locate({
        kind: "area",
        selector: "#p-n1",
        text: "",
        ink: {
          marks: [
            {
              line: "under",
              on: [
                { el: "#p-n1", text: "Hello", ink: 20 },
                { el: "#p-1A2B3C4D", text: "Demonstration of…", ink: 100 },
              ],
            },
          ],
        },
      })
    ).toEqual({
      paragraphId: "p-1A2B3C4D",
      index: 0,
      text: "Demonstration of the agent tools",
      start: 0,
      end: 32,
    });
  });

  it("writes suggestions as tracked changes that survive a save", async () => {
    const doc = await fixture();
    const created = doc.suggest({
      paragraphId: "p-n1",
      find: "world",
      replace: "there",
      author: "Agent",
      date: DATE,
    });
    expect(created.map(({ kind, text, author }) => [kind, text, author])).toEqual([
      ["deletion", "world", "Agent"],
      ["insertion", "there", "Agent"],
    ]);
    doc.suggest({ paragraphId: "p-1A2B3C4D", replace: ".", author: "Agent" });

    const after = await reopened(doc);
    const ids = doc.paragraphs().map((paragraph) => paragraph.id);
    expect(after.paragraphs().map((paragraph) => paragraph.text)).toEqual([
      "Demonstration of the agent tools.",
      "Hello there, hello again",
      "The quick fox",
    ]);
    expect(
      after
        .revisions()
        .revisions.map(({ kind, author, text, paragraphId }) => [kind, author, text, paragraphId])
    ).toEqual([
      ["insertion", "Agent", ".", "p-1A2B3C4D"],
      ["deletion", "Agent", "world", ids[1]],
      ["insertion", "Agent", "there", ids[1]],
      ["deletion", "Bob", " slow", ids[2]],
    ]);
  });

  it("accepts and rejects one revision or all of them", async () => {
    const doc = await fixture();
    doc.suggest({ paragraphId: "p-n1", find: "world", replace: "there", author: "Agent" });
    const bob = doc.revisions().revisions.find((revision) => revision.author === "Bob")!;
    expect(doc.reject(bob.id).revisions.map((revision) => revision.author)).toEqual([
      "Agent",
      "Agent",
    ]);
    expect(doc.paragraphs()[2]!.text).toBe("The quick slow fox");

    const accepted = await fixture();
    accepted.suggest({ paragraphId: "p-n1", find: "world", replace: "there", author: "Agent" });
    expect(accepted.accept("all")).toEqual({ revisions: [], unlisted: [] });
    const acceptedAfter = await reopened(accepted);
    expect(acceptedAfter.revisions()).toEqual({ revisions: [], unlisted: [] });
    expect(acceptedAfter.paragraphs().map((paragraph) => paragraph.text)).toEqual([
      "Demonstration of the agent tools",
      "Hello there, hello again",
      "The quick fox",
    ]);

    const rejected = await fixture();
    rejected.suggest({ paragraphId: "p-n1", find: "world", replace: "there", author: "Agent" });
    expect(rejected.reject("all")).toEqual({ revisions: [], unlisted: [] });
    expect((await reopened(rejected)).paragraphs()[1]!.text).toBe("Hello world, hello again");
  });

  it("comments, replies and resolves across a save", async () => {
    const doc = await fixture();
    const { commentId } = doc.comment({
      paragraphId: "p-n1",
      find: "hello again",
      text: "Too casual?",
      author: "Agent Smith",
      date: DATE,
    });
    const reply = doc.reply({ commentId: 0, text: "Yes, it is.", author: "Agent Smith" });
    doc.resolve(0);

    const after = await reopened(doc);
    const ids = doc.paragraphs().map((paragraph) => paragraph.id);
    expect(after.comments()).toEqual([
      expect.objectContaining({
        id: 0,
        author: "Carol",
        text: "Is this right?",
        resolved: true,
        paragraphId: ids[2],
        quote: "quick",
      }),
      expect.objectContaining({
        id: commentId,
        author: "Agent Smith",
        text: "Too casual?",
        resolved: false,
        paragraphId: ids[1],
        quote: "hello again",
      }),
      expect.objectContaining({
        id: reply.commentId,
        text: "Yes, it is.",
        parentId: 0,
        paragraphId: ids[2],
        quote: "quick",
      }),
    ]);
  });

  it("gives every paragraph a stable paraId id when it saves", async () => {
    const doc = await fixture();
    await doc.save();
    const ids = doc.paragraphs().map((paragraph) => paragraph.id);
    expect(ids[0]).toBe("p-1A2B3C4D");
    for (const id of ids) {
      expect(id).toMatch(/^p-[0-7][0-9A-F]{7}$/);
    }
    expect(new Set(ids).size).toBe(3);

    // Edits keep the ids, in this session and in the file it writes next.
    doc.suggest({ paragraphId: ids[1]!, find: "world", replace: "there", author: "Agent" });
    const after = await reopened(doc);
    expect(after.paragraphs().map((paragraph) => paragraph.id)).toEqual(ids);
    expect(doc.paragraphs().map((paragraph) => paragraph.id)).toEqual(ids);
  });

  it("suggests paragraph breaks that survive a save and reject back to one paragraph", async () => {
    const doc = await fixture();
    const created = doc.suggest({
      paragraphId: "p-n1",
      find: ", ",
      replace: ".\nAnd ",
      author: "Agent",
      date: DATE,
    });
    expect(created.map(({ kind, text }) => [kind, text])).toEqual([
      ["deletion", ", "],
      ["insertion", "."],
      ["paragraph-insertion", ""],
      ["insertion", "And "],
    ]);
    const after = await reopened(doc);
    expect(after.paragraphs().map((paragraph) => paragraph.text)).toEqual([
      "Demonstration of the agent tools",
      "Hello world.",
      "And hello again",
      "The quick fox",
    ]);
    const breaks = after
      .revisions()
      .revisions.filter((revision) => revision.kind === "paragraph-insertion");
    expect(breaks).toHaveLength(1);
    after.reject(breaks[0]!.id);
    expect(after.paragraphs().map((paragraph) => paragraph.text)[1]).toBe(
      "Hello world.And hello again"
    );
    after.reject("all");
    expect(after.paragraphs()[1]!.text).toBe("Hello world, hello again");
  });

  it("lists a break Word marked deleted and joins the paragraphs on accept", async () => {
    const pkg = createMinimalDocxPackage(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="${W_NS}"><w:body>` +
        `<w:p><w:pPr><w:rPr><w:del w:id="9" w:author="Bob" w:date="${DATE}"/></w:rPr></w:pPr><w:r><w:t>One</w:t></w:r></w:p>` +
        `<w:p><w:r><w:t xml:space="preserve"> two</w:t></w:r></w:p>` +
        `</w:body></w:document>`
    );
    const doc = await openDocx(await packageToArrayBuffer(pkg));
    expect(doc.revisions()).toEqual({
      revisions: [
        { id: "9", kind: "paragraph-deletion", author: "Bob", date: DATE, text: "", paragraphId: "p-n0" },
      ],
      unlisted: [],
    });
    doc.accept("9");
    expect((await reopened(doc)).paragraphs().map((paragraph) => paragraph.text)).toEqual(["One two"]);
  });

  it("reports revisions it cannot list instead of claiming none remain", async () => {
    const table =
      `<w:tbl><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc>` +
      `<w:p><w:ins w:id="7" w:author="Bob" w:date="${DATE}"><w:r><w:t>cell</w:t></w:r></w:ins></w:p>` +
      `</w:tc></w:tr></w:tbl>`;
    const pkg = createMinimalDocxPackage(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="${W_NS}"><w:body>` +
        `<w:p><w:r><w:t>Before</w:t></w:r></w:p>${table}<w:p><w:r><w:t>After</w:t></w:r></w:p>` +
        `</w:body></w:document>`
    );
    const doc = await openDocx(await packageToArrayBuffer(pkg));
    const unlisted = [{ index: 1, kind: "table" }];
    expect(doc.revisions()).toEqual({ revisions: [], unlisted });
    expect(doc.accept("all")).toEqual({ revisions: [], unlisted });
  });

  it("reads body tables and each header and footer once", async () => {
    const R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    const part = (tag: string, text: string) =>
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:${tag} xmlns:w="${W_NS}"><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:${tag}>`;
    const cell = (text: string) => `<w:tc><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`;
    const sectionParts =
      `<w:headerReference w:type="default" r:id="rId1"/><w:footerReference w:type="first" r:id="rId2"/>`;
    let pkg = createMinimalDocxPackage(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="${W_NS}" xmlns:r="${R_NS}"><w:body>` +
        `<w:p><w:pPr><w:sectPr>${sectionParts}</w:sectPr></w:pPr><w:r><w:t>Before</w:t></w:r></w:p>` +
        `<w:tbl><w:tr>${cell("Name")}${cell("Total")}</w:tr><w:tr>${cell("Ada")}` +
        `<w:tc><w:p><w:r><w:t>3</w:t></w:r></w:p><w:p><w:r><w:t>4</w:t></w:r></w:p></w:tc></w:tr></w:tbl>` +
        `<w:sectPr>${sectionParts}</w:sectPr></w:body></w:document>`
    );
    pkg = withPart(pkg, { name: "word/header1.xml", content: part("hdr", "Quarterly report") });
    pkg = withPart(pkg, { name: "word/footer1.xml", content: part("ftr", "Confidential") });
    pkg = withPart(pkg, {
      name: "word/_rels/document.xml.rels",
      content:
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rId1" Type="${R_NS}/header" Target="header1.xml"/>` +
        `<Relationship Id="rId2" Type="${R_NS}/footer" Target="footer1.xml"/></Relationships>`,
    });
    const doc = await openDocx(await packageToArrayBuffer(pkg));

    expect(doc.tables()).toEqual([{ index: 1, rows: [["Name", "Total"], ["Ada", "3\n4"]] }]);
    expect(doc.headersFooters()).toEqual([
      { kind: "header", pages: ["default"], text: "Quarterly report" },
      { kind: "footer", pages: ["first"], text: "Confidential" },
    ]);
  });

  it("explains a target it cannot find", async () => {
    const doc = await fixture();
    expect(() =>
      doc.suggest({ paragraphId: "p-n1", find: "absent", replace: "x", author: "Agent" })
    ).toThrow(/"absent" is not in paragraph p-n1/);
    expect(() => doc.accept("99")).toThrow(/no revision 99/);
  });
});
