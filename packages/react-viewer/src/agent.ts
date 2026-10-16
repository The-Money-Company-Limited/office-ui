/**
 * DOM-free document API for agents: open DOCX bytes, find text by the ids
 * the viewer renders, answer with Word tracked changes and comments, and
 * save DOCX bytes. Runs in Node with no browser and no React.
 */
import {
  buildDocModel,
  type DocModel,
  type ParagraphNode,
} from "@extend-ai/react-docx-doc-model";
import {
  acceptAllRevisions,
  acceptParagraphRevision,
  createParagraphComment,
  readParagraphAnnotations,
  rejectAllRevisions,
  rejectParagraphRevision,
  replaceParagraphTextTracked,
  replyToComment,
  setCommentResolved,
  type AnnotationMutationResult,
} from "@extend-ai/react-docx-editor-ops";
import { parseDocx, type OoxmlPackage } from "@extend-ai/react-docx-ooxml-core";
import { serializeDocx } from "@extend-ai/react-docx-serializer";

export interface DocParagraph {
  /** The paragraph's DOM id in the viewer: `p-<w14:paraId>` or `p-n<index>`. */
  id: string;
  /** Index in the document body. */
  index: number;
  text: string;
}

/** Offsets into a paragraph's visible text (tracked deletions excluded). */
export interface TextRange {
  paragraphId: string;
  start: number;
  end: number;
}

export interface FindOptions {
  /** 1-based; returns only that match. */
  occurrence?: number;
  /** Paragraph id to search in. */
  within?: string;
}

/** A comment-layer anchor: a clicked element or an inked area. */
export interface RuntimeAnchor {
  kind: "element" | "area" | string;
  selector?: string;
  text?: string;
  ink?: {
    marks?: Array<{
      line?: "under" | "through" | "around" | string;
      on?: Array<{ el: string; text?: string; ink?: number }>;
    }>;
  };
}

export interface LocatedText {
  paragraphId: string;
  index: number;
  /** The paragraph's text. */
  text: string;
  start?: number;
  end?: number;
}

export interface SuggestInput {
  paragraphId?: string;
  range?: TextRange;
  /** Text to replace inside `paragraphId`; empty or absent inserts at the end. */
  find?: string;
  replace: string;
  author: string;
  date?: string;
}

export interface CommentInput {
  range?: TextRange;
  paragraphId?: string;
  /** Text to comment on inside `paragraphId`; absent covers the paragraph. */
  find?: string;
  text: string;
  author: string;
  date?: string;
}

export interface ReplyInput {
  commentId: number;
  text: string;
  author: string;
  date?: string;
}

export interface DocComment {
  id: number;
  author?: string;
  date?: string;
  text: string;
  resolved: boolean;
  parentId?: number;
  /** Paragraph where the comment's range starts, when it has one. */
  paragraphId?: string;
  /** Text the comment covers in that paragraph. */
  quote?: string;
}

export interface DocRevision {
  /** OOXML `w:id`. */
  id: string;
  kind: "insertion" | "deletion";
  author?: string;
  date?: string;
  text: string;
  paragraphId: string;
}

const PARAGRAPH_PARA_ID = /^<w:p\b[^>]*?\bw14:paraId="([0-9A-Fa-f]{1,8})"/;
/** A viewer paragraph id at the end of a DOM id, without a `-l<line>` segment suffix. */
const PARAGRAPH_DOM_ID = /(?:^|-)(p-(?:n\d+|[0-9A-Fa-f]{1,8}))(?:-l\d+)?$/;

function paragraphId(node: ParagraphNode, index: number): string {
  const paraId = node.sourceXml?.match(PARAGRAPH_PARA_ID)?.[1];
  return paraId ? `p-${paraId.toUpperCase()}` : `p-n${index}`;
}

function paragraphText(node: ParagraphNode): string {
  return node.children
    .map((child) => (child.type === "text" ? child.text : ""))
    .join("");
}

/** `#p-n1 > span` → `p-n1`; `#p-0A1B2C3D-l2` → `p-0A1B2C3D`. */
function selectorParagraphId(selector: string | undefined): string | undefined {
  const id = selector?.trim().match(/^#([^\s>+~:.[#]+)/)?.[1];
  return id?.match(PARAGRAPH_DOM_ID)?.[1];
}

function required<T extends { ok: true; model: DocModel }>(
  result: T | Exclude<AnnotationMutationResult, { ok: true }>,
  action: string
): T {
  if (!result.ok) {
    throw new Error(
      `${action} failed (${result.reason}): ${
        result.reason === "stale"
          ? "the target no longer matches the document; re-read it and retry"
          : "this paragraph's markup is not supported for this edit"
      }`
    );
  }
  return result;
}

export class Doc {
  private constructor(
    private model: DocModel,
    private readonly pkg: OoxmlPackage
  ) {}

  static async open(bytes: Uint8Array | ArrayBuffer): Promise<Doc> {
    const buffer =
      bytes instanceof Uint8Array
        ? (bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer)
        : bytes;
    const pkg = await parseDocx(buffer);
    return new Doc(await buildDocModel(pkg), pkg);
  }

  async save(): Promise<Uint8Array> {
    return new Uint8Array(await serializeDocx(this.model, this.pkg));
  }

  paragraphs(): DocParagraph[] {
    const result: DocParagraph[] = [];
    this.model.nodes.forEach((node, index) => {
      if (node.type === "paragraph") {
        result.push({ id: paragraphId(node, index), index, text: paragraphText(node) });
      }
    });
    return result;
  }

  find(text: string, options: FindOptions = {}): TextRange[] {
    if (!text) {
      return [];
    }
    const matches: TextRange[] = [];
    for (const paragraph of this.paragraphs()) {
      if (options.within !== undefined && paragraph.id !== options.within) {
        continue;
      }
      for (let at = paragraph.text.indexOf(text); at >= 0; at = paragraph.text.indexOf(text, at + text.length)) {
        matches.push({ paragraphId: paragraph.id, start: at, end: at + text.length });
      }
    }
    if (options.occurrence === undefined) {
      return matches;
    }
    const match = matches[options.occurrence - 1];
    return match ? [match] : [];
  }

  locate(anchor: RuntimeAnchor): LocatedText | null {
    const marks = anchor.ink?.marks ?? [];
    if (marks.length > 0) {
      const inkByParagraph = new Map<string, number>();
      for (const mark of marks) {
        for (const on of mark.on ?? []) {
          const id = selectorParagraphId(on.el);
          if (id) {
            inkByParagraph.set(id, (inkByParagraph.get(id) ?? 0) + (on.ink ?? 0));
          }
        }
      }
      const [best] = [...inkByParagraph].sort((a, b) => b[1] - a[1]);
      const paragraph = this.paragraph(best?.[0] ?? selectorParagraphId(anchor.selector));
      if (!paragraph) {
        return null;
      }
      const marksText = marks.some((mark) => mark.line === "under" || mark.line === "through");
      return {
        paragraphId: paragraph.id,
        index: paragraph.index,
        text: paragraph.text,
        ...(marksText ? { start: 0, end: paragraph.text.length } : {}),
      };
    }

    const paragraph = this.paragraph(selectorParagraphId(anchor.selector));
    if (!paragraph) {
      return null;
    }
    const located: LocatedText = {
      paragraphId: paragraph.id,
      index: paragraph.index,
      text: paragraph.text,
    };
    const quote = anchor.text?.trim();
    const start = quote ? paragraph.text.indexOf(quote) : -1;
    return start >= 0 ? { ...located, start, end: start + quote!.length } : located;
  }

  suggest(input: SuggestInput): DocRevision[] {
    const range = this.targetRange(input.range, input.paragraphId, input.find, "end");
    const before = new Set(this.revisions().map(revisionKey));
    this.model = required(
      replaceParagraphTextTracked(this.model, {
        nodeIndex: this.paragraphOrThrow(range.paragraphId).index,
        startOffset: range.start,
        endOffset: range.end,
        text: input.replace,
        author: input.author,
        date: input.date,
      }),
      "suggest"
    ).model;
    return this.revisions().filter((revision) => !before.has(revisionKey(revision)));
  }

  comment(input: CommentInput): { commentId: number } {
    const range = this.targetRange(input.range, input.paragraphId, input.find, "whole");
    const result = required(createParagraphComment(this.model, {
      nodeIndex: this.paragraphOrThrow(range.paragraphId).index,
      startOffset: range.start,
      endOffset: range.end,
      text: input.text,
      author: input.author,
      initials: initials(input.author),
      date: input.date,
    }), "comment");
    this.model = result.model;
    return { commentId: result.commentId };
  }

  reply(input: ReplyInput): { commentId: number } {
    const result = required(replyToComment(this.model, {
      commentId: input.commentId,
      text: input.text,
      author: input.author,
      initials: initials(input.author),
      date: input.date,
    }), "reply");
    this.model = result.model;
    return { commentId: result.commentId };
  }

  resolve(commentId: number, resolved = true): void {
    this.model = required(setCommentResolved(this.model, commentId, resolved), "resolve").model;
  }

  comments(): DocComment[] {
    const anchors = new Map<number, { paragraphId: string; quote: string }>();
    for (const paragraph of this.paragraphs()) {
      for (const range of readParagraphAnnotations(this.model, paragraph.index).comments) {
        anchors.set(range.commentId, { paragraphId: paragraph.id, quote: range.text });
      }
    }
    return (this.model.metadata.comments ?? []).map((comment) => ({
      id: comment.id,
      author: comment.author,
      date: comment.date,
      text: comment.text,
      resolved: comment.resolved ?? false,
      parentId: comment.parentId,
      ...anchors.get(comment.id),
    }));
  }

  revisions(): DocRevision[] {
    return this.paragraphs().flatMap((paragraph) =>
      readParagraphAnnotations(this.model, paragraph.index).revisions.map((revision) => ({
        ...revision,
        paragraphId: paragraph.id,
      }))
    );
  }

  /** Accepts one revision (every paragraph carrying its id) or all; returns what remains. */
  accept(id: string | number): DocRevision[] {
    return this.decide(id, "accept");
  }

  /** Rejects one revision (every paragraph carrying its id) or all; returns what remains. */
  reject(id: string | number): DocRevision[] {
    return this.decide(id, "reject");
  }

  private decide(id: string | number, decision: "accept" | "reject"): DocRevision[] {
    if (id === "all") {
      this.model = (decision === "accept" ? acceptAllRevisions : rejectAllRevisions)(this.model).model;
      return this.revisions();
    }
    const targets = this.revisions().filter((revision) => revision.id === String(id));
    if (targets.length === 0) {
      throw new Error(`${decision} failed: no revision ${id}; list revisions() for current ids`);
    }
    const apply = decision === "accept" ? acceptParagraphRevision : rejectParagraphRevision;
    for (const target of targets) {
      this.model = required(
        apply(this.model, {
          nodeIndex: this.paragraphOrThrow(target.paragraphId).index,
          revisionId: target.id,
          kind: target.kind,
        }),
        decision
      ).model;
    }
    return this.revisions();
  }

  private paragraph(id: string | undefined): DocParagraph | undefined {
    return id === undefined
      ? undefined
      : this.paragraphs().find((paragraph) => paragraph.id.toUpperCase() === id.toUpperCase());
  }

  private paragraphOrThrow(id: string): DocParagraph {
    const paragraph = this.paragraph(id);
    if (!paragraph) {
      throw new Error(`no paragraph ${id}; list paragraphs() for current ids`);
    }
    return paragraph;
  }

  /** A range, `find` inside a paragraph, or (without `find`) its end or whole text. */
  private targetRange(
    range: TextRange | undefined,
    paragraphId: string | undefined,
    find: string | undefined,
    empty: "end" | "whole"
  ): TextRange {
    if (range) {
      return range;
    }
    if (paragraphId === undefined) {
      throw new Error("pass a range or a paragraphId");
    }
    const paragraph = this.paragraphOrThrow(paragraphId);
    if (!find) {
      const end = paragraph.text.length;
      return { paragraphId: paragraph.id, start: empty === "end" ? end : 0, end };
    }
    const [match] = this.find(find, { within: paragraph.id, occurrence: 1 });
    if (!match) {
      throw new Error(`"${find}" is not in paragraph ${paragraph.id}; read its text with paragraphs()`);
    }
    return match;
  }
}

function revisionKey(revision: DocRevision): string {
  return `${revision.paragraphId}:${revision.kind}:${revision.id}`;
}

function initials(author: string): string | undefined {
  const letters = author
    .split(/\s+/)
    .map((word) => word[0] ?? "")
    .join("")
    .toUpperCase();
  return letters || undefined;
}

export function openDocx(bytes: Uint8Array | ArrayBuffer): Promise<Doc> {
  return Doc.open(bytes);
}
