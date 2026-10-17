import type {
  DocModel,
  DocNode,
  DocumentCommentDefinition,
  ParagraphNode,
  TextLigatures,
  TextStyle,
} from "@extend-ai/react-docx-doc-model";
import { cloneDocModel } from "@extend-ai/react-docx-doc-model";

import {
  createRevisionIdAllocator,
  escapeXmlAttribute,
  formattingAnchorAt,
  insertionIndexAt,
  mapBodySourceXml,
  paragraphSourceForModel,
  parseParagraphSource,
  paragraphWithSource,
  revisionWrapperId,
  splitItemsAt,
  textItemIndexAt,
  textItemOffsets,
  visibleText,
  type ParagraphSource,
  type SourceItem,
  type SourceWrapper,
} from "./paragraph-source";
import {
  directChildElements,
  elementsNamed,
  scanBalancedXml,
  setXmlAttribute,
  xmlAttribute,
  type XmlElementRange,
} from "./xml";

export type AnnotationMutationFailureReason =
  | "stale"
  | "unsupported"
  | "unsafe-xml";

export type AnnotationMutationResult =
  | { ok: true; model: DocModel }
  | { ok: false; reason: AnnotationMutationFailureReason };

export type CommentCreationResult =
  | { ok: true; model: DocModel; commentId: number }
  | { ok: false; reason: AnnotationMutationFailureReason };

export type RevisionBatchResult = {
  ok: true;
  model: DocModel;
  /**
   * Body nodes (paragraphs or tables) that still carry revision markup this
   * batch could not resolve: moves, tables, a break before a table, or
   * markup it cannot parse.
   */
  remainingNodeIndexes: number[];
};

export type ParagraphRevisionKind =
  | "insertion"
  | "deletion"
  | "move-from"
  | "move-to"
  | "format-change"
  | "paragraph-format-change"
  /** An inserted paragraph mark: a paragraph break someone added. */
  | "paragraph-insertion"
  /** A deleted paragraph mark: the paragraph joins the next one. */
  | "paragraph-deletion";

export interface ParagraphRevisionTarget {
  /** Top-level body node containing the revision. Table locations are unsupported. */
  nodeIndex: number;
  /** OOXML `w:id` of the `w:ins` or `w:del` element(s). */
  revisionId: string | number;
  kind: ParagraphRevisionKind;
}

export interface RevisionAuthorship {
  author: string;
  /** ISO 8601 date-time written to `w:date`. */
  date?: string;
}

export interface TrackedTextReplacementInput extends RevisionAuthorship {
  /** Top-level body paragraph index. */
  nodeIndex: number;
  /** UTF-16 offsets into the paragraph's visible text (deleted text excluded). */
  startOffset: number;
  endOffset: number;
  /** Replacement text; empty for a pure deletion. */
  text: string;
}

export interface TrackedParagraphSplitInput extends RevisionAuthorship {
  /** Top-level body paragraph index. */
  nodeIndex: number;
  /** UTF-16 offset into the paragraph's visible text where the break goes. */
  offset: number;
}

export interface TrackedParagraphJoinInput extends RevisionAuthorship {
  /** Top-level body paragraph index; it joins the paragraph after it. */
  nodeIndex: number;
}

export interface CreateParagraphCommentInput {
  /** Top-level body paragraph index. */
  nodeIndex: number;
  /** UTF-16 offsets into the paragraph's plain text, matching DOM selection offsets. */
  startOffset: number;
  endOffset: number;
  text: string;
  author?: string;
  initials?: string;
  date?: string;
}

export interface ReplyToCommentInput {
  /** The comment being answered; replies attach to the thread's root comment. */
  commentId: number;
  text: string;
  author?: string;
  initials?: string;
  date?: string;
}

function booleanElementValue(
  scan: ReturnType<typeof scanBalancedXml>,
  tagName: string
): boolean | undefined {
  const element = elementsNamed(scan, tagName)[0];
  if (!element) {
    return undefined;
  }
  const raw = xmlAttribute(element.openTag, "w:val")?.trim().toLowerCase();
  return raw === undefined || !["0", "false", "off", "none"].includes(raw);
}

function normalizedColor(raw?: string): string | undefined {
  const value = raw?.trim();
  if (!value || value.toLowerCase() === "auto") {
    return undefined;
  }
  return /^[0-9a-f]{6}$/i.test(value) ? `#${value.toLowerCase()}` : value;
}

const TEXT_LIGATURE_VALUES: readonly TextLigatures[] = [
  "none", "standard", "contextual", "historical", "discretional",
  "standardContextual", "standardHistorical", "contextualHistorical",
  "standardDiscretional", "contextualDiscretional", "historicalDiscretional",
  "standardContextualHistorical", "standardContextualDiscretional",
  "standardHistoricalDiscretional", "contextualHistoricalDiscretional", "all",
];

function styleFromRunXml(runXml: string): TextStyle | undefined {
  const scan = scanBalancedXml(runXml);
  if (scan.malformed) {
    return undefined;
  }
  const run = elementsNamed(scan, "w:r")[0];
  const properties = run
    ? directChildElements(scan, run).find(
        (candidate) => candidate.tagName === "w:rPr"
      )
    : undefined;
  if (!properties) {
    return undefined;
  }
  const propertyScan = scanBalancedXml(
    runXml.slice(properties.start, properties.end)
  );
  if (propertyScan.malformed) {
    return undefined;
  }

  const fonts = elementsNamed(propertyScan, "w:rFonts")[0]?.openTag;
  const language = elementsNamed(propertyScan, "w:lang")[0]?.openTag;
  const size = elementsNamed(propertyScan, "w:sz")[0]?.openTag;
  const spacing = elementsNamed(propertyScan, "w:spacing")[0]?.openTag;
  const verticalAlign = elementsNamed(propertyScan, "w:vertAlign")[0]?.openTag;
  const color = elementsNamed(propertyScan, "w:color")[0]?.openTag;
  const highlight = elementsNamed(propertyScan, "w:highlight")[0]?.openTag;
  const shading = elementsNamed(propertyScan, "w:shd")[0]?.openTag;
  const border = elementsNamed(propertyScan, "w:bdr")[0]?.openTag;
  const fontFamilyAscii = fonts ? xmlAttribute(fonts, "w:ascii") : undefined;
  const fontFamilyHAnsi = fonts ? xmlAttribute(fonts, "w:hAnsi") : undefined;
  const fontFamilyEastAsia = fonts
    ? xmlAttribute(fonts, "w:eastAsia")
    : undefined;
  const fontFamilyCs = fonts ? xmlAttribute(fonts, "w:cs") : undefined;
  const fontFamily =
    fontFamilyAscii ?? fontFamilyHAnsi ?? fontFamilyEastAsia ?? fontFamilyCs;
  const kerning = elementsNamed(propertyScan, "w:kern")[0]?.openTag;
  const kerningHalfPoints = Number(kerning ? xmlAttribute(kerning, "w:val") : undefined);
  const ligatures = elementsNamed(propertyScan, "w14:ligatures")[0]?.openTag;
  const ligaturesValue = ligatures ? xmlAttribute(ligatures, "w14:val") : undefined;
  const contextualAlternates = elementsNamed(propertyScan, "w14:cntxtAlts")[0]?.openTag;
  const contextualAlternatesValue = contextualAlternates
    ? xmlAttribute(contextualAlternates, "w14:val")
    : undefined;
  const sizeHalfPoints = Number(size ? xmlAttribute(size, "w:val") : undefined);
  const spacingTwips = Number(
    spacing ? xmlAttribute(spacing, "w:val") : undefined
  );
  const borderSize = Number(border ? xmlAttribute(border, "w:sz") : undefined);
  const borderSpace = Number(
    border ? xmlAttribute(border, "w:space") : undefined
  );
  const borderType = border ? xmlAttribute(border, "w:val") : undefined;

  const style: TextStyle = {
    bold: booleanElementValue(propertyScan, "w:b"),
    italic: booleanElementValue(propertyScan, "w:i"),
    underline: booleanElementValue(propertyScan, "w:u"),
    strike: booleanElementValue(propertyScan, "w:strike"),
    color: normalizedColor(color ? xmlAttribute(color, "w:val") : undefined),
    highlight: highlight ? xmlAttribute(highlight, "w:val") : undefined,
    backgroundColor: normalizedColor(
      shading ? xmlAttribute(shading, "w:fill") : undefined
    ),
    fontSizePt:
      Number.isFinite(sizeHalfPoints) && sizeHalfPoints > 0
        ? sizeHalfPoints / 2
        : undefined,
    fontFamily,
    sourceFontFamily: fontFamily,
    fontFamilyAscii,
    fontFamilyHAnsi,
    fontFamilyEastAsia,
    fontFamilyCs,
    fontThemeAscii: fonts ? xmlAttribute(fonts, "w:asciiTheme") : undefined,
    fontThemeHAnsi: fonts ? xmlAttribute(fonts, "w:hAnsiTheme") : undefined,
    fontThemeEastAsia: fonts
      ? xmlAttribute(fonts, "w:eastAsiaTheme")
      : undefined,
    fontThemeCs: fonts ? xmlAttribute(fonts, "w:csTheme") : undefined,
    fontHint: fonts ? xmlAttribute(fonts, "w:hint") : undefined,
    language: language ? xmlAttribute(language, "w:val") : undefined,
    languageEastAsia: language
      ? xmlAttribute(language, "w:eastAsia")
      : undefined,
    languageBidi: language ? xmlAttribute(language, "w:bidi") : undefined,
    rightToLeft: booleanElementValue(propertyScan, "w:rtl"),
    complexScript: booleanElementValue(propertyScan, "w:cs"),
    ligatures: TEXT_LIGATURE_VALUES.includes(ligaturesValue as TextLigatures)
      ? ligaturesValue as TextLigatures
      : undefined,
    contextualAlternates: !contextualAlternates
      ? undefined
      : contextualAlternatesValue === undefined || ["true", "1"].includes(contextualAlternatesValue)
        ? true
        : ["false", "0"].includes(contextualAlternatesValue)
          ? false
          : undefined,
    kerningMinFontSizePt: Number.isFinite(kerningHalfPoints) && kerningHalfPoints >= 0
      ? kerningHalfPoints / 2
      : undefined,
    characterSpacingTwips: Number.isFinite(spacingTwips)
      ? spacingTwips
      : undefined,
    verticalAlign:
      verticalAlign &&
      ["superscript", "subscript"].includes(
        xmlAttribute(verticalAlign, "w:val") ?? ""
      )
        ? (xmlAttribute(verticalAlign, "w:val") as "superscript" | "subscript")
        : undefined,
    runBorder:
      borderType && borderType !== "none" && borderType !== "nil"
        ? {
            type: borderType,
            color: normalizedColor(xmlAttribute(border!, "w:color")),
            sizeEighthPt: Number.isFinite(borderSize) ? borderSize : undefined,
            spacePt: Number.isFinite(borderSpace) ? borderSpace : undefined,
            frame: border ? xmlAttribute(border, "w:frame") === "1" : undefined,
            shadow: border
              ? xmlAttribute(border, "w:shadow") === "1"
              : undefined,
          }
        : undefined,
  };
  return Object.values(style).some((value) => value !== undefined)
    ? style
    : undefined;
}

export function setCommentResolved(
  model: DocModel,
  commentId: number,
  resolved: boolean
): AnnotationMutationResult {
  const index = model.metadata.comments?.findIndex(
    (comment) => comment.id === commentId
  );
  if (index === undefined || index < 0) {
    return { ok: false, reason: "stale" };
  }
  if (model.metadata.comments?.[index]?.resolved === resolved) {
    return { ok: true, model };
  }
  const next = cloneDocModel(model);
  const comment = next.metadata.comments?.[index];
  if (!comment) {
    return { ok: false, reason: "stale" };
  }
  comment.resolved = resolved;
  comment.resolutionDirty = comment.sourceResolved !== resolved;
  return { ok: true, model: next };
}

function nextCommentId(comments: DocumentCommentDefinition[]): number {
  return (
    comments.reduce((largest, comment) => Math.max(largest, comment.id), -1) + 1
  );
}

function allocateExtendedParagraphId(
  comments: DocumentCommentDefinition[],
  commentId: number
): string {
  const used = new Set(
    comments
      .map((comment) => comment.extendedParagraphId?.toUpperCase())
      .filter((value): value is string => Boolean(value))
  );
  let candidate = (0xc0000000 + Math.max(0, commentId)) >>> 0;
  while (used.has(candidate.toString(16).padStart(8, "0").toUpperCase())) {
    candidate = (candidate + 1) >>> 0;
  }
  return candidate.toString(16).padStart(8, "0").toUpperCase();
}

function isValidXml10Text(value: string): boolean {
  for (let index = 0; index < value.length; index += 1) {
    const codeUnit = value.charCodeAt(index);
    if (codeUnit >= 0xd800 && codeUnit <= 0xdbff) {
      const next = value.charCodeAt(index + 1);
      if (!(next >= 0xdc00 && next <= 0xdfff)) {
        return false;
      }
      index += 1;
      continue;
    }
    if (codeUnit >= 0xdc00 && codeUnit <= 0xdfff) {
      return false;
    }
    if (
      codeUnit !== 0x09 &&
      codeUnit !== 0x0a &&
      codeUnit !== 0x0d &&
      (codeUnit < 0x20 || codeUnit === 0xfffe || codeUnit === 0xffff)
    ) {
      return false;
    }
  }
  return true;
}

function isLeapYear(year: bigint): boolean {
  return year % 4n === 0n && (year % 100n !== 0n || year % 400n === 0n);
}

function isValidXmlDateTime(value: string): boolean {
  const match =
    /^(-?)(\d{4,})-(\d{2})-(\d{2})T(\d{2}):(\d{2}):(\d{2})(?:\.(\d+))?(?:(Z)|([+-])(\d{2}):(\d{2}))?$/.exec(
      value
    );
  if (!match || /^0+$/.test(match[2] ?? "")) {
    return false;
  }

  const year = BigInt(`${match[1] ?? ""}${match[2]}`);
  const month = Number(match[3]);
  const day = Number(match[4]);
  const hour = Number(match[5]);
  const minute = Number(match[6]);
  const second = Number(match[7]);
  const fraction = match[8];
  const zoneHour = match[11] === undefined ? undefined : Number(match[11]);
  const zoneMinute = match[12] === undefined ? undefined : Number(match[12]);
  const daysInMonth = [
    31,
    isLeapYear(year) ? 29 : 28,
    31,
    30,
    31,
    30,
    31,
    31,
    30,
    31,
    30,
    31,
  ];

  if (
    month < 1 ||
    month > 12 ||
    day < 1 ||
    day > (daysInMonth[month - 1] ?? 0) ||
    minute > 59 ||
    second > 59
  ) {
    return false;
  }
  if (
    hour > 24 ||
    (hour === 24 &&
      (minute !== 0 || second !== 0 || Boolean(fraction?.replace(/0/g, ""))))
  ) {
    return false;
  }
  if (
    zoneHour !== undefined &&
    (zoneHour > 14 ||
      (zoneHour === 14 && zoneMinute !== 0) ||
      zoneMinute === undefined ||
      zoneMinute > 59)
  ) {
    return false;
  }
  return true;
}

function isUtf16Boundary(text: string, offset: number): boolean {
  if (offset <= 0 || offset >= text.length) {
    return true;
  }
  const previous = text.charCodeAt(offset - 1);
  const next = text.charCodeAt(offset);
  return !(
    previous >= 0xd800 &&
    previous <= 0xdbff &&
    next >= 0xdc00 &&
    next <= 0xdfff
  );
}

type PreparedParagraph =
  | { ok: true; paragraph: ParagraphNode; source: ParagraphSource }
  | { ok: false; reason: AnnotationMutationFailureReason };

function prepareParagraph(model: DocModel, nodeIndex: number): PreparedParagraph {
  const paragraph = model.nodes[nodeIndex];
  if (!paragraph) {
    return { ok: false, reason: "stale" };
  }
  if (paragraph.type !== "paragraph" || !paragraph.sourceXml) {
    return { ok: false, reason: "unsupported" };
  }
  if (scanBalancedXml(paragraph.sourceXml).malformed) {
    return { ok: false, reason: "unsafe-xml" };
  }
  if (!parseParagraphSource(paragraph.sourceXml)) {
    return { ok: false, reason: "unsupported" };
  }
  const source = paragraphSourceForModel(paragraph);
  if (!source) {
    return { ok: false, reason: "stale" };
  }
  return { ok: true, paragraph, source };
}

function withParagraph(
  model: DocModel,
  nodeIndex: number,
  paragraph: ParagraphNode
): DocModel {
  const nodes = [...model.nodes];
  nodes[nodeIndex] = paragraph;
  return { ...model, nodes };
}

let newWrapperCounter = 0;

function revisionWrapper(
  tagName: "w:ins" | "w:del",
  id: number,
  authorship: RevisionAuthorship
): SourceWrapper {
  newWrapperCounter += 1;
  const date =
    authorship.date === undefined
      ? ""
      : ` w:date="${escapeXmlAttribute(authorship.date)}"`;
  return {
    tagName,
    openTag: `<${tagName} w:id="${id}" w:author="${escapeXmlAttribute(
      authorship.author
    )}"${date}>`,
    key: `new:${newWrapperCounter}`,
  };
}

function isInsertionBy(wrapper: SourceWrapper, author: string): boolean {
  return (
    wrapper.tagName === "w:ins" &&
    xmlAttribute(wrapper.openTag, "w:author") === author
  );
}

function withoutFormatChange(propertiesXml: string): string {
  return propertiesXml.replace(
    /<w:rPrChange\b[^>]*?(?:\/>|>[\s\S]*?<\/w:rPrChange>)/g,
    ""
  );
}

function validRange(text: string, start: number, end: number): boolean {
  return (
    Number.isInteger(start) &&
    Number.isInteger(end) &&
    start >= 0 &&
    end >= start &&
    end <= text.length &&
    isUtf16Boundary(text, start) &&
    isUtf16Boundary(text, end)
  );
}

function validAuthorship(authorship: RevisionAuthorship): boolean {
  return (
    authorship.author.trim().length > 0 &&
    isValidXml10Text(authorship.author) &&
    (authorship.date === undefined || isValidXmlDateTime(authorship.date))
  );
}

/**
 * Replaces `[startOffset, endOffset)` of a paragraph's visible text as a
 * tracked change: removed text becomes `w:del`/`w:delText`, new text a
 * `w:ins`, both copying the surrounding run's formatting. Text the same
 * author inserted earlier is removed outright and new text joins that
 * author's insertion, as Word does. Existing revisions and comment anchors
 * are kept in place.
 */
export function replaceParagraphTextTracked(
  model: DocModel,
  input: TrackedTextReplacementInput
): AnnotationMutationResult {
  if (!validAuthorship(input) || !isValidXml10Text(input.text)) {
    return { ok: false, reason: "unsupported" };
  }
  const prepared = prepareParagraph(model, input.nodeIndex);
  if (!prepared.ok) {
    return prepared;
  }
  const { startOffset: start, endOffset: end } = input;
  const items = splitItemsAt(
    splitItemsAt(prepared.source.items, start),
    end
  );
  const current = visibleText(items);
  if (!validRange(current, start, end)) {
    return { ok: false, reason: "stale" };
  }
  if (start === end && input.text === "") {
    return { ok: true, model };
  }

  const allocate = createRevisionIdAllocator(model);
  const offsets = textItemOffsets(items);
  const anchor =
    start < end ? items[textItemIndexAt(items, start)!] : formattingAnchorAt(items, start);
  const insertAt =
    start < end ? textItemIndexAt(items, end - 1)! + 1 : insertionIndexAt(items, start);

  const inserted: SourceItem[] = [];
  if (input.text) {
    const anchorRun = anchor && anchor.kind !== "marker" ? anchor : undefined;
    const anchorWrappers = (anchorRun?.wrappers ?? []).filter(
      (wrapper) => wrapper.tagName !== "w:del"
    );
    const joinsOwnInsertion = anchorWrappers.some((wrapper) =>
      isInsertionBy(wrapper, input.author)
    );
    const run = {
      openTag: anchorRun?.run.openTag ?? "<w:r>",
      propertiesXml: withoutFormatChange(anchorRun?.run.propertiesXml ?? ""),
    };
    inserted.push({
      kind: "text",
      text: input.text,
      run,
      wrappers: joinsOwnInsertion
        ? anchorWrappers
        : [
            ...anchorWrappers.filter((wrapper) => wrapper.tagName !== "w:ins"),
            revisionWrapper("w:ins", allocate(), input),
          ],
      style:
        anchorRun?.kind === "text"
          ? anchorRun.style
          : styleFromRunXml(`${run.openTag}${run.propertiesXml}</w:r>`),
      link: anchorRun?.kind === "text" ? anchorRun.link : undefined,
    });
  }

  let deletion: SourceWrapper | undefined;
  const next: SourceItem[] = [];
  items.forEach((item, index) => {
    if (index === insertAt) {
      next.push(...inserted);
    }
    const offset = offsets.get(index);
    if (item.kind !== "text" || offset === undefined || offset < start || offset >= end) {
      next.push(item);
      return;
    }
    if (item.wrappers.some((wrapper) => isInsertionBy(wrapper, input.author))) {
      return;
    }
    deletion ??= revisionWrapper("w:del", allocate(), input);
    next.push({
      kind: "deleted",
      text: item.text,
      run: item.run,
      wrappers: [...item.wrappers, deletion],
    });
  });
  if (insertAt >= items.length) {
    next.push(...inserted);
  }

  const paragraph = paragraphWithSource(
    prepared.paragraph,
    { ...prepared.source, items: next },
    allocate
  );
  return { ok: true, model: withParagraph(model, input.nodeIndex, paragraph) };
}

/**
 * A paragraph's properties and its mark's properties. `openXml` (a `<w:p>`
 * start tag and its `w:pPr`) is scanned closed, so ranges index into it.
 */
function scanParagraphOpen(openXml: string) {
  return scanParagraphProperties(`${openXml}</w:p>`);
}

function scanParagraphProperties(paragraphXml: string) {
  const scan = scanBalancedXml(paragraphXml);
  const paragraph = scan.elements.find(
    (element) => element.tagName === "w:p" && element.parentStart === undefined
  );
  const properties =
    paragraph &&
    directChildElements(scan, paragraph).find((element) => element.tagName === "w:pPr");
  const mark =
    properties &&
    directChildElements(scan, properties).find((element) => element.tagName === "w:rPr");
  const childrenOf = (parent: XmlElementRange | undefined) =>
    parent ? directChildElements(scan, parent) : [];
  return { properties, mark, childrenOf };
}

type MarkRevisionTag = "w:ins" | "w:del";

const MARK_REVISION_KIND = {
  "w:ins": "paragraph-insertion",
  "w:del": "paragraph-deletion",
} as const;

/**
 * The `w:ins` or `w:del` on a paragraph's mark: the paragraph break was
 * inserted, or deleted (joining the paragraph with the next), as a tracked
 * change.
 */
function paragraphMarkRevision(
  openXml: string,
  tag: MarkRevisionTag
): XmlElementRange | undefined {
  const { mark, childrenOf } = scanParagraphOpen(openXml);
  return childrenOf(mark).find((element) => element.tagName === tag);
}

export interface ParagraphMarkRevision {
  kind: "paragraph-insertion" | "paragraph-deletion";
  /** OOXML `w:id`. */
  id: string;
  author?: string;
  date?: string;
}

/** Tracked changes to a paragraph's break, read from its `w:p` XML. */
export function readParagraphMarkRevisions(paragraphXml: string): ParagraphMarkRevision[] {
  const { mark, childrenOf } = scanParagraphProperties(paragraphXml);
  return childrenOf(mark).flatMap((element) => {
    const id = xmlAttribute(element.openTag, "w:id");
    return (element.tagName === "w:ins" || element.tagName === "w:del") && id !== undefined
      ? [
          {
            kind: MARK_REVISION_KIND[element.tagName],
            id,
            author: xmlAttribute(element.openTag, "w:author"),
            date: xmlAttribute(element.openTag, "w:date"),
          },
        ]
      : [];
  });
}

function withoutRanges(xml: string, ranges: XmlElementRange[]): string {
  return [...ranges]
    .sort((left, right) => right.start - left.start)
    .reduce((rest, range) => rest.slice(0, range.start) + rest.slice(range.end), xml);
}

/** Drops the mark's `tag` revision, and the mark properties if that empties them. */
function withoutParagraphMarkRevision(openXml: string, tag: MarkRevisionTag): string {
  const { mark, childrenOf } = scanParagraphOpen(openXml);
  const children = childrenOf(mark);
  const revision = children.find((element) => element.tagName === tag);
  if (!mark || !revision) {
    return openXml;
  }
  return withoutRanges(openXml, [children.length === 1 ? mark : revision]);
}

/** Adds `revisionXml` (a self-closing `w:ins` or `w:del`) to the paragraph mark. */
function withParagraphMarkRevision(
  openXml: string,
  revisionXml: string,
  tag: MarkRevisionTag
): string {
  const { properties, mark, childrenOf } = scanParagraphOpen(openXml);
  const insertAt = (at: number, xml: string) => openXml.slice(0, at) + xml + openXml.slice(at);
  if (!properties) {
    return `${openXml}<w:pPr><w:rPr>${revisionXml}</w:rPr></w:pPr>`;
  }
  if (properties.selfClosing) {
    return (
      openXml.slice(0, properties.start) +
      `<w:pPr><w:rPr>${revisionXml}</w:rPr></w:pPr>` +
      openXml.slice(properties.end)
    );
  }
  if (mark?.selfClosing) {
    return openXml.slice(0, mark.start) + `<w:rPr>${revisionXml}</w:rPr>` + openXml.slice(mark.end);
  }
  if (mark) {
    // Revision marks lead the mark properties, `w:ins` before `w:del` (CT_ParaRPr).
    const insertion = tag === "w:del" ? paragraphMarkRevision(openXml, "w:ins") : undefined;
    return insertAt(insertion ? insertion.end : mark.openEnd, revisionXml);
  }
  // The mark properties come before a section break or a property change.
  const next = childrenOf(properties).find(
    (element) => element.tagName === "w:sectPr" || element.tagName === "w:pPrChange"
  );
  return insertAt(next ? next.start : properties.closeStart, `<w:rPr>${revisionXml}</w:rPr>`);
}

/**
 * The opening of the first half of a split paragraph: its mark is new, so it
 * carries `insertionXml` (a self-closing `w:ins`) and none of the original
 * mark's insertion, and the section break stays with the original mark.
 */
function insertedParagraphOpen(openXml: string, insertionXml: string): string {
  const { properties, childrenOf } = scanParagraphOpen(openXml);
  const sections = childrenOf(properties).filter((element) => element.tagName === "w:sectPr");
  const cleared = withoutParagraphMarkRevision(withoutRanges(openXml, sections), "w:ins");
  return withParagraphMarkRevision(cleared, insertionXml, "w:ins");
}

/** Paragraph properties a split cannot copy to two paragraphs and keep valid. */
const UNSPLITTABLE_PARAGRAPH_OPEN =
  /<w:(?:del|moveFrom|moveTo|pPrChange|rPrChange|numberingChange)\b/;

// Imported source XML always quotes attributes with `"`.
const PARAGRAPH_IDENTITY_ATTRIBUTES = /\s+w14:(?:paraId|textId)\s*=\s*"[^"]*"/g;

/** Markers that open a range; at a split they go with the text after them. */
const RANGE_START_MARKER = /^<w:(?:commentRangeStart|bookmarkStart|permStart)\b/;

/**
 * Sections after `nodeIndex` start `delta` nodes later, after a paragraph
 * there was split (+1) or joined with the next one (-1).
 */
function withShiftedSections(model: DocModel, nodeIndex: number, delta: number): DocModel {
  const sections = model.metadata.sections;
  if (!sections?.some((section) => section.startNodeIndex > nodeIndex)) {
    return model;
  }
  return {
    ...model,
    metadata: {
      ...model.metadata,
      sections: sections.map((section) =>
        section.startNodeIndex > nodeIndex
          ? { ...section, startNodeIndex: section.startNodeIndex + delta }
          : section
      ),
    },
  };
}

/**
 * Splits a paragraph at a visible-text offset as a tracked change, as Word
 * records Enter while tracking: the first half ends in a new paragraph mark
 * inside `w:ins`, the second keeps the original mark and properties. The
 * first half keeps the paragraph's identity (`w14:paraId`); the second gets
 * none, so saving stamps a fresh one. Revisions that span the break get a
 * fresh id in the second half, so every revision id stays unique.
 */
export function splitParagraphTracked(
  model: DocModel,
  input: TrackedParagraphSplitInput
): AnnotationMutationResult {
  if (!validAuthorship(input)) {
    return { ok: false, reason: "unsupported" };
  }
  const prepared = prepareParagraph(model, input.nodeIndex);
  if (!prepared.ok) {
    return prepared;
  }
  const { paragraph, source } = prepared;
  if (UNSPLITTABLE_PARAGRAPH_OPEN.test(source.openXml)) {
    return { ok: false, reason: "unsupported" };
  }
  const items = splitItemsAt(source.items, input.offset);
  if (!validRange(visibleText(items), input.offset, input.offset)) {
    return { ok: false, reason: "stale" };
  }
  const allocate = createRevisionIdAllocator(model);
  // Zero-width markers at the cut stay with the text they close, except those
  // opening a range, which go with the text they open.
  let at = insertionIndexAt(items, input.offset);
  while (at > 0 && items[at - 1]!.kind === "marker") {
    at -= 1;
  }
  let end = at;
  while (end < items.length && items[end]!.kind === "marker") {
    end += 1;
  }
  const boundary = items.slice(at, end);
  const opensRange = (item: SourceItem) =>
    item.kind === "marker" && RANGE_START_MARKER.test(item.xml);
  const first = [
    ...items.slice(0, at),
    ...boundary.filter((item) => !opensRange(item)),
  ];
  const rest = [...boundary.filter(opensRange), ...items.slice(end)];
  const firstKeys = new Set(first.flatMap((item) => item.wrappers.map((wrapper) => wrapper.key)));
  const renamed = new Map<string, SourceWrapper>();
  const second = rest.map((item) => ({
    ...item,
    wrappers: item.wrappers.map((wrapper) => {
      if (wrapper.tagName === "w:hyperlink" || !firstKeys.has(wrapper.key)) {
        return wrapper;
      }
      const fresh = renamed.get(wrapper.key) ?? {
        ...wrapper,
        openTag: setXmlAttribute(wrapper.openTag, "w:id", String(allocate())),
        key: `${wrapper.key}:split`,
      };
      renamed.set(wrapper.key, fresh);
      return fresh;
    }),
  })) as SourceItem[];

  const insertion = markRevisionXml("w:ins", allocate(), input);
  const firstHalf = paragraphWithSource(
    paragraph,
    { openXml: insertedParagraphOpen(source.openXml, insertion), items: first },
    allocate
  );
  const secondHalf = paragraphWithSource(
    { ...paragraph, blockId: undefined, paraId: undefined },
    { openXml: source.openXml.replace(PARAGRAPH_IDENTITY_ATTRIBUTES, ""), items: second },
    allocate
  );
  const nodes = [...model.nodes];
  nodes.splice(input.nodeIndex, 1, firstHalf, secondHalf);
  return { ok: true, model: withShiftedSections({ ...model, nodes }, input.nodeIndex, 1) };
}

function markRevisionXml(tag: MarkRevisionTag, id: number, authorship: RevisionAuthorship): string {
  return revisionWrapper(tag, id, authorship).openTag.replace(/>$/, "/>");
}

/**
 * Joins a paragraph with the next one, as Word does when a paragraph mark
 * goes: the joined paragraph keeps the next one's mark and properties and
 * takes this paragraph's identity.
 */
function joinWithNext(model: DocModel, nodeIndex: number): AnnotationMutationResult {
  const first = prepareParagraph(model, nodeIndex);
  if (!first.ok) {
    return first;
  }
  const next = prepareParagraph(model, nodeIndex + 1);
  if (!next.ok) {
    return next.reason === "stale" ? { ok: false, reason: "unsupported" } : next;
  }
  const { paragraph, source } = first;
  // The joined paragraph keeps the next one's properties, so this one's
  // section break would be lost.
  if (/<w:sectPr\b/.test(source.openXml)) {
    return { ok: false, reason: "unsupported" };
  }
  const paraId = xmlAttribute(/^<w:p\b[^>]*>/.exec(source.openXml)?.[0] ?? "", "w14:paraId");
  const nextTag = /^<w:p\b[^>]*>/.exec(next.source.openXml)?.[0] ?? "<w:p>";
  const joinedTag = paraId
    ? setXmlAttribute(nextTag.replace(PARAGRAPH_IDENTITY_ATTRIBUTES, ""), "w14:paraId", paraId)
    : nextTag;
  const joined = paragraphWithSource(
    { ...next.paragraph, blockId: paragraph.blockId, paraId: paragraph.paraId },
    {
      openXml: joinedTag + next.source.openXml.slice(nextTag.length),
      items: [
        ...source.items,
        // Keys come from each paragraph's own offsets; keep them apart.
        ...next.source.items.map((item) => ({
          ...item,
          wrappers: item.wrappers.map((wrapper) => ({ ...wrapper, key: `joined:${wrapper.key}` })),
        })),
      ],
    },
    createRevisionIdAllocator(model)
  );
  const nodes = [...model.nodes];
  nodes.splice(nodeIndex, 2, joined);
  return { ok: true, model: withShiftedSections({ ...model, nodes }, nodeIndex, -1) };
}

/**
 * Decides a tracked change to a paragraph's break. Keeping the break (an
 * accepted insertion or a rejected deletion) drops the revision mark;
 * removing it joins the paragraph with the next one.
 */
function decideParagraphMarkRevision(
  model: DocModel,
  target: ParagraphRevisionTarget,
  decision: RevisionDecision
): AnnotationMutationResult {
  const prepared = prepareParagraph(model, target.nodeIndex);
  if (!prepared.ok) {
    return prepared;
  }
  const { paragraph, source } = prepared;
  const tag: MarkRevisionTag = target.kind === "paragraph-insertion" ? "w:ins" : "w:del";
  const revision = paragraphMarkRevision(source.openXml, tag);
  if (
    !revision ||
    xmlAttribute(revision.openTag, "w:id")?.trim() !== String(target.revisionId).trim()
  ) {
    return { ok: false, reason: "stale" };
  }
  if ((tag === "w:ins") === (decision === "accept")) {
    return {
      ok: true,
      model: withParagraph(
        model,
        target.nodeIndex,
        paragraphWithSource(paragraph, {
          ...source,
          openXml: withoutParagraphMarkRevision(source.openXml, tag),
        })
      ),
    };
  }
  return joinWithNext(model, target.nodeIndex);
}

/**
 * Joins a paragraph with the next one as a tracked change, as Word records
 * Backspace at the start of a paragraph while tracking: the paragraph's mark
 * is marked deleted and both paragraphs stay until the change is decided. A
 * break the same author inserted is removed outright.
 */
export function joinParagraphsTracked(
  model: DocModel,
  input: TrackedParagraphJoinInput
): AnnotationMutationResult {
  if (!validAuthorship(input)) {
    return { ok: false, reason: "unsupported" };
  }
  const prepared = prepareParagraph(model, input.nodeIndex);
  if (!prepared.ok) {
    return prepared;
  }
  const next = prepareParagraph(model, input.nodeIndex + 1);
  if (!next.ok) {
    return next.reason === "stale" ? { ok: false, reason: "unsupported" } : next;
  }
  const { paragraph, source } = prepared;
  // A paragraph ending a section keeps its break: the section would go with it.
  if (/<w:sectPr\b/.test(source.openXml)) {
    return { ok: false, reason: "unsupported" };
  }
  // A break already proposed for removal waits for that decision, even when
  // the author inserted it: joining now would drop the other proposal.
  if (paragraphMarkRevision(source.openXml, "w:del")) {
    return { ok: true, model };
  }
  const insertion = paragraphMarkRevision(source.openXml, "w:ins");
  if (insertion && xmlAttribute(insertion.openTag, "w:author") === input.author) {
    return joinWithNext(model, input.nodeIndex);
  }
  const deletion = markRevisionXml("w:del", createRevisionIdAllocator(model)(), input);
  return {
    ok: true,
    model: withParagraph(
      model,
      input.nodeIndex,
      paragraphWithSource(paragraph, {
        ...source,
        openXml: withParagraphMarkRevision(source.openXml, deletion, "w:del"),
      })
    ),
  };
}

type RevisionDecision = "accept" | "reject";

/**
 * Accepts or rejects the selected `w:ins`/`w:del` wrappers. Kept content
 * loses the wrapper (rejected deletions become visible text again); dropped
 * content is removed. Zero-width markers inside a dropped wrapper stay.
 */
function resolveRevisions(
  items: SourceItem[],
  selected: (wrapper: SourceWrapper) => boolean,
  decision: RevisionDecision
): SourceItem[] {
  const next: SourceItem[] = [];
  for (const item of items) {
    const targets = item.wrappers.filter(selected);
    if (targets.length === 0) {
      next.push(item);
      continue;
    }
    const dropsContent = targets.some(
      (wrapper) => (wrapper.tagName === "w:ins") === (decision === "reject")
    );
    const wrappers = item.wrappers.filter((wrapper) => !selected(wrapper));
    if (item.kind === "marker") {
      next.push({ ...item, wrappers });
      continue;
    }
    if (dropsContent) {
      continue;
    }
    if (
      item.kind === "deleted" &&
      !wrappers.some((wrapper) => wrapper.tagName === "w:del")
    ) {
      next.push({
        kind: "text",
        text: item.text,
        run: item.run,
        wrappers,
        style: styleFromRunXml(
          `${item.run.openTag}${item.run.propertiesXml}</w:r>`
        ),
      });
      continue;
    }
    next.push({ ...item, wrappers });
  }
  return next;
}

function decideParagraphRevision(
  model: DocModel,
  target: ParagraphRevisionTarget,
  decision: RevisionDecision
): AnnotationMutationResult {
  if (target.kind === "paragraph-insertion" || target.kind === "paragraph-deletion") {
    return decideParagraphMarkRevision(model, target, decision);
  }
  const tagName =
    target.kind === "insertion"
      ? "w:ins"
      : target.kind === "deletion"
      ? "w:del"
      : undefined;
  if (!tagName) {
    return { ok: false, reason: "unsupported" };
  }
  const prepared = prepareParagraph(model, target.nodeIndex);
  if (!prepared.ok) {
    return prepared;
  }
  const revisionId = String(target.revisionId).trim();
  const selected = (wrapper: SourceWrapper) =>
    wrapper.tagName === tagName && revisionWrapperId(wrapper) === revisionId;
  if (!prepared.source.items.some((item) => item.wrappers.some(selected))) {
    return { ok: false, reason: "stale" };
  }
  const paragraph = paragraphWithSource(prepared.paragraph, {
    ...prepared.source,
    items: resolveRevisions(prepared.source.items, selected, decision),
  });
  return { ok: true, model: withParagraph(model, target.nodeIndex, paragraph) };
}

/** Accepts every `w:ins`/`w:del` element carrying the target id in one paragraph. */
export function acceptParagraphRevision(
  model: DocModel,
  target: ParagraphRevisionTarget
): AnnotationMutationResult {
  return decideParagraphRevision(model, target, "accept");
}

/** Rejects every `w:ins`/`w:del` element carrying the target id in one paragraph. */
export function rejectParagraphRevision(
  model: DocModel,
  target: ParagraphRevisionTarget
): AnnotationMutationResult {
  return decideParagraphRevision(model, target, "reject");
}

const RUN_REVISION_XML = /<w:(?:ins|del|moveFrom|moveTo)\b/;

function decideAllRevisions(
  model: DocModel,
  decision: RevisionDecision
): RevisionBatchResult {
  let nodes = model.nodes;
  model.nodes.forEach((node, nodeIndex) => {
    if (node.type !== "paragraph" || !RUN_REVISION_XML.test(node.sourceXml ?? "")) {
      return;
    }
    // Paragraph-break revisions (in source.openXml) are decided below.
    const source = paragraphSourceForModel(node);
    if (!source) {
      return;
    }
    const isRevision = (wrapper: SourceWrapper) =>
      wrapper.tagName === "w:ins" || wrapper.tagName === "w:del";
    if (!source.items.some((item) => item.wrappers.some(isRevision))) {
      return;
    }
    if (nodes === model.nodes) {
      nodes = [...model.nodes];
    }
    nodes[nodeIndex] = paragraphWithSource(node, {
      ...source,
      items: resolveRevisions(source.items, isRevision, decision),
    });
  });
  // Paragraph breaks next, from the end, so a decision that joins two
  // paragraphs never moves one still to be decided.
  let decided: DocModel = { ...model, nodes };
  for (let nodeIndex = nodes.length - 1; nodeIndex >= 0; nodeIndex -= 1) {
    const count = decided.nodes.length;
    for (;;) {
      const node = decided.nodes[nodeIndex];
      const [revision] =
        node?.type === "paragraph" && node.sourceXml
          ? readParagraphMarkRevisions(node.sourceXml)
          : [];
      if (!revision) {
        break;
      }
      const result = decideParagraphMarkRevision(
        decided,
        { nodeIndex, revisionId: revision.id, kind: revision.kind },
        decision
      );
      if (!result.ok) {
        break;
      }
      decided = result.model;
      if (decided.nodes.length !== count) {
        break;
      }
    }
  }
  const remainingNodeIndexes = decided.nodes.flatMap((node, nodeIndex) =>
    RUN_REVISION_XML.test(node.sourceXml ?? "") ? [nodeIndex] : []
  );
  return { ok: true, model: decided, remainingNodeIndexes };
}

/**
 * Accepts every run-level insertion and deletion and every tracked paragraph
 * break in top-level paragraphs.
 */
export function acceptAllRevisions(model: DocModel): RevisionBatchResult {
  return decideAllRevisions(model, "accept");
}

/**
 * Rejects every run-level insertion and deletion and every tracked paragraph
 * break in top-level paragraphs.
 */
export function rejectAllRevisions(model: DocModel): RevisionBatchResult {
  return decideAllRevisions(model, "reject");
}

function validCommentFields(input: {
  text: string;
  author?: string;
  initials?: string;
  date?: string;
}): boolean {
  const fields = [input.text, input.author, input.initials, input.date].filter(
    (value): value is string => value !== undefined
  );
  return (
    input.text.trim().length > 0 &&
    fields.every(isValidXml10Text) &&
    (input.date === undefined || isValidXmlDateTime(input.date))
  );
}

function newCommentDefinition(
  comments: DocumentCommentDefinition[],
  input: { text: string; author?: string; initials?: string; date?: string },
  parentId?: number
): DocumentCommentDefinition {
  const id = nextCommentId(comments);
  return {
    id,
    author: input.author,
    initials: input.initials,
    date: input.date,
    text: input.text,
    ...(parentId === undefined ? {} : { parentId }),
    resolved: false,
    extendedParagraphId: allocateExtendedParagraphId(comments, id),
    sourceResolved: false,
    resolutionDirty: false,
    isNew: true,
  };
}

/** Wrapper chain shared by both neighbours, so a new marker never splits a wrapper. */
function sharedWrappers(
  left: SourceItem | undefined,
  right: SourceItem | undefined
): SourceWrapper[] {
  if (!left || !right) {
    return [];
  }
  const shared: SourceWrapper[] = [];
  for (const [index, wrapper] of left.wrappers.entries()) {
    if (wrapper.key !== right.wrappers[index]?.key || wrapper.tagName === "w:del") {
      break;
    }
    shared.push(wrapper);
  }
  return shared;
}

function commentReferenceRunXml(id: number): string {
  return `<w:r><w:commentReference w:id="${id}"/></w:r>`;
}

/**
 * Creates a comment over a visible-text range in one top-level paragraph,
 * including paragraphs that already carry revisions or other comments.
 */
export function createParagraphComment(
  model: DocModel,
  input: CreateParagraphCommentInput
): CommentCreationResult {
  if (!validCommentFields(input)) {
    return { ok: false, reason: "unsupported" };
  }
  const prepared = prepareParagraph(model, input.nodeIndex);
  if (!prepared.ok) {
    return prepared;
  }
  const { startOffset: start, endOffset: end } = input;
  const current = visibleText(prepared.source.items);
  if (!Number.isInteger(start) || !Number.isInteger(end) || start < 0 || end <= start) {
    return { ok: false, reason: "unsupported" };
  }
  if (end > current.length) {
    return { ok: false, reason: "stale" };
  }
  if (!isUtf16Boundary(current, start) || !isUtf16Boundary(current, end)) {
    return { ok: false, reason: "unsupported" };
  }

  const comments = model.metadata.comments ?? [];
  const definition = newCommentDefinition(comments, input);
  const items = splitItemsAt(splitItemsAt(prepared.source.items, start), end);
  const startIndex = textItemIndexAt(items, start)!;
  const endIndex = textItemIndexAt(items, end - 1)! + 1;
  const endWrappers = sharedWrappers(items[endIndex - 1], items[endIndex]);
  items.splice(
    endIndex,
    0,
    { kind: "marker", xml: `<w:commentRangeEnd w:id="${definition.id}"/>`, wrappers: endWrappers },
    { kind: "marker", xml: commentReferenceRunXml(definition.id), wrappers: endWrappers }
  );
  items.splice(startIndex, 0, {
    kind: "marker",
    xml: `<w:commentRangeStart w:id="${definition.id}"/>`,
    wrappers: sharedWrappers(items[startIndex - 1], items[startIndex]),
  });

  const paragraph = paragraphWithSource(prepared.paragraph, {
    ...prepared.source,
    items,
  });
  const next = withParagraph(model, input.nodeIndex, paragraph);
  return {
    ok: true,
    model: {
      ...next,
      metadata: { ...next.metadata, comments: [...comments, definition] },
    },
    commentId: definition.id,
  };
}

function markerId(openTag: string): number | undefined {
  const raw = xmlAttribute(openTag, "w:id")?.trim();
  return raw !== undefined && /^-?\d+$/.test(raw) ? Number(raw) : undefined;
}

function replaceXmlRanges(
  xml: string,
  replacements: Array<{ start: number; end: number; value: string }>
): string {
  let result = xml;
  for (const replacement of [...replacements].sort((a, b) => b.start - a.start)) {
    result = result.slice(0, replacement.start) + replacement.value + result.slice(replacement.end);
  }
  return result;
}

function withReplyMarkers(xml: string, parentId: number, replyId: number): string {
  if (!xml.includes("w:comment")) {
    return xml;
  }
  const scan = scanBalancedXml(xml);
  if (scan.malformed) {
    return xml;
  }
  const replacements: Array<{ start: number; end: number; value: string }> = [];
  for (const element of scan.elements) {
    if (markerId(element.openTag) !== parentId) {
      continue;
    }
    if (element.tagName === "w:commentRangeStart" || element.tagName === "w:commentRangeEnd") {
      replacements.push({
        start: element.end,
        end: element.end,
        value: `<${element.tagName} w:id="${replyId}"/>`,
      });
    } else if (element.tagName === "w:commentReference") {
      const run = scan.elements.find((candidate) => candidate.start === element.parentStart);
      const after = run?.tagName === "w:r" ? run.end : element.end;
      replacements.push({ start: after, end: after, value: commentReferenceRunXml(replyId) });
    }
  }
  return replaceXmlRanges(xml, replacements);
}

/**
 * Adds a Word reply to a comment thread: the reply is linked to the thread's
 * root through `commentsExtended` and gets its own range markers and
 * reference next to the root's.
 */
export function replyToComment(
  model: DocModel,
  input: ReplyToCommentInput
): CommentCreationResult {
  if (!validCommentFields(input)) {
    return { ok: false, reason: "unsupported" };
  }
  const comments = model.metadata.comments ?? [];
  const byId = new Map(comments.map((comment) => [comment.id, comment]));
  let root = byId.get(input.commentId);
  if (!root) {
    return { ok: false, reason: "stale" };
  }
  const visited = new Set<number>();
  while (root.parentId !== undefined && byId.has(root.parentId) && !visited.has(root.id)) {
    visited.add(root.id);
    root = byId.get(root.parentId)!;
  }
  const rootId = root.id;
  const definition = newCommentDefinition(comments, input, rootId);
  return {
    ok: true,
    model: {
      ...model,
      nodes: mapBodySourceXml(model.nodes, (xml) =>
        withReplyMarkers(xml, rootId, definition.id)
      ),
      metadata: { ...model.metadata, comments: [...comments, definition] },
    },
    commentId: definition.id,
  };
}

function withoutCommentMarkers(xml: string, ids: Set<number>): string {
  if (!xml.includes("w:comment")) {
    return xml;
  }
  const scan = scanBalancedXml(xml);
  if (scan.malformed) {
    return xml;
  }
  const replacements: Array<{ start: number; end: number; value: string }> = [];
  for (const element of scan.elements) {
    const id = markerId(element.openTag);
    if (id === undefined || !ids.has(id)) {
      continue;
    }
    if (element.tagName === "w:commentRangeStart" || element.tagName === "w:commentRangeEnd") {
      replacements.push({ start: element.start, end: element.end, value: "" });
    } else if (element.tagName === "w:commentReference") {
      const run = scan.elements.find((candidate) => candidate.start === element.parentStart);
      const referenceOnlyRun =
        run?.tagName === "w:r" &&
        directChildElements(scan, run).every(
          (child) => child === element || child.tagName === "w:rPr"
        );
      const removed = referenceOnlyRun ? run : element;
      replacements.push({ start: removed.start, end: removed.end, value: "" });
    }
  }
  return replaceXmlRanges(xml, replacements);
}

function sourceXmlChanges(nodes: DocNode[], transform: (xml: string) => string): boolean {
  return nodes.some(
    (node) =>
      (node.sourceXml !== undefined && transform(node.sourceXml) !== node.sourceXml) ||
      (node.type === "table" &&
        node.rows.some((row) =>
          row.cells.some((cell) => sourceXmlChanges(cell.nodes as DocNode[], transform))
        ))
  );
}

function mapPartSourceXml<T extends { nodes: DocNode[] }>(
  parts: T[],
  transform: (xml: string) => string
): T[] {
  return parts.map((part) => ({ ...part, nodes: mapBodySourceXml(part.nodes, transform) }));
}

/**
 * Deletes a comment with its range markers and reference, in the body,
 * headers and footers. Deleting a thread's root deletes its replies too, as
 * Word does. A comment anchored in a footnote or endnote is refused
 * ("unsupported"): those parts are written back as imported, so its markers
 * would be left dangling.
 */
export function deleteComment(
  model: DocModel,
  commentId: number
): AnnotationMutationResult {
  const comments = model.metadata.comments ?? [];
  if (!comments.some((comment) => comment.id === commentId)) {
    return { ok: false, reason: "stale" };
  }
  const deleted = new Set([commentId]);
  let grew = true;
  while (grew) {
    grew = false;
    for (const comment of comments) {
      if (comment.parentId !== undefined && deleted.has(comment.parentId) && !deleted.has(comment.id)) {
        deleted.add(comment.id);
        grew = true;
      }
    }
  }
  const strip = (xml: string) => withoutCommentMarkers(xml, deleted);
  const notes = [...(model.metadata.footnotes ?? []), ...(model.metadata.endnotes ?? [])];
  if (notes.some((note) => note.nodes && sourceXmlChanges(note.nodes, strip))) {
    return { ok: false, reason: "unsupported" };
  }
  const { metadata } = model;
  return {
    ok: true,
    model: {
      ...model,
      nodes: mapBodySourceXml(model.nodes, strip),
      metadata: {
        ...metadata,
        headerSections: mapPartSourceXml(metadata.headerSections, strip),
        footerSections: mapPartSourceXml(metadata.footerSections, strip),
        ...(metadata.sections
          ? {
              sections: metadata.sections.map((section) => ({
                ...section,
                headerSections: mapPartSourceXml(section.headerSections, strip),
                footerSections: mapPartSourceXml(section.footerSections, strip),
              })),
            }
          : {}),
        comments: comments.filter((comment) => !deleted.has(comment.id)),
      },
    },
  };
}

export interface ParagraphRevisionInfo {
  /** OOXML `w:id` shared by every element of this revision. */
  id: string;
  /**
   * `paragraph-insertion`: the paragraph ends in a break someone inserted;
   * `paragraph-deletion`: someone removed its break, joining the next one.
   */
  kind: "insertion" | "deletion" | "paragraph-insertion" | "paragraph-deletion";
  author?: string;
  date?: string;
  /** Inserted visible text, or the deleted text; empty for a break. */
  text: string;
}

export interface ParagraphCommentRange {
  commentId: number;
  /** Offsets into the paragraph's visible text. */
  startOffset: number;
  endOffset: number;
  /** Visible text the comment covers in this paragraph. */
  text: string;
}

export interface ParagraphAnnotations {
  revisions: ParagraphRevisionInfo[];
  /** Comments whose range starts in this paragraph. */
  comments: ParagraphCommentRange[];
}

/**
 * Reads the run-level revisions, the paragraph break's revisions, and the comment
 * ranges of one top-level paragraph from its source XML. Paragraphs without
 * readable source report none.
 */
export function readParagraphAnnotations(
  model: DocModel,
  nodeIndex: number
): ParagraphAnnotations {
  const node = model.nodes[nodeIndex];
  const source =
    node?.type === "paragraph" && node.sourceXml
      ? parseParagraphSource(node.sourceXml)
      : undefined;
  if (!source) {
    return { revisions: [], comments: [] };
  }

  const revisions = new Map<string, ParagraphRevisionInfo>();
  const open = new Map<number, ParagraphCommentRange>();
  const comments: ParagraphCommentRange[] = [];
  let offset = 0;
  for (const item of source.items) {
    if (item.kind === "marker") {
      const element = /^<(w:commentRange(?:Start|End))\b[^>]*>/.exec(item.xml);
      const id = element ? markerId(element[0]) : undefined;
      if (id === undefined) {
        continue;
      }
      if (element![1] === "w:commentRangeStart") {
        const range = { commentId: id, startOffset: offset, endOffset: offset, text: "" };
        open.set(id, range);
        comments.push(range);
      } else {
        open.delete(id);
      }
      continue;
    }
    for (const wrapper of item.wrappers) {
      const kind =
        wrapper.tagName === "w:ins" && item.kind === "text"
          ? "insertion"
          : wrapper.tagName === "w:del" && item.kind === "deleted"
          ? "deletion"
          : undefined;
      const id = kind && revisionWrapperId(wrapper);
      if (!kind || id === undefined) {
        continue;
      }
      const key = `${kind}:${id}`;
      const revision = revisions.get(key) ?? {
        id,
        kind,
        author: xmlAttribute(wrapper.openTag, "w:author"),
        date: xmlAttribute(wrapper.openTag, "w:date"),
        text: "",
      };
      revision.text += item.text;
      revisions.set(key, revision);
    }
    if (item.kind === "text") {
      offset += item.text.length;
      for (const range of open.values()) {
        range.endOffset = offset;
        range.text += item.text;
      }
    }
  }
  for (const mark of readParagraphMarkRevisions(`${source.openXml}</w:p>`)) {
    revisions.set(`${mark.kind}:${mark.id}`, { ...mark, text: "" });
  }
  return { revisions: [...revisions.values()], comments };
}
