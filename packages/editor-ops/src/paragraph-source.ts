import type {
  DocModel,
  DocNode,
  ParagraphNode,
  TextRunNode,
  TextStyle,
} from "@extend-ai/react-docx-doc-model";

import {
  decodeXmlText,
  escapeXmlText,
  scanBalancedXml,
  setXmlAttribute,
  xmlAttribute,
  type XmlElementRange,
} from "./xml";

/**
 * A paragraph's source XML read as a flat sequence of items. Visible text,
 * deleted text (`w:delText`) and zero-width markers (comment ranges,
 * bookmarks, comment reference runs) keep their order, and each item carries
 * the chain of revision/hyperlink wrappers it sits in. Edits work on this
 * sequence and `emitParagraphSource` writes it back as XML, so revisions and
 * comment anchors that the doc model does not represent survive every edit.
 */

export type WrapperTagName = "w:ins" | "w:del" | "w:hyperlink";

export interface SourceWrapper {
  tagName: WrapperTagName;
  openTag: string;
  /** Identity of one wrapper element; items sharing it are emitted inside one element. */
  key: string;
}

export interface SourceRunTemplate {
  openTag: string;
  propertiesXml: string;
}

export interface SourceTextItem {
  kind: "text";
  text: string;
  run: SourceRunTemplate;
  wrappers: SourceWrapper[];
  style?: TextStyle;
  link?: string;
}

export interface SourceDeletedItem {
  kind: "deleted";
  text: string;
  run: SourceRunTemplate;
  wrappers: SourceWrapper[];
}

export interface SourceMarkerItem {
  kind: "marker";
  xml: string;
  wrappers: SourceWrapper[];
}

export type SourceItem = SourceTextItem | SourceDeletedItem | SourceMarkerItem;

export interface ParagraphSource {
  /** `<w:p ...>` plus its `w:pPr`, written back unchanged. */
  openXml: string;
  items: SourceItem[];
}

const MARKER_ELEMENTS = new Set([
  "w:commentRangeStart",
  "w:commentRangeEnd",
  "w:bookmarkStart",
  "w:bookmarkEnd",
  "w:proofErr",
  "w:permStart",
  "w:permEnd",
]);

const WRAPPER_ELEMENTS = new Set<string>(["w:ins", "w:del", "w:hyperlink"]);

/** Markup the plain-text editing path would drop when it regenerates a paragraph. */
export const ANNOTATION_MARKUP_XML =
  /<w:(?:ins|del|commentRangeStart|commentRangeEnd|commentReference)\b/;

function elementText(xml: string, element: XmlElementRange): string | undefined {
  if (element.selfClosing) {
    return "";
  }
  const content = xml.slice(element.openEnd, element.closeStart);
  return content.includes("<") ? undefined : decodeXmlText(content);
}

/** Parses paragraph source XML, or returns undefined for content it cannot rewrite safely. */
export function parseParagraphSource(
  sourceXml: string
): ParagraphSource | undefined {
  const scan = scanBalancedXml(sourceXml);
  if (scan.malformed) {
    return undefined;
  }
  const children = new Map<number, XmlElementRange[]>();
  const roots: XmlElementRange[] = [];
  for (const element of scan.elements) {
    if (element.parentStart === undefined) {
      roots.push(element);
      continue;
    }
    const siblings = children.get(element.parentStart) ?? [];
    siblings.push(element);
    children.set(element.parentStart, siblings);
  }
  const paragraph = roots[0];
  if (
    roots.length !== 1 ||
    !paragraph ||
    paragraph.tagName !== "w:p" ||
    paragraph.selfClosing ||
    sourceXml.slice(0, paragraph.start).trim() ||
    sourceXml.slice(paragraph.end).trim()
  ) {
    return undefined;
  }
  const childrenOf = (element: XmlElementRange) =>
    children.get(element.start) ?? [];
  const onlyElementsInside = (element: XmlElementRange): boolean => {
    let cursor = element.openEnd;
    for (const child of childrenOf(element)) {
      if (sourceXml.slice(cursor, child.start).trim()) {
        return false;
      }
      cursor = child.end;
    }
    return !sourceXml.slice(cursor, element.closeStart).trim();
  };

  const items: SourceItem[] = [];
  let openEnd = paragraph.openEnd;

  const visitRun = (run: XmlElementRange, wrappers: SourceWrapper[]): boolean => {
    if (run.selfClosing) {
      return true;
    }
    if (!onlyElementsInside(run)) {
      return false;
    }
    const deleted = wrappers.some((wrapper) => wrapper.tagName === "w:del");
    const template: SourceRunTemplate = { openTag: run.openTag, propertiesXml: "" };
    const runChildren = childrenOf(run);
    for (const [index, child] of runChildren.entries()) {
      switch (child.tagName) {
        case "w:rPr":
          if (index !== 0) {
            return false;
          }
          template.propertiesXml = sourceXml.slice(child.start, child.end);
          break;
        case "w:t":
        case "w:delText":
        case "w:tab": {
          if ((child.tagName === "w:delText") !== deleted && child.tagName !== "w:tab") {
            return false;
          }
          const text = child.tagName === "w:tab" ? "\t" : elementText(sourceXml, child);
          if (text === undefined) {
            return false;
          }
          if (text) {
            items.push(
              deleted
                ? { kind: "deleted", text, run: template, wrappers }
                : { kind: "text", text, run: template, wrappers }
            );
          }
          break;
        }
        case "w:commentReference":
        case "w:annotationRef":
          items.push({
            kind: "marker",
            xml: `${run.openTag}${template.propertiesXml}${sourceXml.slice(
              child.start,
              child.end
            )}</w:r>`,
            wrappers,
          });
          break;
        case "w:lastRenderedPageBreak":
          // A layout cache hint Word recomputes; it carries no content.
          break;
        default:
          return false;
      }
    }
    return true;
  };

  const visit = (element: XmlElementRange, wrappers: SourceWrapper[]): boolean => {
    if (!onlyElementsInside(element)) {
      return false;
    }
    for (const child of childrenOf(element)) {
      if (child.tagName === "w:r") {
        if (!visitRun(child, wrappers)) {
          return false;
        }
      } else if (WRAPPER_ELEMENTS.has(child.tagName)) {
        const wrapper: SourceWrapper = {
          tagName: child.tagName as WrapperTagName,
          openTag: child.openTag,
          key: `source:${child.start}`,
        };
        if (!child.selfClosing && !visit(child, [...wrappers, wrapper])) {
          return false;
        }
      } else if (MARKER_ELEMENTS.has(child.tagName)) {
        items.push({
          kind: "marker",
          xml: sourceXml.slice(child.start, child.end),
          wrappers,
        });
      } else if (
        child.tagName === "w:pPr" &&
        element === paragraph &&
        child === childrenOf(paragraph)[0]
      ) {
        openEnd = child.end;
      } else {
        return false;
      }
    }
    return true;
  };

  if (!visit(paragraph, [])) {
    return undefined;
  }
  return { openXml: sourceXml.slice(paragraph.start, openEnd), items };
}

/**
 * Parses a paragraph and attaches the modeled style/link of each visible
 * token. Returns undefined when the model no longer matches its source XML.
 */
export function paragraphSourceForModel(
  paragraph: ParagraphNode
): ParagraphSource | undefined {
  if (!paragraph.sourceXml) {
    return undefined;
  }
  const source = parseParagraphSource(materializedSourceXml(paragraph));
  if (!source) {
    return undefined;
  }
  const runs = paragraph.children;
  if (runs.some((run) => run.type !== "text" || run.noteReference)) {
    return undefined;
  }
  const textRuns = (runs as TextRunNode[]).filter((run) => run.text.length > 0);
  const textItems = source.items.filter(
    (item): item is SourceTextItem => item.kind === "text"
  );
  if (
    textItems.length !== textRuns.length ||
    textItems.some((item, index) => item.text !== textRuns[index]?.text)
  ) {
    return undefined;
  }
  const provenance = paragraph.sourceRunProvenance?.runs;
  if (
    provenance &&
    (provenance.length !== runs.length ||
      (runs as TextRunNode[]).some(
        (run, index) =>
          JSON.stringify(canonical(run.style)) !==
            JSON.stringify(canonical(provenance[index]?.style)) ||
          run.link !== provenance[index]?.link
      ))
  ) {
    return undefined;
  }
  textItems.forEach((item, index) => {
    item.style = textRuns[index]?.style;
    item.link = textRuns[index]?.link;
  });
  return source;
}

function canonical(value: unknown): unknown {
  if (Array.isArray(value)) {
    return value.map(canonical);
  }
  if (value && typeof value === "object") {
    return Object.fromEntries(
      Object.entries(value as Record<string, unknown>)
        .filter(([, entry]) => entry !== undefined)
        .sort(([left], [right]) => left.localeCompare(right))
        .map(([key, entry]) => [key, canonical(entry)])
    );
  }
  return value;
}

/**
 * Source XML with a pending plain-text patch applied: the patch is only
 * recorded for one-`w:t`-per-run paragraphs, so run N takes child N's text.
 */
function materializedSourceXml(paragraph: ParagraphNode): string {
  const sourceXml = paragraph.sourceXml ?? "";
  if (!paragraph.sourceTextPatch) {
    return sourceXml;
  }
  const scan = scanBalancedXml(sourceXml);
  const texts = scan.elements.filter((element) => element.tagName === "w:t");
  let result = sourceXml;
  for (let index = texts.length - 1; index >= 0; index -= 1) {
    const element = texts[index]!;
    const child = paragraph.children[index];
    const text = child?.type === "text" ? child.text : "";
    result =
      result.slice(0, element.start) +
      textElementXml("w:t", element.openTag, text) +
      result.slice(element.end);
  }
  return result;
}

export function escapeXmlAttribute(value: string): string {
  return escapeXmlText(value).replace(/"/g, "&quot;");
}

function textElementXml(tagName: string, openTag: string, text: string): string {
  const open = /^\s|\s$/.test(text)
    ? setXmlAttribute(
        openTag.replace(/\s*\/>$/, ">"),
        "xml:space",
        "preserve"
      )
    : openTag.replace(/\s*\/>$/, ">");
  return `${open}${escapeXmlText(text)}</${tagName}>`;
}

function runContentXml(text: string, deleted: boolean): string {
  const tagName = deleted ? "w:delText" : "w:t";
  return text
    .split(/(\t)/)
    .filter((piece) => piece.length > 0)
    .map((piece) =>
      piece === "\t" ? "<w:tab/>" : textElementXml(tagName, `<${tagName}>`, piece)
    )
    .join("");
}

function sameWrappers(left: SourceWrapper[], right: SourceWrapper[]): boolean {
  return (
    left.length === right.length &&
    left.every((wrapper, index) => wrapper.key === right[index]?.key)
  );
}

function mergeableItems(left: SourceItem, right: SourceItem): boolean {
  if (left.kind === "marker" || left.kind !== right.kind) {
    return false;
  }
  const rightRun = (right as SourceTextItem | SourceDeletedItem).run;
  if (
    left.run.openTag !== rightRun.openTag ||
    left.run.propertiesXml !== rightRun.propertiesXml ||
    !sameWrappers(left.wrappers, right.wrappers)
  ) {
    return false;
  }
  if (left.kind === "text") {
    const rightText = right as SourceTextItem;
    return (
      left.link === rightText.link &&
      JSON.stringify(canonical(left.style)) ===
        JSON.stringify(canonical(rightText.style))
    );
  }
  return true;
}

/** Allocates `w:id` values for revision elements the emitter has to create. */
export type RevisionIdAllocator = () => number;

/**
 * Writes the items back as one `w:p`. Adjacent items with the same run
 * formatting and wrappers share one run; a revision element split by other
 * content is reopened with a fresh `w:id` so every revision id stays unique.
 */
export function emitParagraphSource(
  source: ParagraphSource,
  allocateRevisionId?: RevisionIdAllocator
): { sourceXml: string; children: TextRunNode[] } {
  const merged: SourceItem[] = [];
  for (const item of source.items) {
    const previous = merged[merged.length - 1];
    if (previous && mergeableItems(previous, item)) {
      merged[merged.length - 1] = {
        ...previous,
        text: (previous as SourceTextItem).text + (item as SourceTextItem).text,
      } as SourceItem;
    } else {
      merged.push(item);
    }
  }

  let xml = source.openXml;
  const children: TextRunNode[] = [];
  const open: SourceWrapper[] = [];
  const closedKeys = new Set<string>();
  for (const item of merged) {
    let shared = 0;
    while (
      shared < open.length &&
      shared < item.wrappers.length &&
      open[shared]!.key === item.wrappers[shared]!.key
    ) {
      shared += 1;
    }
    while (open.length > shared) {
      const wrapper = open.pop()!;
      closedKeys.add(wrapper.key);
      xml += `</${wrapper.tagName}>`;
    }
    for (const wrapper of item.wrappers.slice(shared)) {
      const reopened =
        closedKeys.has(wrapper.key) &&
        wrapper.tagName !== "w:hyperlink" &&
        allocateRevisionId;
      xml += reopened
        ? setXmlAttribute(wrapper.openTag, "w:id", String(allocateRevisionId()))
        : wrapper.openTag;
      open.push(wrapper);
    }
    if (item.kind === "marker") {
      xml += item.xml;
      continue;
    }
    xml += `${item.run.openTag}${item.run.propertiesXml}${runContentXml(
      item.text,
      item.kind === "deleted"
    )}</w:r>`;
    if (item.kind === "text") {
      for (const piece of item.text.split(/(\t)/)) {
        if (piece) {
          children.push({
            type: "text",
            text: piece,
            ...(item.style ? { style: { ...item.style } } : {}),
            ...(item.link !== undefined ? { link: item.link } : {}),
          });
        }
      }
    }
  }
  while (open.length > 0) {
    xml += `</${open.pop()!.tagName}>`;
  }
  xml += "</w:p>";
  return {
    sourceXml: xml,
    children: children.length > 0 ? children : [{ type: "text", text: "" }],
  };
}

/** Writes items back into a copy of the paragraph and refreshes its run provenance. */
export function paragraphWithSource(
  paragraph: ParagraphNode,
  source: ParagraphSource,
  allocateRevisionId?: RevisionIdAllocator
): ParagraphNode {
  const { sourceXml, children } = emitParagraphSource(source, allocateRevisionId);
  return {
    ...paragraph,
    children,
    sourceXml,
    sourceTextPatch: undefined,
    sourceRunProvenance: {
      runs: children.map((run) => ({
        style: run.style ? { ...run.style } : undefined,
        link: run.link,
      })),
    },
  };
}

export function visibleText(items: SourceItem[]): string {
  return items
    .map((item) => (item.kind === "text" ? item.text : ""))
    .join("");
}

/** Splits text items so no item straddles `offset` in visible text. */
export function splitItemsAt(items: SourceItem[], offset: number): SourceItem[] {
  const result: SourceItem[] = [];
  let position = 0;
  for (const item of items) {
    if (item.kind !== "text") {
      result.push(item);
      continue;
    }
    const end = position + item.text.length;
    if (offset > position && offset < end) {
      const cut = offset - position;
      result.push({ ...item, text: item.text.slice(0, cut) });
      result.push({ ...item, text: item.text.slice(cut) });
    } else {
      result.push(item);
    }
    position = end;
  }
  return result;
}

/** Index of each visible text item's first character, keyed by item index. */
export function textItemOffsets(items: SourceItem[]): Map<number, number> {
  const offsets = new Map<number, number>();
  let position = 0;
  items.forEach((item, index) => {
    if (item.kind === "text") {
      offsets.set(index, position);
      position += item.text.length;
    }
  });
  return offsets;
}

/** Index of the text item holding visible character `offset`, after splitting there. */
export function textItemIndexAt(
  items: SourceItem[],
  offset: number
): number | undefined {
  for (const [index, start] of textItemOffsets(items)) {
    const item = items[index] as SourceTextItem;
    if (offset >= start && offset < start + item.text.length) {
      return index;
    }
  }
  return undefined;
}

/**
 * Where text inserted at `offset` goes: right after the character before it,
 * or before the first visible character when inserting at the start.
 */
export function insertionIndexAt(items: SourceItem[], offset: number): number {
  if (offset > 0) {
    const left = textItemIndexAt(items, offset - 1);
    if (left !== undefined) {
      return left + 1;
    }
  }
  const right = textItemIndexAt(items, offset);
  return right ?? items.length;
}

/** The run whose formatting new text at `offset` copies. */
export function formattingAnchorAt(
  items: SourceItem[],
  offset: number
): SourceTextItem | SourceDeletedItem | undefined {
  const left = offset > 0 ? textItemIndexAt(items, offset - 1) : undefined;
  const right = textItemIndexAt(items, offset);
  const index = left ?? right;
  if (index !== undefined) {
    return items[index] as SourceTextItem;
  }
  return items.find(
    (item): item is SourceTextItem | SourceDeletedItem => item.kind !== "marker"
  );
}

export function revisionWrapperId(wrapper: SourceWrapper): string | undefined {
  return xmlAttribute(wrapper.openTag, "w:id")?.trim();
}

function forEachSourceXmlNode(
  nodes: DocNode[],
  visit: (node: DocNode) => void
): void {
  for (const node of nodes) {
    visit(node);
    if (node.type === "table") {
      for (const row of node.rows) {
        for (const cell of row.cells) {
          forEachSourceXmlNode(cell.nodes, visit);
        }
      }
    }
  }
}

/** Next unused annotation id across the body's source XML and comments. */
export function createRevisionIdAllocator(model: DocModel): RevisionIdAllocator {
  let largest = -1;
  forEachSourceXmlNode(model.nodes, (node) => {
    for (const match of (node.sourceXml ?? "").matchAll(/\bw:id\s*=\s*"(-?\d+)"/g)) {
      largest = Math.max(largest, Number(match[1]));
    }
  });
  for (const comment of model.metadata.comments ?? []) {
    largest = Math.max(largest, comment.id);
  }
  return () => {
    largest += 1;
    return largest;
  };
}

/**
 * Returns a copy of the body nodes with `transform` applied to every source
 * XML string (top-level paragraphs, tables and their nested paragraphs).
 * Used for zero-width marker edits that never change modeled text.
 */
export function mapBodySourceXml(
  nodes: DocNode[],
  transform: (sourceXml: string) => string
): DocNode[] {
  return nodes.map((node) => {
    const sourceXml = node.sourceXml ? transform(node.sourceXml) : undefined;
    if (node.type === "paragraph") {
      return sourceXml === node.sourceXml ? node : { ...node, sourceXml };
    }
    const rows = node.rows.map((row) => ({
      ...row,
      cells: row.cells.map((cell) => ({
        ...cell,
        nodes: mapBodySourceXml(cell.nodes, transform) as typeof cell.nodes,
      })),
    }));
    return { ...node, sourceXml, rows };
  });
}
