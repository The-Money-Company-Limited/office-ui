import type { DocModel, ParagraphNode } from "@extend-ai/react-docx-doc-model";

import {
  joinParagraphsTracked,
  replaceParagraphTextTracked,
  splitParagraphTracked,
  type AnnotationMutationResult,
  type RevisionAuthorship,
} from "./annotations";

/**
 * Suggesting mode: the editor computes every edit as a plain model change,
 * and this module turns it into Word tracked changes on the model the edit
 * started from. Edits that change the visible text of existing top-level
 * paragraphs, Enter splitting one, and Backspace joining two can be expressed
 * that way; anything else (formatting, tables, headers) fails closed so an
 * untracked change never slips through. That includes an edit that changes
 * text and formatting together, such as a styled paste: the tracked text
 * would lose the formatting.
 */

export interface TextReplacement {
  startOffset: number;
  endOffset: number;
  text: string;
}

function isHighSurrogate(code: number): boolean {
  return code >= 0xd800 && code <= 0xdbff;
}

function isLowSurrogate(code: number): boolean {
  return code >= 0xdc00 && code <= 0xdfff;
}

/**
 * The smallest single replacement turning `before` into `after`: the common
 * prefix and suffix are kept, and neither boundary splits a surrogate pair.
 * Returns undefined when the texts are equal.
 */
export function textReplacementBetween(
  before: string,
  after: string
): TextReplacement | undefined {
  if (before === after) {
    return undefined;
  }
  const limit = Math.min(before.length, after.length);
  let prefix = 0;
  while (prefix < limit && before.charCodeAt(prefix) === after.charCodeAt(prefix)) {
    prefix += 1;
  }
  if (prefix > 0 && isHighSurrogate(before.charCodeAt(prefix - 1))) {
    prefix -= 1;
  }
  let suffix = 0;
  while (
    suffix < limit - prefix &&
    before.charCodeAt(before.length - 1 - suffix) ===
      after.charCodeAt(after.length - 1 - suffix)
  ) {
    suffix += 1;
  }
  if (suffix > 0 && isLowSurrogate(before.charCodeAt(before.length - suffix))) {
    suffix -= 1;
  }
  return {
    startOffset: prefix,
    endOffset: before.length - suffix,
    text: after.slice(prefix, after.length - suffix),
  };
}

function paragraphVisibleText(paragraph: ParagraphNode): string | undefined {
  let text = "";
  for (const child of paragraph.children) {
    if (child.type !== "text") {
      return undefined;
    }
    text += child.text;
  }
  return text;
}

/** JSON with sorted keys and no undefined values, for structural equality. */
function canonicalJson(value: unknown): string {
  return JSON.stringify(value, (_key, inner: unknown) =>
    inner && typeof inner === "object" && !Array.isArray(inner)
      ? Object.fromEntries(
          Object.entries(inner as Record<string, unknown>).sort(([left], [right]) =>
            left < right ? -1 : left > right ? 1 : 0
          )
        )
      : inner
  );
}

/** One entry per UTF-16 unit of visible text: the formatting it carries. */
function characterFormats(paragraph: ParagraphNode): string[] {
  const formats: string[] = [];
  for (const child of paragraph.children) {
    if (child.type !== "text") {
      continue;
    }
    const format = canonicalJson({ style: child.style, link: child.link });
    for (let index = 0; index < child.text.length; index += 1) {
      formats.push(format);
    }
  }
  return formats;
}

/** Whether the tracked paragraph shows the text and formatting the edit produced. */
function sameFormatting(tracked: ParagraphNode, edited: ParagraphNode): boolean {
  const trackedFormats = characterFormats(tracked);
  const editedFormats = characterFormats(edited);
  return (
    canonicalJson(tracked.style) === canonicalJson(edited.style) &&
    trackedFormats.length === editedFormats.length &&
    trackedFormats.every((format, index) => format === editedFormats[index])
  );
}

/**
 * Re-expresses the change from `before` to `after` as tracked changes by
 * `authorship.author`, applied to `before`. Every changed node must be a
 * top-level paragraph whose visible text changed and whose formatting the
 * tracked replacement reproduces; each one becomes one tracked replacement.
 * One paragraph becoming two with the same text and formatting is a tracked
 * paragraph break; two becoming one with their joined text is a tracked
 * removal of the first one's break. Fails with "unsupported" for any other
 * change.
 */
export function trackModelTextEdits(
  before: DocModel,
  after: DocModel,
  authorship: RevisionAuthorship
): AnnotationMutationResult {
  if (after === before) {
    return { ok: true, model: before };
  }
  const keys = new Set([...Object.keys(before), ...Object.keys(after)]);
  for (const key of keys) {
    if (
      key !== "nodes" &&
      before[key as keyof DocModel] !== after[key as keyof DocModel]
    ) {
      return { ok: false, reason: "unsupported" };
    }
  }
  if (after.nodes.length === before.nodes.length + 1) {
    return trackParagraphSplit(before, after, authorship);
  }
  if (after.nodes.length === before.nodes.length - 1) {
    return trackParagraphJoin(before, after, authorship);
  }
  if (before.nodes.length !== after.nodes.length) {
    return { ok: false, reason: "unsupported" };
  }

  let model = before;
  for (let index = 0; index < before.nodes.length; index += 1) {
    const previous = before.nodes[index];
    const next = after.nodes[index];
    if (previous === next) {
      continue;
    }
    if (previous?.type !== "paragraph" || next?.type !== "paragraph") {
      return { ok: false, reason: "unsupported" };
    }
    const previousText = paragraphVisibleText(previous);
    const nextText = paragraphVisibleText(next);
    if (previousText === undefined || nextText === undefined) {
      return { ok: false, reason: "unsupported" };
    }
    const replacement = textReplacementBetween(previousText, nextText);
    if (!replacement) {
      // Same text, different node: a formatting change, which suggesting
      // does not track.
      return { ok: false, reason: "unsupported" };
    }
    const tracked = replaceParagraphTextTracked(model, {
      nodeIndex: index,
      ...replacement,
      ...authorship,
    });
    if (!tracked.ok) {
      return tracked;
    }
    if (!sameFormatting(tracked.model.nodes[index] as ParagraphNode, next)) {
      return { ok: false, reason: "unsupported" };
    }
    model = tracked.model;
  }
  return { ok: true, model };
}

/** Enter: one top-level paragraph became two, with every other node kept. */
function trackParagraphSplit(
  before: DocModel,
  after: DocModel,
  authorship: RevisionAuthorship
): AnnotationMutationResult {
  let index = 0;
  while (index < before.nodes.length && before.nodes[index] === after.nodes[index]) {
    index += 1;
  }
  for (let rest = index + 1; rest < before.nodes.length; rest += 1) {
    if (before.nodes[rest] !== after.nodes[rest + 1]) {
      return { ok: false, reason: "unsupported" };
    }
  }
  const previous = before.nodes[index];
  const first = after.nodes[index];
  const second = after.nodes[index + 1];
  if (
    previous?.type !== "paragraph" ||
    first?.type !== "paragraph" ||
    second?.type !== "paragraph"
  ) {
    return { ok: false, reason: "unsupported" };
  }
  const firstText = paragraphVisibleText(first);
  const secondText = paragraphVisibleText(second);
  if (
    firstText === undefined ||
    secondText === undefined ||
    firstText + secondText !== paragraphVisibleText(previous)
  ) {
    return { ok: false, reason: "unsupported" };
  }
  const tracked = splitParagraphTracked(before, {
    nodeIndex: index,
    offset: firstText.length,
    ...authorship,
  });
  if (!tracked.ok) {
    return tracked;
  }
  const nodes = [...tracked.model.nodes];
  const trackedFirst = nodes[index] as ParagraphNode;
  const trackedSecond = nodes[index + 1] as ParagraphNode;
  if (!sameFormatting(trackedFirst, first) || !sameFormatting(trackedSecond, second)) {
    return { ok: false, reason: "unsupported" };
  }
  // The editor's block ids keep its layout caches and caret on the new halves.
  nodes[index] = { ...trackedFirst, blockId: first.blockId };
  nodes[index + 1] = { ...trackedSecond, blockId: second.blockId };
  return { ok: true, model: { ...tracked.model, nodes } };
}

/**
 * Backspace at a paragraph's start: two top-level paragraphs became one with
 * their joined text, and every other node was kept. The paragraphs stay apart
 * in the tracked model until the removed break is accepted.
 */
function trackParagraphJoin(
  before: DocModel,
  after: DocModel,
  authorship: RevisionAuthorship
): AnnotationMutationResult {
  let index = 0;
  while (index < after.nodes.length && before.nodes[index] === after.nodes[index]) {
    index += 1;
  }
  for (let rest = index + 1; rest < after.nodes.length; rest += 1) {
    if (after.nodes[rest] !== before.nodes[rest + 1]) {
      return { ok: false, reason: "unsupported" };
    }
  }
  const first = before.nodes[index];
  const second = before.nodes[index + 1];
  const joined = after.nodes[index];
  if (
    first?.type !== "paragraph" ||
    second?.type !== "paragraph" ||
    joined?.type !== "paragraph"
  ) {
    return { ok: false, reason: "unsupported" };
  }
  const firstText = paragraphVisibleText(first);
  const secondText = paragraphVisibleText(second);
  if (
    firstText === undefined ||
    secondText === undefined ||
    paragraphVisibleText(joined) !== firstText + secondText ||
    // Accepting the join keeps the second paragraph's properties, so it must
    // show the paragraph style the edit produced.
    canonicalJson(joined.style) !== canonicalJson(second.style)
  ) {
    return { ok: false, reason: "unsupported" };
  }
  return joinParagraphsTracked(before, { nodeIndex: index, ...authorship });
}
