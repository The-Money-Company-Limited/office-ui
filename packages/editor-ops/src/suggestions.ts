import type { DocModel, ParagraphNode } from "@extend-ai/react-docx-doc-model";

import {
  replaceParagraphTextTracked,
  type AnnotationMutationResult,
  type RevisionAuthorship,
} from "./annotations";

/**
 * Suggesting mode: the editor computes every edit as a plain model change,
 * and this module turns it into Word tracked changes on the model the edit
 * started from. Only edits that change the visible text of existing
 * top-level paragraphs can be expressed that way; anything else (splitting
 * or merging paragraphs, formatting, tables, headers) fails closed so an
 * untracked change never slips through.
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

/**
 * Re-expresses the change from `before` to `after` as tracked changes by
 * `authorship.author`, applied to `before`. Every changed node must be a
 * top-level paragraph whose visible text changed; each one becomes one
 * tracked replacement. Fails with "unsupported" for any other change.
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
    model = tracked.model;
  }
  return { ok: true, model };
}
