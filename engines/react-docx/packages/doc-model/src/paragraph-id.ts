import type { DocModel, DocNode } from "./types";

/**
 * Paragraph ids shared by the viewer's DOM anchors (`<prefix>p-…`) and the
 * agent API. A body paragraph is `p-<w14:paraId>` when it carries a paraId
 * that no earlier body paragraph carries, else `p-n<index>` (its index in the
 * body). Tables have no id.
 */
export function bodyParagraphIds(nodes: readonly DocNode[]): Array<string | undefined> {
  const seen = new Set<string>();
  return nodes.map((node, index) => {
    if (node.type !== "paragraph") {
      return undefined;
    }
    const paraId = node.paraId?.toUpperCase();
    if (paraId && !seen.has(paraId)) {
      seen.add(paraId);
      return `p-${paraId}`;
    }
    return `p-n${index}`;
  });
}

/**
 * A DOM id: any `[A-Za-z0-9_-]` prefix (the viewer's `anchorIdPrefix`), the
 * paragraph id, and an optional `-l<line>` suffix the viewer adds to a
 * paragraph's later page segments.
 */
const PARAGRAPH_DOM_ID = /^[A-Za-z0-9_-]*?(p-(?:n\d+|[0-9A-Fa-f]{1,8}))(?:-l\d+)?$/;

/** `doc_p-n1` → `p-n1`; `p-0A1B2C3D-l2` → `p-0A1B2C3D`; anything else → undefined. */
export function paragraphIdFromDomId(domId: string): string | undefined {
  return domId.match(PARAGRAPH_DOM_ID)?.[1];
}

/** `w14:paraId` values must be below 0x80000000 (MS-DOCX). */
const PARA_ID_LIMIT = 0x80000000;

/**
 * Gives every body paragraph a `w14:paraId` no other body paragraph has:
 * paragraphs without one, or repeating an earlier one, get a fresh random
 * id. Returns the same model when nothing changes.
 */
export function withUniqueBodyParaIds(
  model: DocModel,
  random: () => number = Math.random
): DocModel {
  const used = new Set(
    model.nodes.flatMap((node) =>
      node.type === "paragraph" && node.paraId ? [node.paraId.toUpperCase()] : []
    )
  );
  const kept = new Set<string>();
  let changed = false;
  const nodes = model.nodes.map((node) => {
    if (node.type !== "paragraph") {
      return node;
    }
    const paraId = node.paraId?.toUpperCase();
    if (paraId && !kept.has(paraId)) {
      kept.add(paraId);
      return node;
    }
    let fresh: string;
    do {
      fresh = Math.floor(random() * PARA_ID_LIMIT)
        .toString(16)
        .toUpperCase()
        .padStart(8, "0");
    } while (used.has(fresh) || fresh === "00000000");
    used.add(fresh);
    kept.add(fresh);
    changed = true;
    return { ...node, paraId: fresh };
  });
  return changed ? { ...model, nodes } : model;
}
