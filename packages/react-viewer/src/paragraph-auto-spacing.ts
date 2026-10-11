import type { DocModel, ParagraphNode } from "@extend-ai/react-docx-doc-model";

export function bodyAutomaticParagraphSpacingPx(
  model: DocModel,
  nodeIndex: number,
  paragraph: ParagraphNode,
  edge: "before" | "after"
): number | undefined {
  const compatibility = model.metadata?.compatibility;
  const spacing = paragraph.style?.spacing;
  if (
    model.nodes[nodeIndex] !== paragraph ||
    (compatibility?.compatibilityMode !== 12 && compatibility?.compatibilityMode !== 15) ||
    compatibility.useFixedHtmlParagraphSpacing === true ||
    spacing?.[edge === "before" ? "beforeAutoSpacing" : "afterAutoSpacing"] !== true ||
    !paragraph.children.some((child) => child.type === "text" && child.text.trim().length > 0) ||
    paragraph.children.some((child) => child.type !== "text")
  ) {
    return undefined;
  }

  if (edge === "before" && nodeIndex === 0) {
    return 0;
  }

  const neighbor = model.nodes[nodeIndex + (edge === "before" ? -1 : 1)];
  const numbering = paragraph.style?.numbering;
  const neighborNumbering = neighbor?.type === "paragraph" ? neighbor.style?.numbering : undefined;
  if (numbering && numbering.numId > 0 && neighborNumbering && neighborNumbering.numId > 0) {
    const instances = model.metadata?.numberingDefinitions?.instances;
    const abstractId = instances?.find((instance) => instance.numId === numbering.numId)?.abstractNumId;
    const neighborAbstractId = instances?.find((instance) => instance.numId === neighborNumbering.numId)?.abstractNumId;
    if (numbering.numId === neighborNumbering.numId ||
      (abstractId !== undefined && abstractId === neighborAbstractId)) {
      return 0;
    }
    return undefined;
  }

  return undefined;
}
