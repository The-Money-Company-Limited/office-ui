import * as React from "react";
import { createPortal } from "react-dom";
import type {
  DocumentBlock,
  DocxSourceRef,
} from "@extend-ai/react-docx-doc-model";

export interface BlockBoundingBox {
  left: number;
  top: number;
  right: number;
  bottom: number;
}
export interface DocxBlockFragment {
  nodeId: string;
  pageIndex: number;
  pageWidth: number;
  pageHeight: number;
  boundingBox: BlockBoundingBox;
  polygon: Array<{ x: number; y: number }>;
  coordinateSpace: "page-css-px";
}
export interface DocxBlockGeometry {
  status:
    | "rendered"
    | "partially_rendered"
    | "not_mounted"
    | "not_visible"
    | "unresolved";
  scope: "mounted_pages";
  issues: Array<{
    nodeId: string;
    reason: "unsupported_range" | "document_mismatch" | "not_on_page";
  }>;
  fragments: DocxBlockFragment[];
  unresolvedSourceIds: string[];
}

function intersect(a: BlockBoundingBox, b: BlockBoundingBox): BlockBoundingBox {
  return {
    left: Math.max(a.left, b.left),
    top: Math.max(a.top, b.top),
    right: Math.min(a.right, b.right),
    bottom: Math.min(a.bottom, b.bottom),
  };
}

function visibleBox(
  element: HTMLElement,
  surface: HTMLElement
): BlockBoundingBox | undefined {
  const style = getComputedStyle(element);
  if (style.display === "none" || style.visibility === "hidden")
    return undefined;
  const rect = element.getBoundingClientRect();
  let box: BlockBoundingBox = {
    left: rect.left,
    top: rect.top,
    right: rect.right,
    bottom: rect.bottom,
  };
  for (
    let parent = element.parentElement;
    parent && parent !== surface;
    parent = parent.parentElement
  ) {
    const style = getComputedStyle(parent);
    if (style.display === "none" || style.visibility === "hidden")
      return undefined;
    const rect = parent.getBoundingClientRect();
    if (/hidden|clip|scroll|auto/.test(style.overflowX))
      box = {
        ...box,
        left: Math.max(box.left, rect.left),
        right: Math.min(box.right, rect.right),
      };
    if (/hidden|clip|scroll|auto/.test(style.overflowY))
      box = {
        ...box,
        top: Math.max(box.top, rect.top),
        bottom: Math.min(box.bottom, rect.bottom),
      };
  }
  box = intersect(box, surface.getBoundingClientRect());
  return box.right > box.left && box.bottom > box.top ? box : undefined;
}

export function resolveDocxSourceGeometry(
  root: HTMLElement,
  refs: readonly DocxSourceRef[]
): DocxBlockGeometry {
  return resolveSourceGeometry(root, refs);
}

function resolveSourceGeometry(
  root: HTMLElement,
  refs: readonly DocxSourceRef[],
  elementsById?: Map<string, HTMLElement[]>,
  boxCache?: Map<HTMLElement, BlockBoundingBox | undefined>
): DocxBlockGeometry {
  const fragments: DocxBlockFragment[] = [];
  const unresolvedSourceIds: string[] = [];
  let foundElement = false;
  const issues: DocxBlockGeometry["issues"] = [];
  for (const ref of refs) {
    if (ref.start !== undefined || ref.end !== undefined) {
      issues.push({ nodeId: ref.nodeId, reason: "unsupported_range" });
      unresolvedSourceIds.push(ref.nodeId);
      continue;
    }
    const elements = elementsById
      ? elementsById.get(ref.nodeId) ?? []
      : root.querySelectorAll<HTMLElement>(
          `[data-docx-source-id="${CSS.escape(ref.nodeId)}"]`
        );
    let resolved = false;
    for (const element of elements) {
      const documentId = element.closest<HTMLElement>("[data-docx-document-id]")
        ?.dataset.docxDocumentId;
      if (documentId && documentId !== ref.documentId) {
        issues.push({ nodeId: ref.nodeId, reason: "document_mismatch" });
        continue;
      }
      const surface = element.closest<HTMLElement>("[data-docx-page-surface]");
      const wrapper = surface?.closest<HTMLElement>("[data-docx-page-index]");
      if (!surface || !wrapper) {
        issues.push({ nodeId: ref.nodeId, reason: "not_on_page" });
        continue;
      }
      foundElement = true;
      const box = boxCache?.has(element)
        ? boxCache.get(element)
        : visibleBox(element, surface);
      boxCache?.set(element, box);
      if (!box) continue;
      const page = surface.getBoundingClientRect();
      const pageStyle = getComputedStyle(surface);
      const pageWidth = parseFloat(pageStyle.width) || surface.offsetWidth;
      const pageHeight = parseFloat(pageStyle.height) || surface.offsetHeight;
      if (!page.width || !page.height || !pageWidth || !pageHeight) continue;
      const xScale = page.width / pageWidth;
      const yScale = page.height / pageHeight;
      const boundingBox = {
        left: (box.left - page.left) / xScale,
        top: (box.top - page.top) / yScale,
        right: (box.right - page.left) / xScale,
        bottom: (box.bottom - page.top) / yScale,
      };
      const { left, top, right, bottom } = boundingBox;
      fragments.push({
        nodeId: ref.nodeId,
        pageIndex: Number(wrapper.dataset.docxPageIndex),
        pageWidth,
        pageHeight,
        boundingBox,
        polygon: [
          { x: left, y: top },
          { x: right, y: top },
          { x: right, y: bottom },
          { x: left, y: bottom },
        ],
        coordinateSpace: "page-css-px",
      });
      resolved = true;
    }
    if (!resolved) unresolvedSourceIds.push(ref.nodeId);
  }
  return {
    status: fragments.length
      ? unresolvedSourceIds.length
        ? "partially_rendered"
        : "rendered"
      : issues.length
      ? "unresolved"
      : foundElement
      ? "not_visible"
      : "not_mounted",
    scope: "mounted_pages",
    issues,
    fragments,
    unresolvedSourceIds,
  };
}

export function resolveDocxBlockGeometries(
  root: HTMLElement,
  blocks: readonly DocumentBlock[]
): Map<string, DocxBlockGeometry> {
  const elementsById = new Map<string, HTMLElement[]>();
  for (const element of root.querySelectorAll<HTMLElement>("[data-docx-source-id]")) {
    const id = element.dataset.docxSourceId!;
    const elements = elementsById.get(id) ?? [];
    elements.push(element);
    elementsById.set(id, elements);
  }
  const boxCache = new Map<HTMLElement, BlockBoundingBox | undefined>();
  return new Map(blocks.map((block) => [
    block.id,
    resolveSourceGeometry(root, block.sourceRefs, elementsById, boxCache),
  ]));
}

export function revealDocxSource(root: HTMLElement, nodeId: string): void {
  const element = root.querySelector<HTMLElement>(
    `[data-docx-source-id="${CSS.escape(nodeId)}"]`
  );
  if (element) {
    element.scrollIntoView({ block: "center", inline: "nearest" });
    return;
  }
  const viewer = root.querySelector<HTMLElement>("[data-docx-editor-root]");
  (viewer ?? root).dispatchEvent(
    new CustomEvent("docx-reveal-source", { detail: { nodeId } })
  );
}

export interface DocxBlockHighlightsProps {
  rootRef: React.RefObject<HTMLElement | null>;
  block?: DocumentBlock;
  blocks?: readonly DocumentBlock[];
  color?: string;
  colorForBlock?: (block: DocumentBlock) => string;
  onGeometry?: (geometry: DocxBlockGeometry) => void;
}

export function DocxBlockHighlights({
  rootRef,
  block,
  blocks,
  color = "#2563eb",
  colorForBlock,
  onGeometry,
}: DocxBlockHighlightsProps): React.JSX.Element {
  const displayed = React.useMemo(() => {
    const unique = new Map((blocks ?? []).map((item) => [item.id, item]));
    if (block) unique.set(block.id, block);
    return [...unique.values()];
  }, [block, blocks]);
  const [geometries, setGeometries] = React.useState<Map<string, DocxBlockGeometry>>(new Map());
  React.useEffect(() => {
    const root = rootRef.current;
    if (!root || !displayed.length) {
      setGeometries(new Map());
      return;
    }
    let frame = 0;
    let stopped = false;
    const update = () => {
      frame = 0;
      if (stopped) return;
      const next = resolveDocxBlockGeometries(root, displayed);
      setGeometries((previous) =>
        JSON.stringify([...previous]) === JSON.stringify([...next]) ? previous : next
      );
      if (block) onGeometry?.(next.get(block.id)!);
    };
    const schedule = () => {
      if (!frame && !stopped) frame = requestAnimationFrame(update);
    };
    const mutations = new MutationObserver((records) => {
      const relevant = records.some((record) => {
        if (
          record.target instanceof Element &&
          record.target.closest("[data-docx-block-highlight]")
        )
          return false;
        if (record.type === "attributes" || record.type === "characterData")
          return true;
        return [...record.addedNodes, ...record.removedNodes].some(
          (node) =>
            !(
              node instanceof Element &&
              node.hasAttribute("data-docx-block-highlight")
            )
        );
      });
      if (relevant) schedule();
    });
    mutations.observe(root, {
      childList: true,
      characterData: true,
      subtree: true,
      attributes: true,
      attributeFilter: ["style", "class", "data-docx-source-id"],
    });
    const resize = new ResizeObserver(schedule);
    resize.observe(root);
    root
      .querySelectorAll("[data-docx-page-surface]")
      .forEach((page) => resize.observe(page));
    document.fonts?.addEventListener("loadingdone", schedule);
    window.addEventListener("resize", schedule);
    root.addEventListener("scroll", schedule, true);
    void document.fonts?.ready.then(schedule);
    update();
    return () => {
      stopped = true;
      cancelAnimationFrame(frame);
      mutations.disconnect();
      resize.disconnect();
      document.fonts?.removeEventListener("loadingdone", schedule);
      window.removeEventListener("resize", schedule);
      root.removeEventListener("scroll", schedule, true);
    };
  }, [rootRef, displayed, block, onGeometry]);
  return (
    <>
      {displayed.flatMap((item) =>
        (geometries.get(item.id)?.fragments ?? []).map((fragment, index) => {
          const surface = rootRef.current?.querySelector<HTMLElement>(
            `[data-docx-page-index="${fragment.pageIndex}"] [data-docx-page-surface]`
          );
          if (!surface) return null;
          const box = fragment.boundingBox;
          const tint = colorForBlock?.(item) ?? color;
          const active = item.id === block?.id;
          return createPortal(
            <div
              data-docx-block-highlight={item.id}
              data-docx-block-type={item.type}
              data-docx-highlight-active={active || undefined}
              aria-hidden="true"
              style={{
                position: "absolute",
                pointerEvents: "none",
                zIndex: active ? 101 : 100,
                left: box.left,
                top: box.top,
                width: box.right - box.left,
                height: box.bottom - box.top,
                boxSizing: "border-box",
                border: `${active && blocks ? 2 : 1}px solid ${tint}`,
                backgroundColor: `${tint}${active || !blocks ? "20" : "08"}`,
              }}
            />,
            surface,
            `${item.id}-${index}`
          );
        })
      )}
    </>
  );
}

export function normalizedDocxBoundingBox(fragment: DocxBlockFragment): {
  pageIndex: number;
  left: number;
  top: number;
  width: number;
  height: number;
} {
  const { boundingBox: box, pageWidth, pageHeight, pageIndex } = fragment;
  return {
    pageIndex,
    left: box.left / pageWidth,
    top: box.top / pageHeight,
    width: (box.right - box.left) / pageWidth,
    height: (box.bottom - box.top) / pageHeight,
  };
}

export interface DocxPageBlock
  extends Omit<DocumentBlock, "id" | "parentBlockId" | "children"> {
  id: string;
  sourceBlockId: string;
  parentBlockId?: string;
  metadata: { page: { number: number; width: number; height: number } };
  boundingBox: BlockBoundingBox;
  polygon: Array<{ x: number; y: number }>;
  fragments: DocxBlockFragment[];
}

export function toDocxPageBlocks(
  block: DocumentBlock,
  geometry: DocxBlockGeometry
): DocxPageBlock[] {
  const sourceIds = new Set(block.sourceRefs.map((ref) => ref.nodeId));
  const pages = new Map<number, DocxBlockFragment[]>();
  for (const fragment of geometry.fragments) {
    if (!sourceIds.has(fragment.nodeId)) continue;
    const list = pages.get(fragment.pageIndex) ?? [];
    list.push(fragment);
    pages.set(fragment.pageIndex, list);
  }
  return [...pages]
    .sort(([a], [b]) => a - b)
    .map(([pageIndex, fragments]) => {
      const box = {
        left: Math.min(...fragments.map((f) => f.boundingBox.left)),
        top: Math.min(...fragments.map((f) => f.boundingBox.top)),
        right: Math.max(...fragments.map((f) => f.boundingBox.right)),
        bottom: Math.max(...fragments.map((f) => f.boundingBox.bottom)),
      };
      const { children: _children, id, parentBlockId, ...content } = block;
      return {
        ...content,
        id: `${id}/page/${pageIndex}`,
        sourceBlockId: id,
        parentBlockId: parentBlockId
          ? `${parentBlockId}/page/${pageIndex}`
          : undefined,
        metadata: {
          page: {
            number: pageIndex + 1,
            width: fragments[0].pageWidth,
            height: fragments[0].pageHeight,
          },
        },
        boundingBox: box,
        polygon: [
          { x: box.left, y: box.top },
          { x: box.right, y: box.top },
          { x: box.right, y: box.bottom },
          { x: box.left, y: box.bottom },
        ],
        fragments,
      };
    });
}

export type ExtendParserBlockDetails =
  | { type: "text_details" }
  | { type: "table_details"; rowCount: number; columnCount: number }
  | { type: "table_cell_details"; rowIndex: number; columnIndex: number }
  | { type: "figure_details"; imageUrl?: string; figureType?: string }
  | { type: "barcode_details"; imageUrl?: string; decodedValue?: string; format?: string }
  | { type: "key_value_details"; ocrText?: string }
  | { type: "formula_details"; latex: string }
  | Record<string, never>;

export interface ExtendParserBlock
  extends Omit<DocxPageBlock, "details" | "fragments"> {
  details: ExtendParserBlockDetails;
  sourceDetails: DocumentBlock["details"];
  coordinateSpace: "page-points";
  children: ExtendParserBlock[];
}

function extendDetails(block: DocumentBlock): ExtendParserBlockDetails {
  const d = block.details;
  const string = (name: string) => typeof d[name] === "string" ? d[name] as string : undefined;
  const number = (name: string) => typeof d[name] === "number" ? d[name] as number : 0;
  switch (block.type) {
    case "table":
      return { type: "table_details", rowCount: number("rowCount"), columnCount: number("columnCount") };
    case "table_cell":
      return { type: "table_cell_details", rowIndex: number("rowIndex"), columnIndex: number("columnIndex") };
    case "figure":
      return { type: "figure_details", imageUrl: string("imageUrl"), figureType: string("figureType") };
    case "barcode":
      return { type: "barcode_details", imageUrl: string("imageUrl"), decodedValue: string("decodedValue"), format: string("format") };
    case "key_value":
      return { type: "key_value_details", ocrText: string("ocrText") };
    case "formula":
      return string("latex") !== undefined ? { type: "formula_details", latex: string("latex")! } : {};
    default:
      return { type: "text_details" };
  }
}

export function toExtendParserBlocks(
  blocks: readonly DocumentBlock[],
  geometries: ReadonlyMap<string, DocxBlockGeometry>
): ExtendParserBlock[] {
  const occurrences = new Map<string, ExtendParserBlock>();
  const visit = (block: DocumentBlock) => {
    const geometry = geometries.get(block.id);
    if (geometry) {
      for (const page of toDocxPageBlocks(block, geometry)) {
        const { fragments: _fragments, details: sourceDetails, ...value } = page;
        const points = (px: number) => px * 72 / 96;
        occurrences.set(page.id, {
          ...value,
          details: extendDetails(block),
          sourceDetails,
          coordinateSpace: "page-points",
          metadata: { page: {
            number: page.metadata.page.number,
            width: points(page.metadata.page.width),
            height: points(page.metadata.page.height),
          } },
          boundingBox: {
            left: points(page.boundingBox.left), top: points(page.boundingBox.top),
            right: points(page.boundingBox.right), bottom: points(page.boundingBox.bottom),
          },
          polygon: page.polygon.map(({ x, y }) => ({ x: points(x), y: points(y) })),
          children: [],
        });
      }
    }
    block.children.forEach(visit);
  };
  blocks.forEach(visit);
  const roots: ExtendParserBlock[] = [];
  for (const block of occurrences.values()) {
    const parent = block.parentBlockId ? occurrences.get(block.parentBlockId) : undefined;
    if (parent) parent.children.push(block);
    else {
      delete block.parentBlockId;
      roots.push(block);
    }
  }
  return roots;
}
