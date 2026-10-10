import * as React from "react";
import { createPortal } from "react-dom";
import type { DocumentBlock, DocxSourceRef } from "@extend-ai/react-docx-doc-model";

export interface BlockBoundingBox { left: number; top: number; right: number; bottom: number }
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
  status: "rendered" | "partially_rendered" | "not_mounted" | "not_visible";
  fragments: DocxBlockFragment[];
  unresolvedSourceIds: string[];
}

function intersect(a: BlockBoundingBox, b: BlockBoundingBox): BlockBoundingBox {
  return { left: Math.max(a.left, b.left), top: Math.max(a.top, b.top), right: Math.min(a.right, b.right), bottom: Math.min(a.bottom, b.bottom) };
}

function visibleBox(element: HTMLElement, surface: HTMLElement): BlockBoundingBox | undefined {
  const style = getComputedStyle(element);
  if (style.display === "none" || style.visibility === "hidden") return undefined;
  const rect = element.getBoundingClientRect();
  let box: BlockBoundingBox = { left: rect.left, top: rect.top, right: rect.right, bottom: rect.bottom };
  for (let parent = element.parentElement; parent && parent !== surface; parent = parent.parentElement) {
    const style = getComputedStyle(parent);
    if (style.display === "none" || style.visibility === "hidden") return undefined;
    const rect = parent.getBoundingClientRect();
    if (/hidden|clip|scroll|auto/.test(style.overflowX)) box = { ...box, left: Math.max(box.left, rect.left), right: Math.min(box.right, rect.right) };
    if (/hidden|clip|scroll|auto/.test(style.overflowY)) box = { ...box, top: Math.max(box.top, rect.top), bottom: Math.min(box.bottom, rect.bottom) };
  }
  box = intersect(box, surface.getBoundingClientRect());
  return box.right > box.left && box.bottom > box.top ? box : undefined;
}

export function resolveDocxSourceGeometry(root: HTMLElement, refs: readonly DocxSourceRef[]): DocxBlockGeometry {
  const fragments: DocxBlockFragment[] = [];
  const unresolvedSourceIds: string[] = [];
  let foundElement = false;
  for (const ref of refs) {
    const elements = root.querySelectorAll<HTMLElement>(`[data-docx-source-id="${CSS.escape(ref.nodeId)}"]`);
    let resolved = false;
    for (const element of elements) {
      const surface = element.closest<HTMLElement>("[data-docx-page-surface]");
      const wrapper = surface?.closest<HTMLElement>("[data-docx-page-index]");
      if (!surface || !wrapper) continue;
      foundElement = true;
      const box = visibleBox(element, surface);
      if (!box) continue;
      const page = surface.getBoundingClientRect();
      const pageWidth = surface.offsetWidth;
      const pageHeight = surface.offsetHeight;
      if (!page.width || !page.height || !pageWidth || !pageHeight) continue;
      const xScale = page.width / pageWidth;
      const yScale = page.height / pageHeight;
      const boundingBox = { left: (box.left - page.left) / xScale, top: (box.top - page.top) / yScale, right: (box.right - page.left) / xScale, bottom: (box.bottom - page.top) / yScale };
      const { left, top, right, bottom } = boundingBox;
      fragments.push({ nodeId: ref.nodeId, pageIndex: Number(wrapper.dataset.docxPageIndex), pageWidth, pageHeight, boundingBox, polygon: [{x:left,y:top},{x:right,y:top},{x:right,y:bottom},{x:left,y:bottom}], coordinateSpace: "page-css-px" });
      resolved = true;
    }
    if (!resolved) unresolvedSourceIds.push(ref.nodeId);
  }
  return { status: fragments.length ? unresolvedSourceIds.length ? "partially_rendered" : "rendered" : foundElement ? "not_visible" : "not_mounted", fragments, unresolvedSourceIds };
}

export function revealDocxSource(root: HTMLElement, nodeId: string): void {
  const element = root.querySelector<HTMLElement>(`[data-docx-source-id="${CSS.escape(nodeId)}"]`);
  if (element) { element.scrollIntoView({ block: "center", inline: "nearest" }); return; }
  const viewer = root.querySelector<HTMLElement>("[data-docx-editor-root]");
  (viewer ?? root).dispatchEvent(new CustomEvent("docx-reveal-source", { detail: { nodeId } }));
}

export interface DocxBlockHighlightsProps {
  rootRef: React.RefObject<HTMLElement | null>;
  block?: DocumentBlock;
  color?: string;
  onGeometry?: (geometry: DocxBlockGeometry) => void;
}

export function DocxBlockHighlights({ rootRef, block, color = "#2563eb", onGeometry }: DocxBlockHighlightsProps): React.JSX.Element {
  const [geometry, setGeometry] = React.useState<DocxBlockGeometry>({ status: "not_mounted", fragments: [], unresolvedSourceIds: [] });
  React.useLayoutEffect(() => {
    const root = rootRef.current;
    if (!root || !block) { setGeometry({ status: "not_mounted", fragments: [], unresolvedSourceIds: [] }); return; }
    let frame = 0;
    let stopped = false;
    const update = () => {
      frame = 0;
      if (stopped) return;
      const next = resolveDocxSourceGeometry(root, block.sourceRefs);
      setGeometry(previous => JSON.stringify(previous) === JSON.stringify(next) ? previous : next);
      onGeometry?.(next);
    };
    const schedule = () => { if (!frame && !stopped) frame = requestAnimationFrame(update); };
    const mutations = new MutationObserver(records => {
      const relevant = records.some(record => {
        if (record.target instanceof Element && record.target.closest("[data-docx-block-highlight]")) return false;
        if (record.type === "attributes" || record.type === "characterData") return true;
        return [...record.addedNodes, ...record.removedNodes].some(node => !(node instanceof Element && node.hasAttribute("data-docx-block-highlight")));
      });
      if (relevant) schedule();
    });
    mutations.observe(root, { childList: true, characterData: true, subtree: true, attributes: true, attributeFilter: ["style", "class", "data-docx-source-id"] });
    const resize = new ResizeObserver(schedule);
    resize.observe(root);
    root.querySelectorAll("[data-docx-page-surface]").forEach(page => resize.observe(page));
    document.fonts?.addEventListener("loadingdone", schedule);
    window.addEventListener("resize", schedule);
    update();
    return () => { stopped = true; cancelAnimationFrame(frame); mutations.disconnect(); resize.disconnect(); document.fonts?.removeEventListener("loadingdone", schedule); window.removeEventListener("resize", schedule); };
  }, [rootRef, block, onGeometry]);
  return <>{block && geometry.fragments.map((fragment, index) => {
    const surface = rootRef.current?.querySelector<HTMLElement>(`[data-docx-page-index="${fragment.pageIndex}"] [data-docx-page-surface]`);
    if (!surface) return null;
    const box = fragment.boundingBox;
    return createPortal(<div key={index} data-docx-block-highlight={block.id} aria-hidden="true" style={{ position: "absolute", pointerEvents: "none", zIndex: 100, left: box.left, top: box.top, width: box.right - box.left, height: box.bottom - box.top, boxSizing: "border-box", border: `1px solid ${color}`, backgroundColor: `${color}20` }} />, surface, `${block.id}-${index}`);
  })}</>;
}
