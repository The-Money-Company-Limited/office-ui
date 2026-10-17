import * as React from "react";
import { sameParagraphLocation, type DocxTextRange, type DocxTextRangeLocation } from "./editor";

function locationSelector(location: DocxTextRangeLocation): string {
  return location.kind === "paragraph"
    ? `[data-docx-paragraph-kind="paragraph"][data-docx-paragraph-node-index="${location.nodeIndex}"]`
    : `[data-docx-paragraph-kind="table-cell"][data-docx-table-index="${location.tableIndex}"][data-docx-row-index="${location.rowIndex}"][data-docx-cell-index="${location.cellIndex}"][data-docx-paragraph-index="${location.paragraphIndex}"]`;
}

// Pretext fragments carry model offsets; ordinary runs carry contiguous text.
// Paint each text node separately so tracked deletions and list labels stay out.
function matchRanges(
  root: HTMLElement,
  selector: string,
  matches: readonly DocxTextRange[],
  activeMatch: DocxTextRange | undefined,
  ignoredContent: string
): { ranges: Range[]; activeRanges: Range[] } {
  const ranges: Range[] = [];
  const activeRanges: Range[] = [];
  const active = activeMatch && matches.find(match =>
    sameParagraphLocation(match.start.location, activeMatch.start.location) &&
    match.start.offset === activeMatch.start.offset && match.end.offset === activeMatch.end.offset
  );
  for (const host of root.querySelectorAll<HTMLElement>(selector)) {
    const walker = document.createTreeWalker(host, NodeFilter.SHOW_TEXT);
    let offset = 0;
    let fragment: Element | null = null;
    let node: Node | null;
    while ((node = walker.nextNode())) {
      const parent = node.parentElement;
      if (!parent || parent.closest(`${ignoredContent},textarea,[data-docx-form-field]`)) continue;
      const nextFragment = parent.closest("[data-docx-text-start]");
      if (nextFragment && nextFragment !== fragment) {
        offset = Number(nextFragment.getAttribute("data-docx-text-start"));
      }
      fragment = nextFragment;
      const length = node.textContent?.length ?? 0;
      for (const match of matches) {
        const start = Math.max(match.start.offset, offset);
        const end = Math.min(match.end.offset, offset + length);
        if (end > start) {
          const range = document.createRange();
          range.setStart(node, start - offset);
          range.setEnd(node, end - offset);
          ranges.push(range);
          if (match === active) activeRanges.push(range);
        }
      }
      offset += length;
    }
  }
  return { ranges, activeRanges };
}

function visibleRect(range: Range): DOMRect | undefined {
  const host = range.startContainer.parentElement?.closest<HTMLElement>("[data-docx-paragraph-host]");
  const page = host?.closest<HTMLElement>("[data-docx-page-wrapper]");
  if (!host || !page) return undefined;
  const hostRect = host.getBoundingClientRect();
  const pageRect = page.getBoundingClientRect();
  return Array.from(range.getClientRects()).find(rect =>
    rect.width > 0 && rect.height > 0 &&
    rect.bottom > Math.max(hostRect.top, pageRect.top) &&
    rect.top < Math.min(hostRect.bottom, pageRect.bottom)
  );
}

function scrollMatchIntoView(root: HTMLElement, rect: DOMRect): void {
  let scroller: HTMLElement | null = root.parentElement;
  while (scroller && !/(auto|scroll)/.test(getComputedStyle(scroller).overflowY)) {
    scroller = scroller.parentElement;
  }
  const viewport = scroller?.getBoundingClientRect();
  const delta = rect.top + rect.height / 2 - (viewport ? viewport.top + viewport.height / 2 : window.innerHeight / 2);
  if (scroller) scroller.scrollBy({ top: delta, behavior: "auto" });
  else window.scrollBy({ top: delta, behavior: "auto" });
}

export function useSearchHighlights({
  rootRef, matches, activeMatch, revealMatch, selectMatch, ignoredContent,
}: {
  rootRef: React.RefObject<HTMLDivElement | null>;
  matches?: readonly DocxTextRange[];
  activeMatch?: DocxTextRange;
  revealMatch: (match: DocxTextRange, afterPageIndex?: number) => number | undefined;
  selectMatch: (match: DocxTextRange) => void;
  ignoredContent: string;
}): void {
  const id = React.useId().replace(/[^a-zA-Z0-9]/g, "");
  const pending = React.useRef<DocxTextRange | undefined>(undefined);
  const pendingPage = React.useRef<number | undefined>(undefined);
  const callbacks = React.useRef({ revealMatch, selectMatch });
  callbacks.current = { revealMatch, selectMatch };
  React.useEffect(() => {
    pending.current = activeMatch;
    if (activeMatch) {
      pendingPage.current = callbacks.current.revealMatch(activeMatch);
      callbacks.current.selectMatch(activeMatch);
    }
  }, [activeMatch]);

  React.useEffect(() => {
    const root = rootRef.current;
    if (!root || !matches?.length) return;
    const name = `docx-find-${id}`;
    const activeName = `${name}-active`;
    const style = document.createElement("style");
    style.textContent = `::highlight(${name}) { background-color: #fde68a; color: #111827; }
      ::highlight(${activeName}) { background-color: #fb923c; color: #111827; }`;
    document.head.append(style);
    const paragraphs = new Map<string, DocxTextRange[]>();
    for (const match of matches) {
      const selector = locationSelector(match.start.location);
      const group = paragraphs.get(selector);
      if (group) group.push(match);
      else paragraphs.set(selector, [match]);
    }
    let frame = 0;
    const paint = () => {
      const allRanges: Range[] = [];
      const activeRanges: Range[] = [];
      for (const [selector, group] of paragraphs) {
        const painted = matchRanges(root, selector, group, activeMatch, ignoredContent);
        allRanges.push(...painted.ranges);
        activeRanges.push(...painted.activeRanges);
      }
      if (typeof Highlight !== "undefined" && CSS.highlights) {
        CSS.highlights.set(name, new Highlight(...allRanges));
        const highlight = new Highlight(...activeRanges);
        highlight.priority = 1;
        CSS.highlights.set(activeName, highlight);
      }
      if (pending.current === activeMatch && activeMatch) {
        const rect = activeRanges.map(visibleRect).find(Boolean);
        if (rect) {
          pending.current = undefined;
          scrollMatchIntoView(root, rect);
        } else if (pendingPage.current !== undefined && root.querySelector(
          `[data-docx-page-wrapper][data-index="${pendingPage.current}"] [data-docx-page-surface]`
        )) {
          // A table row can span several pages; reveal its next slice only
          // after the previous slice has mounted and proved not to contain it.
          pendingPage.current = callbacks.current.revealMatch(activeMatch, pendingPage.current);
          if (pendingPage.current !== undefined) schedule();
        }
      }
    };
    const schedule = () => {
      cancelAnimationFrame(frame);
      frame = requestAnimationFrame(paint);
    };
    // Virtualization and paragraph layout can mount the target after page reveal.
    const observer = new MutationObserver(schedule);
    observer.observe(root, { subtree: true, childList: true, characterData: true });
    const resize = new ResizeObserver(schedule);
    resize.observe(root);
    schedule();
    return () => {
      observer.disconnect();
      resize.disconnect();
      cancelAnimationFrame(frame);
      CSS.highlights?.delete(name);
      CSS.highlights?.delete(activeName);
      style.remove();
    };
  }, [rootRef, matches, activeMatch, ignoredContent, id]);
}
