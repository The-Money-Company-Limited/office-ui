// Cell anchors: a transparent DOM element per visible cell, so a page comment
// layer that anchors to `event.target` (an element id and its text) can name a
// cell of the canvas grid. The id is `${prefix}c-${A1}`; `locateCellAnchor`
// reads it back from a stored anchor.
import { formatCellRef } from "./comments.ts";

export interface CellAnchorAxisItem {
  start: number;
  size: number;
}

export interface CellAnchorPlacement {
  id: string;
  ref: string;
  row: number;
  col: number;
  /** Rect relative to the first cell of the sheet, or null for a merge (the caller measures the whole merge). */
  rect: { left: number; top: number; width: number; height: number } | null;
}

export function cellAnchorId(prefix: string, ref: string): string {
  return `${prefix}c-${ref}`;
}

/**
 * One placement per visible cell; a merged region gets one, at its top-left
 * cell, even when only part of it is visible.
 */
export function listCellAnchors(
  prefix: string,
  rows: ReadonlyArray<CellAnchorAxisItem & { actualRow: number }>,
  cols: ReadonlyArray<CellAnchorAxisItem & { actualCol: number }>,
  mergeAnchorOf: (row: number, col: number) => { row: number; col: number } | null
): CellAnchorPlacement[] {
  const placements: CellAnchorPlacement[] = [];
  const seenMerges = new Set<string>();
  for (const row of rows) {
    for (const col of cols) {
      const merge = mergeAnchorOf(row.actualRow, col.actualCol);
      const cell = merge ?? { col: col.actualCol, row: row.actualRow };
      if (merge) {
        const key = `${merge.row}:${merge.col}`;
        if (seenMerges.has(key)) {
          continue;
        }
        seenMerges.add(key);
      }
      const ref = formatCellRef(cell);
      placements.push({
        col: cell.col,
        id: cellAnchorId(prefix, ref),
        rect: merge ? null : { height: row.size, left: col.start, top: row.start, width: col.size },
        ref,
        row: cell.row
      });
    }
  }
  return placements;
}

// ---------------------------------------------------------------------------
// Reading a stored comment anchor back.

export interface CommentAnchorTouch {
  el?: string;
  ink?: number;
}

/** The parts of a RuntimeVM comment anchor this reads; other fields are ignored. */
export interface CommentAnchor {
  page?: string;
  selector?: string;
  ink?: { marks?: Array<{ on?: CommentAnchorTouch[] }> } | null;
}

const CELL_ID = /c-([A-Z]{1,3}[1-9]\d{0,6})$/;

function elementIds(selector: string): string[] {
  return Array.from(selector.matchAll(/#([A-Za-z_][A-Za-z0-9_-]*)|\[id="([A-Za-z0-9_-]+)"\]/g), (match) => match[1] ?? match[2]!);
}

/** The cell named by the deepest cell-anchor id in a selector, or null. */
export function cellRefFromSelector(selector: string): string | null {
  const ids = elementIds(selector);
  for (let index = ids.length - 1; index >= 0; index -= 1) {
    const match = CELL_ID.exec(ids[index]!);
    if (match) {
      return match[1]!;
    }
  }
  return null;
}

/** The value of the `sheet` query parameter of a page (`+` is a space), or null. */
export function sheetFromPage(page: string): string | null {
  const queryStart = page.indexOf("?");
  if (queryStart === -1) {
    return null;
  }
  const query = page.slice(queryStart + 1).split("#")[0]!;
  for (const pair of query.split("&")) {
    const [name, value = ""] = pair.split("=");
    if (name === "sheet") {
      try {
        return decodeURIComponent(value.replace(/\+/g, " "));
      } catch {
        return null;
      }
    }
  }
  return null;
}

/** The cell an anchor points at: the most-inked cell of a drawing, else its selector's cell. */
export function cellRefFromAnchor(anchor: CommentAnchor): string | null {
  const inkByRef = new Map<string, number>();
  for (const mark of anchor.ink?.marks ?? []) {
    for (const touch of mark.on ?? []) {
      const ref = touch.el ? cellRefFromSelector(touch.el) : null;
      if (ref) {
        inkByRef.set(ref, (inkByRef.get(ref) ?? 0) + (Number.isFinite(touch.ink) ? touch.ink! : 0));
      }
    }
  }
  let best: string | null = null;
  for (const [ref, ink] of inkByRef) {
    if (best === null || ink > inkByRef.get(best)!) {
      best = ref;
    }
  }
  return best ?? (anchor.selector ? cellRefFromSelector(anchor.selector) : null);
}
