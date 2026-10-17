import { Workbook } from "../vendor/sheets-wasm/duke_sheets_wasm.js";
import { formatCellRef, parseCellRef, type XlsxCommentThread } from "./comments.ts";
import { saveWorkbookBytes } from "./workbook-save.ts";
import type { WorkbookImageAssets, WorkbookImageSheetOrigin } from "./images.ts";

export type XlsxStructuralEdit =
  | { kind: "rows" | "columns"; sheetIndex: number; index: number; count: number; delete: boolean }
  | { kind: "rename"; sheetIndex: number; name: string }
  | { kind: "move"; sheetIndex: number; to: number }
  | { kind: "deleteSheet"; sheetIndex: number };

export function mapStructuralSheetIndex(index: number, edit: XlsxStructuralEdit): number | null {
  if (edit.kind === "deleteSheet") return index === edit.sheetIndex ? null : index > edit.sheetIndex ? index - 1 : index;
  if (edit.kind !== "move") return index;
  if (index === edit.sheetIndex) return edit.to;
  if (edit.sheetIndex < edit.to && index > edit.sheetIndex && index <= edit.to) return index - 1;
  if (edit.sheetIndex > edit.to && index >= edit.to && index < edit.sheetIndex) return index + 1;
  return index;
}

/** Mutate a saved snapshot so invalid edits cannot alter the live workbook or its history. */
export function editWorkbookStructure(
  bytes: Uint8Array,
  threads: XlsxCommentThread[] | null,
  edit: XlsxStructuralEdit,
  preservation?: { assets: WorkbookImageAssets | null; sheetOrigins: Array<WorkbookImageSheetOrigin | null> }
) {
  const workbook = Workbook.fromBytes(bytes);
  try {
    if (!Number.isInteger(edit.sheetIndex) || edit.sheetIndex < 0 || edit.sheetIndex >= workbook.sheetCount) {
      throw new Error("Sheet index is out of range.");
    }
    const sheet = workbook.getSheet(edit.sheetIndex).name;
    switch (edit.kind) {
      case "rows":
      case "columns": {
        const limit = edit.kind === "rows" ? 1_048_576 : 16_384;
        if (!Number.isInteger(edit.index) || !Number.isInteger(edit.count) || edit.index < 0 || edit.count < 1 || edit.index + edit.count > limit) {
          throw new Error("The row or column range is invalid.");
        }
        if (edit.kind === "rows") {
          if (edit.delete) workbook.deleteRows(edit.sheetIndex, edit.index, edit.count);
          else workbook.insertRows(edit.sheetIndex, edit.index, edit.count);
        } else {
          if (edit.delete) workbook.deleteColumns(edit.sheetIndex, edit.index, edit.count);
          else workbook.insertColumns(edit.sheetIndex, edit.index, edit.count);
        }
        break;
      }
      case "rename": workbook.renameSheet(edit.sheetIndex, edit.name); break;
      case "move":
        if (!Number.isInteger(edit.to) || edit.to < 0 || edit.to >= workbook.sheetCount) throw new Error("Sheet index is out of range.");
        workbook.moveSheet(edit.sheetIndex, edit.to);
        break;
      case "deleteSheet":
        if (workbook.sheetCount <= 1) throw new Error("Keep at least one worksheet.");
        workbook.removeSheet(edit.sheetIndex);
        break;
    }
    const nextThreads = threads?.flatMap((thread) => {
      if (thread.sheet !== sheet) return [thread];
      if (edit.kind === "rename") return [{ ...thread, sheet: edit.name }];
      if (edit.kind === "deleteSheet") return [];
      if (edit.kind === "move") return [thread];
      const cell = parseCellRef(thread.ref);
      if (!cell) return [thread];
      const axis = edit.kind === "rows" ? "row" : "col";
      const index = cell[axis];
      if (index < edit.index) return [thread];
      if (edit.delete && index < edit.index + edit.count) return [];
      cell[axis] += edit.delete ? -edit.count : edit.count;
      return [{ ...thread, ref: formatCellRef(cell) }];
    }) ?? null;
    const origins: Array<WorkbookImageSheetOrigin | null> = Array(workbook.sheetCount).fill(null);
    preservation?.sheetOrigins.forEach((origin, index) => {
      const nextIndex = mapStructuralSheetIndex(index, edit);
      if (nextIndex !== null) origins[nextIndex] = origin;
    });
    return saveWorkbookBytes(workbook, preservation?.assets ?? null, origins, nextThreads, true);
  } finally {
    workbook.free();
  }
}
