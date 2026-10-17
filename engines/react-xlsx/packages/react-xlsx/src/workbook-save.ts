import type { Workbook } from "../vendor/sheets-wasm/duke_sheets_wasm.js";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import { applyCommentThreads, type XlsxCommentThread } from "./comments.ts";
import { mergeWorkbookImageAssets, type WorkbookImageAssets, type WorkbookImageSheetOrigin } from "./images.ts";

export function saveWorkbookBytes(
  workbook: Workbook,
  assets: WorkbookImageAssets | null,
  sheetOrigins: Array<WorkbookImageSheetOrigin | null>,
  threads: XlsxCommentThread[] | null,
  updateDrawingStructure = false
): Uint8Array {
  const archive = unzipSync(workbook.saveXlsxBytes());
  if (archive["xl/styles.xml"]) {
    archive["xl/styles.xml"] = strToU8(strFromU8(archive["xl/styles.xml"])
      .replace(/&amp;quot;/g, "&quot;")
      .replace(/&amp;apos;/g, "&apos;"));
  }
  const bytes = mergeWorkbookImageAssets(zipSync(archive, { level: 6 }), assets, sheetOrigins, updateDrawingStructure);
  return threads ? applyCommentThreads(bytes, threads) : bytes;
}
