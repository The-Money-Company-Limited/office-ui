// Headless workbook API for an agent that answers comments by changing the
// workbook: no React, no DOM, no wasm. It edits the XLSX package in place, so
// every part it does not touch (images, charts, styles) is saved byte for byte.
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import {
  addCommentThread,
  closeSuggestion,
  DEFAULT_REVIEWER,
  escapeXml,
  formatCellRef,
  formatSuggestionText,
  listSuggestions,
  normalizeCellRef,
  parseCellRef,
  readCommentThreads,
  readRelationships,
  readWorkbookSheets,
  removeCommentThread,
  replyToCommentThread,
  requireOpenSuggestion,
  resolveTarget,
  setCommentThreadResolved,
  unescapeXml,
  writeCommentThreads,
  xmlElements,
  xmlRunText,
  type XlsxArchive,
  type XlsxCommentThread,
  type XlsxSuggestion,
  type XlsxSuggestionValue,
  type XlsxWorkbookSheet
} from "./comments.ts";
import { cellRefFromAnchor, sheetFromPage, type CommentAnchor } from "./cell-anchors.ts";
import { normalizeZipEntryNames } from "./zip-entry-names.ts";

export type {
  XlsxComment,
  XlsxCommentThread,
  XlsxSuggestion,
  XlsxSuggestionStatus,
  XlsxSuggestionValue
} from "./comments.ts";
export { SUGGESTION_MARKER, formatSuggestionText, parseSuggestionText } from "./comments.ts";
export type { CommentAnchor } from "./cell-anchors.ts";

export interface XlsxAgentCell {
  ref: string;
  /** Cached value as saved in the file; for a formula, its last calculated result. */
  value: string | number | boolean | null;
  /** Formula without the leading "=", when the cell has one. */
  formula?: string;
}

export interface XlsxAgentWorkbook {
  sheets(): string[];
  /** Non-empty cells of a sheet, optionally limited to an A1 range like `A1:D20`. */
  cells(sheet: string, range?: string): XlsxAgentCell[];
  comments(): XlsxCommentThread[];
  suggestions(): XlsxSuggestion[];
  comment(input: { sheet: string; ref: string; text: string; author: string }): XlsxCommentThread;
  /** Returns the thread id, which changes when a legacy note becomes a thread. */
  reply(input: { id: string; text: string; author: string }): string;
  resolve(input: { id: string; resolved?: boolean }): string;
  delete(input: { id: string }): void;
  /** Proposes a cell value as a threaded comment; `previous` is filled from the cell. */
  suggest(input: { sheet: string; ref: string; value: XlsxSuggestionValue; author: string; note?: string }): XlsxCommentThread;
  /** Writes the proposed value(s), replies "Accepted" and resolves. Returns the thread ids. */
  accept(id: string | "all", author?: string): string[];
  /** Replies "Rejected" and resolves without changing the cell. Returns the thread ids. */
  reject(id: string | "all", author?: string): string[];
  save(): Uint8Array;
  /**
   * The cell a page comment points at, for a viewer rendered with `cellAnchorIdPrefix`: the sheet
   * from the page's `?sheet=` query (else the workbook's active sheet), the cell from the anchor's
   * most-inked cell or its selector. Null when the anchor names no cell of a known sheet.
   */
  locate(anchor: CommentAnchor, options?: { page?: string }): XlsxAgentLocation | null;
}

export interface XlsxAgentLocation {
  sheet: string;
  ref: string;
  value: XlsxAgentCell["value"];
  formula?: string;
}

export function openXlsx(bytes: Uint8Array): XlsxAgentWorkbook {
  const archive: XlsxArchive = unzipSync(normalizeZipEntryNames(bytes));
  const workbookSheets = readWorkbookSheets(archive);
  if (workbookSheets.length === 0) {
    throw new Error("Not an XLSX workbook: xl/workbook.xml has no worksheets.");
  }
  let threads = readCommentThreads(archive);
  let sharedStrings: string[] | null = null;

  const sheetByName = (name: string): XlsxWorkbookSheet => {
    const sheet = workbookSheets.find((entry) => entry.name === name)
      ?? workbookSheets.find((entry) => entry.name.toLowerCase() === name.toLowerCase());
    if (!sheet) {
      throw new Error(`No sheet named "${name}". Sheets: ${workbookSheets.map((entry) => entry.name).join(", ")}.`);
    }
    return sheet;
  };

  const readSharedStrings = () => {
    if (!sharedStrings) {
      const relationship = readRelationships(archive, "xl/workbook.xml").find((entry) => entry.type.endsWith("/sharedStrings"));
      const entry = relationship ? archive[resolveTarget("xl/workbook.xml", relationship.target)] : undefined;
      sharedStrings = entry ? xmlElements(strFromU8(entry), "si").map((element) => xmlRunText(element.body)) : [];
    }
    return sharedStrings;
  };

  const cells = (sheetName: string, range?: string): XlsxAgentCell[] => {
    const sheet = sheetByName(sheetName);
    const bounds = range ? parseRange(range) : null;
    const xml = strFromU8(archive[sheet.path]!);
    return xmlElements(sheetDataOf(xml).body, "c").flatMap(({ attributes, body }) => {
      const position = parseCellRef(attributes.r ?? "");
      if (!position || (bounds && !inRange(position, bounds))) {
        return [];
      }
      const cell = readCell(attributes, body, readSharedStrings());
      return cell.value === null && cell.formula === undefined ? [] : [{ ref: formatCellRef(position), ...cell }];
    });
  };

  const suggestionsToClose = (id: string) => (
    id === "all"
      ? listSuggestions(threads).filter((suggestion) => suggestion.status === "open")
      : [requireOpenSuggestion(threads, id)]
  );

  return {
    accept(id, author = DEFAULT_REVIEWER) {
      return suggestionsToClose(id).map((suggestion) => {
        writeCellValue(archive, sheetByName(suggestion.sheet).path, suggestion.ref, suggestion.value);
        threads = closeSuggestion(threads, suggestion.threadId, "accepted", author);
        return suggestion.threadId;
      });
    },
    cells,
    comment(input) {
      const result = addCommentThread(threads, { ...input, sheet: sheetByName(input.sheet).name });
      threads = result.threads;
      return result.thread;
    },
    comments: () => threads,
    delete({ id }) {
      threads = removeCommentThread(threads, id);
    },
    reject(id, author = DEFAULT_REVIEWER) {
      return suggestionsToClose(id).map((suggestion) => {
        threads = closeSuggestion(threads, suggestion.threadId, "rejected", author);
        return suggestion.threadId;
      });
    },
    reply({ id, text, author }) {
      const index = threads.findIndex((thread) => thread.id === id);
      threads = replyToCommentThread(threads, id, text, author);
      return threads[index]!.id;
    },
    resolve({ id, resolved = true }) {
      const index = threads.findIndex((thread) => thread.id === id);
      threads = setCommentThreadResolved(threads, id, resolved);
      return threads[index]!.id;
    },
    save() {
      const output = { ...archive };
      writeCommentThreads(output, threads);
      return zipSync(output, { level: 6 });
    },
    locate(anchor, options = {}) {
      const ref = cellRefFromAnchor(anchor);
      const page = options.page ?? anchor.page ?? "";
      const sheetName = sheetFromPage(page) ?? activeSheetName(archive) ?? workbookSheets[0]!.name;
      const sheet = workbookSheets.find((entry) => entry.name === sheetName);
      if (!ref || !sheet) {
        return null;
      }
      const cell = cells(sheet.name, ref)[0];
      return { ...(cell?.formula === undefined ? {} : { formula: cell.formula }), ref, sheet: sheet.name, value: cell?.value ?? null };
    },
    sheets: () => workbookSheets.map((sheet) => sheet.name),
    suggest({ sheet: sheetName, ref, value, author, note }) {
      const sheet = sheetByName(sheetName);
      const normalizedRef = normalizeCellRef(ref);
      const current = cells(sheet.name, normalizedRef)[0];
      const previous = current?.formula !== undefined ? `=${current.formula}` : current?.value ?? null;
      const result = addCommentThread(threads, {
        author,
        ref: normalizedRef,
        sheet: sheet.name,
        text: formatSuggestionText(value, previous, note)
      });
      threads = result.threads;
      return result.thread;
    },
    suggestions: () => listSuggestions(threads)
  };
}

/** Name of the tab the workbook opens on (`workbookView activeTab`, counting every sheet kind). */
function activeSheetName(archive: XlsxArchive): string | null {
  const entry = archive["xl/workbook.xml"];
  const xml = entry ? strFromU8(entry) : "";
  const activeTab = Number(xmlElements(xml, "workbookView")[0]?.attributes.activeTab ?? 0);
  return xmlElements(xml, "sheet")[activeTab]?.attributes.name ?? null;
}

// ---------------------------------------------------------------------------
// Worksheet cells.

interface CellRange {
  start: { row: number; col: number };
  end: { row: number; col: number };
}

function parseRange(range: string): CellRange {
  const [first, second] = range.split(":");
  const start = parseCellRef(first ?? "");
  const end = parseCellRef(second ?? first ?? "");
  if (!start || !end) {
    throw new Error(`"${range}" is not a range like A1:D20.`);
  }
  return {
    end: { col: Math.max(start.col, end.col), row: Math.max(start.row, end.row) },
    start: { col: Math.min(start.col, end.col), row: Math.min(start.row, end.row) }
  };
}

function inRange(position: { row: number; col: number }, range: CellRange): boolean {
  return position.row >= range.start.row && position.row <= range.end.row
    && position.col >= range.start.col && position.col <= range.end.col;
}

function sheetDataOf(xml: string) {
  const sheetData = xmlElements(xml, "sheetData")[0];
  if (!sheetData) {
    throw new Error("Worksheet has no sheetData element.");
  }
  return sheetData;
}

function readCell(attributes: Record<string, string>, body: string, sharedStrings: string[]): Omit<XlsxAgentCell, "ref"> {
  const formulaElement = xmlElements(body, "f")[0];
  const formula = formulaElement ? unescapeXml(formulaElement.body) : undefined;
  const raw = xmlElements(body, "v")[0];
  const rawText = raw ? unescapeXml(raw.body) : null;
  let value: XlsxAgentCell["value"];
  switch (attributes.t) {
    case "s":
      value = rawText === null ? null : sharedStrings[Number(rawText)] ?? null;
      break;
    case "inlineStr":
      value = xmlRunText(xmlElements(body, "is")[0]?.body ?? "");
      break;
    case "b":
      value = rawText === null ? null : rawText === "1";
      break;
    case "str":
    case "e":
      value = rawText;
      break;
    default:
      value = rawText === null || rawText === "" ? null : Number(rawText);
  }
  return formula === undefined ? { value } : { formula, value };
}

function cellXml(ref: string, styleAttribute: string, value: XlsxSuggestionValue | null): string {
  if (value === null || value === "") {
    return `<c r="${ref}"${styleAttribute}/>`;
  }
  if (typeof value === "string" && value.startsWith("=") && value.length > 1) {
    return `<c r="${ref}"${styleAttribute}><f>${escapeXml(value.slice(1))}</f></c>`;
  }
  if (typeof value === "number") {
    return `<c r="${ref}"${styleAttribute}><v>${value}</v></c>`;
  }
  if (typeof value === "boolean") {
    return `<c r="${ref}"${styleAttribute} t="b"><v>${value ? 1 : 0}</v></c>`;
  }
  return `<c r="${ref}"${styleAttribute} t="inlineStr"><is><t xml:space="preserve">${escapeXml(value)}</t></is></c>`;
}

/**
 * Sets one cell in the worksheet XML, keeping its style. Formulas are written
 * without a cached value and the workbook is marked for a full recalculation
 * on open, so dependent cells are correct in Excel and in the viewer.
 */
function writeCellValue(archive: XlsxArchive, sheetPath: string, ref: string, value: XlsxSuggestionValue | null) {
  const position = parseCellRef(ref)!;
  const xml = strFromU8(archive[sheetPath]!);
  const sheetData = sheetDataOf(xml);
  const rowElements = xmlElements(sheetData.body, "row");
  const row = rowElements.find((element) => Number(element.attributes.r) === position.row + 1);
  const rowBody = row?.body ?? "";
  const cellElements = xmlElements(rowBody, "c");
  const existing = cellElements.find((element) => parseCellRef(element.attributes.r ?? "")?.col === position.col);
  if (existing && /<(?:\w+:)?f\b[^>]*\bref="/.test(existing.body)) {
    throw new Error(`${ref} anchors a shared or array formula; edit it in Excel.`);
  }
  const style = existing?.attributes.s ? ` s="${existing.attributes.s}"` : "";
  const nextCell = cellXml(ref, style, value);

  let nextRowBody: string;
  if (existing) {
    nextRowBody = `${rowBody.slice(0, existing.start)}${nextCell}${rowBody.slice(existing.end)}`;
  } else {
    const after = cellElements.find((element) => (parseCellRef(element.attributes.r ?? "")?.col ?? -1) > position.col);
    const at = after ? after.start : rowBody.length;
    nextRowBody = `${rowBody.slice(0, at)}${nextCell}${rowBody.slice(at)}`;
  }

  let nextSheetDataBody: string;
  if (row) {
    const openTag = /^<[^>]*?(?=\/?>)/.exec(sheetData.body.slice(row.start, row.end))![0];
    const rowName = /^<([\w:]+)/.exec(openTag)![1];
    const nextRow = `${openTag}>${nextRowBody}</${rowName}>`;
    nextSheetDataBody = `${sheetData.body.slice(0, row.start)}${nextRow}${sheetData.body.slice(row.end)}`;
  } else {
    const after = rowElements.find((element) => Number(element.attributes.r) > position.row + 1);
    const at = after ? after.start : sheetData.body.length;
    nextSheetDataBody = `${sheetData.body.slice(0, at)}<row r="${position.row + 1}">${nextRowBody}</row>${sheetData.body.slice(at)}`;
  }

  const sheetDataName = /^<([\w:]+)/.exec(xml.slice(sheetData.start))![1];
  const nextXml = `${xml.slice(0, sheetData.start)}<${sheetDataName}>${nextSheetDataBody}</${sheetDataName}>${xml.slice(sheetData.end)}`;
  archive[sheetPath] = strToU8(nextXml);
  requestFullRecalculation(archive);
}

function requestFullRecalculation(archive: XlsxArchive) {
  const workbookXml = strFromU8(archive["xl/workbook.xml"]!);
  const calcPr = xmlElements(workbookXml, "calcPr")[0];
  let next: string;
  if (calcPr) {
    const tag = workbookXml.slice(calcPr.start, calcPr.end).replace(/\s+fullCalcOnLoad="[^"]*"/, "");
    next = `${workbookXml.slice(0, calcPr.start)}${tag.replace(/^<([\w:]+)/, "<$1 fullCalcOnLoad=\"1\"")}${workbookXml.slice(calcPr.end)}`;
  } else {
    // CT_Workbook puts calcPr right after definedNames (or sheets when there are none).
    const anchors = Array.from(workbookXml.matchAll(/<\/(?:\w+:)?definedNames>|<\/(?:\w+:)?sheets>/g));
    const last = anchors[anchors.length - 1];
    const at = last ? last.index! + last[0].length : workbookXml.lastIndexOf("</");
    next = `${workbookXml.slice(0, at)}<calcPr fullCalcOnLoad="1"/>${workbookXml.slice(at)}`;
  }
  archive["xl/workbook.xml"] = strToU8(next);
  dropCalcChain(archive);
}

/** Excel rebuilds the calculation chain; a stale one can trigger repair prompts. */
function dropCalcChain(archive: XlsxArchive) {
  const relationships = readRelationships(archive, "xl/workbook.xml");
  const calcChain = relationships.find((entry) => entry.type.endsWith("/calcChain"));
  if (!calcChain) {
    return;
  }
  const path = resolveTarget("xl/workbook.xml", calcChain.target);
  delete archive[path];
  const relsXml = strFromU8(archive["xl/_rels/workbook.xml.rels"]!);
  archive["xl/_rels/workbook.xml.rels"] = strToU8(relsXml.replace(new RegExp(`<Relationship\\b[^>]*Id="${calcChain.id}"[^>]*/>`), ""));
  const contentTypes = archive["[Content_Types].xml"];
  if (contentTypes) {
    archive["[Content_Types].xml"] = strToU8(strFromU8(contentTypes).replace(new RegExp(`<Override\\b[^>]*PartName="/${path}"[^>]*/>`), ""));
  }
}
