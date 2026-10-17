// Excel threaded comments and suggestions, read from and written to the XLSX
// package directly. DOM-free so the headless agent API can use it in Node.
//
// The wasm engine keeps legacy notes but drops xl/threadedComments and
// xl/persons, so every save runs `applyCommentThreads` over the saved bytes,
// and the threads are the single source of truth for all comment parts.
//
// A suggestion is a thread whose root comment ends with one marker line:
//
//   Optional explanation for the reviewer.
//   [suggestion] {"value":1200,"previous":1000}
//
// `value` is what the cell becomes on accept: a string starting with "=" is a
// formula, any other JSON string, number or boolean is a literal value.
// Accepting or rejecting replies "Accepted"/"Rejected" and resolves the thread.
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";

export type XlsxArchive = Record<string, Uint8Array>;

export interface XlsxComment {
  id: string;
  /** Root comment id for replies. */
  parentId?: string;
  personId: string;
  author: string;
  text: string;
  /** Excel timestamp (`YYYY-MM-DDTHH:MM:SS.00`), empty when the file has none. */
  date: string;
}

export interface XlsxCommentThread {
  id: string;
  /** Worksheet name. */
  sheet: string;
  /** A1 cell reference, e.g. `B2`. */
  ref: string;
  /** Excel's "Resolve thread" state (the `done` attribute). */
  resolved: boolean;
  /**
   * True for a legacy note (no threaded comment). Replying to or resolving it
   * converts it into a threaded comment, as Excel does.
   */
  legacy: boolean;
  comments: XlsxComment[];
}

export type XlsxSuggestionValue = string | number | boolean;

export type XlsxSuggestionStatus = "open" | "accepted" | "rejected" | "resolved";

export interface XlsxSuggestion {
  threadId: string;
  sheet: string;
  ref: string;
  /** Proposed cell content; a string starting with "=" is a formula. */
  value: XlsxSuggestionValue;
  previous: XlsxSuggestionValue | null;
  /** The root comment text without the marker line. */
  note: string;
  author: string;
  date: string;
  status: XlsxSuggestionStatus;
}

export const SUGGESTION_MARKER = "[suggestion]";
export const ACCEPTED_REPLY = "Accepted";
export const REJECTED_REPLY = "Rejected";
/** Reply author when whoever accepts or rejects a suggestion is not named. */
export const DEFAULT_REVIEWER = "Reviewer";

const SPREADSHEET_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const THREADED_NS = "http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments";
const RELATIONSHIPS_NS = "http://schemas.openxmlformats.org/package/2006/relationships";
const OFFICE_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const REL_COMMENTS = `${OFFICE_REL_NS}/comments`;
const REL_VML = `${OFFICE_REL_NS}/vmlDrawing`;
const REL_THREADED = "http://schemas.microsoft.com/office/2017/10/relationships/threadedComment";
const REL_PERSON = "http://schemas.microsoft.com/office/2017/10/relationships/person";
const CT_COMMENTS = "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml";
const CT_THREADED = "application/vnd.ms-excel.threadedcomments+xml";
const CT_PERSON = "application/vnd.ms-excel.person+xml";
const CT_VML = "application/vnd.openxmlformats-officedocument.vmlDrawing";
const XML_DECLARATION = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n";
const THREADED_NOTE_PREAMBLE = "[Threaded comment]\n\nYour version of Excel allows you to read this threaded comment; however, any edits to it will get removed if the file is opened in a newer version of Excel. Learn more: https://go.microsoft.com/fwlink/?linkid=870924\n\n";

// ---------------------------------------------------------------------------
// Minimal XML helpers for the parts this module owns.

export function escapeXml(text: string): string {
  return text
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

function escapeAttribute(text: string): string {
  return escapeXml(text).replace(/\r/g, "&#13;").replace(/\n/g, "&#10;").replace(/\t/g, "&#9;");
}

export function unescapeXml(text: string): string {
  return text.replace(/&(#x[0-9a-fA-F]+|#\d+|lt|gt|quot|apos|amp);/g, (_, entity: string) => {
    switch (entity) {
      case "lt": return "<";
      case "gt": return ">";
      case "quot": return "\"";
      case "apos": return "'";
      case "amp": return "&";
      default: return String.fromCodePoint(entity[1] === "x" ? parseInt(entity.slice(2), 16) : parseInt(entity.slice(1), 10));
    }
  });
}

export interface XmlElement {
  attributes: Record<string, string>;
  /** Inner XML; empty for a self-closing element. */
  body: string;
  start: number;
  end: number;
}

function parseAttributes(source: string): Record<string, string> {
  const attributes: Record<string, string> = {};
  for (const match of source.matchAll(/([\w:.-]+)\s*=\s*(?:"([^"]*)"|'([^']*)')/g)) {
    const name = match[1]!;
    attributes[name.includes(":") && !name.startsWith("xmlns") ? name.slice(name.indexOf(":") + 1) : name] = unescapeXml(match[2] ?? match[3] ?? "");
  }
  return attributes;
}

/**
 * Elements with this local name (any prefix), in document order. Assumes the
 * element does not nest inside itself, which holds for every part read here.
 */
export function xmlElements(xml: string, localName: string): XmlElement[] {
  const pattern = new RegExp(`<(?:[\\w.-]+:)?${localName}(?=[\\s/>])([^>]*?)(\\/>|>([\\s\\S]*?)<\\/(?:[\\w.-]+:)?${localName}>)`, "g");
  return Array.from(xml.matchAll(pattern), (match) => ({
    attributes: parseAttributes(match[1] ?? ""),
    body: match[3] ?? "",
    start: match.index!,
    end: match.index! + match[0].length
  }));
}

/** Concatenated text of every `<t>` run in an element body (rich text or plain). */
export function xmlRunText(body: string): string {
  return xmlElements(body.replace(/<(?:\w+:)?rPh\b[\s\S]*?<\/(?:\w+:)?rPh>/g, ""), "t")
    .map((element) => unescapeXml(element.body.replace(/<[^>]*>/g, "")))
    .join("");
}

function readText(archive: XlsxArchive, path: string): string | null {
  const entry = archive[path];
  return entry ? strFromU8(entry) : null;
}

function writeText(archive: XlsxArchive, path: string, text: string) {
  archive[path] = strToU8(text);
}

// ---------------------------------------------------------------------------
// Package structure: relationships, content types, workbook sheets.

interface Relationship {
  id: string;
  type: string;
  target: string;
  targetMode?: string;
}

function relsPathFor(partPath: string): string {
  const slash = partPath.lastIndexOf("/");
  return `${partPath.slice(0, slash + 1)}_rels/${partPath.slice(slash + 1)}.rels`;
}

/** Archive path of a relationship target, resolved against its source part. */
export function resolveTarget(sourcePart: string, target: string): string {
  const segments = target.startsWith("/")
    ? []
    : sourcePart.split("/").slice(0, -1);
  for (const segment of target.replace(/^\//, "").split("/")) {
    if (segment === "..") {
      segments.pop();
    } else if (segment && segment !== ".") {
      segments.push(segment);
    }
  }
  return segments.join("/");
}

function relativeTarget(sourcePart: string, path: string): string {
  const from = sourcePart.split("/").slice(0, -1);
  const to = path.split("/");
  let shared = 0;
  while (shared < from.length && shared < to.length - 1 && from[shared] === to[shared]) {
    shared += 1;
  }
  return [...from.slice(shared).map(() => ".."), ...to.slice(shared)].join("/");
}

export function readRelationships(archive: XlsxArchive, partPath: string): Relationship[] {
  const xml = readText(archive, relsPathFor(partPath));
  if (!xml) {
    return [];
  }
  return xmlElements(xml, "Relationship").map(({ attributes }) => ({
    id: attributes.Id ?? "",
    target: attributes.Target ?? "",
    targetMode: attributes.TargetMode,
    type: attributes.Type ?? ""
  }));
}

function writeRelationships(archive: XlsxArchive, partPath: string, relationships: Relationship[]) {
  const path = relsPathFor(partPath);
  if (relationships.length === 0) {
    delete archive[path];
    return;
  }
  const body = relationships.map((relationship) => (
    `<Relationship Id="${escapeAttribute(relationship.id)}" Type="${escapeAttribute(relationship.type)}" Target="${escapeAttribute(relationship.target)}"${relationship.targetMode ? ` TargetMode="${escapeAttribute(relationship.targetMode)}"` : ""}/>`
  )).join("");
  writeText(archive, path, `${XML_DECLARATION}<Relationships xmlns="${RELATIONSHIPS_NS}">${body}</Relationships>`);
}

function addRelationship(relationships: Relationship[], sourcePart: string, type: string, path: string): string {
  const ids = new Set(relationships.map((relationship) => relationship.id));
  let index = 1;
  while (ids.has(`rIdCmt${index}`)) {
    index += 1;
  }
  const id = `rIdCmt${index}`;
  relationships.push({ id, target: relativeTarget(sourcePart, path), type });
  return id;
}

function setContentTypeOverride(archive: XlsxArchive, path: string, contentType: string | null) {
  const xml = readText(archive, "[Content_Types].xml");
  if (!xml) {
    return;
  }
  const partName = `/${path}`;
  let next = xml.replace(new RegExp(`<Override\\b[^>]*PartName="${partName.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")}"[^>]*\\/>`, "g"), "");
  if (contentType) {
    next = next.replace(/<\/Types>\s*$/, `<Override PartName="${escapeAttribute(partName)}" ContentType="${contentType}"/></Types>`);
  }
  writeText(archive, "[Content_Types].xml", next);
}

function ensureContentTypeDefault(archive: XlsxArchive, extension: string, contentType: string) {
  const xml = readText(archive, "[Content_Types].xml");
  if (!xml || new RegExp(`<Default\\b[^>]*Extension="${extension}"`, "i").test(xml)) {
    return;
  }
  writeText(archive, "[Content_Types].xml", xml.replace(/(<Types\b[^>]*>)/, `$1<Default Extension="${extension}" ContentType="${contentType}"/>`));
}

function removePart(archive: XlsxArchive, path: string) {
  delete archive[path];
  delete archive[relsPathFor(path)];
  setContentTypeOverride(archive, path, null);
}

function unusedPath(archive: XlsxArchive, prefix: string, extension: string): string {
  let index = 1;
  while (archive[`${prefix}${index}${extension}`]) {
    index += 1;
  }
  return `${prefix}${index}${extension}`;
}

export interface XlsxWorkbookSheet {
  name: string;
  /** Archive path of the worksheet part. */
  path: string;
}

/** Worksheets in workbook order (chartsheets and other sheet kinds excluded). */
export function readWorkbookSheets(archive: XlsxArchive): XlsxWorkbookSheet[] {
  const workbookXml = readText(archive, "xl/workbook.xml");
  if (!workbookXml) {
    return [];
  }
  const relationships = new Map(readRelationships(archive, "xl/workbook.xml").map((relationship) => [relationship.id, relationship]));
  return xmlElements(workbookXml, "sheet").flatMap(({ attributes }) => {
    const relationship = relationships.get(attributes.id ?? "");
    if (!relationship || !relationship.type.endsWith("/worksheet")) {
      return [];
    }
    return [{ name: attributes.name ?? "", path: resolveTarget("xl/workbook.xml", relationship.target) }];
  });
}

// ---------------------------------------------------------------------------
// Cell references.

export interface XlsxCellPosition {
  row: number;
  col: number;
}

/** Normalizes `$b$2` to `B2`; throws on anything that is not a single cell. */
export function normalizeCellRef(ref: string): string {
  const normalized = ref.replace(/\$/g, "").trim().toUpperCase();
  const position = parseCellRef(normalized);
  if (!position) {
    throw new Error(`"${ref}" is not a cell reference like B2.`);
  }
  return normalized;
}

/** Zero-based row and column of an A1 reference, or null. */
export function parseCellRef(ref: string): XlsxCellPosition | null {
  const match = /^\$?([A-Za-z]{1,3})\$?([1-9]\d*)$/.exec(ref.trim());
  if (!match) {
    return null;
  }
  let col = 0;
  for (const letter of match[1]!.toUpperCase()) {
    col = col * 26 + letter.charCodeAt(0) - 64;
  }
  const row = Number(match[2]) - 1;
  if (col > 16384 || row >= 1048576) {
    return null;
  }
  return { col: col - 1, row };
}

export function formatCellRef({ row, col }: XlsxCellPosition): string {
  let letters = "";
  for (let value = col + 1; value > 0; value = Math.floor((value - 1) / 26)) {
    letters = String.fromCharCode(65 + ((value - 1) % 26)) + letters;
  }
  return `${letters}${row + 1}`;
}

function compareRefs(left: string, right: string): number {
  const a = parseCellRef(left) ?? { col: 0, row: 0 };
  const b = parseCellRef(right) ?? { col: 0, row: 0 };
  return a.row - b.row || a.col - b.col;
}

// ---------------------------------------------------------------------------
// Reading.

function readPersons(archive: XlsxArchive): { path: string | null; persons: Map<string, string>; xml: string | null } {
  const relationship = readRelationships(archive, "xl/workbook.xml").find((entry) => entry.type === REL_PERSON);
  const path = relationship ? resolveTarget("xl/workbook.xml", relationship.target) : null;
  const xml = path ? readText(archive, path) : null;
  const persons = new Map<string, string>();
  for (const { attributes } of xml ? xmlElements(xml, "person") : []) {
    if (attributes.id) {
      persons.set(attributes.id, attributes.displayName ?? "");
    }
  }
  return { path, persons, xml };
}

function legacyDrawingRelationshipId(sheetXml: string): string | null {
  return xmlElements(sheetXml, "legacyDrawing")[0]?.attributes.id ?? null;
}

/** Threads from xl/threadedComments, plus legacy notes that have no thread. */
export function readCommentThreads(archive: XlsxArchive): XlsxCommentThread[] {
  const { persons } = readPersons(archive);
  const threads: XlsxCommentThread[] = [];

  for (const sheet of readWorkbookSheets(archive)) {
    const relationships = readRelationships(archive, sheet.path);
    const sheetThreads = new Map<string, XlsxCommentThread>();

    for (const relationship of relationships.filter((entry) => entry.type === REL_THREADED)) {
      const xml = readText(archive, resolveTarget(sheet.path, relationship.target));
      for (const { attributes, body } of xml ? xmlElements(xml, "threadedComment") : []) {
        const textElement = xmlElements(body, "text")[0];
        const comment: XlsxComment = {
          author: persons.get(attributes.personId ?? "") ?? "",
          date: attributes.dT ?? "",
          id: attributes.id ?? "",
          personId: attributes.personId ?? "",
          text: textElement ? unescapeXml(textElement.body) : ""
        };
        const parent = attributes.parentId ? sheetThreads.get(attributes.parentId) : undefined;
        if (attributes.parentId) {
          comment.parentId = attributes.parentId;
          parent?.comments.push(comment);
          continue;
        }
        sheetThreads.set(comment.id, {
          comments: [comment],
          id: comment.id,
          legacy: false,
          ref: normalizeCellRef(attributes.ref ?? "A1"),
          resolved: attributes.done === "1" || attributes.done === "true",
          sheet: sheet.name
        });
      }
    }

    const threadedRefs = new Set(Array.from(sheetThreads.values(), (thread) => thread.ref));
    const commentsRelationship = relationships.find((entry) => entry.type === REL_COMMENTS);
    const commentsXml = commentsRelationship ? readText(archive, resolveTarget(sheet.path, commentsRelationship.target)) : null;
    if (commentsXml) {
      const authors = xmlElements(commentsXml, "author").map((element) => unescapeXml(element.body));
      for (const { attributes, body } of xmlElements(commentsXml, "comment")) {
        const ref = parseCellRef(attributes.ref ?? "") ? normalizeCellRef(attributes.ref!) : null;
        const author = authors[Number(attributes.authorId ?? 0)] ?? "";
        if (!ref || threadedRefs.has(ref) || author.startsWith("tc=")) {
          continue;
        }
        const id = `note:${sheet.name}!${ref}`;
        sheetThreads.set(id, {
          comments: [{ author, date: "", id, personId: "", text: xmlRunText(xmlElements(body, "text")[0]?.body ?? "") }],
          id,
          legacy: true,
          ref,
          resolved: false,
          sheet: sheet.name
        });
      }
    }

    threads.push(...Array.from(sheetThreads.values()).sort((left, right) => compareRefs(left.ref, right.ref)));
  }

  return threads;
}

// ---------------------------------------------------------------------------
// Writing.

function legacyNoteText(thread: XlsxCommentThread): string {
  if (thread.legacy) {
    return thread.comments[0]?.text ?? "";
  }
  const lines = thread.comments.map((comment, index) => `${index === 0 ? "Comment" : "Reply"}:\n    ${comment.text}`);
  return `${THREADED_NOTE_PREAMBLE}${lines.join("\n")}`;
}

function buildCommentsXml(threads: XlsxCommentThread[]): string {
  const authors: string[] = [];
  const authorIndex = (author: string) => {
    const existing = authors.indexOf(author);
    if (existing >= 0) {
      return existing;
    }
    authors.push(author);
    return authors.length - 1;
  };
  const comments = threads.map((thread) => {
    const author = thread.legacy ? thread.comments[0]?.author ?? "" : `tc=${thread.id}`;
    const uid = thread.legacy ? "" : ` xr:uid="${escapeAttribute(thread.id)}"`;
    return `<comment ref="${thread.ref}" authorId="${authorIndex(author)}" shapeId="0"${uid}><text><t xml:space="preserve">${escapeXml(legacyNoteText(thread))}</t></text></comment>`;
  }).join("");
  return `${XML_DECLARATION}<comments xmlns="${SPREADSHEET_NS}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="xr" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision"><authors>${authors.map((author) => `<author>${escapeXml(author)}</author>`).join("")}</authors><commentList>${comments}</commentList></comments>`;
}

function buildThreadedCommentsXml(threads: XlsxCommentThread[]): string {
  const body = threads.flatMap((thread) => thread.comments.map((comment, index) => {
    const date = comment.date ? ` dT="${escapeAttribute(comment.date)}"` : "";
    const parent = index > 0 ? ` parentId="${escapeAttribute(thread.id)}"` : "";
    const done = index === 0 && thread.resolved ? " done=\"1\"" : "";
    return `<threadedComment ref="${thread.ref}"${date} personId="${escapeAttribute(comment.personId)}" id="${escapeAttribute(comment.id)}"${parent}${done}><text>${escapeXml(comment.text)}</text></threadedComment>`;
  })).join("");
  return `${XML_DECLARATION}<ThreadedComments xmlns="${THREADED_NS}" xmlns:x="${SPREADSHEET_NS}">${body}</ThreadedComments>`;
}

function writePersons(archive: XlsxArchive, threads: XlsxCommentThread[]) {
  const { path: existingPath, persons, xml } = readPersons(archive);
  const missing = new Map<string, string>();
  for (const thread of threads) {
    for (const comment of thread.legacy ? [] : thread.comments) {
      if (comment.personId && !persons.has(comment.personId)) {
        missing.set(comment.personId, comment.author);
      }
    }
  }
  if (missing.size === 0) {
    return;
  }
  const added = Array.from(missing, ([id, name]) => (
    `<person displayName="${escapeAttribute(name)}" id="${escapeAttribute(id)}" userId="${escapeAttribute(name)}" providerId="None"/>`
  )).join("");
  const path = existingPath ?? "xl/persons/person.xml";
  if (xml && /<\/(?:\w+:)?personList>/.test(xml)) {
    writeText(archive, path, xml.replace(/(<\/(?:\w+:)?personList>)/, `${added}$1`));
  } else {
    writeText(archive, path, `${XML_DECLARATION}<personList xmlns="${THREADED_NS}" xmlns:x="${SPREADSHEET_NS}">${added}</personList>`);
  }
  if (!existingPath) {
    const relationships = readRelationships(archive, "xl/workbook.xml");
    addRelationship(relationships, "xl/workbook.xml", REL_PERSON, path);
    writeRelationships(archive, "xl/workbook.xml", relationships);
  }
  setContentTypeOverride(archive, path, CT_PERSON);
}

const VML_NOTE_SHAPE = /<v:shape\b[^>]*>[\s\S]*?<\/v:shape>/g;

function vmlIdMaps(xml: string): number[] {
  return Array.from(xml.matchAll(/<o:idmap\b[^>]*\bdata="([^"]*)"/g), (match) => match[1]!.split(",").map(Number))
    .flat()
    .filter(Number.isFinite);
}

function newVmlDrawing(idMap: number): string {
  return `<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">\n <o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="${idMap}"/></o:shapelayout>\n</xml>`;
}

const NOTE_SHAPE_TYPE = "<v:shapetype id=\"_x0000_t202\" coordsize=\"21600,21600\" o:spt=\"202\" path=\"m,l,21600r21600,l21600,xe\"><v:stroke joinstyle=\"miter\"/><v:path gradientshapeok=\"t\" o:connecttype=\"rect\"/></v:shapetype>";

function noteShape(id: number, zIndex: number, ref: string): string {
  const { row, col } = parseCellRef(ref)!;
  return `<v:shape id="_x0000_s${id}" type="#_x0000_t202" style="position:absolute;margin-left:59.25pt;margin-top:1.5pt;width:108pt;height:59.25pt;z-index:${zIndex};visibility:hidden" fillcolor="#ffffe1" o:insetmode="auto">`
    + "<v:fill color2=\"#ffffe1\"/><v:shadow on=\"t\" color=\"black\" obscured=\"t\"/><v:path o:connecttype=\"none\"/>"
    + "<v:textbox style=\"mso-direction-alt:auto\"><div style=\"text-align:left\"></div></v:textbox>"
    + `<x:ClientData ObjectType="Note"><x:MoveWithCells/><x:SizeWithCells/><x:Anchor>${col + 1}, 15, ${row}, 2, ${col + 3}, 15, ${row + 4}, 16</x:Anchor><x:AutoFill>False</x:AutoFill><x:Row>${row}</x:Row><x:Column>${col}</x:Column></x:ClientData></v:shape>`;
}

/** Replaces the note shapes of a VML drawing, keeping every other shape (form controls). */
function withNoteShapes(xml: string, refs: string[]): string {
  const kept = xml.replace(VML_NOTE_SHAPE, (shape) => (/ObjectType="Note"/.test(shape) ? "" : shape));
  if (refs.length === 0) {
    return kept;
  }
  const withType = kept.includes("id=\"_x0000_t202\"") ? kept : kept.replace(/(<xml\b[^>]*>)/, `$1\n ${NOTE_SHAPE_TYPE}`);
  const idMap = vmlIdMaps(withType)[0] ?? 1;
  const usedIds = Array.from(withType.matchAll(/id="_x0000_s(\d+)"/g), (match) => Number(match[1]));
  const firstId = Math.max(idMap * 1024, ...usedIds) + 1;
  const shapes = refs.map((ref, index) => noteShape(firstId + index, index + 1, ref)).join("\n");
  return withType.replace(/<\/xml>\s*$/, `${shapes}\n</xml>`);
}

const LEGACY_DRAWING_SUCCESSORS = ["legacyDrawingHF", "drawingHF", "picture", "oleObjects", "controls", "webPublishItems", "tableParts", "extLst"];

function insertLegacyDrawing(sheetXml: string, relationshipId: string): string {
  let xml = sheetXml;
  let prefix = /xmlns:(\w+)="http:\/\/schemas\.openxmlformats\.org\/officeDocument\/2006\/relationships"/.exec(xml)?.[1];
  if (!prefix) {
    prefix = "r";
    xml = xml.replace(/<((?:\w+:)?worksheet)\b/, `<$1 xmlns:r="${OFFICE_REL_NS}"`);
  }
  const element = `<legacyDrawing ${prefix}:id="${relationshipId}"/>`;
  const sheetDataEnd = Math.max(xml.indexOf("</sheetData>"), xml.indexOf("<sheetData/>"));
  const successor = new RegExp(`<(?:${LEGACY_DRAWING_SUCCESSORS.join("|")})[\\s/>]`, "g");
  successor.lastIndex = Math.max(0, sheetDataEnd);
  const match = successor.exec(xml);
  const at = match ? match.index : xml.lastIndexOf("</");
  return `${xml.slice(0, at)}${element}${xml.slice(at)}`;
}

function removeLegacyDrawing(sheetXml: string): string {
  return sheetXml.replace(/<(?:\w+:)?legacyDrawing\b[^>]*?(?:\/>|>[\s\S]*?<\/(?:\w+:)?legacyDrawing>)/g, "");
}

/**
 * Rewrites every comment part of the package from `threads`: threaded comments,
 * persons, and the legacy note plus VML shape Excel pairs with each thread.
 * Sheets with no threads lose their notes; other VML shapes are kept.
 */
export function writeCommentThreads(archive: XlsxArchive, threads: XlsxCommentThread[]): void {
  writePersons(archive, threads);
  const allIdMaps = Object.keys(archive)
    .filter((path) => path.endsWith(".vml"))
    .flatMap((path) => vmlIdMaps(readText(archive, path) ?? ""));
  let nextIdMap = Math.max(0, ...allIdMaps) + 1;

  for (const sheet of readWorkbookSheets(archive)) {
    const sheetThreads = threads
      .filter((thread) => thread.sheet === sheet.name && thread.comments.length > 0)
      .sort((left, right) => compareRefs(left.ref, right.ref));
    let relationships = readRelationships(archive, sheet.path);
    const ownedRelationships = relationships.filter((entry) => entry.type === REL_COMMENTS || entry.type === REL_THREADED);
    let sheetXml = readText(archive, sheet.path);
    if (!sheetXml || (sheetThreads.length === 0 && ownedRelationships.length === 0)) {
      continue;
    }

    for (const relationship of ownedRelationships) {
      removePart(archive, resolveTarget(sheet.path, relationship.target));
    }
    relationships = relationships.filter((entry) => !ownedRelationships.includes(entry));

    const vmlRelationshipId = legacyDrawingRelationshipId(sheetXml);
    const vmlRelationship = relationships.find((entry) => entry.id === vmlRelationshipId && entry.type === REL_VML);
    let vmlPath = vmlRelationship ? resolveTarget(sheet.path, vmlRelationship.target) : null;
    let vmlXml = vmlPath ? readText(archive, vmlPath) : null;

    if (sheetThreads.length > 0) {
      const threaded = sheetThreads.filter((thread) => !thread.legacy);
      if (threaded.length > 0) {
        const path = unusedPath(archive, "xl/threadedComments/threadedComment", ".xml");
        writeText(archive, path, buildThreadedCommentsXml(threaded));
        setContentTypeOverride(archive, path, CT_THREADED);
        addRelationship(relationships, sheet.path, REL_THREADED, path);
      }
      const commentsPath = unusedPath(archive, "xl/comments", ".xml");
      writeText(archive, commentsPath, buildCommentsXml(sheetThreads));
      setContentTypeOverride(archive, commentsPath, CT_COMMENTS);
      addRelationship(relationships, sheet.path, REL_COMMENTS, commentsPath);

      if (!vmlPath || vmlXml === null) {
        vmlPath = unusedPath(archive, "xl/drawings/vmlDrawing", ".vml");
        vmlXml = newVmlDrawing(nextIdMap);
        nextIdMap += 1;
        ensureContentTypeDefault(archive, "vml", CT_VML);
        sheetXml = insertLegacyDrawing(removeLegacyDrawing(sheetXml), addRelationship(relationships, sheet.path, REL_VML, vmlPath));
      }
      writeText(archive, vmlPath, withNoteShapes(vmlXml, sheetThreads.map((thread) => thread.ref)));
    } else if (vmlPath && vmlXml !== null) {
      const remaining = withNoteShapes(vmlXml, []);
      if (/<v:shape\b/.test(remaining)) {
        writeText(archive, vmlPath, remaining);
      } else {
        removePart(archive, vmlPath);
        relationships = relationships.filter((entry) => entry !== vmlRelationship);
        sheetXml = removeLegacyDrawing(sheetXml);
      }
    }

    writeRelationships(archive, sheet.path, relationships);
    writeText(archive, sheet.path, sheetXml);
  }
}

/** Saved XLSX bytes with their comment parts rewritten from `threads`. */
export function applyCommentThreads(bytes: Uint8Array, threads: XlsxCommentThread[]): Uint8Array {
  const archive = unzipSync(bytes);
  writeCommentThreads(archive, threads);
  return zipSync(archive, { level: 6 });
}

// ---------------------------------------------------------------------------
// Editing: pure functions over a thread list, shared by the controller and the
// headless agent API.

export function createCommentId(): string {
  const bytes = new Uint8Array(16);
  globalThis.crypto.getRandomValues(bytes);
  bytes[6] = (bytes[6]! & 0x0f) | 0x40;
  bytes[8] = (bytes[8]! & 0x3f) | 0x80;
  const hex = Array.from(bytes, (byte) => byte.toString(16).padStart(2, "0")).join("").toUpperCase();
  return `{${hex.slice(0, 8)}-${hex.slice(8, 12)}-${hex.slice(12, 16)}-${hex.slice(16, 20)}-${hex.slice(20)}}`;
}

function commentDate(date = new Date()): string {
  return `${date.toISOString().slice(0, 19)}.00`;
}

function personIdFor(threads: XlsxCommentThread[], author: string): string {
  for (const thread of threads) {
    const match = thread.comments.find((comment) => comment.author === author && comment.personId);
    if (match) {
      return match.personId;
    }
  }
  return createCommentId();
}

function requireThread(threads: XlsxCommentThread[], threadId: string): XlsxCommentThread {
  const thread = threads.find((entry) => entry.id === threadId);
  if (!thread) {
    throw new Error(`No comment thread with id ${threadId}.`);
  }
  return thread;
}

function replaceThread(threads: XlsxCommentThread[], threadId: string, next: XlsxCommentThread): XlsxCommentThread[] {
  return threads.map((thread) => (thread.id === threadId ? next : thread));
}

/** A legacy note becomes a threaded comment before anyone replies to or resolves it. */
function promoteNote(threads: XlsxCommentThread[], thread: XlsxCommentThread): XlsxCommentThread {
  if (!thread.legacy) {
    return thread;
  }
  const id = createCommentId();
  const root = thread.comments[0]!;
  const author = root.author || "Unknown";
  return {
    ...thread,
    comments: [{ ...root, author, id, personId: personIdFor(threads, author) }],
    id,
    legacy: false
  };
}

export interface XlsxAddCommentInput {
  sheet: string;
  ref: string;
  text: string;
  author: string;
}

export function addCommentThread(threads: XlsxCommentThread[], input: XlsxAddCommentInput): { thread: XlsxCommentThread; threads: XlsxCommentThread[] } {
  const ref = normalizeCellRef(input.ref);
  if (threads.some((thread) => thread.sheet === input.sheet && thread.ref === ref)) {
    throw new Error(`${input.sheet}!${ref} already has a comment thread; reply to it instead.`);
  }
  const id = createCommentId();
  const thread: XlsxCommentThread = {
    comments: [{ author: input.author, date: commentDate(), id, personId: personIdFor(threads, input.author), text: input.text }],
    id,
    legacy: false,
    ref,
    resolved: false,
    sheet: input.sheet
  };
  return { thread, threads: [...threads, thread] };
}

export function replyToCommentThread(threads: XlsxCommentThread[], threadId: string, text: string, author: string): XlsxCommentThread[] {
  const original = requireThread(threads, threadId);
  const thread = promoteNote(threads, original);
  const reply: XlsxComment = {
    author,
    date: commentDate(),
    id: createCommentId(),
    parentId: thread.id,
    personId: personIdFor(threads, author),
    text
  };
  return replaceThread(threads, threadId, { ...thread, comments: [...thread.comments, reply] });
}

export function setCommentThreadResolved(threads: XlsxCommentThread[], threadId: string, resolved: boolean): XlsxCommentThread[] {
  const thread = promoteNote(threads, requireThread(threads, threadId));
  return replaceThread(threads, threadId, { ...thread, resolved });
}

export function removeCommentThread(threads: XlsxCommentThread[], threadId: string): XlsxCommentThread[] {
  requireThread(threads, threadId);
  return threads.filter((thread) => thread.id !== threadId);
}

// ---------------------------------------------------------------------------
// Suggestions.

export function formatSuggestionText(value: XlsxSuggestionValue, previous: XlsxSuggestionValue | null | undefined, note?: string): string {
  const marker = `${SUGGESTION_MARKER} ${JSON.stringify(previous === undefined ? { value } : { previous, value })}`;
  return note?.trim() ? `${note.trim()}\n${marker}` : marker;
}

function isSuggestionValue(value: unknown): value is XlsxSuggestionValue {
  return typeof value === "string" || typeof value === "boolean" || (typeof value === "number" && Number.isFinite(value));
}

export function parseSuggestionText(text: string): { value: XlsxSuggestionValue; previous: XlsxSuggestionValue | null; note: string } | null {
  const lines = text.replace(/\r\n?/g, "\n").trimEnd().split("\n");
  const last = lines[lines.length - 1]!.trim();
  if (!last.startsWith(`${SUGGESTION_MARKER} `)) {
    return null;
  }
  try {
    const payload = JSON.parse(last.slice(SUGGESTION_MARKER.length + 1)) as { previous?: unknown; value?: unknown };
    if (!isSuggestionValue(payload.value)) {
      return null;
    }
    return {
      note: lines.slice(0, -1).join("\n").trim(),
      previous: isSuggestionValue(payload.previous) ? payload.previous : null,
      value: payload.value
    };
  } catch {
    return null;
  }
}

function suggestionStatus(thread: XlsxCommentThread): XlsxSuggestionStatus {
  if (!thread.resolved) {
    return "open";
  }
  const decision = thread.comments.slice(1).reverse().map((comment) => comment.text.trim())
    .find((text) => text === ACCEPTED_REPLY || text === REJECTED_REPLY);
  if (decision === ACCEPTED_REPLY) {
    return "accepted";
  }
  return decision === REJECTED_REPLY ? "rejected" : "resolved";
}

export function listSuggestions(threads: XlsxCommentThread[]): XlsxSuggestion[] {
  return threads.flatMap((thread) => {
    const root = thread.comments[0];
    const parsed = root && !thread.legacy ? parseSuggestionText(root.text) : null;
    return parsed ? [{
      author: root!.author,
      date: root!.date,
      note: parsed.note,
      previous: parsed.previous,
      ref: thread.ref,
      sheet: thread.sheet,
      status: suggestionStatus(thread),
      threadId: thread.id,
      value: parsed.value
    }] : [];
  });
}

export function requireOpenSuggestion(threads: XlsxCommentThread[], threadId: string): XlsxSuggestion {
  const suggestion = listSuggestions(threads).find((entry) => entry.threadId === threadId);
  if (!suggestion) {
    throw new Error(`Comment thread ${threadId} is not a suggestion.`);
  }
  if (suggestion.status !== "open") {
    throw new Error(`Suggestion ${threadId} is already ${suggestion.status}.`);
  }
  return suggestion;
}

/** Replies "Accepted" or "Rejected" and resolves the thread. The caller changes the cell. */
export function closeSuggestion(threads: XlsxCommentThread[], threadId: string, decision: "accepted" | "rejected", author: string): XlsxCommentThread[] {
  requireOpenSuggestion(threads, threadId);
  const replied = replyToCommentThread(threads, threadId, decision === "accepted" ? ACCEPTED_REPLY : REJECTED_REPLY, author);
  return setCommentThreadResolved(replied, threadId, true);
}
