export interface FontTableEntry {
  family: string;
  /** Child markup of the `w:font` element. */
  xml: string;
}

/**
 * Lists `w:font` elements that have child markup. A self-closing element such
 * as `<w:font w:name="Arial"/>` declares no embedded faces, so it must never
 * borrow the children of the following entry.
 */
export function fontTableEntries(fontTableXml: string): FontTableEntry[] {
  const entries: FontTableEntry[] = [];
  const openTag = /<w:font\b([^>]*)>/gi;
  let match: RegExpExecArray | null;
  while ((match = openTag.exec(fontTableXml))) {
    const attributes = match[1] ?? "";
    if (/\/\s*$/.test(attributes)) {
      continue;
    }
    const close = fontTableXml.indexOf("</w:font>", openTag.lastIndex);
    if (close < 0) {
      break;
    }
    const family = attributes.match(/\bw:name="([^"]*)"/i)?.[1]?.trim();
    if (family) {
      entries.push({ family, xml: fontTableXml.slice(openTag.lastIndex, close) });
    }
    openTag.lastIndex = close + "</w:font>".length;
  }
  return entries;
}
