/** A tag's attributes, where a quoted value may hold a `>` (`x:fmla="=IF(A1>0,1)"`). */
const ATTRIBUTES = `(?:[^>"']|"[^"]*"|'[^']*')*`;
const ROW = new RegExp(`<tr\\b(${ATTRIBUTES})>([\\s\\S]*?)<\\/tr>`, "gi");
const CELL = new RegExp(`<t[dh]\\b(${ATTRIBUTES})>`, "gi");

/**
 * The numbers desktop Excel puts in the clipboard's HTML beside the text it shows. A number cell
 * carries its value as `x:num="46331"` (a date's serial, an amount without its currency sign),
 * written the same whatever the computer's regional settings, so a date copied where it shows as
 * 11/5/2026 pastes as the same day whatever the editor's regional format.
 *
 * Gives the numbers by the rows and columns of `grid`, the clipboard's text, `undefined` where a
 * cell carries none (text, or a number Excel shows as written). `null` when the HTML has no such
 * numbers, or its table doesn't line up with the text: cells spanning rows, or another shape.
 */
export function excelClipboardNumbers(html: string, grid: readonly (readonly string[])[]): (number | undefined)[][] | null {
  if (!/\bx:num\b/i.test(html)) return null;
  const rows: (number | undefined)[][] = [];
  for (const [, rowAttributes, row] of html.matchAll(ROW)) {
    // Excel ends the table with a hidden row of column widths.
    if (/display\s*:\s*none/i.test(rowAttributes)) continue;
    const cells: (number | undefined)[] = [];
    for (const [, attributes] of row.matchAll(CELL)) {
      if (Number(attribute(attributes, "rowspan") ?? 1) > 1) return null;
      const raw = attribute(attributes, "x:num")?.trim();
      const number = raw ? Number(raw) : NaN;
      cells.push(Number.isFinite(number) ? number : undefined);
      // A merged cell is one cell in the HTML and several in the text.
      for (let extra = 1; extra < Number(attribute(attributes, "colspan") ?? 1); extra += 1) cells.push(undefined);
    }
    rows.push(cells);
  }
  const fits = rows.length === grid.length && rows.every((cells, index) => cells.length === grid[index]?.length);
  return fits ? rows : null;
}

const ATTRIBUTE_PATTERNS = new Map<string, RegExp>();

/** An HTML attribute's value, quoted or not; `undefined` when it is missing or has none. */
function attribute(attributes: string, name: string): string | undefined {
  let pattern = ATTRIBUTE_PATTERNS.get(name);
  if (!pattern) {
    pattern = new RegExp(`(?:^|\\s)${name}\\s*=\\s*(?:"([^"]*)"|'([^']*)'|([^\\s"'>]+))`, "i");
    ATTRIBUTE_PATTERNS.set(name, pattern);
  }
  return pattern.exec(attributes)?.slice(1).find((value) => value !== undefined);
}
