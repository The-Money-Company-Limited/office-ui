/**
 * The numbers desktop Excel puts in the clipboard's HTML beside the text it shows. A number cell
 * carries its value as `x:num="46331"` (a date's serial, an amount without its currency sign),
 * written the same whatever the computer's regional settings, so a date copied where it shows as
 * 11/5/2026 pastes as the same day whatever the editor's regional format.
 *
 * Gives the table's rows of cells, `undefined` where a cell carries no number (text, or a number
 * Excel shows as written). `null` when the HTML has no such numbers, or has cells spanning rows,
 * which leave its rows unlike the clipboard's text.
 */
export function excelClipboardNumbers(html: string): (number | undefined)[][] | null {
  if (!/\bx:num\b/i.test(html) || /\browspan\s*=\s*["']?\s*([2-9]|\d{2,})/i.test(html)) return null;
  const rows: (number | undefined)[][] = [];
  for (const [, row] of html.matchAll(/<tr\b[^>]*>([\s\S]*?)<\/tr>/gi)) {
    const cells: (number | undefined)[] = [];
    for (const [, attributes] of row.matchAll(/<t[dh]\b([^>]*)>/gi)) {
      const raw = /\bx:num\s*=\s*(?:"([^"]*)"|'([^']*)'|([^\s"'>]+))/i.exec(attributes)?.slice(1).find((v) => v !== undefined)?.trim();
      const number = raw ? Number(raw) : NaN;
      cells.push(Number.isFinite(number) ? number : undefined);
      // A merged cell is one cell in the HTML and several in the text.
      const span = Number(/\bcolspan\s*=\s*["']?(\d+)/i.exec(attributes)?.[1] ?? 1);
      for (let extra = 1; extra < span; extra += 1) cells.push(undefined);
    }
    rows.push(cells);
  }
  return rows.length > 0 ? rows : null;
}
