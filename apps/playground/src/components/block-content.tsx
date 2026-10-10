import * as React from "react";
import { renderToStaticMarkup } from "react-dom/server";
import type { DocModel, DocNode, DocumentBlock, ImageRunNode, ParagraphStyle, TextStyle } from "@extend-ai/react-docx";
import { mathContent } from "../../../../packages/react-viewer/src/math-content";
import { resolveRenderableImageSource } from "../../../../packages/react-viewer/src/image-render";

type Inline = { type: string; text?: string; style?: TextStyle | null; link?: string; blockId?: string };
type BlockImages = ReadonlyMap<string, ImageRunNode>;

export function collectBlockImages(model: DocModel): BlockImages {
  const images = new Map<string, ImageRunNode>();
  const visit = (nodes: DocNode[]) => {
    for (const node of nodes) {
      if (node.type === "table") {
        for (const row of node.rows) for (const cell of row.cells) visit(cell.nodes);
      } else {
        for (const child of node.children) if (child.type === "image") {
          if (child.blockId) images.set(child.blockId, child);
          if (child.contentNodes) visit(child.contentNodes);
        }
      }
    }
  };
  visit(model.nodes);
  const metadata = model.metadata;
  for (const story of [
    ...metadata.headerSections, ...metadata.footerSections,
    ...(metadata.sections ?? []).flatMap(section => [...section.headerSections, ...section.footerSections]),
    ...(metadata.footnotes ?? []), ...(metadata.endnotes ?? []), ...(metadata.comments ?? []),
  ]) if (story.nodes) visit(story.nodes);
  return images;
}

function blockImage(block: DocumentBlock, images?: BlockImages): ImageRunNode | undefined {
  return images?.get(block.id) ?? block.sourceRefs.map(ref => images?.get(ref.nodeId)).find(Boolean);
}

const highlights: Record<string, string> = {
  yellow: "#fff59d", green: "#bbf7d0", cyan: "#a5f3fc", magenta: "#f5d0fe",
  red: "#fecaca", blue: "#bfdbfe", black: "#111827", white: "#ffffff",
};

function color(value?: string): string | undefined {
  if (!value || value === "auto" || value === "none") return undefined;
  if (/^[a-f\d]{6}$/i.test(value)) return `#${value}`;
  return /^(#[a-f\d]{3,8}|[a-z]+)$/i.test(value) ? value : undefined;
}

function safeLink(value?: string): string | undefined {
  return value && /^(https?:|mailto:|tel:|#)/i.test(value.trim()) ? value.trim() : undefined;
}

function inlineStyle(value?: TextStyle | null): React.CSSProperties {
  const style = value ?? {};
  return {
    fontWeight: style.bold ? 700 : undefined,
    fontStyle: style.italic ? "italic" : undefined,
    textDecoration: [style.underline && "underline", style.strike && "line-through"].filter(Boolean).join(" ") || undefined,
    color: color(style.color),
    backgroundColor: color(style.backgroundColor) ?? highlights[style.highlight ?? ""] ?? color(style.highlight),
    fontSize: style.fontSizePt ? `${style.fontSizePt * (style.verticalAlign ? 0.65 : 1)}pt` : style.verticalAlign ? "0.65em" : undefined,
    fontFamily: style.fontFamily,
    letterSpacing: style.characterSpacingTwips === undefined ? undefined : `${style.characterSpacingTwips / 20}pt`,
    verticalAlign: style.verticalAlign === "superscript" ? "super" : style.verticalAlign === "subscript" ? "sub" : undefined,
    direction: style.rightToLeft ? "rtl" : undefined,
  };
}

function fieldInline(block: DocumentBlock): Inline {
  const field = block.details.field as { style?: TextStyle; link?: string } | undefined;
  return { type: "text", text: block.content, style: field?.style, link: field?.link };
}

function inlines(block: DocumentBlock): Inline[] | undefined {
  if (block.type === "key_value" && block.details.kind === "form_field") return [fieldInline(block)];
  if (!Array.isArray(block.details.inlines)) return undefined;
  return (block.details.inlines as Inline[]).map(run => {
    const field = run.type === "form_field" ? block.children.find(child => child.id === run.blockId) : undefined;
    return field ? { ...run, ...fieldInline(field) } : run;
  });
}
const escapeHtml = (text: string) => text.replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll(">", "&gt;").replaceAll('"', "&quot;");
const escapeMarkdown = (text: string) => text.replace(/([\\`*_{}[\]<>|])/g, "\\$1");

function markdownRun(run: Inline): string {
  if (!run.text) return "";
  const [, before, value, after] = /^(\s*)([\s\S]*?)(\s*)$/.exec(run.text)!;
  let text = escapeMarkdown(value).replaceAll("\n", "<br />");
  const style = run.style ?? {};
  if (style.bold && value) text = `**${text}**`;
  if (style.italic && value) text = `*${text}*`;
  if (style.strike && value) text = `~~${text}~~`;
  text = `${before}${text}${after}`.replaceAll("\n", "<br />");
  if (style.underline) text = `<u>${text}</u>`;
  if (style.verticalAlign) {
    const tag = style.verticalAlign === "superscript" ? "sup" : "sub";
    text = `<${tag}>${text}</${tag}>`;
  }
  const css = inlineStyle(style);
  if (css.backgroundColor) text = `<mark style="background-color:${css.backgroundColor}">${text}</mark>`;
  const retained = Object.entries({ color: css.color, "font-family": css.fontFamily ? JSON.stringify(css.fontFamily) : undefined, "font-size": css.fontSize, "letter-spacing": css.letterSpacing, direction: css.direction }).filter(([, value]) => value !== undefined);
  if (retained.length) text = `<span style="${escapeHtml(retained.map(([key, value]) => `${key}:${value}`).join(";"))}">${text}</span>`;
  const href = safeLink(run.link);
  return href ? `[${text}](${encodeURI(href).replaceAll("(", "%28").replaceAll(")", "%29")})` : text;
}

function tableCells(block: DocumentBlock): DocumentBlock[] {
  return block.children.flatMap(child => child.type === "table_head" ? child.children : child.type === "table_cell" ? [child] : []);
}

function tableRows(block: DocumentBlock): Array<{ index: number; cells: DocumentBlock[] }> {
  const rows = new Map<number, DocumentBlock[]>();
  for (const cell of tableCells(block)) {
    const index = Number(cell.details.rowIndex);
    const cells = rows.get(index) ?? [];
    cells.push(cell);
    rows.set(index, cells);
  }
  const count = Number(block.details.rowCount) || Math.max(0, ...[...rows.keys()].map(index => index + 1));
  return Array.from({ length: count }, (_, index) => ({ index, cells: (rows.get(index) ?? []).sort((a, b) => Number(a.details.columnIndex) - Number(b.details.columnIndex)) }));
}

function formattedTableRows(block: DocumentBlock) {
  const occupied = new Map<number, number>();
  return tableRows(block).map(row => {
    const cells: Array<{ cell?: DocumentBlock; column: number }> = [];
    const width = Number(block.details.columnCount) || row.cells.length;
    for (let column = 0; column < width;) {
      if ((occupied.get(column) ?? -1) >= row.index) { column++; continue; }
      const cell = row.cells.find(cell => Number(cell.details.columnIndex) === column);
      cells.push({ cell, column });
      const span = Number(cell?.details.columnSpan) || 1;
      if (cell) for (let index = column; index < column + span; index++)
        occupied.set(index, row.index + (Number(cell.details.rowSpan) || 1) - 1);
      column += span;
    }
    return { index: row.index, cells };
  });
}

export function blockMarkdown(block: DocumentBlock, images?: BlockImages): string {
  const image = blockImage(block, images);
  const asset = image?.partName ?? (typeof block.details.asset === "string" ? block.details.asset : undefined);
  if (asset && !image?.syntheticTextBox) {
    const alt = escapeMarkdown(image?.alt ?? block.content).replaceAll("\n", " ");
    return `![${alt}](${encodeURI(asset).replaceAll("(", "%28").replaceAll(")", "%29")})`;
  }
  if (block.type === "table") {
    const rows = tableRows(block);
    const merged = tableCells(block).some(c => Number(c.details.rowSpan) > 1 || Number(c.details.columnSpan) > 1);
    if (merged) return renderToStaticMarkup(<FormattedBlock block={block} images={images} />);
    const width = Number(block.details.columnCount) || Math.max(1, ...rows.map(r => r.cells.length));
    const values = rows.map(row => {
      const cells = Array.from({ length: width }, () => "");
      for (const cell of row.cells) cells[Number(cell.details.columnIndex)] = blockMarkdown(cell, images).replaceAll("\n", "<br />");
      return `| ${cells.join(" | ")} |`;
    });
    if (!values.length) return block.content;
    values.splice(1, 0, `| ${Array.from({ length: width }, () => "---").join(" | ")} |`);
    return values.join("\n");
  }
  const runs = inlines(block);
  let text = runs ? runs.map(run => run.type === "block"
    ? block.children.filter(child => child.id === run.blockId).map(child => blockMarkdown(child, images)).join("\n")
    : markdownRun(run)).join("") : block.children.length
    ? block.children.map(child => blockMarkdown(child, images)).join("\n\n")
    : escapeMarkdown(block.content);
  if (block.type === "heading" || block.type === "section_heading")
    text = `${"#".repeat(Math.min(6, Math.max(1, Number(block.details.headingLevel) || 1)))} ${text}`;
  if (block.type === "formula" && typeof block.details.mathMl === "string") text = block.details.mathMl;
  return text;
}

function FigurePreview({ image, block }: { image: ImageRunNode; block: DocumentBlock }): React.JSX.Element {
  const src = React.useMemo(() => resolveRenderableImageSource(image), [image]);
  const [failedSrc, setFailedSrc] = React.useState<string>();
  const alt = image.alt || block.content || "Figure";
  if (!src || failedSrc === src) return <div className="whitespace-pre-wrap break-words text-muted-foreground">{alt} · Image preview unavailable</div>;
  const width = image.widthPx;
  const height = image.heightPx;
  const hasFrame = !!width && !!height && width > 0 && height > 0;
  const fraction = (value = 0) => Math.min(1, Math.max(0, value));
  const left = fraction(image.crop?.leftFraction);
  const top = fraction(image.crop?.topFraction);
  const visibleWidth = 1 - left - fraction(image.crop?.rightFraction);
  const visibleHeight = 1 - top - fraction(image.crop?.bottomFraction);
  const crop = hasFrame && !!image.crop && visibleWidth > 0 && visibleHeight > 0;
  return <div data-block-figure="true" style={{ position: "relative", overflow: "hidden", maxWidth: "100%", width: hasFrame ? width : undefined, aspectRatio: hasFrame ? `${width} / ${height}` : undefined }}>
    <img src={src} alt={alt} onError={() => setFailedSrc(src)} draggable={false}
      style={{ display: "block", maxWidth: crop ? "none" : "100%", width: crop ? `${100 / visibleWidth}%` : hasFrame ? "100%" : undefined,
        height: crop ? `${100 / visibleHeight}%` : hasFrame ? "100%" : "auto", position: hasFrame ? "absolute" : undefined,
        left: crop ? `${-left / visibleWidth * 100}%` : undefined, top: crop ? `${-top / visibleHeight * 100}%` : undefined,
        filter: image.cssFilter, opacity: image.cssOpacity }} />
  </div>;
}

function FormattedBlock({ block, images }: { block: DocumentBlock; images?: BlockImages }): React.JSX.Element {
  if (block.type === "table") return (
    <div className="max-w-full overflow-x-auto">
      <table style={{ borderCollapse: "collapse", width: "100%" }}>
        <tbody>{formattedTableRows(block).map(row => (
          <tr key={row.index}>{row.cells.map(({ cell, column }) => {
            if (!cell) return <td key={`empty-${column}`} style={{ border: "1px solid #b8b8b8", padding: 4 }} />;
            const style = cell.details.style as { backgroundColor?: string; verticalAlign?: string } | undefined;
            return <td key={cell.id} rowSpan={Number(cell.details.rowSpan) || 1} colSpan={Number(cell.details.columnSpan) || 1}
              style={{ border: "1px solid #b8b8b8", padding: 4, backgroundColor: color(style?.backgroundColor), verticalAlign: style?.verticalAlign === "center" ? "middle" : style?.verticalAlign === "bottom" ? "bottom" : "top" }}>
              <FormattedBlock block={cell} images={images} />
            </td>;
          })}</tr>
        ))}</tbody>
      </table>
    </div>
  );
  if (block.type === "formula" && typeof block.details.mathMl === "string")
    return <div>{mathContent({ type: "image", mathMl: block.details.mathMl, alt: block.content })}</div>;
  const image = blockImage(block, images);
  if (image && !image.syntheticTextBox) return <FigurePreview block={block} image={image} />;
  const runs = inlines(block);
  if (runs) {
    const style = block.details.style as ParagraphStyle | undefined;
    return <div style={{ textAlign: style?.align as React.CSSProperties["textAlign"], backgroundColor: color(style?.backgroundColor), whiteSpace: "pre-wrap", overflowWrap: "anywhere" }}>
      {runs.map((run, index) => {
        if (run.type === "block") return block.children.filter(child => child.id === run.blockId).map(child => <FormattedBlock key={child.id} block={child} images={images} />);
        const value = <span key={index} data-block-formatted-run="true" style={inlineStyle(run.style)}>{run.text}</span>;
        const href = safeLink(run.link);
        return href ? <a key={index} href={href} target="_blank" rel="noreferrer" onClick={event => event.stopPropagation()}>{value}</a> : value;
      })}
    </div>;
  }
  if (block.children.length) return <>{block.children.map(child => <FormattedBlock key={child.id} block={child} images={images} />)}</>;
  return <div className="whitespace-pre-wrap break-words">{block.content || (block.type === "key_value" ? "Form control" : block.type.replaceAll("_", " "))}</div>;
}

export function BlockContent({ block, mode, images }: { block: DocumentBlock; mode: "formatted" | "markdown"; images?: BlockImages }): React.JSX.Element {
  return mode === "markdown"
    ? <pre data-block-markdown="true" className="whitespace-pre-wrap break-words text-xs leading-5 [overflow-wrap:anywhere]">{blockMarkdown(block, images)}</pre>
    : <div data-block-formatted="true" className="rounded-sm bg-white p-2 text-sm leading-normal text-[#111827]"><FormattedBlock block={block} images={images} /></div>;
}
