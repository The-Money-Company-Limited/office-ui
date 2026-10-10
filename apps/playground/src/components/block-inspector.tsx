import * as React from "react";
import { useVirtualizer } from "@tanstack/react-virtual";
import { Braces, Download, Loader2 } from "lucide-react";
import {
  DOCUMENT_BLOCK_TYPES, DocxBlockHighlights, extractDocumentBlocks,
  revealDocxSource, type DocModel, type DocumentBlock, type DocumentBlockType,
  type DocxBlockGeometry, type StructuredDocument,
} from "@extend-ai/react-docx";
import { Button } from "./ui/button";

const colors: Record<DocumentBlockType, string> = {
  text: "#2563eb", heading: "#7c3aed", section_heading: "#7c3aed", figure: "#db2777",
  table: "#059669", table_head: "#059669", table_cell: "#0d9488", key_value: "#d97706",
  page_number: "#64748b", barcode: "#0891b2", formula: "#9333ea", header: "#64748b",
  footer: "#64748b", watermark: "#64748b", legend: "#c2410c",
};

function rows(blocks: DocumentBlock[], depth = 0): Array<{ block: DocumentBlock; depth: number }> {
  return blocks.flatMap(block => [{ block, depth }, ...rows(block.children, depth + 1)]);
}

export function BlockInspector({ model, rootRef }: { model: DocModel; rootRef: React.RefObject<HTMLDivElement | null> }): React.JSX.Element {
  const [output, setOutput] = React.useState<StructuredDocument>();
  const [error, setError] = React.useState<string>();
  const [busy, setBusy] = React.useState(true);
  const [activeId, setActiveId] = React.useState<string>();
  const [geometry, setGeometry] = React.useState<DocxBlockGeometry>();
  const [type, setType] = React.useState<string>("all");
  const [showJson, setShowJson] = React.useState(false);
  const scrollRef = React.useRef<HTMLDivElement>(null);
  React.useEffect(() => {
    let cancelled = false;
    setBusy(true); setError(undefined); setActiveId(undefined); setGeometry(undefined);
    void extractDocumentBlocks(model).then(result => {
      if (!cancelled) { setOutput(result); setBusy(false); }
    }).catch(error => { if (!cancelled) { setError(String(error)); setBusy(false); setOutput(undefined); } });
    return () => { cancelled = true; };
  }, [model]);
  const allRows = React.useMemo(() => rows(output?.blocks ?? []), [output]);
  const filtered = React.useMemo(() => allRows.filter(row => type === "all" || row.block.type === type), [allRows, type]);
  const active = allRows.find(row => row.block.id === activeId)?.block;
  const virtualizer = useVirtualizer({ count: filtered.length, getScrollElement: () => scrollRef.current, estimateSize: () => 108, overscan: 6 });
  const download = () => {
    if (!output) return;
    const url = URL.createObjectURL(new Blob([JSON.stringify(output, null, 2)], { type: "application/json" }));
    const link = document.createElement("a"); link.href = url; link.download = "document-blocks.json"; link.click();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  };
  return <>
    <aside aria-label="Parsed blocks" className="flex max-h-[45vh] min-h-0 w-full shrink-0 flex-col overflow-hidden rounded-lg border bg-background lg:max-h-none lg:w-[360px]">
      <div className="flex items-center gap-2 border-b px-4 py-3">
        <h2 className="text-sm font-semibold">Blocks</h2>
        <span className="text-xs text-muted-foreground">{busy ? <Loader2 className="size-3 animate-spin" /> : allRows.length}</span>
        <div className="ml-auto flex gap-1">
          <Button variant="ghost" size="icon-sm" aria-label="Show block JSON" aria-pressed={showJson} onClick={() => setShowJson(!showJson)}><Braces className="size-4" /></Button>
          <Button variant="ghost" size="icon-sm" aria-label="Download block output" disabled={!output || busy} onClick={download}><Download className="size-4" /></Button>
        </div>
      </div>
      <div className="border-b px-4 py-2">
        <select aria-label="Filter block type" className="w-full bg-background text-xs outline-offset-4" value={type} onChange={event => setType(event.target.value)}>
          <option value="all">All block types</option>
          {DOCUMENT_BLOCK_TYPES.map(type => <option key={type} value={type}>{type.replaceAll("_", " ")}</option>)}
        </select>
      </div>
      {error ? <p role="alert" className="p-4 text-sm text-destructive">{error}</p> : null}
      <div ref={scrollRef} className="min-h-0 flex-1 overflow-auto" onMouseLeave={() => setActiveId(undefined)}>
        {showJson ? <pre className="whitespace-pre-wrap break-all p-4 text-[11px] leading-5">{JSON.stringify(active ? { ...active, geometry } : output, null, 2)}</pre> :
          <div style={{ height: virtualizer.getTotalSize(), position: "relative" }}>
            {virtualizer.getVirtualItems().map(item => {
              const { block, depth } = filtered[item.index];
              const selected = block.id === activeId;
              const color = colors[block.type];
              return <button key={block.id} type="button" data-block-output-id={block.id} data-index={item.index} ref={virtualizer.measureElement}
                className="absolute left-0 top-0 w-full border-b px-4 py-3 text-left hover:bg-muted/50 focus-visible:z-10 focus-visible:outline-2 focus-visible:outline-primary"
                style={{ transform: `translateY(${item.start}px)`, paddingLeft: 16 + Math.min(depth, 4) * 12, backgroundColor: selected ? `${color}0d` : undefined }}
                onMouseEnter={() => { setGeometry(undefined); setActiveId(block.id); }}
                onFocus={() => { setGeometry(undefined); setActiveId(block.id); }}
                onBlur={() => setActiveId(undefined)}
                onClick={() => { if (rootRef.current) revealDocxSource(rootRef.current, block.sourceRefs[0]?.nodeId ?? block.id); }}>
                <div className="mb-2 flex items-center gap-2 text-[11px] font-medium">
                  <span className="size-2 rounded-sm" style={{ backgroundColor: color }} />
                  <span className="capitalize" style={{ color }}>{block.type.replaceAll("_", " ")}</span>
                  <span className="ml-auto font-normal text-muted-foreground">{selected && geometry?.fragments.length ? [...new Set(geometry.fragments.map(f => `p. ${f.pageIndex + 1}`))].join(", ") : block.sourceRefs[0]?.story}</span>
                </div>
                <p className="line-clamp-3 whitespace-pre-wrap break-words text-xs leading-5">{block.content || (block.children.length ? `${block.children.length} child blocks` : "Empty paragraph")}</p>
              </button>;
            })}
          </div>}
      </div>
      {output?.warnings.length ? <details className="border-t px-4 py-2 text-xs text-muted-foreground"><summary>{output.warnings.length} import diagnostics</summary><ul className="max-h-40 list-disc overflow-auto pl-4 pt-2">{output.warnings.map((warning, i) => <li key={i}>{warning}</li>)}</ul></details> : null}
      {active && geometry && geometry.status !== "rendered" ? <p className="border-t px-4 py-2 text-xs text-muted-foreground">{geometry.status === "not_mounted" ? "No visible fragment on mounted pages. Click to reveal." : geometry.status === "not_visible" ? "This content is not visible in the current view." : "Some source fragments are outside the current view."}</p> : null}
    </aside>
    <DocxBlockHighlights rootRef={rootRef} block={active} color={active ? colors[active.type] : undefined} onGeometry={setGeometry} />
  </>;
}
