import { wasmExtractDocumentBlocks } from "@extend-ai/react-docx-wasm";
import type { DocModel } from "./types";

export const DOCUMENT_BLOCK_TYPES = [
  "text", "heading", "section_heading", "figure", "table", "table_head", "table_cell",
  "key_value", "page_number", "barcode", "formula", "header", "footer", "watermark", "legend",
] as const;

export type DocumentBlockType = (typeof DOCUMENT_BLOCK_TYPES)[number];

export interface DocxSourceRef {
  documentId: string;
  nodeId: string;
  partName: string;
  story: string;
  start?: number;
  end?: number;
  offsetEncoding: "utf-16";
}

export interface DocumentBlock {
  object: "block";
  id: string;
  type: DocumentBlockType;
  content: string;
  details: Record<string, unknown>;
  sourceRefs: DocxSourceRef[];
  parentBlockId?: string;
  children: DocumentBlock[];
}

export interface StructuredDocument {
  schemaVersion: 1;
  documentId: string;
  blocks: DocumentBlock[];
  warnings: string[];
  geometryStatus: "requires_layout";
}

export async function extractDocumentBlocks(model: DocModel): Promise<StructuredDocument> {
  return await wasmExtractDocumentBlocks(model) as StructuredDocument;
}

export function flattenDocumentBlocks(blocks: readonly DocumentBlock[]): DocumentBlock[] {
  return blocks.flatMap(block => [block, ...flattenDocumentBlocks(block.children)]);
}
