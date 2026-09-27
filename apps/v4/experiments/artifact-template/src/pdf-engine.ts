import { createPdfiumEngine } from "@embedpdf/engines/pdfium-worker-engine"
import type { PdfEngine } from "@embedpdf/models"
import pdfiumWasmUrl from "@embedpdf/pdfium/pdfium.wasm?url"

let enginePromise: Promise<PdfEngine> | null = null

export function loadSharedPdfEngine() {
  enginePromise ??= Promise.resolve(createPdfiumEngine(pdfiumWasmUrl, {}))
  return enginePromise
}
