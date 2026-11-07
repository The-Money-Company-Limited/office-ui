type SheetsWasmModule = typeof import("../vendor/sheets-wasm/duke_sheets_wasm.js");

let wasmModulePromise: Promise<SheetsWasmModule> | null = null;

export type XlsxWasmSource =
  | string
  | URL
  | Request
  | Response
  | BufferSource
  | WebAssembly.Module;

export type WorkerWasmSource = string | ArrayBuffer | WebAssembly.Module;

let hasConfiguredWasmSource = false;
let configuredWasmSource: XlsxWasmSource | undefined;
let configuredWorkerWasmSource: WorkerWasmSource | undefined;

function bufferSourceToArrayBuffer(source: ArrayBuffer | ArrayBufferView<ArrayBufferLike>): ArrayBuffer {
  if (source instanceof ArrayBuffer) {
    return source.slice(0);
  }

  const bytes = new Uint8Array(source.buffer, source.byteOffset, source.byteLength);
  const copy = new Uint8Array(bytes);
  return copy.buffer;
}

function sourceToWorkerSource(source: XlsxWasmSource): WorkerWasmSource | undefined {
  if (typeof source === "string") {
    return source;
  }
  if (typeof URL !== "undefined" && source instanceof URL) {
    return source.href;
  }
  if (typeof Request !== "undefined" && source instanceof Request) {
    return source.url;
  }
  if (source instanceof ArrayBuffer || ArrayBuffer.isView(source)) {
    return bufferSourceToArrayBuffer(source);
  }
  if (typeof WebAssembly !== "undefined" && source instanceof WebAssembly.Module) {
    return source;
  }

  return undefined;
}

export function setWasmSource(source: XlsxWasmSource): void {
  hasConfiguredWasmSource = true;
  configuredWasmSource = source;
  configuredWorkerWasmSource = sourceToWorkerSource(source);
}

export function initWasm(source?: XlsxWasmSource) {
  if (source !== undefined) {
    setWasmSource(source);
  }

  return getSheetsWasmModule();
}

export function canUseConfiguredWasmSourceInWorker(): boolean {
  return !hasConfiguredWasmSource || configuredWorkerWasmSource !== undefined;
}

export function getConfiguredWorkerWasmSource(): WorkerWasmSource | undefined {
  return configuredWorkerWasmSource;
}

// Formulas, typed entries and the grid follow a regional format, standing in for the regional
// settings Excel takes from the computer: 5/11/2026 is 5 November in en-GB, and de-DE shows 1.5
// as 1,5. It is the browser's language unless the host picks another.
let regionalFormat = globalThis.navigator?.language || "en-US";
let loadedModule: SheetsWasmModule | null = null;

/** The language tag workbooks opened from now on take their regional settings from. */
export function getRegionalFormat(): string {
  return regionalFormat;
}

/**
 * Sets the language tag (`en-GB`, `de-DE`) workbooks opened from now on take their regional
 * settings from. A controller's `setRegionalFormat` also re-applies it to its open workbook.
 */
export function setRegionalFormat(tag: string): void {
  regionalFormat = tag;
  loadedModule?.setLocale(tag);
}

/**
 * How a short date and a number show in a regional format (`05/11/2026 · 1,234.5` in en-GB), for
 * a picker to show beside it; `null` until the engine has loaded.
 */
export function getRegionalFormatSample(tag: string): string | null {
  return loadedModule?.regionalSample(tag) ?? null;
}

export function getSheetsWasmModule() {
  if (!wasmModulePromise) {
    wasmModulePromise = import("../vendor/sheets-wasm/duke_sheets_wasm.js").then(async (mod) => {
      if (configuredWasmSource !== undefined) {
        await mod.default({ module_or_path: configuredWasmSource });
      } else {
        await mod.default();
      }
      mod.setLocale(regionalFormat);
      loadedModule = mod;
      return mod;
    });
  }

  return wasmModulePromise;
}
