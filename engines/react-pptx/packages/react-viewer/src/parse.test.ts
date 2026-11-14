import { existsSync, readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { PptxViewerError } from './errors';
import { parsePresentation, presentationParsingDefaults } from './parse';
import type { PresentationDocument } from '@extend-ai/react-pptx-model';
import type { ParsePresentationOptions } from './types';
import { parseWithWasm } from './wasm';

vi.mock('./wasm', () => ({ parseWithWasm: vi.fn() }));

const nativeDocument: PresentationDocument = {
  format: 'pptx',
  size: { widthEmu: 12_192_000, heightEmu: 6_858_000 },
  slides: [{ id: 'slide-1', index: 0, nodes: [] }],
  masters: [],
  layouts: [],
  themes: [],
  assets: {},
  warnings: [],
};

describe('parsePresentation', () => {
  beforeEach(() => vi.mocked(parseWithWasm).mockReset());

  it('accepts an already normalized presentation without a DOM parser', async () => {
    const document: PresentationDocument = {
      format: 'ppt',
      size: { widthEmu: 9_144_000, heightEmu: 6_858_000 },
      slides: [],
      masters: [],
      layouts: [],
      themes: [],
      assets: {},
      warnings: [],
    };
    const result = await parsePresentation(document);
    expect(result.document).toBe(document);
    expect(result.kind).toBe('parsed-presentation');
  });

  it('rejects unknown magic bytes before loading the renderer', async () => {
    await expect(parsePresentation(new Uint8Array([1, 2, 3, 4]))).rejects.toMatchObject({
      code: 'unsupported-format',
    } satisfies Partial<PptxViewerError>);
  });

  it('uses the owned Wasm model for PPTX input', async () => {
    vi.mocked(parseWithWasm).mockResolvedValueOnce(nativeDocument);

    const result = await parsePresentation(new Uint8Array([0x50, 0x4b, 0x03, 0x04]));

    expect(parseWithWasm).toHaveBeenCalledWith(
      expect.any(Uint8Array),
      expect.objectContaining({ formatHint: 'pptx' }),
    );
    expect(result).toEqual({
      kind: 'parsed-presentation',
      document: nativeDocument,
      warnings: [],
    });
    expect(result).not.toHaveProperty('rendererData');
  });

  it('reports native parsing performance', async () => {
    vi.mocked(parseWithWasm).mockResolvedValueOnce(nativeDocument);
    const measurements: string[] = [];

    await parsePresentation(new Uint8Array([0x50, 0x4b, 0x03, 0x04]), {
      onPerformanceMeasurement: (measurement) => measurements.push(measurement.phase),
    });

    expect(measurements).toEqual(['wasm-parsing']);
  });

  it('reports a native PPTX decode failure without loading a second parser', async () => {
    vi.mocked(parseWithWasm).mockResolvedValueOnce(null);

    await expect(parsePresentation(new Uint8Array([0x50, 0x4b, 0x03, 0x04]))).rejects.toMatchObject(
      { code: 'parse-failed' },
    );
  });

  it('enforces the default input ceiling', () => {
    expect(presentationParsingDefaults).toEqual({ maxInputBytes: 100 * 1024 * 1024 });
  });

  it('exposes only parser options implemented by the owned parser boundary', () => {
    const supportedOptions: Record<keyof ParsePresentationOptions, true> = {
      signal: true,
      formatHint: true,
      maxInputBytes: true,
      fetchInit: true,
      onPerformanceMeasurement: true,
      onDownloadProgress: true,
    };

    expect(Object.keys(supportedOptions)).toEqual([
      'signal',
      'formatHint',
      'maxInputBytes',
      'fetchInit',
      'onPerformanceMeasurement',
      'onDownloadProgress',
    ]);
  });
});

const fixturePath = [
  resolve(process.cwd(), 'tests/fixtures/pretext-layout-cases.pptx'),
  resolve(process.cwd(), '../../tests/fixtures/pretext-layout-cases.pptx'),
].find(existsSync);

if (!fixturePath) throw new Error('PPTX fixture was not found from the workspace root.');

const fixture = new Uint8Array(readFileSync(fixturePath));

/** Serves the fixture in 4 KiB chunks, as a network response would arrive. */
function respondWithFixture(headers: Record<string, string>): void {
  vi.stubGlobal(
    'fetch',
    vi.fn(async () => {
      let offset = 0;
      const body = new ReadableStream<Uint8Array>({
        pull(controller) {
          if (offset >= fixture.byteLength) {
            controller.close();
            return;
          }
          controller.enqueue(fixture.slice(offset, offset + 4096));
          offset += 4096;
        },
      });
      return new Response(body, { status: 200, headers });
    }),
  );
}

async function downloadWithProgress(options: ParsePresentationOptions = {}) {
  const progress: Array<[number, number | undefined]> = [];
  const result = await parsePresentation('https://example.test/deck.pptx', {
    ...options,
    onDownloadProgress: (loadedBytes, totalBytes) => progress.push([loadedBytes, totalBytes]),
  });
  return { progress, result };
}

describe('parsePresentation download progress', () => {
  beforeEach(() => {
    vi.mocked(parseWithWasm).mockReset();
    vi.mocked(parseWithWasm).mockResolvedValue(nativeDocument);
  });
  afterEach(() => vi.unstubAllGlobals());

  it('reports the bytes received against the declared length and parses the whole file', async () => {
    respondWithFixture({ 'content-length': String(fixture.byteLength) });

    const { progress } = await downloadWithProgress();

    expect(progress.length).toBe(Math.ceil(fixture.byteLength / 4096));
    expect(progress.every(([, total]) => total === fixture.byteLength)).toBe(true);
    expect(progress.at(-1)).toEqual([fixture.byteLength, fixture.byteLength]);
    expect(vi.mocked(parseWithWasm).mock.calls[0]?.[0]).toEqual(fixture);
  });

  it('reports no total when the server declares no length', async () => {
    respondWithFixture({});

    const { progress } = await downloadWithProgress();

    expect(progress.at(-1)).toEqual([fixture.byteLength, undefined]);
    expect(progress.every(([, total]) => total === undefined)).toBe(true);
    expect(vi.mocked(parseWithWasm).mock.calls[0]?.[0]).toEqual(fixture);
  });

  it('reports no total for a compressed response, whose length counts encoded bytes', async () => {
    respondWithFixture({
      'content-encoding': 'gzip',
      'content-length': String(fixture.byteLength - 1000),
    });

    const { progress } = await downloadWithProgress();

    expect(progress.every(([, total]) => total === undefined)).toBe(true);
    expect(progress.at(-1)?.[0]).toBe(fixture.byteLength);
  });

  it('stops a download of undeclared length once it passes the input limit', async () => {
    respondWithFixture({});
    const maxInputBytes = 10_000;

    await expect(downloadWithProgress({ maxInputBytes })).rejects.toMatchObject({
      code: 'resource-limit',
    } satisfies Partial<PptxViewerError>);
    expect(parseWithWasm).not.toHaveBeenCalled();
  });
});
