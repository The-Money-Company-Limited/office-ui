import { describe, expect, it } from "vitest";
import { createClippedBodyRectReader } from "../../packages/react-viewer/src/body-measurement";

function fixture() {
  let rectReads = 0;
  let styleReads = 0;
  type MockElement = HTMLElement & { mockStyle: { overflowX: string; overflowY: string } };
  const view = {
    getComputedStyle: (element: MockElement) => {
      styleReads += 1;
      return element.mockStyle;
    },
  };
  const element = (
    parent: HTMLElement | null,
    box: [number, number, number, number],
    overflowX = "visible",
    overflowY = "visible"
  ): HTMLElement => {
    const [left, top, width, height] = box;
    return {
      parentElement: parent,
      ownerDocument: { defaultView: view },
      mockStyle: { overflowX, overflowY },
      offsetWidth: width, offsetHeight: height,
      clientWidth: width, clientHeight: height, clientLeft: 0, clientTop: 0,
      getBoundingClientRect: () => {
        rectReads += 1;
        return { left, top, width, height, right: left + width, bottom: top + height };
      },
    } as unknown as HTMLElement;
  };
  const dimensions = (node: HTMLElement, values: Record<string, number>) => {
    for (const [key, value] of Object.entries(values)) {
      Object.defineProperty(node, key, { value, configurable: true });
    }
  };
  return { element, dimensions, reads: () => ({ rectReads, styleReads }) };
}

describe("body measurement clipping", () => {
  it("limits translated content to the local fragment viewport", () => {
    const { element } = fixture();
    const body = element(null, [0, 0, 500, 500]);
    const clip = element(body, [20, 100, 200, 40], "hidden", "hidden");
    const content = element(clip, [20, -300, 200, 1200]);
    expect(createClippedBodyRectReader(body)(content)).toEqual({
      left: 20, top: 100, right: 220, bottom: 140, width: 200, height: 40,
    });
  });

  it("removes completely clipped descendants from the visible extent", () => {
    const { element } = fixture();
    const body = element(null, [0, 0, 500, 500]);
    const clip = element(body, [20, 100, 200, 40], "visible", "hidden");
    const tail = element(clip, [20, 900, 200, 20]);
    expect(createClippedBodyRectReader(body)(tail)).toMatchObject({ width: 0, height: 0 });
  });

  it("intersects nested clipping independently on both axes", () => {
    const { element } = fixture();
    const body = element(null, [0, 0, 500, 500]);
    const vertical = element(body, [0, 60, 400, 100], "visible", "auto");
    const horizontal = element(vertical, [40, 0, 90, 300], "clip", "visible");
    const content = element(horizontal, [10, 10, 300, 200]);
    expect(createClippedBodyRectReader(body)(content)).toEqual({
      left: 40, top: 60, right: 130, bottom: 160, width: 90, height: 100,
    });
  });

  it("retains true overflow beyond the body and physical page", () => {
    const { element } = fixture();
    const page = element(null, [0, 0, 500, 500], "hidden", "hidden");
    const body = element(page, [0, 20, 500, 400], "hidden", "hidden");
    const flow = element(body, [0, 700, 400, 30]);
    expect(createClippedBodyRectReader(body)(flow).bottom).toBe(730);
  });

  it("retains a clipping container border but clips its descendants to the client box", () => {
    const { element, dimensions } = fixture();
    const body = element(null, [0, 0, 500, 500]);
    const clip = element(body, [20, 100, 200, 40], "visible", "hidden");
    dimensions(clip, { clientTop: 2, clientHeight: 36 });
    const content = element(clip, [20, 90, 200, 100]);
    const read = createClippedBodyRectReader(body);
    expect(read(clip).bottom).toBe(140);
    expect(read(content)).toMatchObject({ top: 102, bottom: 138 });
  });

  it("maps the client box to scaled viewport coordinates", () => {
    const { element, dimensions } = fixture();
    const body = element(null, [0, 0, 500, 500]);
    const clip = element(body, [20, 100, 200, 80], "hidden", "hidden");
    dimensions(clip, { offsetWidth: 100, offsetHeight: 40, clientLeft: 2,
      clientTop: 2, clientWidth: 96, clientHeight: 36 });
    const content = element(clip, [0, 0, 500, 500]);
    expect(createClippedBodyRectReader(body)(content)).toEqual({
      left: 24, top: 104, right: 216, bottom: 176, width: 192, height: 72,
    });
  });

  it("caches each rectangle and inherited clip for one measurement pass", () => {
    const { element, reads } = fixture();
    const body = element(null, [0, 0, 500, 500]);
    const clip = element(body, [20, 100, 200, 40], "visible", "hidden");
    const first = element(clip, [20, 100, 200, 20]);
    const second = element(clip, [20, 120, 200, 20]);
    const read = createClippedBodyRectReader(body);
    read(first); read(second); read(first); read(clip);
    expect(reads()).toEqual({ rectReads: 3, styleReads: 1 });
    createClippedBodyRectReader(body)(first);
    expect(reads()).toEqual({ rectReads: 5, styleReads: 2 });
  });
});
