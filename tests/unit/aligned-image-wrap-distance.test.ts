import { describe, expect, it } from "vitest";
import type { ImageRunNode } from "../../packages/doc-model/src";
import {
  absoluteFloatingImageStyle,
  resolveDualWrappedFloatingImageGeometry,
  wrappedFloatingImageStyle,
} from "../../packages/react-viewer/src/editor";

function image(
  align: "left" | "right",
  distance: number,
  tight = false
): ImageRunNode {
  return {
    type: "image",
    widthPx: 120,
    heightPx: 48,
    floating: {
      horizontalAlign: align,
      horizontalRelativeTo: "margin",
      verticalRelativeTo: "paragraph",
      yPx: 0,
      distLPx: distance,
      distRPx: distance,
      wrapType: tight ? "tight" : "square",
      wrapText: "bothSides",
      behindDocument: tight,
    },
  };
}

describe("aligned image frames and text clearance", () => {
  it.each([false, true])("keeps a left frame fixed while expanding text clearance, tight=%s", (tight) => {
    for (const distance of [0, 12]) {
      const value = image("left", distance, tight);
      const geometry = resolveDualWrappedFloatingImageGeometry(value, 624)!;
      expect(geometry.imageLeftPx).toBe(0);
      expect(geometry.exclusion).toMatchObject({ left: 0, right: 120 + distance });
      expect(wrappedFloatingImageStyle(value, { containerWidthPx: 624 })).toMatchObject({
        marginLeft: 0,
        marginRight: distance,
      });
      expect(absoluteFloatingImageStyle(value).left).toBe(0);
    }
  });

  it.each([false, true])("keeps a right frame fixed while expanding text clearance, tight=%s", (tight) => {
    for (const distance of [0, 12]) {
      const value = image("right", distance, tight);
      const geometry = resolveDualWrappedFloatingImageGeometry(value, 624)!;
      expect(geometry.imageLeftPx).toBe(504);
      expect(geometry.exclusion).toMatchObject({ left: 504 - distance, right: 624 });
      expect(wrappedFloatingImageStyle(value, { containerWidthPx: 624 })).toMatchObject({
        marginLeft: distance,
        marginRight: 0,
      });
      expect(absoluteFloatingImageStyle(value).right).toBe(0);
    }
  });

  it("preserves explicit and preview offsets independently of the alignment", () => {
    for (const xPx of [-6, 0, 120]) {
      const value = image("left", 12);
      value.floating!.xPx = xPx;
      expect(resolveDualWrappedFloatingImageGeometry(value, 624)!.imageLeftPx).toBe(xPx);
      expect(absoluteFloatingImageStyle(value, { marginOriginLeft: 96 }).left).toBe(96 + xPx);
    }
    const preview = resolveDualWrappedFloatingImageGeometry(image("left", 12), 624, { baseLeftPx: 214 })!;
    expect(preview.imageLeftPx).toBe(214);
    expect(preview.exclusion).toMatchObject({ left: 202, right: 346 });
  });

  it.each(["left", "right"] as const)("uses horizontal margin bounds independently of the vertical anchor, align=%s", (align) => {
    for (const verticalRelativeTo of ["page", "paragraph"]) {
      for (const distance of [0, 12]) {
        const value = image(align, distance);
        value.floating!.verticalRelativeTo = verticalRelativeTo;
        value.floating!.yPx = 240;
        const style = absoluteFloatingImageStyle(value, {
          marginOriginLeft: 72,
          marginOriginRight: 108,
          pageOriginTop: 0,
          paragraphOriginTop: 48,
        });
        expect(align === "left" ? style.left : style.right).toBe(align === "left" ? 72 : 108);
        expect(style.top).toBe(verticalRelativeTo === "page" ? 240 : 288);
      }
    }
  });

  it("keeps other positioning contexts unchanged", () => {
    for (const horizontalRelativeTo of ["page", "column"]) {
      const value = image("left", 12);
      value.floating!.horizontalRelativeTo = horizontalRelativeTo;
      expect(resolveDualWrappedFloatingImageGeometry(value, 624)!.imageLeftPx).toBe(12);
      expect(wrappedFloatingImageStyle(value, { containerWidthPx: 624 }).marginLeft).toBe(12);
      expect(absoluteFloatingImageStyle(value).left).toBe(12);
    }
  });

  it.each(["left", "right"] as const)("preserves fractional committed offsets for align=%s", (align) => {
    for (const xPx of [-6.25, 0, 120.25, 120.75]) {
      const value = image(align, 12);
      value.floating!.xPx = xPx;
      const geometry = resolveDualWrappedFloatingImageGeometry(value, 624)!;
      expect(geometry.imageLeftPx).toBe(xPx);
      expect(geometry.exclusion.left).toBe(xPx < 12 ? 0 : xPx - 12);
      expect(geometry.exclusion.right).toBe(xPx + 132);
      expect(absoluteFloatingImageStyle(value, { marginOriginLeft: 96 }).left).toBe(96 + xPx);
    }
  });

  it("bounds clearance while preserving an offset beyond the right edge", () => {
    const value = image("right", 12);
    value.floating!.xPx = 600.25;
    const geometry = resolveDualWrappedFloatingImageGeometry(value, 624)!;
    expect(geometry.imageLeftPx).toBe(600.25);
    expect(geometry.exclusion).toEqual({ left: 588.25, right: 624, top: 0, bottom: 48 });
  });

  it("retains explicit preview and drag-delta snapping", () => {
    const value = image("right", 12);
    value.floating!.xPx = 120.25;
    const geometry = resolveDualWrappedFloatingImageGeometry(value, 624, {
      baseLeftPx: 214.25,
      baseTopPx: 12.75,
      deltaX: 2.75,
      deltaY: -2.25,
    })!;
    expect(geometry.imageLeftPx).toBe(217);
    expect(geometry.imageTopPx).toBe(11);
    expect(geometry.exclusion).toEqual({ left: 205, right: 349, top: 11, bottom: 59 });
  });

  it.each([undefined, Number.NaN, Number.POSITIVE_INFINITY])("retains right and top fallback for offset=%s", (xPx) => {
    const value = image("right", 12);
    value.floating!.xPx = xPx;
    value.floating!.yPx = undefined;
    const geometry = resolveDualWrappedFloatingImageGeometry(value, 624)!;
    expect(geometry.imageLeftPx).toBe(504);
    expect(geometry.imageTopPx).toBe(0);
    expect(geometry.exclusion).toEqual({ left: 492, right: 624, top: 0, bottom: 48 });
  });
});
