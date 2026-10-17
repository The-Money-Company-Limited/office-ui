import assert from "node:assert/strict";
import { test } from "node:test";
import * as drawings from "./images.ts";
import type { XlsxImageAnchor } from "./types.ts";

const marker = { col: 0, row: 0, colOffsetEmu: 0, rowOffsetEmu: 0 };
const anchors: XlsxImageAnchor[] = [
  { kind: "absolute", positionEmu: { x: 0, y: 0 }, sizeEmu: { cx: 1, cy: 1 } },
  { kind: "one-cell", from: marker, sizeEmu: { cx: 1, cy: 1 } },
  { kind: "two-cell", from: marker, to: marker }
];
for (const anchor of anchors) test(`${anchor.kind}: ratio resize maps to cells and offsets`, () => {
  const rect = drawings.resizeImageRect({ left: 115, top: 55, width: 200, height: 100 }, "nw", -40, -5, 16, true);
  assert.deepEqual(rect, { left: 75, top: 35, width: 240, height: 120 });
  const result = drawings.rectToImageAnchor(rect, anchor, {
    contentOffsetLeft: 40, contentOffsetTop: 24,
    getColumnWidthPx: (col) => [30, 0, 70][col] ?? 50,
    getRowHeightPx: (row) => [10, 0, 30][row] ?? 20
  });
  assert.equal(result.kind, anchor.kind);
  if (result.kind === "absolute") assert.deepEqual(result.positionEmu, { x: 35 * 9525, y: 11 * 9525 });
  else assert.deepEqual(result.from, { col: 2, row: 2, colOffsetEmu: 5 * 9525, rowOffsetEmu: 9525 });
  if (result.kind === "two-cell") assert.deepEqual(result.to, { col: 6, row: 7, colOffsetEmu: 25 * 9525, rowOffsetEmu: 11 * 9525 });
  else assert.deepEqual(result.sizeEmu, { cx: 240 * 9525, cy: 120 * 9525 });
});
test("ratio resize keeps the opposite corner and minimum size", () => {
  assert.deepEqual(drawings.resizeImageRect({ left: 40, top: 24, width: 200, height: 100 }, "nw", 500, 500, 16, true),
    { left: 208, top: 108, width: 32, height: 16 });
});
test("edge ratio resize keeps the opposite edge and centers the other axis", () => {
  assert.deepEqual(drawings.resizeImageRect({ left: 40, top: 24, width: 200, height: 100 }, "e", 40, 0, 16, true),
    { left: 40, top: 14, width: 240, height: 120 });
});
test("moving snaps either edge to the nearest grid line without resizing", () => {
  assert.deepEqual(drawings.snapImageRect({ left: 98, top: 47, width: 153, height: 40 }, [40, 100, 160, 220, 280], [24, 44, 64, 84, 104], 5),
    { left: 100, top: 44, width: 153, height: 40 });
  assert.equal(drawings.snapImageRect({ left: 68, top: 31, width: 153, height: 40 }, [40, 100, 160, 220], [24, 44], 5).left, 67);
});
