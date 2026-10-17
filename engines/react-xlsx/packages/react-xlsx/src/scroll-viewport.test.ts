import assert from "node:assert/strict";
import test from "node:test";
import { resolveWindowScrollMetrics } from "./scroll-viewport.ts";

const page = { gridHeight: 5000, gridWidth: 3000, stickyTop: 48, viewportHeight: 800, viewportWidth: 1200 };

test("a grid just below the toolbar reports no scroll and the viewport below the toolbar", () => {
  const metrics = resolveWindowScrollMetrics({ ...page, gridLeft: 0, gridTop: 48 });
  assert.equal(metrics.scrollTop, 0);
  assert.equal(metrics.scrollLeft, 0);
  assert.equal(metrics.clientHeight, 752);
  assert.equal(metrics.clientWidth, 1200);
  assert.deepEqual(metrics.rect, { bottom: 800, height: 752, left: 0, right: 1200, top: 48, width: 1200 });
});

test("scrolling the window past the toolbar becomes the grid's scroll offset", () => {
  const metrics = resolveWindowScrollMetrics({ ...page, gridLeft: -300, gridTop: -952 });
  assert.equal(metrics.scrollTop, 1000);
  assert.equal(metrics.scrollLeft, 300);
  assert.equal(metrics.rect.top, 48);
  assert.equal(metrics.rect.left, 0);
});

test("a grid lower on the page is partly visible and not scrolled", () => {
  const metrics = resolveWindowScrollMetrics({ ...page, gridLeft: 20, gridTop: 500 });
  assert.equal(metrics.scrollTop, 0);
  assert.equal(metrics.clientHeight, 300);
  assert.equal(metrics.clientWidth, 1180);
  assert.equal(metrics.rect.top, 500);
});

test("the visible box never extends past the grid's end", () => {
  const metrics = resolveWindowScrollMetrics({ ...page, gridLeft: -2500, gridTop: -4600 });
  assert.equal(metrics.scrollTop + metrics.clientHeight, page.gridHeight);
  assert.equal(metrics.scrollLeft + metrics.clientWidth, page.gridWidth);
});

test("a grid scrolled entirely above the viewport has an empty client box", () => {
  const metrics = resolveWindowScrollMetrics({ ...page, gridLeft: 0, gridTop: -6000 });
  assert.equal(metrics.clientHeight, 0);
  assert.equal(metrics.scrollTop, page.gridHeight);
});
