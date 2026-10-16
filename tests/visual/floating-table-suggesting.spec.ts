import path from "node:path";
import { expect, test, type Page } from "@playwright/test";

async function expectWrapped(page: Page) {
  await expect.poll(() => page.evaluate(() => {
    const scroll = document.querySelector('[data-testid="float-scroll"]')!;
    const table = scroll.querySelector('[data-docx-floating-table-host] table');
    const paragraph = Array.from(scroll.querySelectorAll('[data-docx-paragraph-host]'))
      .find(element => element.textContent?.startsWith("Tables in Word"));
    if (!table || !paragraph) return null;
    const tableRect = table.getBoundingClientRect();
    const walker = document.createTreeWalker(paragraph, NodeFilter.SHOW_TEXT);
    const rects: DOMRect[] = [];
    while (walker.nextNode()) {
      const node = walker.currentNode;
      for (let i = 0; i < (node.textContent?.length ?? 0); i++) {
        if (!node.textContent![i].trim()) continue;
        const range = document.createRange();
        range.setStart(node, i);
        range.setEnd(node, i + 1);
        rects.push(range.getBoundingClientRect());
      }
    }
    const beside = rects.filter(rect => rect.bottom > tableRect.top && rect.top < tableRect.bottom);
    const below = rects.filter(rect => rect.top > tableRect.bottom + 5);
    return {
      hasTextBesideTable: beside.length > 0,
      clearsTable: beside.every(rect => rect.left >= tableRect.right + 10),
      resumesFullWidth: below.some(rect => rect.left < tableRect.left + 10),
    };
  })).toEqual({ hasTextBesideTable: true, clearsTable: true, resumesFullWidth: true });
}

for (const initialMode of ["read-only", "suggest"] as const) {
  test(`wraps the calibre floating table when entering Suggesting from ${initialMode}`, async ({ page }) => {
    await page.goto("/");
    const modulePath = `/@fs${path.resolve(process.cwd(), "tests/visual/fixtures/floating-table-suggesting.tsx")}`;
    await page.evaluate(async ({ modulePath, initialMode }) => {
      const fixture = await import(modulePath);
      await fixture.mount(initialMode);
    }, { modulePath, initialMode });
    const scroll = page.getByTestId("float-scroll");
    await expect(scroll.locator('[data-docx-floating-table-host]')).toBeAttached();
    // Enter Suggesting before scrolling to the Tables page, as in the report.
    await page.getByRole("button", { name: "suggest", exact: true }).click();
    await scroll.locator('[data-docx-floating-table-host]').scrollIntoViewIfNeeded();
    await expectWrapped(page);
    await expect(scroll.locator('[data-docx-tracked-change="insertion"]')).toHaveText("Existing suggestion");
    for (const mode of ["edit", "read-only", "suggest"]) {
      await page.getByRole("button", { name: mode, exact: true }).click();
      await scroll.evaluate(element => { element.scrollTop = 0; });
      await scroll.locator('[data-docx-floating-table-host]').scrollIntoViewIfNeeded();
      await expectWrapped(page);
    }
  });
}
