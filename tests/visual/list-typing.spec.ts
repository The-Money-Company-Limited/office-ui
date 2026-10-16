import { expect, test, type Page } from "@playwright/test";
import { createZip } from "../unit/helpers/zip";

function typingDocx(): Buffer {
  return Buffer.from(createZip([
    { name: "[Content_Types].xml", content: `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>` },
    { name: "_rels/.rels", content: `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>` },
    { name: "word/document.xml", content: `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
      <w:p><w:r><w:t>First paragraph.</w:t></w:r></w:p>
      <w:p><w:r><w:t>Second paragraph.</w:t></w:r></w:p>
      <w:p><w:pPr><w:ind w:left="360"/><w:rPr><w:sz w:val="24"/></w:rPr></w:pPr><w:r><w:t>Paragraph 3 ends here.</w:t></w:r></w:p>
      <w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>
    </w:body></w:document>` },
  ]));
}

async function paragraphText(page: Page, index: number): Promise<string> {
  return page.evaluate((index) => (window as any).__DOCX_TEST_HOOKS__.getParagraphText(index), index);
}

async function focusState(page: Page) {
  return page.evaluate(() => {
    const element = document.activeElement;
    return {
      tag: element?.tagName,
      paragraph: element?.closest("[data-docx-paragraph-node-index]")?.getAttribute("data-docx-paragraph-node-index"),
    };
  });
}

test.beforeEach(async ({ page }) => {
  await page.goto("/");
  await page.locator('input[type="file"][accept*=".doc"]').setInputFiles({
    name: "list-typing.docx",
    mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    buffer: typingDocx(),
  });
  await expect.poll(() => paragraphText(page, 2)).toBe("Paragraph 3 ends here.");
  await page.waitForTimeout(500);
  await page.locator('[data-docx-paragraph-node-index="2"]').click();
  await page.waitForTimeout(300);
  await expect.poll(() => focusState(page)).toEqual({ tag: "DIV", paragraph: "2" });
  await page.keyboard.press("End");
  await page.waitForTimeout(300);
  await page.keyboard.press("Enter");
  await expect(page.locator('[data-docx-paragraph-node-index="3"]')).toBeVisible();
  await expect.poll(() => paragraphText(page, 3)).toBe("");
  await expect.poll(() => focusState(page)).toEqual({ tag: "TEXTAREA", paragraph: "3" });
});

test("keeps one typing surface through list conversion and insertion", async ({ page }) => {
  for (const [index, text] of [[3, "- apples"], [4, "pears"]] as const) {
    for (let offset = 0; offset < text.length; offset += 1) {
      await page.keyboard.type(text[offset]!);
      await page.waitForTimeout(300);
      const expected = index === 3 && offset > 0 ? text.slice(2, offset + 1) : text.slice(0, offset + 1);
      expect(await paragraphText(page, index)).toBe(expected);
      expect(await focusState(page)).toEqual({ tag: "TEXTAREA", paragraph: String(index) });
    }
    await page.keyboard.press("Enter");
    await page.waitForTimeout(300);
  }
});

test("keeps fast plain typing after Enter", async ({ page }) => {
  await page.keyboard.type("Plain paragraph", { delay: 0 });
  await expect.poll(() => paragraphText(page, 3)).toBe("Plain paragraph");
  await expect.poll(() => focusState(page)).toEqual({ tag: "TEXTAREA", paragraph: "3" });
  await page.keyboard.press("Enter");
  await page.keyboard.type("Next paragraph", { delay: 0 });
  await expect.poll(() => paragraphText(page, 4)).toBe("Next paragraph");
});

test("keeps the textarea through a literal tab and subsequent typing", async ({ page }) => {
  await page.keyboard.type("before");
  await page.keyboard.press("Tab");
  await page.keyboard.type("after");
  await expect.poll(() => paragraphText(page, 3)).toBe("before\tafter");
  await expect.poll(() => focusState(page)).toEqual({ tag: "TEXTAREA", paragraph: "3" });
});

async function listState(page: Page, index: number) {
  return page.evaluate((index) => {
    const element = document.querySelector(`[data-docx-paragraph-node-index="${index}"]`)!;
    return {
      listType: (window as any).__DOCX_TEST_HOOKS__.getLocationState({ kind: "paragraph", nodeIndex: index }).listType,
      markers: element.querySelectorAll('[data-docx-numbering-label="true"]').length,
      marginLeft: getComputedStyle(element).marginLeft,
    };
  }, index);
}

for (const nested of [true, false]) {
  test(`empty Enter ${nested ? "outdents a nested item" : "exits level zero without list indent"} across save/reopen`, async ({ page, context }, testInfo) => {
    await page.keyboard.type("- apples");
    await page.keyboard.press("Enter");
    if (nested) {
      await page.keyboard.type("pears");
      await page.keyboard.press("Tab");
      await page.keyboard.press("Enter");
    }
    const index = nested ? 5 : 4;
    await expect.poll(async () => (await listState(page, index)).listType).toBe("unordered");
    const before = await listState(page, index);
    await page.keyboard.press("Enter");
    const expectedType = nested ? "unordered" : null;
    await expect.poll(async () => (await listState(page, index)).listType).toBe(expectedType);
    const after = await listState(page, index);
    expect(after.markers).toBe(nested ? 1 : 0);
    if (nested) {
      expect(parseFloat(after.marginLeft)).toBeLessThan(parseFloat(before.marginLeft));
    } else {
      expect(after.marginLeft).toBe((await listState(page, 0)).marginLeft);
    }
    await expect.poll(() => focusState(page)).toEqual({ tag: "TEXTAREA", paragraph: String(index) });
    await expect.poll(() => paragraphText(page, index)).toBe("");

    const downloadPromise = page.waitForEvent("download");
    await page.evaluate(() => (window as any).__DOCX_TEST_HOOKS__.exportDocx());
    const download = await downloadPromise;
    const exportPath = testInfo.outputPath("empty-list-enter.docx");
    await download.saveAs(exportPath);
    const reopened = await context.newPage();
    await reopened.goto("/");
    await reopened.locator('input[type="file"][accept*=".doc"]').setInputFiles(exportPath);
    await expect.poll(() => paragraphText(reopened, 3)).toBe("apples");
    await reopened.evaluate((index) => (window as any).__DOCX_TEST_HOOKS__.selectParagraph(index), index);
    await expect.poll(() => listState(reopened, index)).toEqual(after);
    await expect.poll(() => paragraphText(reopened, index)).toBe("");
  });
}
