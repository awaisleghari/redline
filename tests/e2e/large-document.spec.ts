/**
 * A 100-page contract must compare in the browser, in the Web Worker, in a
 * reasonable time, and the page must stay responsive while it does.
 */
import { expect, test } from '@playwright/test';
import { readFile } from 'node:fs/promises';
import { buildLargeContract } from '../fixtures/large';
import { readTrackedChanges } from '../helpers/tracked-changes';

const DOCX_MIME = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';

test('a 100-page contract compares in the browser', async ({ page }) => {
  const contract = await buildLargeContract();
  expect(contract.words).toBeGreaterThan(45_000); // about 100 pages at 450 words a page

  await page.goto('/app');
  await page.locator('#original').setInputFiles({ name: 'original.docx', mimeType: DOCX_MIME, buffer: Buffer.from(contract.original) });
  await page.locator('#modified').setInputFiles({ name: 'modified.docx', mimeType: DOCX_MIME, buffer: Buffer.from(contract.modified) });

  // While the worker is busy, the page's own thread keeps ticking: count 50 ms timer ticks.
  await page.evaluate(() => {
    const w = window as unknown as { ticks: number; longestGapMs: number };
    w.ticks = 0;
    w.longestGapMs = 0;
    let last = performance.now();
    setInterval(() => {
      const now = performance.now();
      w.longestGapMs = Math.max(w.longestGapMs, now - last);
      last = now;
      w.ticks++;
    }, 50);
  });

  const started = Date.now();
  await page.locator('#create').click();
  await expect(page.locator('#status')).toHaveAttribute('data-state', 'done', { timeout: 60_000 });
  // 'done' means the independent Checker also read all 100 pages and passed the output.
  await expect(page.locator('#report-tracked li[data-passed="true"]')).toHaveCount(6);
  await expect(page.locator('#report-static li[data-passed="true"]')).toHaveCount(9);
  await expect(page.locator('#doc').getByText('Comparison of documents')).toBeVisible({ timeout: 30_000 });
  const wallMs = Date.now() - started;
  const engineMs = Number(await page.locator('#status').getAttribute('data-engine-ms'));
  const longestGapMs = await page.evaluate(() => (window as unknown as { longestGapMs: number }).longestGapMs);
  console.log(`100-page contract: engine ${engineMs} ms, click-to-done (both Word files made, checked and shown) ${wallMs} ms, longest main-thread pause ${Math.round(longestGapMs)} ms`);

  // Generous ceilings: this guards against a collapse in speed, not small drifts.
  expect(engineMs).toBeLessThan(30_000);
  // The page itself never froze for long while the engine worked.
  expect(longestGapMs).toBeLessThan(1_000);

  // The PDF of the whole contract: converted by the server, read back and checked page by page in the browser.
  const pdfStarted = Date.now();
  const pdfDownload = page.waitForEvent('download', { timeout: 180_000 });
  await page.locator('#download').click();
  expect((await pdfDownload).suggestedFilename()).toBe('redline.pdf');
  await expect(page.locator('#pdf-status')).toHaveText('The PDF passed its checks.');
  await expect(page.locator('#report-pdf li[data-passed="true"]')).toHaveCount(2);
  // The pages view breaks the redline where the PDF does: one sheet per page of the PDF, each with words on it.
  const pagesStarted = Date.now();
  await page.locator('#view-pages').click();
  const sheets = page.locator('#doc .rl-sheet');
  await expect(sheets.first()).toContainText('page 1 of');
  const count = await sheets.count();
  console.log(`100-page contract: PDF made, checked and downloaded in ${pagesStarted - pdfStarted} ms; ${count} pages laid out in ${Date.now() - pagesStarted} ms`);
  expect(count).toBeGreaterThan(40);
  const lengths = await sheets.evaluateAll((all) => all.map((sheet) => sheet.textContent!.length));
  expect(lengths.filter((length) => length < 200).length).toBeLessThanOrEqual(2);
  // No words were lost or repeated by cutting the document into pages.
  const paged = await page.locator('#doc').evaluate((doc) => Array.from(doc.querySelectorAll('.rl-sheet')).map((sheet) => { const copy = sheet.cloneNode(true) as HTMLElement; copy.querySelector('.rl-pagelabel')?.remove(); return copy.textContent; }).join('').replace(/\s+/g, ''));
  await page.locator('#view-continuous').click();
  const whole = await page.locator('#doc').evaluate((doc) => doc.textContent!.replace(/\s+/g, ''));
  expect(paged).toBe(whole);

  await page.locator('#download-more').click();
  await page.locator('#download-menu [data-format="tracked"]').click();
  const downloadPromise = page.waitForEvent('download');
  await page.locator('#download').click();
  const redline = new Uint8Array(await readFile(await (await downloadPromise).path()));
  const found = await readTrackedChanges(redline);
  // Every edit made to the modified version shows up as some tracked change.
  expect(found.parts.body.insertions.length).toBeGreaterThanOrEqual(contract.edits.wordChanges);
  expect(found.parts.body.deletions.length).toBeGreaterThanOrEqual(contract.edits.wordChanges);
  expect(found.parts.body.movedFrom.length).toBeGreaterThanOrEqual(contract.edits.movedParagraphs);
  expect(found.parts.body.movedTo.length).toBe(found.parts.body.movedFrom.length);
  // The generated edits include a block of consecutive paragraphs moved together.
  expect(found.parts.body.movedFrom.length).toBeGreaterThan(contract.edits.movedParagraphs);
});
