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

  await page.goto('/');
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
  await page.getByRole('button', { name: 'Compare' }).click();
  await expect(page.locator('#status')).toHaveAttribute('data-state', 'done', { timeout: 60_000 });
  // 'done' means the independent Checker also read all 100 pages and passed the output.
  await expect(page.locator('#report-tracked li[data-passed="true"]')).toHaveCount(6);
  await expect(page.locator('#report-static li[data-passed="true"]')).toHaveCount(9);
  await expect(page.frameLocator('#viewer').getByText('Comparison of documents')).toBeVisible({ timeout: 30_000 });
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
  await page.getByRole('button', { name: 'Make PDF' }).click();
  await expect(page.locator('#pdf-status')).toHaveText('The PDF passed its checks.', { timeout: 120_000 });
  console.log(`100-page contract: PDF made, checked and drawn in ${Date.now() - pdfStarted} ms (${await page.locator('#pdf-viewer canvas').count()} pages)`);
  await expect(page.locator('#report-pdf li[data-passed="true"]')).toHaveCount(2);

  const downloadPromise = page.waitForEvent('download');
  await page.locator('#download-tracked').click();
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
