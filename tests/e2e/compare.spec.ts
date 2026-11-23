/**
 * Browser test of the Stage 1 development page: pick two files, press Compare,
 * download the tracked-changes .docx, and check the downloaded file contains
 * exactly the expected tracked changes.
 */
import { expect, test, type Page } from '@playwright/test';
import { readFile } from 'node:fs/promises';
import { PAIR_SPECS, buildPair } from '../fixtures/pairs';
import { fullExpectation, readTrackedChanges } from '../helpers/tracked-changes';

const DOCX_MIME = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';

async function compareInPage(page: Page, original: Uint8Array, modified: Uint8Array): Promise<Uint8Array> {
  await page.locator('#original').setInputFiles({ name: 'original.docx', mimeType: DOCX_MIME, buffer: Buffer.from(original) });
  await page.locator('#modified').setInputFiles({ name: 'modified.docx', mimeType: DOCX_MIME, buffer: Buffer.from(modified) });
  await page.getByRole('button', { name: 'Compare' }).click();
  await expect(page.locator('#status')).toHaveAttribute('data-state', 'done');
  const downloadPromise = page.waitForEvent('download');
  await page.getByRole('link', { name: /Download Word file/ }).click();
  const download = await downloadPromise;
  expect(download.suggestedFilename()).toBe('redline-tracked-changes.docx');
  return new Uint8Array(await readFile(await download.path()));
}

test('the page loads with Compare disabled until both files are chosen', async ({ page }) => {
  await page.goto('/');
  await expect(page.getByRole('heading', { name: 'Redline' })).toBeVisible();
  const button = page.getByRole('button', { name: 'Compare' });
  await expect(button).toBeDisabled();
  const pair = await buildPair(PAIR_SPECS[0]!);
  await page.locator('#original').setInputFiles({ name: 'original.docx', mimeType: DOCX_MIME, buffer: Buffer.from(pair.original) });
  await expect(button).toBeDisabled();
  await page.locator('#modified').setInputFiles({ name: 'modified.docx', mimeType: DOCX_MIME, buffer: Buffer.from(pair.modified) });
  await expect(button).toBeEnabled();
});

for (const spec of PAIR_SPECS) {
  test(`${spec.id}: the downloaded file has exactly the expected tracked changes`, async ({ page }) => {
    const pair = await buildPair(spec);
    await page.goto('/');
    const redline = await compareInPage(page, pair.original, pair.modified);
    const found = await readTrackedChanges(redline);
    expect(found.parts).toEqual(fullExpectation(spec.expected));
  });
}

test('the engine runs in a Web Worker, not on the page itself', async ({ page }) => {
  const pair = await buildPair(PAIR_SPECS[0]!);
  await page.goto('/');
  await compareInPage(page, pair.original, pair.modified);

  const loadedEngine = (): boolean =>
    performance.getEntriesByType('resource').some((entry) => entry.name.includes('/wasm/_framework/') && entry.name.endsWith('.wasm'));
  const workers = page.workers();
  expect(workers).toHaveLength(1);
  // The engine's WebAssembly files were fetched by the worker…
  expect(await workers[0]!.evaluate(loadedEngine)).toBe(true);
  // …and never by the page's own (main) thread.
  expect(await page.evaluate(loadedEngine)).toBe(false);
});

test('a file that is not a Word document gives a plain error and no download', async ({ page }) => {
  const pair = await buildPair(PAIR_SPECS[0]!);
  await page.goto('/');
  await page.locator('#original').setInputFiles({ name: 'original.docx', mimeType: DOCX_MIME, buffer: Buffer.from('this is not a docx file') });
  await page.locator('#modified').setInputFiles({ name: 'modified.docx', mimeType: DOCX_MIME, buffer: Buffer.from(pair.modified) });
  await page.getByRole('button', { name: 'Compare' }).click();
  await expect(page.locator('#status')).toHaveAttribute('data-state', 'error');
  await expect(page.locator('#status')).toContainText('The comparison failed');
  await expect(page.getByRole('link', { name: /Download Word file/ })).toBeHidden();
});
