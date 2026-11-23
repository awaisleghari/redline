/**
 * Browser test of the development page: pick two files, press Compare, see the
 * Checker's report, and download the tracked-changes .docx, which is offered
 * only when every check passed.
 */
import { expect, test, type Page } from '@playwright/test';
import { compareDocuments, initialize } from 'docxodus/core';
import { readFile } from 'node:fs/promises';
import { BROKEN_CASES, breakOutput } from '../fixtures/broken';
import { ALL_PAIR_SPECS, PAIR_SPECS, buildPair, pairById } from '../fixtures/pairs';
import { fullExpectation, readTrackedChanges } from '../helpers/tracked-changes';

const DOCX_MIME = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
const CHECKS = ['accept-all', 'reject-all', 'completeness', 'structure'];

async function chooseFiles(page: Page, original: Uint8Array | Buffer, modified: Uint8Array | Buffer): Promise<void> {
  await page.locator('#original').setInputFiles({ name: 'original.docx', mimeType: DOCX_MIME, buffer: Buffer.from(original) });
  await page.locator('#modified').setInputFiles({ name: 'modified.docx', mimeType: DOCX_MIME, buffer: Buffer.from(modified) });
}

async function downloadFrom(page: Page, linkName: RegExp): Promise<{ name: string; bytes: Uint8Array }> {
  const downloadPromise = page.waitForEvent('download');
  await page.getByRole('link', { name: linkName }).click();
  const download = await downloadPromise;
  return { name: download.suggestedFilename(), bytes: new Uint8Array(await readFile(await download.path())) };
}

async function compareInPage(page: Page, original: Uint8Array, modified: Uint8Array): Promise<Uint8Array> {
  await chooseFiles(page, original, modified);
  await page.getByRole('button', { name: 'Compare' }).click();
  await expect(page.locator('#status')).toHaveAttribute('data-state', 'done');
  const file = await downloadFrom(page, /Download Word file/);
  expect(file.name).toBe('redline-tracked-changes.docx');
  return file.bytes;
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

for (const spec of ALL_PAIR_SPECS) {
  test(`${spec.id}: passes every check and the download has exactly the expected tracked changes`, async ({ page }) => {
    const pair = await buildPair(spec);
    await page.goto('/');
    const redline = await compareInPage(page, pair.original, pair.modified);

    // The check report lists every check, all passed.
    const items = page.locator('#report li');
    await expect(items).toHaveCount(CHECKS.length);
    for (const check of CHECKS) await expect(page.locator(`#report li[data-check="${check}"]`)).toHaveAttribute('data-passed', 'true');

    const found = await readTrackedChanges(redline);
    expect(found.parts).toEqual(fullExpectation(spec.expected));
  });
}

test('the check report reads in plain English and can be downloaded as a text file', async ({ page }) => {
  const pair = await buildPair(PAIR_SPECS[0]!);
  await page.goto('/');
  await compareInPage(page, pair.original, pair.modified);
  await expect(page.locator('#report li[data-check="accept-all"]')).toHaveText('Accepting every change reproduces the modified document: passed');
  await expect(page.locator('#report li[data-check="reject-all"]')).toHaveText('Rejecting every change reproduces the original document: passed');

  const report = await downloadFrom(page, /Download check report/);
  expect(report.name).toBe('redline-check-report.txt');
  const text = new TextDecoder().decode(report.bytes);
  expect(text).toContain('Original: original.docx');
  expect(text).toContain('Modified: modified.docx');
  expect(text).toContain('Overall: all checks passed');
  expect(text).toContain('Accepting every change reproduces the modified document: passed');
  expect(text).toContain('The Word file is structurally sound: passed');
});

test('the engine and the Checker each run in their own Web Worker, not on the page', async ({ page }) => {
  const pair = await buildPair(PAIR_SPECS[0]!);
  await page.goto('/');
  await compareInPage(page, pair.original, pair.modified);

  const loadedEngine = (): boolean =>
    performance.getEntriesByType('resource').some((entry) => entry.name.includes('/wasm/_framework/') && entry.name.endsWith('.wasm'));
  const workers = page.workers();
  expect(workers).toHaveLength(2);
  const compareWorker = workers.find((w) => w.url().includes('compare.worker'));
  const checkerWorker = workers.find((w) => w.url().includes('checker.worker'));
  expect(compareWorker).toBeDefined();
  expect(checkerWorker).toBeDefined();
  // The engine's WebAssembly files were fetched by the compare worker only:
  expect(await compareWorker!.evaluate(loadedEngine)).toBe(true);
  // never by the Checker (it must not use the engine), and never by the page's own thread.
  expect(await checkerWorker!.evaluate(loadedEngine)).toBe(false);
  expect(await page.evaluate(loadedEngine)).toBe(false);
});

test('a file that is not a Word document gives a plain error and no download', async ({ page }) => {
  const pair = await buildPair(PAIR_SPECS[0]!);
  await page.goto('/');
  await chooseFiles(page, Buffer.from('this is not a docx file'), pair.modified);
  await page.getByRole('button', { name: 'Compare' }).click();
  await expect(page.locator('#status')).toHaveAttribute('data-state', 'error');
  await expect(page.locator('#status')).toContainText('The comparison failed');
  await expect(page.getByRole('link', { name: /Download Word file/ })).toBeHidden();
});

/**
 * Swap the Producer's genuine output for a damaged one on its way from the
 * compare worker to the page. This is done entirely from the test (by wrapping
 * the browser's Worker class before the page loads); the product has no switch
 * for it.
 */
async function tamperWithOutput(page: Page, replacement: Uint8Array): Promise<void> {
  await page.addInitScript((base64: string) => {
    const RealWorker = window.Worker;
    window.Worker = class extends RealWorker {
      constructor(url: string | URL, options?: WorkerOptions) {
        super(url, options);
        if (options?.name !== 'redline-compare') return;
        // Registered before the page's own handler, so it runs first.
        this.addEventListener('message', (event: MessageEvent) => {
          if (event.data?.type !== 'compared') return;
          const binary = atob(base64);
          const bytes = new Uint8Array(binary.length);
          for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
          event.data.trackedChanges = bytes.buffer;
        });
      }
    };
  }, Buffer.from(replacement).toString('base64'));
}

for (const caseId of ['dropped-word', 'changed-digit-in-insertion', 'broken-comment-anchor', 'unreferenced-part-with-text', 'not-a-word-file']) {
  const brokenCase = BROKEN_CASES.find((c) => c.id === caseId)!;
  test(`a broken output is withheld: ${caseId}`, async ({ page }) => {
    const pair = await buildPair(pairById(brokenCase.pairId));
    await initialize();
    const genuine = await compareDocuments(pair.original, pair.modified, { authorName: 'Redline' });
    await tamperWithOutput(page, await breakOutput(genuine, brokenCase));

    await page.goto('/');
    await chooseFiles(page, pair.original, pair.modified);
    await page.getByRole('button', { name: 'Compare' }).click();
    await expect(page.locator('#status')).toHaveAttribute('data-state', 'withheld');
    await expect(page.locator('#status')).toContainText('withheld');

    // The report names exactly the failed checks, each with an explanation…
    for (const check of CHECKS) {
      const item = page.locator(`#report li[data-check="${check}"]`);
      await expect(item).toHaveAttribute('data-passed', String(!brokenCase.caughtBy.includes(check as never)));
      if (brokenCase.caughtBy.includes(check as never)) await expect(item.locator('.failure')).not.toBeEmpty();
    }
    // …and the Word file is not offered: no visible link, and no address behind it.
    const link = page.locator('#download');
    await expect(link).toBeHidden();
    expect(await link.getAttribute('href')).toBeNull();

    // The report is still downloadable, and says the output was withheld.
    const report = await downloadFrom(page, /Download check report/);
    expect(new TextDecoder().decode(report.bytes)).toContain('FAILED - this output was withheld');
  });
}

test('a withheld output shows the first paragraph where it went wrong', async ({ page }) => {
  const brokenCase = BROKEN_CASES.find((c) => c.id === 'dropped-word')!;
  const pair = await buildPair(pairById(brokenCase.pairId));
  await initialize();
  const genuine = await compareDocuments(pair.original, pair.modified, { authorName: 'Redline' });
  await tamperWithOutput(page, await breakOutput(genuine, brokenCase));
  await page.goto('/');
  await chooseFiles(page, pair.original, pair.modified);
  await page.getByRole('button', { name: 'Compare' }).click();
  await expect(page.locator('#status')).toHaveAttribute('data-state', 'withheld');
  const failure = page.locator('#report li[data-check="accept-all"] .failure');
  await expect(failure).toContainText('Where: main text, paragraph 2');
  // An excerpt around the first difference is shown, not the whole paragraph.
  await expect(failure).toContainText(/Expected: .*provide the Services with exceptional skill and care\./);
  await expect(failure).toContainText(/Found: .*provide the Services with exceptional and care\./);
});
