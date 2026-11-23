/**
 * Browser tests of the interface: upload two files, compare, see the redline in
 * the viewer, read the check report, download both Word files; and the
 * refusals and withheld outputs.
 */
import { expect, test, type Page } from '@playwright/test';
import { initialize } from 'docxodus/core';
import { readFile } from 'node:fs/promises';
import { checkStaticRedline, checkTrackedChanges, CHECKS_FOR } from '../../src/checker';
import { unzip, zip } from '../../src/common/zip';
import { makeStaticRedline } from '../../src/producer/make-static-redline';
import { produceTrackedChanges } from '../../src/producer/produce';
import { BROKEN_CASES, breakOutput, expectedFailures } from '../fixtures/broken';
import { BROKEN_STATIC_CASES, expectedStaticFailures } from '../fixtures/broken-static';
import { ALL_PAIR_SPECS, PAIR_SPECS, buildDocx, buildPair, pairById } from '../fixtures/pairs';
import { fullExpectation, readTrackedChanges } from '../helpers/tracked-changes';

const DOCX_MIME = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';

async function chooseFiles(page: Page, original: Uint8Array | Buffer, modified: Uint8Array | Buffer, names = ['original.docx', 'modified.docx']): Promise<void> {
  await page.locator('#original').setInputFiles({ name: names[0]!, mimeType: DOCX_MIME, buffer: Buffer.from(original) });
  await page.locator('#modified').setInputFiles({ name: names[1]!, mimeType: DOCX_MIME, buffer: Buffer.from(modified) });
}

async function downloadFrom(page: Page, selector: string): Promise<{ name: string; bytes: Uint8Array }> {
  const downloadPromise = page.waitForEvent('download');
  await page.locator(selector).click();
  const download = await downloadPromise;
  return { name: download.suggestedFilename(), bytes: new Uint8Array(await readFile(await download.path())) };
}

async function compare(page: Page, original: Uint8Array, modified: Uint8Array, expectedState = 'done'): Promise<void> {
  await page.goto('/');
  await chooseFiles(page, original, modified);
  await page.getByRole('button', { name: 'Compare' }).click();
  await expect(page.locator('#status')).toHaveAttribute('data-state', expectedState, { timeout: 60_000 });
}

test('the page loads with Compare disabled until both files are chosen, and works from the keyboard', async ({ page }) => {
  await page.goto('/');
  await expect(page.getByRole('heading', { name: 'Redline', level: 1 })).toBeVisible();
  await expect(page.getByLabel('Original')).toBeVisible();
  await expect(page.getByLabel('Modified')).toBeVisible();
  const button = page.getByRole('button', { name: 'Compare' });
  await expect(button).toBeDisabled();
  const pair = await buildPair(PAIR_SPECS[0]!);
  await page.locator('#original').setInputFiles({ name: 'original.docx', mimeType: DOCX_MIME, buffer: Buffer.from(pair.original) });
  await expect(button).toBeDisabled();
  await page.locator('#modified').setInputFiles({ name: 'modified.docx', mimeType: DOCX_MIME, buffer: Buffer.from(pair.modified) });
  await expect(button).toBeEnabled();
  await expect(page.locator('#name-original')).toHaveText('original.docx');
  // The footer links to the source and says where the comparison happens.
  await expect(page.getByRole('link', { name: 'Source code' })).toHaveAttribute('href', /github\.com/);
  await expect(page.locator('footer')).toContainText('compared inside your browser');
  // Keyboard: focus the button and press Enter.
  await button.focus();
  await page.keyboard.press('Enter');
  await expect(page.locator('#status')).toHaveAttribute('data-state', 'done', { timeout: 60_000 });
});

for (const spec of ALL_PAIR_SPECS) {
  test(`${spec.id}: both Word files pass every check, show in the viewer and download`, async ({ page }) => {
    const pair = await buildPair(spec);
    await compare(page, pair.original, pair.modified);

    // The check report: every check of both outputs, all passed.
    for (const [kind, output] of [['tracked', 'tracked-changes'], ['static', 'static-redline']] as const) {
      await expect(page.locator(`#report-${kind} li`)).toHaveCount(CHECKS_FOR[output].length);
      await expect(page.locator(`#report-${kind} li[data-passed="true"]`)).toHaveCount(CHECKS_FOR[output].length);
    }

    // The viewer shows the static redline, cover page first.
    const viewer = page.frameLocator('#viewer');
    await expect(viewer.getByText('Comparison of documents')).toBeVisible();
    await expect(viewer.getByText('Original: original.docx')).toBeVisible();

    // The summary bar shows the counts, which are the ones on the cover page.
    if (spec.staticCounts) {
      for (const [key, value] of Object.entries(spec.staticCounts)) await expect(page.locator(`#summary [data-count="${key}"]`)).toHaveText(String(value));
      await expect(viewer.getByText(`Insertions: ${spec.staticCounts.insertions}`)).toBeVisible();
    }

    // Word, tracked changes: exactly the expected changes, each kind under its own author.
    const tracked = await downloadFrom(page, '#download-tracked');
    expect(tracked.name).toBe('redline-tracked-changes.docx');
    expect((await readTrackedChanges(tracked.bytes)).parts).toEqual(fullExpectation(spec.expected));
    // Word, static redline: passes the Checker again here, outside the browser.
    const staticRedline = await downloadFrom(page, '#download-static');
    expect(staticRedline.name).toBe('redline-static.docx');
    const coverTime = (await viewer.getByText(/^Compared: /).textContent())!.replace('Compared: ', '');
    const recheck = await checkStaticRedline({ original: pair.original, modified: pair.modified, output: staticRedline.bytes, originalName: 'original.docx', modifiedName: 'modified.docx', comparedAt: coverTime });
    expect(recheck.checks.filter((c) => !c.passed)).toEqual([]);
    expect((await checkTrackedChanges({ original: pair.original, modified: pair.modified, output: tracked.bytes })).passed).toBe(true);
  });
}

test('each kind of tracked change has its own author, so Word colours them differently', async ({ page }) => {
  const pair = await buildPair(pairById('moved-paragraph'));
  await compare(page, pair.original, pair.modified);
  const tracked = await downloadFrom(page, '#download-tracked');
  const xml = new TextDecoder().decode((await unzip(tracked.bytes)).get('word/document.xml')!);
  expect(xml).toMatch(/<w:moveFrom [^>]*w:author="Redline \(moved\)"/);
  expect(xml).toMatch(/<w:moveTo [^>]*w:author="Redline \(moved\)"/);
  const word = await buildPair(pairById('one-word-change'));
  await compare(page, word.original, word.modified);
  const xml2 = new TextDecoder().decode((await unzip((await downloadFrom(page, '#download-tracked')).bytes)).get('word/document.xml')!);
  expect(xml2).toMatch(/<w:ins [^>]*w:author="Redline \(inserted\)"/);
  expect(xml2).toMatch(/<w:del [^>]*w:author="Redline \(deleted\)"/);
});

test('a link inside a document cannot be followed from the viewer, and the page requests nothing from outside', async ({ page }) => {
  const outside: string[] = [];
  page.on('request', (request) => {
    if (!request.url().startsWith('http://localhost:4173/') && !request.url().startsWith('blob:') && !request.url().startsWith('data:') && !request.url().startsWith('about:')) outside.push(request.url());
  });
  const pair = await buildPair(pairById('hyperlink-address-change'));
  await compare(page, pair.original, pair.modified);
  const viewer = page.frameLocator('#viewer');
  await expect(viewer.getByText('our website').first()).toBeVisible();
  // The link's words are shown, but there is nowhere for a click to go.
  expect(await viewer.locator('a[href]').count()).toBe(0);
  await viewer.getByText('our website').first().click({ force: true });
  await expect(viewer.getByText('Comparison of documents')).toBeVisible();
  expect(outside).toEqual([]);
});

test('making the viewer inert never changes the words shown', async ({ page }) => {
  // Words that look like link attributes must survive untouched in the viewer.
  const sentence = 'The corrective action = termination; set href="none" and ping=5 within thirty days.';
  const block = (text: string) => ({ blocks: [{ kind: 'p' as const, text: 'Schedule of remedies.' }, { kind: 'p' as const, text }] });
  const original = await buildDocx(block(sentence));
  const modified = await buildDocx(block(sentence.replace('thirty', 'sixty')));
  await compare(page, original, modified);
  const viewer = page.frameLocator('#viewer');
  await expect(viewer.getByText('The corrective action = termination; set href="none" and ping=5 within')).toBeVisible();
  await expect(viewer.getByText('sixty')).toBeVisible();
});

test('files cannot be changed while a comparison runs, and choosing a new file clears the old results', async ({ page }) => {
  const pair = await buildPair(PAIR_SPECS[0]!);
  await page.goto('/');
  await chooseFiles(page, pair.original, pair.modified);
  await page.getByRole('button', { name: 'Compare' }).click();
  // While it runs: both file inputs and the button are disabled.
  await expect(page.locator('#original')).toBeDisabled();
  await expect(page.locator('#modified')).toBeDisabled();
  await expect(page.getByRole('button', { name: 'Compare' })).toBeDisabled();
  await expect(page.locator('#status')).toHaveAttribute('data-state', 'done', { timeout: 60_000 });
  await expect(page.locator('#results')).toBeVisible();
  // Choosing a different file afterwards: the previous results, links and status go.
  await page.locator('#original').setInputFiles({ name: 'another.docx', mimeType: DOCX_MIME, buffer: Buffer.from(pair.modified) });
  await expect(page.locator('#results')).toBeHidden();
  await expect(page.locator('#status')).toHaveText('');
  expect(await page.locator('#download-tracked').getAttribute('href')).toBeNull();
  expect(await page.locator('#viewer').getAttribute('srcdoc')).toBeNull();
});

test('the page fits a narrow phone screen without scrolling sideways', async ({ page }) => {
  await page.setViewportSize({ width: 360, height: 800 });
  await page.goto('/');
  expect(await page.evaluate(() => document.documentElement.scrollWidth <= window.innerWidth)).toBe(true);
});

test('the check report reads in plain English and can be downloaded as a text file', async ({ page }) => {
  const pair = await buildPair(PAIR_SPECS[0]!);
  await compare(page, pair.original, pair.modified);
  await expect(page.locator('#report-tracked li[data-check="accept-all"]')).toHaveText('Accepting every change reproduces the modified document: passed');
  await expect(page.locator('#report-static li[data-check="cover-page"]')).toContainText('its counts match the redline: passed');
  await expect(page.locator('#verdict-static')).toHaveText('(all checks passed)');

  const report = await downloadFrom(page, '#download-report');
  expect(report.name).toBe('redline-check-report.txt');
  const text = new TextDecoder().decode(report.bytes);
  expect(text).toContain('Original: original.docx');
  expect(text).toContain('Output: Word file with the static redline');
  expect(text).toContain('Output: Word file with tracked changes');
  expect(text).toContain('Overall: all checks passed');
  expect(text).toContain('Rejecting every change reproduces the original document: passed');
  expect(text).toContain('Counted independently: 1 insertions, 1 deletions, 0 moves, 0 formatting changes.');
});

test('the PDF button is present but not yet available, with its privacy notice', async ({ page }) => {
  await page.goto('/');
  const pair = await buildPair(PAIR_SPECS[0]!);
  await compare(page, pair.original, pair.modified);
  await expect(page.getByRole('button', { name: 'PDF' })).toBeDisabled();
  await expect(page.locator('#pdf-note')).toContainText('sends the static redline to our server for conversion, and it is deleted immediately');
});

test('the Producer and the Checker each run in their own Web Worker, not on the page', async ({ page }) => {
  const pair = await buildPair(PAIR_SPECS[0]!);
  await compare(page, pair.original, pair.modified);
  const loadedEngine = (): boolean =>
    performance.getEntriesByType('resource').some((entry) => entry.name.includes('/wasm/_framework/') && entry.name.endsWith('.wasm'));
  const workers = page.workers();
  expect(workers).toHaveLength(2);
  const compareWorker = workers.find((w) => w.url().includes('compare.worker'));
  const checkerWorker = workers.find((w) => w.url().includes('checker.worker'));
  expect(compareWorker).toBeDefined();
  expect(checkerWorker).toBeDefined();
  expect(await compareWorker!.evaluate(loadedEngine)).toBe(true);
  // The Checker never loads the engine; nor does the page's own thread.
  expect(await checkerWorker!.evaluate(loadedEngine)).toBe(false);
  expect(await page.evaluate(loadedEngine)).toBe(false);
});

test.describe('refusals, in plain English', () => {
  const refuse = async (page: Page, original: { name: string; buffer: Buffer }, expected: RegExp): Promise<void> => {
    const pair = await buildPair(PAIR_SPECS[0]!);
    await page.goto('/');
    await page.locator('#original').setInputFiles({ name: original.name, mimeType: 'application/octet-stream', buffer: original.buffer });
    await page.locator('#modified').setInputFiles({ name: 'modified.docx', mimeType: DOCX_MIME, buffer: Buffer.from(pair.modified) });
    await page.getByRole('button', { name: 'Compare' }).click();
    await expect(page.locator('#status')).toHaveAttribute('data-state', 'refused');
    await expect(page.locator('#status')).toHaveText(expected);
    await expect(page.locator('#results')).toBeHidden();
    await expect(page.locator('#download-tracked')).toBeHidden();
  };

  test('an old .doc file', async ({ page }) => {
    await refuse(page, { name: 'contract.doc', buffer: Buffer.from('old word file') }, /old-format Word file \(\.doc\).*save it as a Word Document \(\.docx\)/);
  });
  test('a file that is not a Word document', async ({ page }) => {
    await refuse(page, { name: 'contract.pdf', buffer: Buffer.from('%PDF-1.7') }, /is not a Word document\. Only \.docx files can be compared\./);
  });
  test('a file over 25 MB', async ({ page }) => {
    await refuse(page, { name: 'huge.docx', buffer: Buffer.alloc(25 * 1024 * 1024 + 1) }, /larger than 25 MB/);
  });
  test('a .docx that is not really one', async ({ page }) => {
    await refuse(page, { name: 'fake.docx', buffer: Buffer.from('this is not a docx file') }, /could not be read as a Word document/);
  });
  test('a document that already contains tracked changes', async ({ page }) => {
    const pair = await buildPair(PAIR_SPECS[0]!);
    await initialize();
    const withChanges = await produceTrackedChanges(pair.original, pair.modified);
    await refuse(page, { name: 'marked-up.docx', buffer: Buffer.from(withChanges) }, /already contains tracked changes\. Please accept or reject the existing changes first, then compare\./);
  });
});

/**
 * Swap one of the Producer's genuine outputs for a damaged one on its way from
 * the Producer's worker to the page. Done entirely from the test, by wrapping
 * the browser's Worker class before the page loads; the product has no switch
 * for it.
 */
async function tamper(page: Page, messageType: 'compared' | 'static-made', field: 'trackedChanges' | 'staticRedline', replacement: Uint8Array): Promise<void> {
  await page.addInitScript(
    ({ base64, messageType, field }) => {
      const RealWorker = window.Worker;
      window.Worker = class extends RealWorker {
        constructor(url: string | URL, options?: WorkerOptions) {
          super(url, options);
          if (options?.name !== 'redline-compare') return;
          this.addEventListener('message', (event: MessageEvent) => {
            if (event.data?.type !== messageType) return;
            const binary = atob(base64);
            const bytes = new Uint8Array(binary.length);
            for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
            event.data[field] = bytes.buffer;
          });
        }
      };
    },
    { base64: Buffer.from(replacement).toString('base64'), messageType, field },
  );
}

for (const caseId of ['dropped-word', 'changed-digit-in-insertion', 'broken-comment-anchor', 'not-a-word-file']) {
  const brokenCase = BROKEN_CASES.find((c) => c.id === caseId)!;
  test(`a broken tracked-changes file is withheld: ${caseId}`, async ({ page }) => {
    const pair = await buildPair(pairById(brokenCase.pairId));
    await initialize();
    await tamper(page, 'compared', 'trackedChanges', await breakOutput(await produceTrackedChanges(pair.original, pair.modified), brokenCase));
    await compare(page, pair.original, pair.modified, 'withheld');

    for (const check of CHECKS_FOR['tracked-changes']) {
      const item = page.locator(`#report-tracked li[data-check="${check}"]`);
      const shouldFail = expectedFailures(brokenCase).includes(check);
      await expect(item).toHaveAttribute('data-passed', String(!shouldFail));
      if (shouldFail) await expect(item.locator('.failure')).not.toBeEmpty();
    }
    // Not offered: no visible link, and no address behind it.
    await expect(page.locator('#download-tracked')).toBeHidden();
    expect(await page.locator('#download-tracked').getAttribute('href')).toBeNull();
    await expect(page.locator('#withheld-tracked')).toBeVisible();
    await expect(page.locator('#verdict-tracked')).toHaveText('(failed: withheld)');
    const report = await downloadFrom(page, '#download-report');
    expect(new TextDecoder().decode(report.bytes)).toContain('FAILED - this output was withheld');
  });
}

for (const caseId of ['static-dropped-word', 'static-wrong-cover-count', 'static-leftover-tracked-change', 'static-format-change-not-shaded', 'static-old-number-wrong']) {
  const brokenCase = BROKEN_STATIC_CASES.find((c) => c.id === caseId)!;
  test(`a broken static redline is withheld and not shown: ${caseId}`, async ({ page }) => {
    const pair = await buildPair(pairById(brokenCase.pairId));
    await initialize();
    // The page prints its own clock time on the cover page; freeze the clock so the damaged file's cover matches.
    await page.clock.setFixedTime(new Date('2026-11-23T14:05:00'));
    const comparedAt = await page.evaluate(() => 0).then(() => '23 November 2026 at 14:05');
    const offsetMinutes = -new Date('2026-11-23T14:05:00').getTimezoneOffset();
    const two = (n: number): string => String(n).padStart(2, '0');
    const zone = `UTC${offsetMinutes < 0 ? '-' : '+'}${two(Math.floor(Math.abs(offsetMinutes) / 60))}:${two(Math.abs(offsetMinutes) % 60)}`;
    const cover = { originalName: 'original.docx', modifiedName: 'modified.docx', comparedAt: `${comparedAt} (${zone})` };
    const genuine = (await makeStaticRedline(await produceTrackedChanges(pair.original, pair.modified), pair.original, cover)).docx;
    const files = await unzip(genuine);
    brokenCase.damage(files);
    await tamper(page, 'static-made', 'staticRedline', await zip(files));
    await compare(page, pair.original, pair.modified, 'withheld');

    for (const check of CHECKS_FOR['static-redline']) {
      const shouldFail = expectedStaticFailures(brokenCase).includes(check);
      await expect(page.locator(`#report-static li[data-check="${check}"]`)).toHaveAttribute('data-passed', String(!shouldFail));
    }
    await expect(page.locator('#download-static')).toBeHidden();
    expect(await page.locator('#download-static').getAttribute('href')).toBeNull();
    await expect(page.locator('#withheld-static')).toBeVisible();
    // The viewer does not show a redline that failed, and no unverified counts are shown.
    await expect(page.locator('#viewer')).toBeHidden();
    expect(await page.locator('#viewer').getAttribute('srcdoc')).toBeNull();
    await expect(page.locator('#viewer-message')).toContainText('did not pass its checks, so it is not shown');
    await expect(page.locator('#summary')).toBeHidden();
    // The tracked-changes file was fine and is still offered.
    await expect(page.locator('#download-tracked')).toBeVisible();
  });
}

test('a withheld output shows the first paragraph where it went wrong', async ({ page }) => {
  const brokenCase = BROKEN_CASES.find((c) => c.id === 'dropped-word')!;
  const pair = await buildPair(pairById(brokenCase.pairId));
  await initialize();
  await tamper(page, 'compared', 'trackedChanges', await breakOutput(await produceTrackedChanges(pair.original, pair.modified), brokenCase));
  await compare(page, pair.original, pair.modified, 'withheld');
  const failure = page.locator('#report-tracked li[data-check="accept-all"] .failure');
  await expect(failure).toContainText('Where: main text, paragraph 2');
  await expect(failure).toContainText(/Expected: .*provide the Services with exceptional skill and care\./);
  await expect(failure).toContainText(/Found: .*provide the Services with exceptional and care\./);
});
