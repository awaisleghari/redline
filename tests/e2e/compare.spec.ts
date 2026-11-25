/**
 * Browser tests of the workspace (/app): add two files, create the redline, see
 * it in the viewer, read the checks, download each format; and the refusals and
 * withheld outputs. The landing page and the docs are in pages.spec.ts.
 */
import { expect, test, type Page } from '@playwright/test';
import { initialize } from 'docxodus/core';
import { readFile } from 'node:fs/promises';
import { checkPdf, checkStaticRedline, checkTrackedChanges, CHECKS_FOR } from '../../src/checker';
import { unzip, zip } from '../../src/common/zip';
import { makeStaticRedline } from '../../src/producer/make-static-redline';
import { produceTrackedChanges } from '../../src/producer/produce';
import { BROKEN_CASES, breakOutput, expectedFailures } from '../fixtures/broken';
import { BROKEN_PDF_CASES } from '../fixtures/broken-pdf';
import { BROKEN_STATIC_CASES, expectedStaticFailures } from '../fixtures/broken-static';
import { convertToPdf } from '../pdf/container';
import { buildLargeContract } from '../fixtures/large';
import { ALL_PAIR_SPECS, PAIR_SPECS, buildDocx, buildPair, pairById } from '../fixtures/pairs';
import { fullExpectation, readTrackedChanges } from '../helpers/tracked-changes';

const DOCX_MIME = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';

type Format = 'pdf' | 'static' | 'tracked';

async function chooseFiles(page: Page, original: Uint8Array | Buffer, modified: Uint8Array | Buffer, names = ['original.docx', 'modified.docx']): Promise<void> {
  await page.locator('#original').setInputFiles({ name: names[0]!, mimeType: DOCX_MIME, buffer: Buffer.from(original) });
  await expect(page.locator('#file-original')).toBeVisible();
  await page.locator('#modified').setInputFiles({ name: names[1]!, mimeType: DOCX_MIME, buffer: Buffer.from(modified) });
  await expect(page.locator('#file-modified')).toBeVisible();
}

/** Choose a format in the export menu. */
async function pickFormat(page: Page, format: Format): Promise<void> {
  await page.locator('#download-more').click();
  await page.locator(`#download-menu [data-format="${format}"]`).click();
  await expect(page.locator('#download-menu')).toBeHidden();
  await expect(page.locator('#download')).toHaveAttribute('data-format', format);
}

/** Download a format through the export button, as a user would. */
async function download(page: Page, format: Format): Promise<{ name: string; bytes: Uint8Array }> {
  if ((await page.locator('#download').getAttribute('data-format')) !== format) await pickFormat(page, format);
  const downloadPromise = page.waitForEvent('download', { timeout: 90_000 });
  await page.locator('#download').click();
  const file = await downloadPromise;
  return { name: file.suggestedFilename(), bytes: new Uint8Array(await readFile(await file.path())) };
}

async function downloadReport(page: Page): Promise<string> {
  if ((await page.locator('#checks-toggle').getAttribute('aria-expanded')) !== 'true') await page.locator('#checks-toggle').click();
  const downloadPromise = page.waitForEvent('download');
  await page.locator('#download-report').click();
  const file = await downloadPromise;
  expect(file.suggestedFilename()).toBe('redline-check-report.txt');
  return new TextDecoder().decode(await readFile(await file.path()));
}

async function compare(page: Page, original: Uint8Array, modified: Uint8Array, expectedState = 'done'): Promise<void> {
  await page.goto('/app');
  await chooseFiles(page, original, modified);
  await page.getByRole('button', { name: 'create redline' }).click();
  await expect(page.locator('#status')).toHaveAttribute('data-state', expectedState, { timeout: 60_000 });
}

/** The page must never log an error or throw. */
function watchForErrors(page: Page): string[] {
  const errors: string[] = [];
  page.on('console', (message) => {
    if (message.type() === 'error') errors.push(message.text());
  });
  page.on('pageerror', (error) => errors.push(error.message));
  return errors;
}

test('the workspace opens empty, with create disabled until both files are added, and works from the keyboard', async ({ page }) => {
  const errors = watchForErrors(page);
  await page.goto('/app');
  await expect(page.getByRole('heading', { name: 'create a redline' })).toBeVisible();
  await expect(page.locator('#viewer-title')).toHaveText('new redline');
  await expect(page.getByText('.docx only, up to 25 MB each')).toBeVisible();
  await expect(page.getByText('Changes, downloads and checks appear here once the redline is made.')).toBeVisible();
  // Two empty slots of the same size.
  await expect(page.getByRole('button', { name: /original\s*Drop the earlier version, or browse/ })).toBeVisible();
  await expect(page.getByRole('button', { name: /modified\s*Drop the later version, or browse/ })).toBeVisible();
  const button = page.locator('#create');
  await expect(button).toBeDisabled();
  const pair = await buildPair(PAIR_SPECS[0]!);
  await page.locator('#original').setInputFiles({ name: 'Services Agreement v1.docx', mimeType: DOCX_MIME, buffer: Buffer.from(pair.original) });
  await expect(page.locator('#file-original')).toBeVisible();
  await expect(button).toBeDisabled();
  await page.locator('#modified').setInputFiles({ name: 'A very long name for the second version of the Services Agreement, as marked up by the other side.docx', mimeType: DOCX_MIME, buffer: Buffer.from(pair.modified) });
  await expect(button).toBeEnabled();
  // The name is shown without its extension; the full name is there on hover. Both slots stay the same size.
  await expect(page.locator('#name-original')).toHaveText('Services Agreement v1');
  await expect(page.locator('#name-original')).toHaveAttribute('title', 'Services Agreement v1.docx');
  await expect(page.locator('#file-original .rl-filemeta')).toHaveText(/^original · docx · \d+ KB$/);
  const [first, second] = [await page.locator('#file-original').boundingBox(), await page.locator('#file-modified').boundingBox()];
  expect(first!.height).toBe(64);
  expect([second!.width, second!.height]).toEqual([first!.width, first!.height]);
  // The footer links to the source and says where the comparison happens.
  await expect(page.getByRole('link', { name: 'Source code' })).toHaveAttribute('href', /github\.com\/awaisleghari\/redline/);
  await expect(page.locator('.rl-foot')).toContainText('Compared in your browser. Nothing is uploaded, except to make a PDF, which is deleted at once. No AI is used.');
  // Keyboard: Ctrl+Enter (⌘+Enter on a Mac) creates the redline.
  await expect(page.locator('#create-hint')).toHaveText(/^(Ctrl|⌘) ↵$/);
  await page.keyboard.press((await page.locator('#create-hint').textContent())!.startsWith('⌘') ? 'Meta+Enter' : 'Control+Enter');
  await expect(page.locator('#status')).toHaveAttribute('data-state', 'done', { timeout: 60_000 });
  await expect(page.getByRole('button', { name: 'create redline again' })).toBeVisible();
  await expect(page.locator('#viewer-title')).toHaveText('A very long name for the second version of the Services Agreement, as marked up by the other side');
  // The + button starts again from nothing.
  await page.getByRole('button', { name: 'Start a new redline' }).click();
  await expect(page.getByRole('heading', { name: 'create a redline' })).toBeVisible();
  await expect(page.locator('#pick-original')).toBeVisible();
  await expect(page.locator('#pick-modified')).toBeVisible();
  await expect(page.locator('#doc')).toBeEmpty();
  expect(errors).toEqual([]);
});

test('the creating screen is shown while the redline is made', async ({ page }) => {
  const pair = await buildPair(PAIR_SPECS[0]!);
  await page.goto('/app');
  await chooseFiles(page, pair.original, pair.modified);
  await page.locator('#create').click();
  await expect(page.getByRole('status').filter({ hasText: 'Comparing two documents in your browser' })).toBeVisible();
  await expect(page.getByRole('heading', { name: 'creating redline' })).toBeVisible();
  await expect(page.locator('#create-label')).toHaveText('creating redline');
  await expect(page.locator('#create-hint')).toBeHidden();
  await expect(page.locator('#status')).toHaveAttribute('data-state', 'done', { timeout: 60_000 });
  await expect(page.locator('#loading')).toBeHidden();
});

for (const spec of ALL_PAIR_SPECS) {
  test(`${spec.id}: both Word files pass every check, show in the viewer and download`, async ({ page }) => {
    const errors = watchForErrors(page);
    const pair = await buildPair(spec);
    await compare(page, pair.original, pair.modified);

    // The checks: every check of both outputs, all passed.
    for (const [kind, output] of [['tracked', 'tracked-changes'], ['static', 'static-redline']] as const) {
      await expect(page.locator(`#report-${kind} li`)).toHaveCount(CHECKS_FOR[output].length);
      await expect(page.locator(`#report-${kind} li[data-passed="true"]`)).toHaveCount(CHECKS_FOR[output].length);
    }
    const total = CHECKS_FOR['tracked-changes'].length + CHECKS_FOR['static-redline'].length;
    await expect(page.locator('#checks-summary')).toHaveText(`${total} of ${total} checks passed`);

    // The viewer shows the static redline, cover page first.
    const viewer = page.locator('#doc');
    await expect(viewer.getByText('Comparison of documents')).toBeVisible();
    await expect(viewer.getByText('Original: original.docx')).toBeVisible();

    // The sidebar shows the counts, which are the ones on the cover page; the change counter steps through as many.
    if (spec.staticCounts) {
      let sum = 0;
      for (const [key, value] of Object.entries(spec.staticCounts)) {
        sum += value;
        await expect(page.locator(`#summary [data-count="${key}"]`)).toHaveText(String(value));
      }
      await expect(page.locator('#changes-total')).toHaveText(String(sum));
      await expect(viewer.getByText(`Insertions: ${spec.staticCounts.insertions}`)).toBeVisible();
      await expect(page.locator('#change-counter')).toHaveText(`${sum ? 1 : 0} / ${sum}`);
    }

    // Word, tracked changes: exactly the expected changes, each kind under its own author.
    const tracked = await download(page, 'tracked');
    expect(tracked.name).toBe('redline-tracked-changes.docx');
    expect((await readTrackedChanges(tracked.bytes)).parts).toEqual(fullExpectation(spec.expected));
    // Word, static redline: passes the Checker again here, outside the browser.
    const staticRedline = await download(page, 'static');
    expect(staticRedline.name).toBe('redline-static.docx');
    const coverTime = (await viewer.getByText(/^Compared: /).textContent())!.replace('Compared: ', '');
    const recheck = await checkStaticRedline({ original: pair.original, modified: pair.modified, output: staticRedline.bytes, originalName: 'original.docx', modifiedName: 'modified.docx', comparedAt: coverTime });
    expect(recheck.checks.filter((c) => !c.passed)).toEqual([]);
    expect((await checkTrackedChanges({ original: pair.original, modified: pair.modified, output: tracked.bytes })).passed).toBe(true);
    expect(errors).toEqual([]);
  });
}

test('each kind of tracked change has its own author, so Word colours them differently', async ({ page }) => {
  const pair = await buildPair(pairById('moved-paragraph'));
  await compare(page, pair.original, pair.modified);
  const tracked = await download(page, 'tracked');
  const xml = new TextDecoder().decode((await unzip(tracked.bytes)).get('word/document.xml')!);
  expect(xml).toMatch(/<w:moveFrom [^>]*w:author="Redline \(moved\)"/);
  expect(xml).toMatch(/<w:moveTo [^>]*w:author="Redline \(moved\)"/);
  const word = await buildPair(pairById('one-word-change'));
  await compare(page, word.original, word.modified);
  const xml2 = new TextDecoder().decode((await unzip((await download(page, 'tracked')).bytes)).get('word/document.xml')!);
  expect(xml2).toMatch(/<w:ins [^>]*w:author="Redline \(inserted\)"/);
  expect(xml2).toMatch(/<w:del [^>]*w:author="Redline \(deleted\)"/);
});

test('a link inside a document cannot be followed from the viewer, and the page requests nothing from outside', async ({ page, baseURL }) => {
  const outside: string[] = [];
  const site = new URL(baseURL!).origin + '/';
  page.on('request', (request) => {
    if (!request.url().startsWith(site) && !request.url().startsWith('blob:') && !request.url().startsWith('data:') && !request.url().startsWith('about:')) outside.push(request.url());
  });
  const pair = await buildPair(pairById('hyperlink-address-change'));
  await compare(page, pair.original, pair.modified);
  const viewer = page.locator('#doc');
  await expect(viewer.getByText('our website').first()).toBeVisible();
  // The link's words are shown, but there is nowhere for a click to go.
  expect(await viewer.locator('a[href]').count()).toBe(0);
  await viewer.getByText('our website').first().click({ force: true });
  await expect(viewer.getByText('Comparison of documents')).toBeVisible();
  expect(page.url()).toMatch(/\/app$/);
  expect(outside).toEqual([]);
});

test('the viewer keeps only the words and structure of what the engine hands it: no scripts, links, frames or outside pictures', async ({ page, baseURL }) => {
  const outside: string[] = [];
  const site = new URL(baseURL!).origin + '/';
  page.on('request', (request) => {
    if (!request.url().startsWith(site) && !request.url().startsWith('blob:') && !request.url().startsWith('data:')) outside.push(request.url());
  });
  const hostile =
    '<html xmlns="http://www.w3.org/1999/xhtml"><head><style>p.x { color: red; background: url(https://outside.example/a.png); position: fixed; font-weight: bold }</style>' +
    '<script>window.hacked = 1</script></head><body>' +
    '<p class="x" onclick="window.hacked = 2" style="background-image: url(https://outside.example/b.png)">Clause one <a href="https://outside.example/" onmouseover="window.hacked = 3">a link</a></p>' +
    '<img src="https://outside.example/c.png" onerror="window.hacked = 4" /><img src="x" onerror="window.hacked = 5" />' +
    '<iframe src="https://outside.example/"></iframe><object data="https://outside.example/d"></object><form action="https://outside.example/"><input name="q" /></form>' +
    '<svg xmlns="http://www.w3.org/2000/svg" onload="window.hacked = 6"><script>window.hacked = 7</script></svg><p>Clause two</p></body></html>';
  await tamper(page, 'rendered', 'html', hostile);
  const pair = await buildPair(PAIR_SPECS[0]!);
  await compare(page, pair.original, pair.modified);
  const viewer = page.locator('#doc');
  await expect(viewer.getByText('Clause one a link')).toBeVisible();
  await expect(viewer.getByText('Clause two')).toBeVisible();
  await viewer.getByText('a link').click({ force: true });
  await viewer.getByText('a link').hover({ force: true });
  expect(await viewer.locator('script, iframe, object, form, input, svg, img, a[href], [onclick], [onerror], [onmouseover], [onload]').count()).toBe(0);
  expect(await viewer.evaluate((root) => Array.from(root.querySelectorAll('*')).some((el) => Array.from(el.attributes).some((a) => /^on|^href$|^src$|^action$/.test(a.name) || /url\(/.test(a.value))))).toBe(false);
  expect(await page.evaluate(() => document.getElementById('redline-document-style')!.textContent)).not.toMatch(/url\(|fixed|outside/);
  expect(await page.evaluate(() => (window as unknown as { hacked?: number }).hacked)).toBeUndefined();
  expect(page.url()).toMatch(/\/app$/);
  expect(outside).toEqual([]);
});

test('showing the redline in the page never changes its words', async ({ page }) => {
  // Words that look like link attributes must survive untouched in the viewer.
  const sentence = 'The corrective action = termination; set href="none" and ping=5 within thirty days.';
  const block = (text: string) => ({ blocks: [{ kind: 'p' as const, text: 'Schedule of remedies.' }, { kind: 'p' as const, text }] });
  const original = await buildDocx(block(sentence));
  const modified = await buildDocx(block(sentence.replace('thirty', 'sixty')));
  await compare(page, original, modified);
  const viewer = page.locator('#doc');
  await expect(viewer.getByText('The corrective action = termination; set href="none" and ping=5 within')).toBeVisible();
  await expect(viewer.locator('.rl-ins', { hasText: 'sixty' })).toBeVisible();
  await expect(viewer.locator('.rl-del', { hasText: 'thirty' })).toBeVisible();
});

test('a deletion and a move of underlined text are marked and counted like any other', async ({ page }) => {
  const under = (text: string) => ({ kind: 'rich' as const, text, runs: [{ text, underline: true }] });
  const moved = under('Each party shall keep the Confidential Information of the other party secret.');
  const filler = ['The Supplier shall provide the Services with reasonable skill and care.', 'The Customer shall pay the Fees within thirty days of the date of each invoice.', 'This Agreement is governed by the laws of England and Wales.'].map((text) => ({ kind: 'p' as const, text }));
  const original = await buildDocx({ blocks: [under('Governing law and jurisdiction of the courts'), moved, ...filler] });
  const modified = await buildDocx({ blocks: [under('Governing law and jurisdiction'), ...filler, moved] });
  await compare(page, original, modified);
  await expect(page.locator('#doc').getByText('Comparison of documents')).toBeVisible();
  const total = Number(await page.locator('#changes-total').textContent());
  expect(total).toBeGreaterThanOrEqual(2);
  await expect(page.locator('#change-counter')).toHaveText(`1 / ${total}`);
  await expect(page.locator('#doc .rl-del[data-chg]', { hasText: 'of the courts' })).toHaveCSS('color', 'rgb(255, 142, 144)');
  // Both ends of the move are found, and they are one change.
  const from = page.locator('#doc .rl-mvf[data-chg]', { hasText: 'Each party shall keep' });
  const to = page.locator('#doc .rl-mvt[data-chg]', { hasText: 'Each party shall keep' });
  await expect(from).toHaveCSS('color', 'rgb(111, 221, 174)');
  await expect(to).toHaveCSS('color', 'rgb(111, 221, 174)');
  expect(await from.first().getAttribute('data-chg')).toBe(await to.first().getAttribute('data-chg'));
});

test('files cannot be changed while a comparison runs, and choosing a new file clears the old results', async ({ page }) => {
  const pair = await buildPair(PAIR_SPECS[0]!);
  await page.goto('/app');
  await chooseFiles(page, pair.original, pair.modified);
  await page.locator('#create').click();
  // While it runs: both file inputs, the remove buttons and the create button are disabled.
  await expect(page.locator('#original')).toBeDisabled();
  await expect(page.locator('#modified')).toBeDisabled();
  await expect(page.locator('#remove-original')).toBeDisabled();
  await expect(page.locator('#create')).toBeDisabled();
  await expect(page.locator('#status')).toHaveAttribute('data-state', 'done', { timeout: 60_000 });
  await expect(page.locator('#viewer')).toBeVisible();
  await expect(page.locator('#section-export')).toBeVisible();
  // Choosing a different file afterwards: the previous results, links and status go.
  await page.locator('#original').setInputFiles({ name: 'another.docx', mimeType: DOCX_MIME, buffer: Buffer.from(pair.modified) });
  await expect(page.locator('#viewer')).toBeHidden();
  await expect(page.getByRole('heading', { name: 'create a redline' })).toBeVisible();
  await expect(page.locator('#status')).toHaveText('');
  await expect(page.locator('#section-export')).toBeHidden();
  await expect(page.locator('#section-checks')).toBeHidden();
  await expect(page.locator('#doc')).toBeEmpty();
  expect(await page.locator('#download-report').getAttribute('href')).toBeNull();
  await expect(page.locator('#name-original')).toHaveText('another');
});

test('the workspace fits a narrow phone screen without scrolling sideways, before and after a redline', async ({ page }) => {
  await page.setViewportSize({ width: 360, height: 800 });
  await page.goto('/app');
  expect(await page.evaluate(() => document.documentElement.scrollWidth <= window.innerWidth)).toBe(true);
  const pair = await buildPair(pairById('several-changes'));
  await compare(page, pair.original, pair.modified);
  expect(await page.evaluate(() => document.documentElement.scrollWidth <= window.innerWidth)).toBe(true);
  // The legend gives way; every control is still on the screen.
  await expect(page.locator('.rl-legend')).toBeHidden();
  for (const id of ['view-pages', 'page-light', 'prev-change', 'next-change']) {
    const box = (await page.locator(`#${id}`).boundingBox())!;
    expect(box.x).toBeGreaterThanOrEqual(0);
    expect(box.x + box.width).toBeLessThanOrEqual(360);
  }
  // The arrows still find the change: the whole page scrolls to it.
  await page.locator('#next-change').click();
  await expect(page.locator('#change-counter')).toHaveText(/^2 \/ \d+$/);
  await expect(page.locator('#doc .rl-cur').first()).toBeInViewport();
});

test('the checks read in plain English, open and close, and can be downloaded as a text file', async ({ page }) => {
  const pair = await buildPair(PAIR_SPECS[0]!);
  await compare(page, pair.original, pair.modified);
  const toggle = page.getByRole('button', { name: '15 of 15 checks passed' });
  await expect(toggle).toHaveAttribute('aria-expanded', 'false');
  await expect(page.locator('#rl-checklist')).toBeHidden();
  await toggle.click();
  await expect(toggle).toHaveAttribute('aria-expanded', 'true');
  await expect(page.locator('#report-tracked li[data-check="accept-all"]')).toHaveText('Accepting every change reproduces the modified document');
  await expect(page.locator('#report-tracked li[data-check="accept-all"]').getByRole('img', { name: 'Passed' })).toBeVisible();
  await expect(page.locator('#report-static li[data-check="cover-page"]')).toContainText('its counts match the redline');
  await expect(page.locator('#group-static')).toContainText('static redline · pdf');
  await expect(page.locator('#verdict-static')).toHaveText('all passed');
  await expect(page.locator('#verdict-tracked')).toHaveText('all passed');

  const text = await downloadReport(page);
  expect(text).toContain('Original: original.docx');
  expect(text).toContain('Output: Word file with the static redline');
  expect(text).toContain('Output: Word file with tracked changes');
  expect(text).toContain('Overall: all checks passed');
  expect(text).toContain('Rejecting every change reproduces the original document: passed');
  expect(text).toContain('Counted independently: 1 insertions, 1 deletions, 0 moves, 0 formatting changes.');
  await toggle.click();
  await expect(page.locator('#rl-checklist')).toBeHidden();
});

test.describe('reading the redline', () => {
  test('the arrows step through the changes, wrap around, ring the current one and bring it into view', async ({ page }) => {
    const pair = await buildPair(pairById('several-changes'));
    await compare(page, pair.original, pair.modified);
    const counter = page.locator('#change-counter');
    const total = Number((await page.locator('#changes-total').textContent())!);
    expect(total).toBe(16);
    await expect(counter).toHaveText(`1 / ${total}`);
    await expect(page.locator('#doc [data-chg="1"]')).toHaveClass(/rl-cur/);
    await page.getByRole('button', { name: 'Next change' }).click();
    await expect(counter).toHaveText(`2 / ${total}`);
    await expect(page.locator('#doc .rl-cur')).toHaveCount(1);
    await expect(page.locator('#doc [data-chg="2"]')).toHaveClass(/rl-cur/);
    await page.getByRole('button', { name: 'Previous change' }).click();
    await page.getByRole('button', { name: 'Previous change' }).click();
    await expect(counter).toHaveText(`${total} / ${total}`);
    await expect(page.locator(`#doc [data-chg="${total}"]`)).toBeInViewport();
    await page.getByRole('button', { name: 'Next change' }).click();
    await expect(counter).toHaveText(`1 / ${total}`);
    // The marks carry the design's classes, and the legend on the cover page is not a change.
    await expect(page.locator('#doc .rl-ins[data-chg]').first()).toBeVisible();
    await expect(page.locator('#doc .rl-del[data-chg]').first()).toBeVisible();
    expect(await page.locator('#doc [data-legend][data-chg]').count()).toBe(0);
  });

  test('a move is one change, and both of its ends are ringed', async ({ page }) => {
    const pair = await buildPair(pairById('moved-paragraph'));
    await compare(page, pair.original, pair.modified);
    await expect(page.locator('#change-counter')).toHaveText('1 / 1');
    await expect(page.locator('#doc .rl-mvf.rl-cur')).toHaveCount(1);
    await expect(page.locator('#doc .rl-mvt.rl-cur')).toHaveCount(1);
    await expect(page.locator('#doc .rl-mvf:not([data-legend])')).toHaveAttribute('data-chg', '1');
    await expect(page.locator('#doc .rl-mvt:not([data-legend])')).toHaveAttribute('data-chg', '1');
  });

  test('the counter follows scrolling by hand, without a ring, and the next arrow carries on from there', async ({ page }) => {
    const contract = await buildLargeContract();
    await page.goto('/app');
    await chooseFiles(page, contract.original, contract.modified);
    await page.locator('#create').click();
    await expect(page.locator('#status')).toHaveAttribute('data-state', 'done', { timeout: 90_000 });
    const counter = page.locator('#change-counter');
    await expect(counter).toHaveText(/^1 \/ \d+$/);
    // The viewer takes the keyboard: it can be focused and scrolled with the keys.
    await expect(page.locator('#viewer')).toHaveAttribute('aria-label', 'Redline of modified');
    await page.locator('#viewer').focus();
    await expect(page.locator('#viewer')).toBeFocused();
    await page.keyboard.press('PageDown');
    await expect.poll(() => page.locator('#viewer').evaluate((viewer) => viewer.scrollTop)).toBeGreaterThan(100);
    // Scroll a long way down by hand.
    await page.locator('#viewer').evaluate((viewer) => viewer.scrollTo(0, viewer.scrollHeight / 2));
    await expect(counter).not.toHaveText(/^1 \//);
    await expect(page.locator('#doc .rl-cur')).toHaveCount(0);
    const here = Number((await counter.textContent())!.split(' / ')[0]);
    expect(here).toBeGreaterThan(10);
    // It is the change nearest the top third of the view (worked out here independently).
    const nearest = await page.evaluate(() => {
      const viewer = document.getElementById('viewer')!;
      const line = viewer.getBoundingClientRect().top + viewer.clientHeight / 3;
      const firstOf = new Map<number, number>();
      for (const mark of Array.from(document.querySelectorAll<HTMLElement>('#doc [data-chg]'))) {
        const number = Number(mark.dataset.chg);
        if (!firstOf.has(number)) firstOf.set(number, Math.abs(mark.getBoundingClientRect().top - line));
      }
      return [...firstOf.entries()].sort((a, b) => a[1] - b[1])[0]![0];
    });
    expect(here).toBe(nearest);
    await page.getByRole('button', { name: 'Next change' }).click();
    await expect(counter).toHaveText(new RegExp(`^${here + 1} /`));
    await expect(page.locator(`#doc [data-chg="${here + 1}"]`).first()).toHaveClass(/rl-cur/);
    await expect(page.locator(`#doc [data-chg="${here + 1}"]`).first()).toBeInViewport();
  });

  test('the light page and the layout are remembered after a reload; the light page is soft paper and the app stays dark', async ({ page }) => {
    const pair = await buildPair(pairById('several-changes'));
    await compare(page, pair.original, pair.modified);
    await expect(page.getByRole('button', { name: 'Dark page' })).toHaveAttribute('aria-pressed', 'true');
    await expect(page.getByRole('button', { name: 'continuous' })).toHaveAttribute('aria-pressed', 'true');
    const sheet = page.locator('#doc .rl-sheet').first();
    await expect(sheet).toHaveCSS('background-color', 'rgb(22, 23, 27)');
    await page.getByRole('button', { name: 'Light page' }).click();
    await expect(sheet).toHaveCSS('background-color', 'rgb(230, 228, 220)');
    await expect(page.locator('#doc .rl-ins').first()).toHaveCSS('color', 'rgb(35, 70, 181)');
    await expect(page.locator('.rl-main')).toHaveCSS('background-color', 'rgb(17, 18, 21)');
    // The legend in the bar stays in its dark-mode colours.
    await expect(page.locator('.rl-legend .rl-ins')).toHaveCSS('color', 'rgb(169, 193, 255)');
    await page.getByRole('button', { name: 'pages' }).click();
    await pickFormat(page, 'static');
    await expect(page.locator('#download')).toHaveText('download word · static');

    await page.reload();
    await chooseFiles(page, pair.original, pair.modified);
    await page.locator('#create').click();
    await expect(page.locator('#status')).toHaveAttribute('data-state', 'done', { timeout: 60_000 });
    await expect(page.getByRole('button', { name: 'Light page' })).toHaveAttribute('aria-pressed', 'true');
    await expect(page.getByRole('button', { name: 'pages' })).toHaveAttribute('aria-pressed', 'true');
    await expect(page.locator('#doc .rl-sheet').first()).toHaveCSS('background-color', 'rgb(230, 228, 220)');
    await expect(page.locator('#download')).toHaveText('download word · static');
  });

  test('the export menu opens and closes from the keyboard, and shows which format is chosen', async ({ page }) => {
    const pair = await buildPair(PAIR_SPECS[0]!);
    await compare(page, pair.original, pair.modified);
    const more = page.getByRole('button', { name: 'More download formats' });
    const menu = page.getByRole('menu', { name: 'Download formats' });
    await expect(page.locator('#download')).toHaveText('download pdf');
    await expect(more).toHaveAttribute('aria-expanded', 'false');
    await more.focus();
    await page.keyboard.press('Enter');
    await expect(menu).toBeVisible();
    await expect(more).toHaveAttribute('aria-expanded', 'true');
    await expect(menu.getByRole('menuitemradio')).toHaveText([/pdf\s*Static redline, ready to share/, /word · static redline\s*Marked up, no tracked changes/, /word · tracked changes\s*Accept or reject each change in Word/]);
    await expect(menu).toContainText('A format is offered only when every check on it has passed. Making a PDF sends the static redline to our server, where it is deleted straight after conversion.');
    await expect(menu.getByRole('menuitemradio', { name: /^PDF|^pdf/ }).first()).toBeFocused();
    await page.keyboard.press('ArrowDown');
    await page.keyboard.press('ArrowDown');
    await expect(menu.getByRole('menuitemradio', { name: /word · tracked changes/ })).toBeFocused();
    await page.keyboard.press('Escape');
    await expect(menu).toBeHidden();
    await expect(more).toBeFocused();
    // An outside click closes it too.
    await more.click();
    await expect(menu).toBeVisible();
    await page.locator('#viewer-title').click();
    await expect(menu).toBeHidden();
    // Choosing a row makes it the button's format.
    await more.click();
    await menu.getByRole('menuitemradio', { name: /word · tracked changes/ }).click();
    await expect(menu).toBeHidden();
    await expect(page.locator('#download')).toHaveText('download word · tracked');
    await more.click();
    await expect(menu.getByRole('menuitemradio', { name: /word · tracked changes/ })).toHaveAttribute('aria-checked', 'true');
    await expect(menu.getByRole('menuitemradio', { name: /word · static redline/ })).toHaveAttribute('aria-checked', 'false');
  });
});

test.describe('the PDF', () => {
  test('nothing is sent to the server until the PDF is asked for; then the PDF is checked, downloaded, and sets the pages', async ({ page }) => {
    const sent: string[] = [];
    page.on('request', (request) => {
      if (request.method() !== 'GET') sent.push(`${request.method()} ${new URL(request.url()).pathname}`);
    });
    await page.goto('/app');
    // Before a comparison there is nothing to download.
    await expect(page.locator('#section-export')).toBeHidden();
    await expect(page.locator('#download')).toBeDisabled();

    const pair = await buildPair(pairById('several-changes'));
    await compare(page, pair.original, pair.modified);
    await expect(page.locator('#download')).toBeEnabled();
    await expect(page.locator('#download')).toHaveText('download pdf');
    // The notice is in view, whole, before the button is pressed, even in a short window.
    await page.setViewportSize({ width: 1366, height: 650 });
    await expect(page.locator('.rl-foot')).toBeInViewport({ ratio: 1 });
    await expect(page.locator('.rl-foot').getByText('Compared in your browser. Nothing is uploaded, except to make a PDF, which is deleted at once. No AI is used.')).toBeVisible();
    await expect(page.locator('#download')).toBeInViewport({ ratio: 1 });
    // The fuller sentence is in the menu the button opens.
    await page.locator('#download-more').click();
    await expect(page.locator('#pdf-note')).toBeVisible();
    await expect(page.locator('#pdf-note')).toContainText('Making a PDF sends the static redline to our server, where it is deleted straight after conversion.');
    await page.keyboard.press('Escape');
    await page.setViewportSize({ width: 1280, height: 720 });
    // Choosing "pages" sends nothing either: it asks first.
    await page.getByRole('button', { name: 'pages' }).click();
    await expect(page.locator('#pages-notice')).toContainText('Making a PDF sends the static redline to our server, where it is deleted straight after conversion.');
    await expect(page.locator('#doc .rl-sheet')).toHaveCount(1);
    // Both Word files: still nothing has left the browser.
    await download(page, 'static');
    await download(page, 'tracked');
    expect(sent).toEqual([]);

    const pdf = await download(page, 'pdf');
    expect(pdf.name).toBe('redline.pdf');
    expect(sent).toEqual(['POST /api/pdf']);
    await expect(page.locator('#pdf-status')).toHaveText('The PDF passed its checks.');
    await expect(page.locator('#report-pdf li[data-check="pdf-text"]')).toHaveAttribute('data-passed', 'true');
    await expect(page.locator('#report-pdf li[data-check="pdf-colours"]')).toHaveAttribute('data-passed', 'true');
    await expect(page.locator('#verdict-pdf')).toHaveText('all passed');
    await expect(page.locator('#checks-summary')).toHaveText('17 of 17 checks passed');

    // The pages view now breaks where the PDF does: the cover page, then the document.
    await expect(page.locator('#pages-notice')).toBeHidden();
    await expect(page.locator('#doc .rl-sheet')).toHaveCount(2);
    await expect(page.locator('#doc .rl-sheet').nth(0)).toContainText('Comparison of documents');
    await expect(page.locator('#doc .rl-sheet').nth(0)).toContainText('page 1 of 2');
    await expect(page.locator('#doc .rl-sheet').nth(0)).not.toContainText('This Services Agreement is made between');
    await expect(page.locator('#doc .rl-sheet').nth(1)).toContainText('This Services Agreement is made between');
    await expect(page.locator('#doc .rl-sheet').nth(1)).toContainText('page 2 of 2');
    // Every change is still there to step through, and a short page keeps the paper's shape.
    await expect(page.locator('#change-counter')).toHaveText(/ \/ 16$/);
    const box = (await page.locator('#doc .rl-sheet').nth(0).boundingBox())!;
    expect(box.height / box.width).toBeGreaterThan(1.25);
    await page.getByRole('button', { name: 'continuous' }).click();
    await expect(page.locator('#doc .rl-sheet')).toHaveCount(1);
    await expect(page.locator('#doc')).not.toContainText('page 1 of 2');

    // A second press downloads the same PDF without sending anything again.
    const again = await download(page, 'pdf');
    expect(sent).toEqual(['POST /api/pdf']);
    expect(Buffer.from(again.bytes).equals(Buffer.from(pdf.bytes))).toBe(true);
    // The downloaded PDF passes check 9 again, outside the browser, against the downloaded static redline.
    const staticRedline = await download(page, 'static');
    expect((await checkPdf({ staticRedline: staticRedline.bytes, pdf: pdf.bytes })).passed).toBe(true);
    // And the check report now covers the PDF too.
    const report = await downloadReport(page);
    expect(report).toContain('Output: PDF of the static redline');
    expect(report).toContain("The PDF's text matches the static redline, complete and in order: passed");
  });

  test('the pages view can make the PDF itself, when asked, and downloads nothing by doing so', async ({ page }) => {
    const sent: string[] = [];
    let downloads = 0;
    page.on('request', (request) => {
      if (request.method() !== 'GET') sent.push(`${request.method()} ${new URL(request.url()).pathname}`);
    });
    page.on('download', () => downloads++);
    const pair = await buildPair(pairById('several-changes'));
    await compare(page, pair.original, pair.modified);
    await page.getByRole('button', { name: 'pages' }).click();
    expect(sent).toEqual([]);
    await page.getByRole('button', { name: 'make the pdf' }).click();
    await expect(page.locator('#doc .rl-sheet')).toHaveCount(2, { timeout: 60_000 });
    expect(sent).toEqual(['POST /api/pdf']);
    expect(downloads).toBe(0);
  });

  test('a table row that runs over several pages is split across them, with its columns kept and no page left blank', async ({ page }) => {
    // (The second column of the long row is left empty: the Checker at present withholds the PDF of a row with words in two columns that runs over a page.)
    const sentence = (n: number): string => `Item ${n}: the Supplier shall deliver the services described in schedule ${n} within ${n + 10} Business Days of the request. `;
    const long = (from: number): string => Array.from({ length: 70 }, (_, i) => sentence(from + i)).join('');
    const blocks = (fee: string) => ({ blocks: [{ kind: 'p' as const, text: 'Schedule of services.' }, { kind: 'table' as const, rows: [['Services', 'Notes'], [long(1) + long(500), ''], ['Monthly fee', fee]] }, { kind: 'p' as const, text: 'End of schedule.' }] });
    const original = await buildDocx(blocks('USD 1,500'));
    const modified = await buildDocx(blocks('USD 2,500'));
    await compare(page, original, modified);
    const whole = await page.locator('#doc').evaluate((doc) => doc.textContent!.replace(/\s+/g, ''));
    await page.getByRole('button', { name: 'pages' }).click();
    await page.getByRole('button', { name: 'make the pdf' }).click();
    const sheets = page.locator('#doc .rl-sheet');
    await expect(sheets.first()).toContainText('page 1 of', { timeout: 60_000 });
    const count = await sheets.count();
    expect(count).toBeGreaterThanOrEqual(4); // the cover page, then a row several pages long
    const pages = await sheets.evaluateAll((all) =>
      all.map((sheet) => {
        const copy = sheet.cloneNode(true) as HTMLElement;
        copy.querySelector('.rl-pagelabel')?.remove();
        return { text: copy.textContent!.replace(/\s+/g, ''), cells: Array.from(copy.querySelectorAll('tr')).map((row) => row.children.length) };
      }),
    );
    // Every page has words on it, every row on every page has both columns, and the pages together are the document.
    for (const one of pages) expect(one.text.length).toBeGreaterThan(100);
    for (const one of pages.slice(1)) expect(one.cells.length).toBeGreaterThan(0);
    expect(pages.flatMap((one) => one.cells).every((cells) => cells === 2)).toBe(true);
    expect(pages.map((one) => one.text).join('')).toBe(whole);
    await expect(sheets.last()).toContainText('End of schedule.');
    await expect(page.locator('#change-counter')).toHaveText('1 / 2');
  });

  test('a wrong PDF from the server is withheld: not downloaded, not offered, and no pages are drawn from it', async ({ page }) => {
    // The server answers with the PDF of a different document.
    const other = await buildPair(pairById('moved-paragraph'));
    await initialize();
    const cover = { originalName: 'original.docx', modifiedName: 'modified.docx', comparedAt: 'another time' };
    const wrongPdf = await convertToPdf((await makeStaticRedline(await produceTrackedChanges(other.original, other.modified), other.original, cover)).docx);
    await page.route('**/api/pdf', (route) => route.fulfill({ status: 200, contentType: 'application/pdf', body: Buffer.from(wrongPdf) }));
    let downloads = 0;
    page.on('download', () => downloads++);

    const pair = await buildPair(pairById('one-word-change'));
    await compare(page, pair.original, pair.modified);
    await page.locator('#download').click();
    await expect(page.locator('#pdf-status')).toContainText('did not pass its checks, so it has been withheld', { timeout: 60_000 });
    await expect(page.locator('#pdf-status')).toBeVisible();
    await expect(page.locator('#report-pdf li[data-check="pdf-text"]')).toHaveAttribute('data-passed', 'false');
    await expect(page.locator('#report-pdf li[data-check="pdf-text"] .failure')).not.toBeEmpty();
    await expect(page.locator('#verdict-pdf')).toHaveText(/failed$/);
    await expect(page.locator('#checks-summary')).toHaveText(/^1[56] of 17 checks passed$/);
    expect(downloads).toBe(0);
    // The PDF is no longer offered; the button falls back to the first format that is.
    await expect(page.locator('#download')).toHaveText('download word · static');
    await page.locator('#download-more').click();
    const item = page.locator('#download-menu [data-format="pdf"]');
    await expect(item).toBeDisabled();
    await expect(item).toContainText('Not offered: a check failed');
    await expect(item).toHaveCSS('opacity', '0.55');
    await page.keyboard.press('Escape');
    // No pages are drawn from a PDF that failed.
    await page.getByRole('button', { name: 'pages' }).click();
    await expect(page.locator('#pages-notice')).toContainText('Pages are not available for this redline');
    await expect(page.getByRole('button', { name: 'make the pdf' })).toBeHidden();
    await expect(page.locator('#doc .rl-sheet')).toHaveCount(1);
    // The Word files are unaffected.
    expect((await download(page, 'static')).name).toBe('redline-static.docx');
    expect(downloads).toBe(1);
  });

  test('a PDF with one colour wrong is withheld', async ({ page }) => {
    const brokenCase = BROKEN_PDF_CASES.find((c) => c.id === 'pdf-deletion-in-black')!;
    const pair = await buildPair(pairById(brokenCase.pairId));
    await initialize();
    await page.clock.setFixedTime(new Date('2026-11-23T14:05:00'));
    const offsetMinutes = -new Date('2026-11-23T14:05:00').getTimezoneOffset();
    const two = (n: number): string => String(n).padStart(2, '0');
    const zone = `UTC${offsetMinutes < 0 ? '-' : '+'}${two(Math.floor(Math.abs(offsetMinutes) / 60))}:${two(Math.abs(offsetMinutes) % 60)}`;
    const cover = { originalName: 'original.docx', modifiedName: 'modified.docx', comparedAt: `23 November 2026 at 14:05 (${zone})` };
    const files = await unzip((await makeStaticRedline(await produceTrackedChanges(pair.original, pair.modified), pair.original, cover)).docx);
    brokenCase.damage(files);
    const wrongPdf = await convertToPdf(await zip(files));
    await page.route('**/api/pdf', (route) => route.fulfill({ status: 200, contentType: 'application/pdf', body: Buffer.from(wrongPdf) }));
    let downloads = 0;
    page.on('download', () => downloads++);
    await compare(page, pair.original, pair.modified);
    await page.locator('#download').click();
    await expect(page.locator('#pdf-status')).toContainText('withheld', { timeout: 60_000 });
    await expect(page.locator('#report-pdf li[data-check="pdf-text"]')).toHaveAttribute('data-passed', 'true');
    await expect(page.locator('#report-pdf li[data-check="pdf-colours"]')).toHaveAttribute('data-passed', 'false');
    await expect(page.locator('#download-menu [data-format="pdf"]')).toBeDisabled();
    expect(downloads).toBe(0);
  });

  test('when the server cannot convert, the page says so plainly, offers no PDF, and lets the user try again', async ({ page }) => {
    await page.route('**/api/pdf', (route) => route.fulfill({ status: 422, contentType: 'text/plain', body: 'The document could not be converted to PDF.' }));
    let downloads = 0;
    page.on('download', () => downloads++);
    const pair = await buildPair(pairById('one-word-change'));
    await compare(page, pair.original, pair.modified);
    await page.locator('#download').click();
    await expect(page.locator('#pdf-status')).toHaveText('The PDF could not be made. (The document could not be converted to PDF.) You can try again.');
    await expect(page.locator('#pdf-status')).toBeVisible();
    // Nothing came back to check, so the checks are as they were.
    await expect(page.locator('#checks-summary')).toHaveText('15 of 15 checks passed');
    await expect(page.locator('#group-pdf')).toBeHidden();
    expect(downloads).toBe(0);
    // Trying again, with the server working: the PDF is made, checked and downloaded.
    await page.unroute('**/api/pdf');
    await expect(page.locator('#download')).toHaveText('download pdf');
    expect((await download(page, 'pdf')).name).toBe('redline.pdf');
    await expect(page.locator('#verdict-pdf')).toHaveText('all passed');
  });

  test('no PDF can be asked for when the static redline itself was withheld', async ({ page }) => {
    const brokenCase = BROKEN_STATIC_CASES.find((c) => c.id === 'static-leftover-tracked-change')!;
    const pair = await buildPair(pairById(brokenCase.pairId));
    await initialize();
    const files = await unzip((await makeStaticRedline(await produceTrackedChanges(pair.original, pair.modified), pair.original, { originalName: 'original.docx', modifiedName: 'modified.docx', comparedAt: 'x' })).docx);
    brokenCase.damage(files);
    await tamper(page, 'static-made', 'staticRedline', await zip(files));
    let asked = false;
    await page.route('**/api/pdf', (route) => {
      asked = true;
      return route.abort();
    });
    await compare(page, pair.original, pair.modified, 'withheld');
    await expect(page.locator('#download-menu [data-format="pdf"]')).toBeDisabled();
    await expect(page.locator('#download-menu [data-format="static"]')).toBeDisabled();
    // The button offers the one file that passed, and pressing it asks the server for nothing.
    await expect(page.locator('#download')).toHaveText('download word · tracked');
    expect((await download(page, 'tracked')).name).toBe('redline-tracked-changes.docx');
    await expect(page.getByRole('button', { name: 'make the pdf' })).toBeHidden();
    expect(asked).toBe(false);
  });
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

test.describe('refusals, in plain English, under the slot the file was added to', () => {
  const refuse = async (page: Page, original: { name: string; buffer: Buffer }, expected: RegExp): Promise<void> => {
    const pair = await buildPair(PAIR_SPECS[0]!);
    await page.goto('/app');
    await page.locator('#modified').setInputFiles({ name: 'modified.docx', mimeType: DOCX_MIME, buffer: Buffer.from(pair.modified) });
    await page.locator('#original').setInputFiles({ name: original.name, mimeType: 'application/octet-stream', buffer: original.buffer });
    await expect(page.locator('#status')).toHaveAttribute('data-state', 'refused');
    await expect(page.locator('#error-original')).toBeVisible();
    await expect(page.locator('#error-original')).toHaveText(expected);
    await expect(page.locator('#error-original')).toHaveCSS('color', 'rgb(255, 142, 144)');
    // The file is not kept, so nothing can be created from it.
    await expect(page.locator('#pick-original')).toBeVisible();
    await expect(page.locator('#file-original')).toBeHidden();
    await expect(page.locator('#create')).toBeDisabled();
    await expect(page.locator('#section-export')).toBeHidden();
    await expect(page.locator('#error-modified')).toBeHidden();
    // A good file in its place clears the message.
    await page.locator('#original').setInputFiles({ name: 'original.docx', mimeType: DOCX_MIME, buffer: Buffer.from(pair.original) });
    await expect(page.locator('#error-original')).toBeHidden();
    await expect(page.locator('#create')).toBeEnabled();
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
  test('a file dropped onto a slot is checked the same way, and a good one is taken', async ({ page }) => {
    const pair = await buildPair(PAIR_SPECS[0]!);
    await page.goto('/app');
    const drop = async (name: string, bytes: Uint8Array | Buffer): Promise<void> => {
      const data = await page.evaluateHandle(
        ({ name, base64 }) => {
          const transfer = new DataTransfer();
          transfer.items.add(new File([Uint8Array.from(atob(base64), (c) => c.charCodeAt(0))], name));
          return transfer;
        },
        { name, base64: Buffer.from(bytes).toString('base64') },
      );
      await page.locator('#slot-modified').dispatchEvent('dragover', { dataTransfer: data });
      await expect(page.locator('#slot-modified')).toHaveClass(/over/);
      await page.locator('#slot-modified').dispatchEvent('drop', { dataTransfer: data });
    };
    await drop('notes.txt', Buffer.from('plain text'));
    await expect(page.locator('#error-modified')).toHaveText(/is not a Word document/);
    await drop('Later draft.docx', pair.modified);
    await expect(page.locator('#name-modified')).toHaveText('Later draft');
    await expect(page.locator('#error-modified')).toBeHidden();
  });
});

test.describe('an upload that already contains tracked changes is compared as if they had been accepted', () => {
  const NOTICE = {
    original: 'The original already contained tracked changes. It was compared as if they had all been accepted.',
    modified: 'The modified document already contained tracked changes. It was compared as if they had all been accepted.',
    both: 'Both documents already contained tracked changes. They were compared as if those had all been accepted.',
  };

  test('as the original, as the modified document and as both: a notice, both Word files pass, and the cover page says so', async ({ page }) => {
    const pair = await buildPair(pairById('several-changes'));
    await initialize();
    const withChanges = await produceTrackedChanges(pair.original, pair.modified); // reads as the modified document
    for (const [original, modified, which] of [[withChanges, pair.original, 'original'], [pair.original, withChanges, 'modified'], [withChanges, withChanges, 'both']] as const) {
      await compare(page, original, modified);
      await expect(page.locator('#tracked-notice')).toBeVisible();
      await expect(page.locator('#tracked-notice')).toHaveText(NOTICE[which]);
      await expect(page.locator('#verdict-tracked')).toHaveText('all passed');
      await expect(page.locator('#verdict-static')).toHaveText('all passed');
      const staticRedline = await download(page, 'static');
      const tracked = await download(page, 'tracked');
      // Checked again outside the browser, against the uploads as uploaded.
      expect((await checkTrackedChanges({ original, modified, output: tracked.bytes })).passed).toBe(true);
      const xml = new TextDecoder().decode((await unzip(staticRedline.bytes)).get('word/document.xml')!);
      expect(xml).toContain(which === 'both' ? 'Note: both documents contained tracked changes.' : `Note: the ${which === 'original' ? 'original' : 'modified document'} contained tracked changes.`);
      // The check report the user can keep says how the upload was read.
      expect(await downloadReport(page)).toContain('contained tracked changes of its own; it was read as if they had all been accepted.');
      // The viewer shows the cover page with the note.
      await expect(page.locator('#doc').getByText('contained tracked changes', { exact: false }).first()).toBeVisible();
    }
  });

  test('an ordinary comparison shows no such notice, and a new comparison clears an earlier one', async ({ page }) => {
    const pair = await buildPair(pairById('several-changes'));
    await initialize();
    const withChanges = await produceTrackedChanges(pair.original, pair.modified);
    await compare(page, withChanges, pair.original);
    await expect(page.locator('#tracked-notice')).toBeVisible();
    await compare(page, pair.original, pair.modified);
    await expect(page.locator('#tracked-notice')).toBeHidden();
    await expect(page.locator('#verdict-static')).toHaveText('all passed');
  });
});

/**
 * Swap one of the Producer's genuine outputs for a damaged one on its way from
 * the Producer's worker to the page. Done entirely from the test, by wrapping
 * the browser's Worker class before the page loads; the product has no switch
 * for it.
 */
async function tamper(page: Page, messageType: 'compared' | 'static-made' | 'rendered', field: 'trackedChanges' | 'staticRedline' | 'html', replacement: Uint8Array | string): Promise<void> {
  await page.addInitScript(
    ({ base64, text, messageType, field }) => {
      const RealWorker = window.Worker;
      window.Worker = class extends RealWorker {
        constructor(url: string | URL, options?: WorkerOptions) {
          super(url, options);
          if (options?.name !== 'redline-compare') return;
          this.addEventListener('message', (event: MessageEvent) => {
            if (event.data?.type !== messageType) return;
            if (text !== undefined) {
              event.data[field] = text;
              return;
            }
            const binary = atob(base64);
            const bytes = new Uint8Array(binary.length);
            for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
            event.data[field] = bytes.buffer;
          });
        }
      };
    },
    { base64: typeof replacement === 'string' ? '' : Buffer.from(replacement).toString('base64'), text: typeof replacement === 'string' ? replacement : undefined, messageType, field },
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
    // The summary turns red and counts what passed (the static redline is made from the same damaged file, so its checks may fail too).
    const failed = await page.locator('#report-tracked li[data-passed="false"]').count();
    expect(failed).toBeGreaterThanOrEqual(expectedFailures(brokenCase).length);
    const passedInAll = await page.locator('#rl-checklist li[data-passed="true"]').count();
    await expect(page.locator('#checks-summary')).toHaveText(`${passedInAll} of 15 checks passed`);
    await expect(page.locator('#checks-summary')).toHaveCSS('color', 'rgb(255, 142, 144)');
    await expect(page.locator('#verdict-tracked')).toHaveText(`${failed} failed`);
    await expect(page.locator('#verdict-tracked')).toHaveCSS('color', 'rgb(255, 142, 144)');
    await page.locator('#checks-toggle').click();
    await expect(page.locator('#report-tracked li[data-passed="false"]').first().getByRole('img', { name: 'Failed' })).toBeVisible();
    // The page says so in plain sight, beside the downloads.
    await expect(page.locator('#withheld-note')).toBeVisible();
    await expect(page.locator('#withheld-note')).toHaveText(/did not pass its checks.*has been withheld\. See the checks\./);
    // Not offered: the row is greyed out and cannot be chosen, and the button downloads something else.
    await page.locator('#download-more').click();
    const item = page.locator('#download-menu [data-format="tracked"]');
    await expect(item).toBeDisabled();
    await expect(item).toContainText('Not offered: a check failed');
    await item.click({ force: true });
    await page.keyboard.press('Escape');
    await expect(page.locator('#download')).not.toHaveAttribute('data-format', 'tracked');
    expect(await downloadReport(page)).toContain('FAILED - this output was withheld');
  });
}

test('a remembered format that is withheld is not downloaded: the button falls back to one that passed', async ({ page }) => {
  const brokenCase = BROKEN_CASES.find((c) => c.id === 'dropped-word')!;
  const pair = await buildPair(pairById(brokenCase.pairId));
  await initialize();
  await page.addInitScript(() => localStorage.setItem('redline.format', 'tracked'));
  await tamper(page, 'compared', 'trackedChanges', await breakOutput(await produceTrackedChanges(pair.original, pair.modified), brokenCase));
  await compare(page, pair.original, pair.modified, 'withheld');
  await expect(page.locator('#download-menu [data-format="tracked"]')).toBeDisabled();
  await expect(page.locator('#download')).not.toHaveAttribute('data-format', 'tracked');
  // Whatever the button now offers passed its checks; if nothing did, the button is switched off.
  let downloads = 0;
  page.on('download', () => downloads++);
  const offered = await page.locator('#download').getAttribute('data-format');
  if (offered) await expect(page.locator(`#download-menu [data-format="${offered}"]`)).toBeEnabled();
  else {
    await expect(page.locator('#download')).toBeDisabled();
    await page.locator('#download').click({ force: true });
    for (const format of ['pdf', 'static', 'tracked']) await expect(page.locator(`#download-menu [data-format="${format}"]`)).toBeDisabled();
    expect(downloads).toBe(0);
  }
});

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
    // Neither the static redline nor the PDF made from it is offered.
    for (const format of ['static', 'pdf']) {
      await expect(page.locator(`#download-menu [data-format="${format}"]`)).toBeDisabled();
      await expect(page.locator(`#download-menu [data-format="${format}"]`)).toContainText('Not offered: a check failed');
    }
    // The viewer does not show a redline that failed, and no unverified counts are shown.
    await expect(page.locator('#viewer')).toBeHidden();
    await expect(page.locator('#doc')).toBeEmpty();
    await expect(page.locator('#viewer-tools')).toBeHidden();
    await expect(page.getByRole('heading', { name: 'redline withheld' })).toBeVisible();
    await expect(page.locator('#viewer-message')).toContainText('did not pass its checks, so it is not shown');
    await expect(page.locator('#withheld-note')).toBeVisible();
    await expect(page.locator('#withheld-note')).toContainText('has been withheld');
    await expect(page.locator('#section-changes')).toBeHidden();
    // The tracked-changes file was fine and is still offered.
    await expect(page.locator('#download')).toHaveText('download word · tracked');
    expect((await download(page, 'tracked')).name).toBe('redline-tracked-changes.docx');
  });
}

test('a withheld output shows the first paragraph where it went wrong', async ({ page }) => {
  const brokenCase = BROKEN_CASES.find((c) => c.id === 'dropped-word')!;
  const pair = await buildPair(pairById(brokenCase.pairId));
  await initialize();
  await tamper(page, 'compared', 'trackedChanges', await breakOutput(await produceTrackedChanges(pair.original, pair.modified), brokenCase));
  await compare(page, pair.original, pair.modified, 'withheld');
  await page.locator('#checks-toggle').click();
  const failure = page.locator('#report-tracked li[data-check="accept-all"] .failure');
  await expect(failure).toBeVisible();
  await expect(failure).toContainText('Where: main text, paragraph 2');
  await expect(failure).toContainText(/Expected: .*provide the Services with exceptional skill and care\./);
  await expect(failure).toContainText(/Found: .*provide the Services with exceptional and care\./);
});
