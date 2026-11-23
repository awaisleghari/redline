/**
 * Checks a deployed site from outside, without a browser (the browser tests do
 * the rest: REDLINE_URL=<address> npm run test:e2e).
 *
 * Usage: npm run test:live -- https://redline.<account>.workers.dev
 *
 * It checks that the page and its security headers are served, measures how
 * much a first visit downloads, makes real PDFs through the live converter and
 * runs the Checker (check 9) on each, confirms the refusals, and reads the
 * converter's state report to confirm that it holds no file and runs no
 * process afterwards. Needs `npm run build` first (it compares against dist/).
 * Needs curl. Behind a proxy, run with NODE_USE_ENV_PROXY=1.
 */
import { execFileSync } from 'node:child_process';
import { readdirSync, readFileSync, statSync } from 'node:fs';
import { join } from 'node:path';
import { initialize } from 'docxodus/core';
import { checkPdf } from '../src/checker';
import { makeStaticRedline } from '../src/producer/make-static-redline';
import { produceTrackedChanges } from '../src/producer/produce';
import { buildLargeContract } from '../tests/fixtures/large';
import { buildPair, pairById } from '../tests/fixtures/pairs';

const site = process.argv[2]?.replace(/\/+$/, '');
if (!site || !/^https?:\/\//.test(site)) throw new Error('usage: live-check.ts <site address>');
const DOCX_TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
const COUNTS = /^conversions=(\d+) files=(\d+) office_processes=(\d+)$/;

let failures = 0;
const result = (passed: boolean, what: string, detail = ''): void => {
  if (!passed) failures++;
  console.log(`${passed ? 'PASS' : 'FAIL'}  ${what}${detail ? `  (${detail})` : ''}`);
};
const mb = (bytes: number): string => `${(bytes / 1024 / 1024).toFixed(1)} MB`;

// 1. The page and its headers.
const page = await fetch(`${site}/`);
const html = await page.text();
result(page.status === 200 && html.includes('<title>Redline'), 'the page is served');
result(html.includes('Content-Security-Policy'), "the page carries its content-security policy (nothing loaded from or sent to another site)");
for (const [name, value] of [['x-content-type-options', 'nosniff'], ['x-frame-options', 'DENY'], ['referrer-policy', 'no-referrer']] as const) {
  result(page.headers.get(name) === value, `header ${name}: ${value}`, `got ${page.headers.get(name)}`);
}
for (const path of ['/wrangler.jsonc', '/worker/index.ts', '/.git/config', '/api/other']) {
  const reply = await fetch(site + path);
  await reply.arrayBuffer();
  result(reply.status === 404, `${path} is not served`, `status ${reply.status}`);
}

// 2. What a first visit downloads: every built file, as the host sends it (compressed) against its size on disk.
const walk = (folder: string): string[] => readdirSync(folder).flatMap((name) => (statSync(join(folder, name)).isDirectory() ? walk(join(folder, name)) : [join(folder, name)]));
let onDisk = 0;
let onWire = 0;
let wrong = 0;
const files = walk('dist').filter((file) => !file.endsWith('_headers'));
for (const file of files) {
  const path = '/' + file.slice('dist/'.length).replaceAll('\\', '/');
  const address = site + (path === '/index.html' ? '/' : path);
  const reply = await fetch(address);
  const body = new Uint8Array(await reply.arrayBuffer());
  const local = readFileSync(file);
  if (reply.status !== 200 || body.byteLength !== local.byteLength || !body.every((byte, index) => byte === local[index])) wrong++;
  onDisk += local.byteLength;
  // The bytes actually sent when a browser asks for the compressed form (curl counts them; fetch would decompress silently).
  onWire += Number(execFileSync('curl', ['-s', '-o', '/dev/null', '-H', 'Accept-Encoding: br, gzip', '-w', '%{size_download}', address], { encoding: 'utf8' }));
}
result(wrong === 0, `all ${files.length} built files are served, byte for byte as built`, wrong ? `${wrong} differ or are missing` : mb(onDisk));
console.log(`INFO  first-visit download: ${mb(onDisk)} as built; ${mb(onWire)} as actually sent to a browser (compressed by the host)`);

// 3. Real PDFs through the live converter, each checked by the Checker.
await initialize();
const state = async (): Promise<string[]> => ((await (await fetch(`${site}/api/converter-state`)).json()) as { converters: string[] }).converters;
const conversionsIn = (lines: string[]): number => lines.reduce((sum, line) => sum + Number(COUNTS.exec(line)?.[1] ?? 0), 0);
const before = await state();
console.log(`INFO  converter state before: ${JSON.stringify(before)}`);
const cover = { originalName: 'original.docx', modifiedName: 'modified.docx', comparedAt: '23 November 2026 at 14:05 (UTC+00:00)' };
const documents: [string, Uint8Array][] = [];
for (const id of ['several-changes', 'moved-paragraph', 'footnote-change', 'numbered-list-change', 'header-change']) {
  const pair = await buildPair(pairById(id));
  documents.push([id, (await makeStaticRedline(await produceTrackedChanges(pair.original, pair.modified), pair.original, cover)).docx]);
}
const large = await buildLargeContract();
documents.push(['100-page contract', (await makeStaticRedline(await produceTrackedChanges(large.original, large.modified), large.original, cover)).docx]);
const askForPdf = (docx: Uint8Array, headers: Record<string, string> = {}): Promise<Response> =>
  fetch(`${site}/api/pdf`, { method: 'POST', headers: { 'Content-Type': DOCX_TYPE, Origin: site, ...headers }, body: new Blob([docx.slice().buffer]) });
let made = 0;
for (const [id, docx] of documents) {
  const started = performance.now();
  const reply = await askForPdf(docx);
  const seconds = ((performance.now() - started) / 1000).toFixed(1);
  if (reply.status !== 200) {
    result(false, `PDF of "${id}"`, `status ${reply.status}: ${await reply.text()}`);
    continue;
  }
  made++;
  const report = await checkPdf({ staticRedline: docx, pdf: new Uint8Array(await reply.arrayBuffer()) });
  const failed = report.checks.filter((check) => !check.passed).map((check) => check.failure?.message);
  result(report.passed && reply.headers.get('cache-control') === 'no-store', `PDF of "${id}" made on the live converter passes check 9 (text and colours), reply not cacheable`, failed.length ? failed.join('; ') : `${seconds} s`);
}

// 4. Refusals.
const sample = documents[0]![1];
result((await askForPdf(sample, { Origin: 'https://elsewhere.example' })).status === 403, 'another website cannot use the converter');
result((await fetch(`${site}/api/pdf`)).status === 405, 'the converter answers only when a document is sent');
result((await askForPdf(new TextEncoder().encode('not a document'))).status === 400, 'a file that is not a Word document is refused');

// 5. The converter afterwards: it did the work, and holds nothing.
const after = await state();
console.log(`INFO  converter state after:  ${JSON.stringify(after)}`);
const awake = after.filter((line) => line !== 'asleep');
result(after.length === 2 && awake.length > 0 && awake.every((line) => COUNTS.test(line)), 'the state report gives counts only');
result(conversionsIn(after) - conversionsIn(before) >= made, `the copies of the converter now awake are the ones that did these ${made} conversions`, `${conversionsIn(after) - conversionsIn(before)} more counted than before`);
result(awake.every((line) => COUNTS.exec(line)?.[2] === '0'), 'after converting, the converter holds no file that was not there when it started');
result(awake.every((line) => COUNTS.exec(line)?.[3] === '0'), 'after converting, no LibreOffice process is left running');

console.log(failures === 0 ? '\nAll live checks passed.' : `\n${failures} live check(s) FAILED.`);
process.exit(failures === 0 ? 0 : 1);
