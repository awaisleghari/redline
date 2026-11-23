/**
 * The deployed shape of the product, run locally: Cloudflare's runtime serving
 * the built site, the Worker's two routes, and the converter container behind
 * it. Everything here goes through the Worker, as a visitor's browser would.
 */
import { execFileSync, spawnSync } from 'node:child_process';
import { initialize } from 'docxodus/core';
import { beforeAll, describe, expect, it } from 'vitest';
import { checkPdf } from '../../src/checker';
import { makeStaticRedline } from '../../src/producer/make-static-redline';
import { produceTrackedChanges } from '../../src/producer/produce';
import { buildPair, pairById } from '../fixtures/pairs';
import { converterContainers, WORKER_URL } from './wrangler-dev';

const DOCX_TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
const COUNTS = /^conversions=(\d+) files=(\d+) office_processes=(\d+)$/;

async function staticRedline(pairId: string): Promise<Uint8Array> {
  const pair = await buildPair(pairById(pairId));
  const cover = { originalName: 'original.docx', modifiedName: 'modified.docx', comparedAt: '23 November 2026 at 14:05 (UTC+00:00)' };
  return (await makeStaticRedline(await produceTrackedChanges(pair.original, pair.modified), pair.original, cover)).docx;
}
const askForPdf = (docx: Uint8Array, headers: Record<string, string> = {}): Promise<Response> =>
  fetch(`${WORKER_URL}/api/pdf`, { method: 'POST', headers: { 'Content-Type': DOCX_TYPE, Origin: WORKER_URL, ...headers }, body: new Blob([docx.slice().buffer]) });
const converterState = async (): Promise<string[]> => ((await (await fetch(`${WORKER_URL}/api/converter-state`)).json()) as { converters: string[] }).converters;

beforeAll(async () => {
  await initialize();
});

describe('the site, as Cloudflare serves it', () => {
  it('serves the page with its security headers, and the engine files', async () => {
    const page = await fetch(`${WORKER_URL}/`);
    expect(page.status).toBe(200);
    expect(page.headers.get('Content-Type')).toContain('text/html');
    expect(page.headers.get('X-Content-Type-Options')).toBe('nosniff');
    expect(page.headers.get('X-Frame-Options')).toBe('DENY');
    expect(page.headers.get('Referrer-Policy')).toBe('no-referrer');
    const html = await page.text();
    expect(html).toContain('Content-Security-Policy');
    const script = /src="(\/assets\/[^"]+\.js)"/.exec(html)![1]!;
    const asset = await fetch(WORKER_URL + script);
    expect(asset.status).toBe(200);
    expect(asset.headers.get('Cache-Control')).toContain('immutable');
    const engine = await fetch(`${WORKER_URL}/wasm/_framework/dotnet.js`);
    expect(engine.status).toBe(200);
    await engine.arrayBuffer();
  });

  it('answers an address that is not a file with a plain 404, not with the page', async () => {
    for (const path of ['/nothing-here', '/api/other', '/wrangler.jsonc', '/worker/index.ts', '/.git/config']) {
      const reply = await fetch(WORKER_URL + path);
      expect(reply.status, path).toBe(404);
      await reply.arrayBuffer();
    }
  });
});

describe('the PDF route, through the Worker to the converter', () => {
  it('turns a static redline into a PDF that passes check 9, and the reply cannot be cached', async () => {
    for (const id of ['several-changes', 'footnote-change', 'numbered-list-change']) {
      const docx = await staticRedline(id);
      const reply = await askForPdf(docx);
      expect(reply.status, id).toBe(200);
      expect(reply.headers.get('Content-Type')).toBe('application/pdf');
      expect(reply.headers.get('Cache-Control')).toBe('no-store');
      const report = await checkPdf({ staticRedline: docx, pdf: new Uint8Array(await reply.arrayBuffer()) });
      expect(report.checks.filter((check) => !check.passed), id).toEqual([]);
      expect(report.passed, id).toBe(true);
    }
  });

  it('refuses another site, anything but POST, and what is not a Word document', async () => {
    const docx = await staticRedline('one-word-change');
    expect((await askForPdf(docx, { Origin: 'https://elsewhere.example' })).status).toBe(403);
    expect((await fetch(`${WORKER_URL}/api/pdf`)).status).toBe(405);
    expect((await askForPdf(docx, { 'Content-Type': 'application/pdf' })).status).toBe(415);
    expect((await askForPdf(new TextEncoder().encode('not a document'))).status).toBe(400);
    // A zip that is not a document reaches the converter, which cannot convert it; its internal reason is not passed on.
    const refused = await askForPdf(new Uint8Array([0x50, 0x4b, 3, 4, 0, 0, 0, 0]));
    expect(refused.status).toBe(422);
    expect(await refused.text()).toBe('The document could not be converted to PDF.');
  });

  it('several conversions at once all succeed', async () => {
    const docx = await staticRedline('table-cell-change');
    const replies = await Promise.all(Array.from({ length: 4 }, () => askForPdf(docx)));
    expect(replies.map((reply) => reply.status)).toEqual([200, 200, 200, 200]);
    await Promise.all(replies.map((reply) => reply.arrayBuffer()));
  });
});

describe('the converter keeps nothing and says nothing, as seen through the Worker', () => {
  it('the state report gives counts only, and after all the conversions above: no file, no process', async () => {
    const states = await converterState();
    expect(states).toHaveLength(2);
    let conversions = 0;
    for (const line of states) {
      if (line === 'asleep') continue;
      const counts = COUNTS.exec(line);
      expect(counts, line).not.toBeNull();
      conversions += Number(counts![1]);
      expect(Number(counts![2]), 'files held').toBe(0);
      expect(Number(counts![3]), 'LibreOffice processes').toBe(0);
    }
    // 3 good + 1 unconvertible + 4 at once reached a converter; the refusals before that never did.
    expect(conversions).toBe(8);
    const reply = await fetch(`${WORKER_URL}/api/converter-state`);
    expect(reply.headers.get('Cache-Control')).toBe('no-store');
    expect((await fetch(`${WORKER_URL}/api/converter-state`, { method: 'POST' })).status).toBe(405);
  });

  it('the report agrees with a direct search of each converter as root, and the converters have logged nothing', () => {
    const names = converterContainers();
    expect(names.length).toBeGreaterThan(0);
    for (const name of names) {
      // Everything changed on the container's disk since it started, anywhere, searched as root.
      const changed = execFileSync('docker', ['diff', name], { encoding: 'utf8' }).split('\n').filter(Boolean);
      expect(changed.filter((line) => !/^[CA] \/(tmp|run|home|home\/converter)$/.test(line)), name).toEqual([]);
      const office = execFileSync('docker', ['exec', '-u', '0', name, 'sh', '-c', 'for p in /proc/[0-9]*; do tr "\\0" " " < $p/cmdline 2>/dev/null; echo; done | grep "[s]office" || true'], { encoding: 'utf8' });
      expect(office.trim(), name).toBe('');
      const log = spawnSync('docker', ['logs', name], { encoding: 'utf8' });
      expect(log.stdout + log.stderr, name).toBe('');
    }
  });
});
