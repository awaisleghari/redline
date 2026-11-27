import { mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { basename } from 'node:path';
import { initialize } from 'docxodus/core';
import { checkStaticRedline, checkTrackedChanges } from '../../../src/checker';
import { acceptTrackedChanges } from '../../../src/producer/accept-changes';
import { makeStaticRedline } from '../../../src/producer/make-static-redline';
import { produceTrackedChanges } from '../../../src/producer/produce';
import { mutate } from '../mutate';
const [mode, path, out] = process.argv.slice(2);
await initialize();
const file = basename(path!);
const seed = [...file].reduce((h, c) => (Math.imul(h, 31) + c.charCodeAt(0)) >>> 0, 7);
const accepted = await acceptTrackedChanges(new Uint8Array(readFileSync(path!)));
const first = await mutate(accepted, seed);
const second = await mutate(first.docx, seed + 1);
const original = mode === 'tracked-original' ? await produceTrackedChanges(accepted, first.docx) : accepted;
const modified = mode === 'tracked-original' ? second.docx : await produceTrackedChanges(first.docx, second.docx);
const COVER = { originalName: 'original.docx', modifiedName: 'modified.docx', comparedAt: '24 November 2026 at 12:00 (UTC+00:00)' };
const t0 = Date.now();
const tracked = await produceTrackedChanges(original, modified);
console.log('engine ms', Date.now() - t0, 'rss MB', Math.round(process.memoryUsage().rss / 1e6));
const made = await makeStaticRedline(tracked, original, { ...COVER, trackedChangesIn: mode === 'tracked-original' ? 'original' : 'modified' } as never);
console.log('static done, rss MB', Math.round(process.memoryUsage().rss / 1e6));
if (out) { mkdirSync(out, { recursive: true }); for (const [n, b] of [['original', original], ['modified', modified], ['tracked', tracked], ['static', made.docx]] as const) writeFileSync(`${out}/${n}.docx`, b); }
for (const [label, report] of [['tracked', await checkTrackedChanges({ original, modified, output: tracked })], ['static', await checkStaticRedline({ original, modified, output: made.docx, ...COVER })]] as const)
  for (const c of report.checks.filter((c) => !c.passed)) console.log(label, c.id, JSON.stringify(c.failure).slice(0, 600));
console.log('done, rss MB', Math.round(process.memoryUsage().rss / 1e6));
