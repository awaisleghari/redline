/**
 * Stage 7 tooling: run one contract and print every failing check in full.
 * Usage: npx tsx scripts/corpus/one.ts <file.docx> [output folder]
 */
import { mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { basename } from 'node:path';
import { initialize } from 'docxodus/core';
import { checkStaticRedline, checkTrackedChanges } from '../../src/checker';
import { acceptTrackedChanges } from '../../src/producer/accept-changes';
import { makeStaticRedline } from '../../src/producer/make-static-redline';
import { produceTrackedChanges } from '../../src/producer/produce';
import { mutate } from './mutate';

const [path, out] = process.argv.slice(2);
await initialize();
const file = basename(path!);
const seed = [...file].reduce((h, c) => (Math.imul(h, 31) + c.charCodeAt(0)) >>> 0, 7);
const original = await acceptTrackedChanges(new Uint8Array(readFileSync(path!)));
const mutation = await mutate(original, seed);
const COVER = { originalName: 'original.docx', modifiedName: 'modified.docx', comparedAt: '24 November 2026 at 12:00 (UTC+00:00)' };
const tracked = await produceTrackedChanges(original, mutation.docx);
const made = await makeStaticRedline(tracked, original, COVER);
if (out) {
  mkdirSync(out, { recursive: true });
  for (const [name, bytes] of [['original', original], ['modified', mutation.docx], ['tracked', tracked], ['static', made.docx]] as const) writeFileSync(`${out}/${name}.docx`, bytes);
  writeFileSync(`${out}/edits.json`, JSON.stringify(mutation.edits, null, 1));
}
for (const [label, report] of [['tracked', await checkTrackedChanges({ original, modified: mutation.docx, output: tracked })], ['static', await checkStaticRedline({ original, modified: mutation.docx, output: made.docx, ...COVER })]] as const) {
  for (const check of report.checks.filter((c) => !c.passed)) console.log(label, check.id, JSON.stringify(check.failure).slice(0, 700));
}
console.log('done', JSON.stringify(made.counts));
