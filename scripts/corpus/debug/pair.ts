// Re-run a saved pair of uploads: npx tsx pair.ts <dir with original.docx, modified.docx> [tracked-in: original|modified]
import { readFileSync, writeFileSync } from 'node:fs';
import { initialize } from 'docxodus/core';
import { checkStaticRedline, checkTrackedChanges } from '../../../src/checker';
import { makeStaticRedline } from '../../../src/producer/make-static-redline';
import { produceTrackedChanges } from '../../../src/producer/produce';
const [dir, where] = process.argv.slice(2);
await initialize();
const original = new Uint8Array(readFileSync(`${dir}/original.docx`));
const modified = new Uint8Array(readFileSync(`${dir}/modified.docx`));
const COVER = { originalName: 'original.docx', modifiedName: 'modified.docx', comparedAt: '24 November 2026 at 12:00 (UTC+00:00)' };
const tracked = await produceTrackedChanges(original, modified);
const made = await makeStaticRedline(tracked, original, { ...COVER, trackedChangesIn: where } as never);
writeFileSync(`${dir}/tracked.docx`, tracked); writeFileSync(`${dir}/static.docx`, made.docx);
for (const [label, report] of [['tracked', await checkTrackedChanges({ original, modified, output: tracked })], ['static', await checkStaticRedline({ original, modified, output: made.docx, ...COVER })]] as const)
  for (const c of report.checks.filter((c) => !c.passed)) console.log(label, c.id, JSON.stringify(c.failure).slice(0, 600));
console.log('done, rss MB', Math.round(process.memoryUsage().rss / 1e6));
