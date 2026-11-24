/**
 * Stage 7 tooling: runs every contract in a folder through the product and
 * scores the result. For each contract: apply 150+ known edits (mutate.ts),
 * make both Word files as the website does, run the Checker on both, and
 * measure how much the redline marks against how much really changed.
 *
 * Usage: npx tsx scripts/corpus/run.ts <folder of .docx> <results.jsonl> [only-matching-text]
 * Real documents are never committed: the folder is normally tests/fixtures/private/corpus/.
 */
import { appendFileSync, existsSync, readdirSync, readFileSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { initialize } from 'docxodus/core';
import { checkStaticRedline, checkTrackedChanges } from '../../src/checker';
import { acceptTrackedChanges } from '../../src/producer/accept-changes';
import { makeStaticRedline } from '../../src/producer/make-static-redline';
import { produceTrackedChanges } from '../../src/producer/produce';
import { markedIn, trueDifference } from './measure';
import { mutate } from './mutate';

const [folder, resultsPath, only] = process.argv.slice(2);
if (!folder || !resultsPath) throw new Error('usage: run.ts <folder> <results.jsonl> [only]');
await initialize();
const COVER = { originalName: 'original.docx', modifiedName: 'modified.docx', comparedAt: '24 November 2026 at 12:00 (UTC+00:00)' };
const done = new Set(existsSync(resultsPath) ? readFileSync(resultsPath, 'utf8').split('\n').filter(Boolean).map((line) => (JSON.parse(line) as { file: string }).file) : []);
const seedOf = (name: string): number => [...name].reduce((h, c) => (Math.imul(h, 31) + c.charCodeAt(0)) >>> 0, 7);
const failuresOf = (report: { checks: { id: string; passed: boolean; failure?: { message: string; where?: string } }[] }): { id: string; message: string; where?: string }[] =>
  report.checks.filter((c) => !c.passed).map((c) => ({ id: c.id, message: c.failure?.message ?? '', where: c.failure?.where }));

for (const file of readdirSync(folder).filter((name) => name.toLowerCase().endsWith('.docx')).sort()) {
  if (done.has(file) || (only && !file.includes(only))) continue;
  const row: Record<string, unknown> = { file };
  try {
    // An original that itself carries tracked changes is taken as accepted, so that the edits are the only difference.
    const original = await acceptTrackedChanges(new Uint8Array(readFileSync(join(folder, file))));
    const mutation = await mutate(original, seedOf(file));
    row.edits = mutation.edits.length;
    row.tally = mutation.tally;
    const truth = await trueDifference(original, mutation.docx);
    row.words = truth.words;
    row.trueDeleted = truth.deleted;
    row.trueInserted = truth.inserted;
    if (process.env.KEEP) writeFileSync(join(process.env.KEEP, file.replace(/\.docx$/i, ' (modified).docx')), mutation.docx);

    const started = performance.now();
    const tracked = await produceTrackedChanges(original, mutation.docx);
    row.engineSeconds = Number(((performance.now() - started) / 1000).toFixed(1));
    const marked = await markedIn(tracked);
    row.marked = marked;
    const trackedReport = await checkTrackedChanges({ original, modified: mutation.docx, output: tracked });
    row.trackedPassed = trackedReport.passed;
    row.trackedFailures = failuresOf(trackedReport);
    try {
      const made = await makeStaticRedline(tracked, original, COVER);
      row.counts = made.counts;
      const staticReport = await checkStaticRedline({ original, modified: mutation.docx, output: made.docx, ...COVER });
      row.staticPassed = staticReport.passed;
      row.staticFailures = failuresOf(staticReport);
      if (process.env.KEEP) writeFileSync(join(process.env.KEEP, file.replace(/\.docx$/i, ' (static redline).docx')), made.docx);
    } catch (error) {
      row.staticPassed = false;
      row.staticError = String(error).slice(0, 300);
    }
    row.totalSeconds = Number(((performance.now() - started) / 1000).toFixed(1));
  } catch (error) {
    row.error = String(error).slice(0, 300);
  }
  appendFileSync(resultsPath, JSON.stringify(row) + '\n');
  const m = row.marked as { inserted: number; deleted: number; movedFrom: number; movedTo: number } | undefined;
  const truth = (row.trueDeleted as number) + (row.trueInserted as number);
  console.log(`${file.slice(0, 50).padEnd(50)} edits ${row.edits ?? '-'} true ${truth || '-'} marked ${m ? m.inserted + m.deleted + m.movedFrom + m.movedTo : '-'} tracked ${row.trackedPassed ?? '-'} static ${row.staticPassed ?? '-'} ${row.error ?? ''} ${row.totalSeconds ?? ''}s`);
}
