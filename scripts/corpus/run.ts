/**
 * Stage 7 tooling: runs every contract in a folder through the product and
 * scores the result. For each contract: apply 150+ known edits (mutate.ts),
 * make both Word files as the website does, run the Checker on both, and
 * measure how much the redline marks against how much really changed.
 *
 * Usage: npx tsx scripts/corpus/run.ts <folder of .docx> <results.jsonl> [only-matching-text]
 *
 * MODE (environment variable) chooses what is uploaded, to test the hard kinds on purpose:
 *   plain (default)    the contract with its own tracked changes accepted, against that with the edits applied
 *   raw                the contract exactly as published (tracked changes and all), against the edited copy
 *   tracked-original   the original upload carries tracked changes: a tracked-changes file of a first round of
 *                      edits, against a copy with a second round applied on top
 *   tracked-modified   the modified upload carries tracked changes: the contract against a tracked-changes file
 *                      whose accepted reading has two rounds of edits
 *   pdf                the edited copy is turned into a PDF and back into Word (as when a counterparty sends a
 *                      PDF and it is converted) before it is uploaded. Needs PDF_URL (the converter's /convert
 *                      address) and the Python package pdf2docx.
 * Real documents are never committed: the folder is normally tests/fixtures/private/corpus/.
 */
import { execFileSync } from 'node:child_process';
import { appendFileSync, existsSync, mkdtempSync, readdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { initialize } from 'docxodus/core';
import { checkStaticRedline, checkTrackedChanges } from '../../src/checker';
import { acceptTrackedChanges, hasTrackedChanges } from '../../src/producer/accept-changes';
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

const MODE = process.env.MODE ?? 'plain';
if (!['plain', 'raw', 'tracked-original', 'tracked-modified', 'pdf'].includes(MODE)) throw new Error(`unknown MODE ${MODE}`);

/** Word to PDF (the product's own converter) and back to Word (pdf2docx, an open-source converter). */
async function throughPdf(docx: Uint8Array): Promise<Uint8Array> {
  const url = process.env.PDF_URL;
  if (!url) throw new Error('MODE=pdf needs PDF_URL');
  const response = await fetch(url, { method: 'POST', headers: { 'Content-Type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' }, body: new Blob([docx.slice().buffer]) });
  if (!response.ok) throw new Error(`to PDF: status ${response.status}`);
  const dir = mkdtempSync(join(tmpdir(), 'roundtrip-'));
  try {
    writeFileSync(join(dir, 'in.pdf'), new Uint8Array(await response.arrayBuffer()));
    execFileSync('python3', ['-I', '-c', 'import sys,logging\nlogging.disable(logging.CRITICAL)\nfrom pdf2docx import Converter\nc=Converter(sys.argv[1]); c.convert(sys.argv[2]); c.close()', join(dir, 'in.pdf'), join(dir, 'out.docx')], { stdio: 'ignore', timeout: 1_800_000 });
    return new Uint8Array(readFileSync(join(dir, 'out.docx')));
  } finally {
    rmSync(dir, { recursive: true, force: true });
  }
}

/** The two uploads for a contract, in the chosen mode, and how many edits went in. */
async function uploadsFor(published: Uint8Array, seed: number): Promise<{ original: Uint8Array; modified: Uint8Array; edits: number; tally: Record<string, number> }> {
  const accepted = await acceptTrackedChanges(published);
  const first = await mutate(accepted, seed);
  const add = (a: Record<string, number>, b: Record<string, number>): Record<string, number> => Object.fromEntries([...new Set([...Object.keys(a), ...Object.keys(b)])].map((k) => [k, (a[k] ?? 0) + (b[k] ?? 0)]));
  if (MODE === 'plain') return { original: accepted, modified: first.docx, edits: first.edits.length, tally: first.tally };
  if (MODE === 'raw') return { original: published, modified: first.docx, edits: first.edits.length, tally: first.tally };
  if (MODE === 'pdf') return { original: accepted, modified: await throughPdf(first.docx), edits: first.edits.length, tally: first.tally };
  const second = await mutate(first.docx, seed + 1);
  const both = { edits: first.edits.length + second.edits.length, tally: add(first.tally, second.tally) };
  if (MODE === 'tracked-original') return { original: await produceTrackedChanges(accepted, first.docx), modified: second.docx, ...both };
  return { original: accepted, modified: await produceTrackedChanges(first.docx, second.docx), ...both };
}

for (const file of readdirSync(folder).filter((name) => name.toLowerCase().endsWith('.docx')).sort()) {
  if (done.has(file) || (only && !file.includes(only))) continue;
  const row: Record<string, unknown> = { file };
  try {
    const uploads = await uploadsFor(new Uint8Array(readFileSync(join(folder, file))), seedOf(file));
    const { original } = uploads;
    const mutation = { docx: uploads.modified };
    row.mode = MODE;
    row.edits = uploads.edits;
    row.tally = uploads.tally;
    // What really changed is measured between the two uploads as the product reads them: tracked changes accepted.
    const inOriginal = await hasTrackedChanges(original);
    const inModified = await hasTrackedChanges(mutation.docx);
    row.trackedChangesIn = inOriginal && inModified ? 'both' : inOriginal ? 'original' : inModified ? 'modified' : undefined;
    const truth = await trueDifference(inOriginal ? await acceptTrackedChanges(original) : original, inModified ? await acceptTrackedChanges(mutation.docx) : mutation.docx);
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
      const made = await makeStaticRedline(tracked, original, { ...COVER, trackedChangesIn: row.trackedChangesIn as 'original' | 'modified' | 'both' | undefined });
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
