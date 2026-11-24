/**
 * Stage 7 tooling: looks at every downloaded file and keeps the ones that are
 * usable: a real Word document, long enough (at least 15 pages as LibreOffice
 * lays it out), and able to take at least 150 edits.
 *
 * Usage: npx tsx scripts/corpus/select.ts <downloaded folder> <corpus folder> <report.json>
 * Needs the PDF converter running (docker run --rm -p 127.0.0.1:8791:8080 redline-pdf) to count pages.
 */
import { copyFileSync, mkdirSync, readdirSync, readFileSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { readPdf } from '../../src/checker/pdf';
import { unzip } from '../../src/common/zip';
import { acceptTrackedChanges } from '../../src/producer/accept-changes';
import { wordsByPart } from './measure';
import { mutate } from './mutate';

const [downloaded, corpus, reportPath] = process.argv.slice(2);
if (!downloaded || !corpus || !reportPath) throw new Error('usage: select.ts <downloaded> <corpus> <report.json>');
mkdirSync(corpus, { recursive: true });
const DOCX = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
const report: Record<string, unknown>[] = [];
for (const file of readdirSync(downloaded).filter((name) => name.toLowerCase().endsWith('.docx')).sort()) {
  const row: Record<string, unknown> = { file };
  try {
    const bytes = new Uint8Array(readFileSync(join(downloaded, file)));
    const parts = await unzip(bytes);
    if (![...parts.keys()].some((name) => /^word\/document\d*\.xml$/.test(name))) throw new Error('not a Word document');
    const clean = await acceptTrackedChanges(bytes);
    row.words = [...(await wordsByPart(clean)).values()].reduce((sum, words) => sum + words.filter((w) => w !== '¶').length, 0);
    if ((row.words as number) < 3500) throw new Error('too short');
    const reply = await fetch('http://127.0.0.1:8791/convert', { method: 'POST', headers: { 'Content-Type': DOCX }, body: new Blob([clean.slice().buffer]) });
    if (!reply.ok) throw new Error(`could not be laid out (${reply.status})`);
    row.pages = (await readPdf(new Uint8Array(await reply.arrayBuffer()))).pages;
    if ((row.pages as number) < 15) throw new Error('fewer than 15 pages');
    row.edits = (await mutate(clean, 1)).edits.length;
    if ((row.edits as number) < 150) throw new Error('cannot take 150 edits');
    copyFileSync(join(downloaded, file), join(corpus, file));
    row.kept = true;
  } catch (error) {
    row.kept = false;
    row.reason = String(error).replace(/^Error: /, '').slice(0, 120);
  }
  report.push(row);
  console.log(`${row.kept ? 'KEEP' : 'drop'} ${file.slice(0, 60).padEnd(60)} ${row.words ?? '-'} words, ${row.pages ?? '-'} pages, ${row.edits ?? '-'} edits ${row.reason ?? ''}`);
}
writeFileSync(reportPath, JSON.stringify(report, null, 1));
console.log(`kept ${report.filter((r) => r.kept).length} of ${report.length}`);
