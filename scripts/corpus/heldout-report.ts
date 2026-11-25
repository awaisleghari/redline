/**
 * Stage 7 tooling: the held-out test's summary. For each results file given (one per kind of upload), how many
 * documents had both Word files released, and how many of those mark close to what really changed; then the
 * same by source, and the reasons for withholding.
 * Usage: npx tsx scripts/corpus/heldout-report.ts <label>=<results.jsonl>[,<more.jsonl>] ...
 */
import { readFileSync } from 'node:fs';

interface Row {
  file: string; error?: string; trueDeleted?: number; trueInserted?: number; trackedPassed?: boolean; staticPassed?: boolean; staticError?: string;
  marked?: { inserted: number; deleted: number; movedFrom: number; movedTo: number };
  trackedFailures?: { id: string; message: string }[]; staticFailures?: { id: string; message: string }[];
}
const CLOSE = 1.5;
const pct = (n: number, of: number): string => `${n} of ${of} (${of ? ((100 * n) / of).toFixed(1) : '0'}%)`;
const released = (r: Row): boolean => !!r.trackedPassed && !!r.staticPassed;
const ratio = (r: Row): number => {
  const truth = (r.trueDeleted ?? 0) + (r.trueInserted ?? 0);
  const m = r.marked;
  return truth && m ? (m.inserted + m.deleted + m.movedFrom + m.movedTo) / truth : Infinity;
};
const sourceOf = (file: string): string => file.split('__')[0]!;

for (const argument of process.argv.slice(2)) {
  const [label, paths] = argument.split('=');
  const rows: Row[] = paths!.split(',').flatMap((path) => readFileSync(path, 'utf8').split('\n').filter(Boolean).map((line) => JSON.parse(line) as Row));
  const ok = rows.filter(released);
  console.log(`\n== ${label}: ${rows.length} documents`);
  console.log(`  both Word files released:                       ${pct(ok.length, rows.length)}`);
  console.log(`  released and marking within ${CLOSE} times the real change: ${pct(ok.filter((r) => ratio(r) <= CLOSE).length, rows.length)}`);
  console.log(`  tracked-changes file alone released:            ${pct(rows.filter((r) => r.trackedPassed).length, rows.length)}`);
  console.log(`  engine failed outright:                         ${pct(rows.filter((r) => r.error).length, rows.length)}`);
  const ratios = ok.map(ratio).sort((a, b) => a - b);
  if (ratios.length) console.log(`  marking, released documents: median ${ratios[Math.floor(ratios.length / 2)]!.toFixed(2)}, worst ${ratios[ratios.length - 1]!.toFixed(1)}`);
  console.log('  by source (released / run):');
  const sources = [...new Set(rows.map((r) => sourceOf(r.file)))].sort();
  for (const source of sources) {
    const mine = rows.filter((r) => sourceOf(r.file) === source);
    console.log(`    ${source.padEnd(34)} ${mine.filter(released).length} / ${mine.length}`);
  }
  const reasons = new Map<string, number>();
  for (const r of rows.filter((x) => !released(x))) {
    const first = [...(r.trackedFailures ?? []), ...(r.staticFailures ?? [])][0];
    const reason = r.error ? `engine: ${r.error.slice(0, 70)}` : r.staticError ? `static redline could not be made: ${r.staticError.slice(0, 60)}` : first ? `${first.id}: ${first.message.slice(0, 80)}` : 'unknown';
    reasons.set(reason, (reasons.get(reason) ?? 0) + 1);
  }
  console.log('  why withheld (first failing check), most common first:');
  for (const [reason, count] of [...reasons].sort((a, b) => b[1] - a[1])) console.log(`    ${String(count).padStart(3)}  ${reason}`);
}
