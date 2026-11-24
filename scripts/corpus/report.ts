/**
 * Stage 7 tooling: turns the runner's results into a plain summary.
 * Usage: npx tsx scripts/corpus/report.ts <results.jsonl> [more.jsonl ...]
 */
import { readFileSync } from 'node:fs';

interface Row {
  file: string; error?: string; edits?: number; tally?: Record<string, number>; words?: number; trueDeleted?: number; trueInserted?: number;
  marked?: { inserted: number; deleted: number; movedFrom: number; movedTo: number; blocks: number; formattingRecords: number };
  counts?: { insertions: number; deletions: number; moves: number; formattingChanges: number };
  trackedPassed?: boolean; staticPassed?: boolean; staticError?: string; engineSeconds?: number; totalSeconds?: number;
  trackedFailures?: { id: string; message: string; where?: string }[]; staticFailures?: { id: string; message: string; where?: string }[];
}
const rows: Row[] = process.argv.slice(2).flatMap((path) => readFileSync(path, 'utf8').split('\n').filter(Boolean).map((line) => JSON.parse(line) as Row));
const n = rows.length;
const pct = (count: number, of = n): string => `${count} of ${of} (${of ? Math.round((100 * count) / of) : 0}%)`;
const median = (values: number[]): number => [...values].sort((a, b) => a - b)[Math.floor(values.length / 2)] ?? 0;
const sum = (values: number[]): number => values.reduce((a, b) => a + b, 0);

console.log(`Documents run: ${n}`);
console.log(`  words per document: median ${median(rows.map((r) => r.words ?? 0))}, total ${sum(rows.map((r) => r.words ?? 0))}`);
console.log(`  edits applied: median ${median(rows.map((r) => r.edits ?? 0))}, total ${sum(rows.map((r) => r.edits ?? 0))}; documents with fewer than 150: ${rows.filter((r) => (r.edits ?? 0) < 150).length}`);
const kinds: Record<string, number> = {};
for (const r of rows) for (const [k, v] of Object.entries(r.tally ?? {})) kinds[k] = (kinds[k] ?? 0) + v;
console.log(`  by kind: ${Object.entries(kinds).sort((a, b) => b[1] - a[1]).map(([k, v]) => `${k} ${v}`).join(', ')}`);

const crashed = rows.filter((r) => r.error);
const both = rows.filter((r) => r.trackedPassed && r.staticPassed);
console.log(`\nOutcome`);
console.log(`  engine failed outright:            ${pct(crashed.length)}`);
console.log(`  tracked-changes file passed:       ${pct(rows.filter((r) => r.trackedPassed).length)}`);
console.log(`  static redline passed:             ${pct(rows.filter((r) => r.staticPassed).length)}`);
console.log(`  both Word files released:          ${pct(both.length)}`);

const reasons = new Map<string, string[]>();
for (const r of rows) {
  const first = r.error ? `engine: ${r.error.replace(/\d+/g, 'N').slice(0, 90)}` : r.trackedFailures?.[0] ? `tracked/${r.trackedFailures[0].id}: ${r.trackedFailures[0].message.slice(0, 90)}` : r.staticError ? `static maker: ${r.staticError.slice(0, 90)}` : r.staticFailures?.[0] ? `static/${r.staticFailures[0].id}: ${r.staticFailures[0].message.slice(0, 90)}` : undefined;
  if (first) reasons.set(first, [...(reasons.get(first) ?? []), r.file]);
}
console.log(`\nWhy documents were withheld (first failing check of each), most common first`);
for (const [reason, files] of [...reasons].sort((a, b) => b[1].length - a[1].length)) console.log(`  ${String(files.length).padStart(3)}  ${reason}\n         e.g. ${files[0]}`);

const scored = rows.filter((r) => r.marked && r.trueDeleted !== undefined);
const ratio = (r: Row): number => (r.marked!.inserted + r.marked!.deleted + r.marked!.movedFrom + r.marked!.movedTo) / Math.max(1, r.trueDeleted! + r.trueInserted!);
const band = (low: number, high: number, set: Row[]): number => set.filter((r) => ratio(r) >= low && ratio(r) < high).length;
for (const [label, set] of [['all documents the engine produced a redline for', scored], ['documents whose Word files were both released', both.filter((r) => r.marked)]] as const) {
  console.log(`\nHow much the redline marks, against the words that really changed (1.0 = exactly the real changes): ${label}`);
  console.log(`  up to 1.1 times:   ${pct(band(0, 1.1, set), set.length)}`);
  console.log(`  1.1 to 1.5 times:  ${pct(band(1.1, 1.5, set), set.length)}`);
  console.log(`  1.5 to 3 times:    ${pct(band(1.5, 3, set), set.length)}`);
  console.log(`  3 to 10 times:     ${pct(band(3, 10, set), set.length)}`);
  console.log(`  over 10 times:     ${pct(band(10, Infinity, set), set.length)}`);
  console.log(`  median ${median(set.map(ratio)).toFixed(2)}, worst ${Math.max(0, ...set.map(ratio)).toFixed(1)}`);
}
console.log(`\nWorst ten for over-marking`);
for (const r of [...scored].sort((a, b) => ratio(b) - ratio(a)).slice(0, 10)) console.log(`  ${ratio(r).toFixed(1).padStart(6)} times  really changed ${r.trueDeleted! + r.trueInserted!} words, marked ${r.marked!.inserted + r.marked!.deleted + r.marked!.movedFrom + r.marked!.movedTo}  ${r.file.slice(0, 70)}`);

const moved = sum(rows.map((r) => r.tally?.['clause moved'] ?? 0));
const formatted = sum(rows.map((r) => r.tally?.['formatting changed'] ?? 0));
console.log(`\nMoved clauses: ${moved} made; shown as moves in released redlines: ${sum(both.map((r) => r.counts?.moves ?? 0))} (of ${sum(both.map((r) => r.tally?.['clause moved'] ?? 0))} made in those documents)`);
console.log(`Formatting changes: ${formatted} made; counted in released redlines: ${sum(both.map((r) => r.counts?.formattingChanges ?? 0))} (of ${sum(both.map((r) => r.tally?.['formatting changed'] ?? 0))} made in those documents)`);
console.log(`\nTime per document (compare, both Word files, all checks): median ${median(rows.map((r) => r.totalSeconds ?? 0))} s, slowest ${Math.max(0, ...rows.map((r) => r.totalSeconds ?? 0))} s`);
