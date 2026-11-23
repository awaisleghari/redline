/**
 * Early checks, step 1 (see PLAN.md, "Early checks"): write every test pair,
 * a few extra pairs and the 100-page contract to a folder, together with the
 * redline the npm (WebAssembly) engine produces for each, and time the engine
 * in Node.
 *
 * Usage: npx tsx scripts/early-checks/prepare.ts <output folder>
 */
import { Document, Packer, Paragraph, Table, TableCell, TableRow, TextRun, WidthType } from 'docx';
import { compareDocuments, initialize } from 'docxodus/core';
import { mkdirSync, writeFileSync } from 'node:fs';
import { PAIR_SPECS, buildPair } from '../../tests/fixtures/pairs';
import { buildLargeContract } from '../../tests/fixtures/large';

const out = process.argv[2];
if (!out) throw new Error('usage: prepare.ts <output folder>');
mkdirSync(out, { recursive: true });

async function pack(children: (Paragraph | Table)[]): Promise<Uint8Array> {
  return new Uint8Array(await Packer.toBuffer(new Document({ sections: [{ children }] })));
}
const para = (...runs: TextRun[]): Paragraph => new Paragraph({ children: runs });
const table = (rows: string[][]): Table =>
  new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: rows.map((cells) => new TableRow({ children: cells.map((text) => new TableCell({ children: [para(new TextRun(text))] })) })),
  });

const pairs: { id: string; original: Uint8Array; modified: Uint8Array }[] = [];
for (const spec of PAIR_SPECS) {
  const pair = await buildPair(spec);
  pairs.push({ id: spec.id, original: pair.original, modified: pair.modified });
}
// Extra pairs for things the nine standard pairs do not exercise.
pairs.push({
  id: 'x-format-change',
  original: await pack([para(new TextRun('The Supplier shall provide the Services with reasonable skill and care.'))]),
  modified: await pack([
    para(new TextRun('The Supplier shall provide the Services with '), new TextRun({ text: 'reasonable skill', bold: true }), new TextRun(' and care.')),
  ]),
});
pairs.push({
  id: 'x-table-row-added',
  original: await pack([para(new TextRun('Fees:')), table([['Service', 'Fee'], ['Hosting', 'USD 4,000']]), para(new TextRun('End.'))]),
  modified: await pack([para(new TextRun('Fees:')), table([['Service', 'Fee'], ['Hosting', 'USD 4,000'], ['Support', 'USD 1,500']]), para(new TextRun('End.'))]),
});
const large = await buildLargeContract();
console.log(`large contract: ${large.paragraphs} paragraphs, ${large.words} words, edits ${JSON.stringify(large.edits)}`);
pairs.push({ id: 'x-large-contract', original: large.original, modified: large.modified });

let started = performance.now();
await initialize();
console.log(`engine start-up (Node): ${Math.round(performance.now() - started)} ms`);
for (const pair of pairs) {
  writeFileSync(`${out}/${pair.id}-original.docx`, pair.original);
  writeFileSync(`${out}/${pair.id}-modified.docx`, pair.modified);
  started = performance.now();
  const redline = await compareDocuments(pair.original, pair.modified, { authorName: 'Redline' });
  console.log(`${pair.id}: ${Math.round(performance.now() - started)} ms (Node, WebAssembly)`);
  writeFileSync(`${out}/${pair.id}-npm.docx`, redline);
}
