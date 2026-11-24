/**
 * Stage 7 tooling: for a folder written by one.ts, print the Checker's reading of the main text around the first
 * paragraph where an output differs from the upload it should reproduce.
 * Usage: npx tsx scripts/corpus/texts.ts <folder> <tracked|static> <new|old>
 */
import { readFileSync } from 'node:fs';
import { DocxPackage } from '../../src/checker/package';
import { readStories } from '../../src/checker/stories';
const [dir, which, side] = process.argv.slice(2);
const open = async (name: string) => DocxPackage.open(new Uint8Array(readFileSync(`${dir}/${name}.docx`)));
const body = (r: ReturnType<typeof readStories>) => r.stories.find((s) => s.key === 'body')!.paragraphs;
const want = body(readStories(await open(side === 'new' ? 'modified' : 'original'), 'accepted'));
const got = body(readStories(await open(which!), side === 'new' ? 'accepted' : 'rejected', which === 'static' ? 'static' : 'tracked'));
let at = 0;
while (at < want.length && want[at] === got[at]) at++;
console.log('first difference at', at, '; paragraphs', want.length, got.length);
for (const [label, list] of [['wanted', want], ['found', got]] as const) { console.log(`== ${label}`); for (let i = Math.max(0, at - 3); i <= at + 4 && i < list.length; i++) console.log(`  ${i} ${JSON.stringify(list[i]!.slice(0, 80))}`); }
