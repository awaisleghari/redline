/**
 * Stage 7 tooling: for a folder written by one.ts, print the clause numbers the Checker reads around the
 * first paragraph whose number differs between the modified document and an output.
 * Usage: npx tsx scripts/corpus/numbers.ts <folder> <tracked|static>
 */
import { readFileSync } from 'node:fs';
import { DocxPackage } from '../../src/checker/package';
import { readStories } from '../../src/checker/stories';
const [dir, which] = process.argv.slice(2);
const open = async (name: string) => DocxPackage.open(new Uint8Array(readFileSync(`${dir}/${name}.docx`)));
const body = (r: ReturnType<typeof readStories>) => r.stories.find((s) => s.key === 'body')!.details;
const modified = body(readStories(await open('modified'), 'accepted'));
const outNew = body(readStories(await open(which!), 'accepted', which === 'static' ? 'static' : 'tracked'));
const outAll = body(readStories(await open(which!), 'all', which === 'static' ? 'static' : 'tracked'));
let at = 0;
while (at < modified.length && modified[at]!.number === outNew[at]?.number) at++;
console.log('first difference at', at, 'of', modified.length, outNew.length);
const show = (label: string, list: typeof modified, from: number, to: number) => { console.log(`== ${label}`); for (let i = Math.max(0, from); i <= to && i < list.length; i++) console.log(`  ${i} [${list[i]!.number}]${list[i]!.listed ? ' L' : ''}${list[i]!.removed ? ' REMOVED' : ''} ${JSON.stringify(list[i]!.text.slice(0, 46))}`); };
show('modified', modified, at - 6, at + 1);
show(`${which} as accepted`, outNew, at - 6, at + 1);
const text = modified[at - 1]?.text.slice(0, 30) ?? '';
const j = outAll.findIndex((p) => p.text.startsWith(text) && text.length > 5);
show(`${which} with everything shown`, outAll, j - 8, j + 3);
