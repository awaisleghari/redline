// story.ts <docx> <accepted|rejected> <tracked|static> <storyKey> <from> <to>: the Checker's reading of a story, with structure marks.
import { readFileSync } from 'node:fs';
import { DocxPackage } from '../../../src/checker/package';
import { readStories } from '../../../src/checker/stories';
const [file, view, markup, key, from, to] = process.argv.slice(2);
const pkg = await DocxPackage.open(new Uint8Array(readFileSync(file!)));
const reading = readStories(pkg, view as never, markup as never);
for (const story of reading.stories) {
  if (!story.key.startsWith(key!)) continue;
  let n = 0;
  console.log('==', story.key, story.label, story.details.length);
  for (const p of story.details) {
    const mark = /^[⟦]/.test(p.text) && p.text.length < 30;
    if (!mark) n++;
    if (n >= Number(from) && n <= Number(to)) console.log(mark ? '      ' : String(n).padStart(5) + ' ', JSON.stringify(p.text.slice(0, 70)), p.number, p.removed ? 'REMOVED' : '', p.oldNumber);
  }
}
