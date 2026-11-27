// pdf-dump.ts <redline.pdf> [text to find, without spaces] [lines of context]: prints the PDF's text line by line, in
// the order the file draws it, with where each line sits (page, height, left edge), its colour and whether the file
// marks it as page furniture. With a text to find: only the lines around each place it occurs.
import { readFileSync } from 'node:fs';
import { readPdf } from '../../../src/checker/pdf';
const [path, needle, contextArg] = process.argv.slice(2);
const pdf = await readPdf(new Uint8Array(readFileSync(path!)));
interface Line { page: number; y: number; x: number; text: string; colours: Set<string>; artifact: boolean; first: number }
const lines: Line[] = [];
pdf.chars.forEach((c, index) => {
  const last = lines[lines.length - 1];
  if (last && last.page === c.page && Math.abs(last.y - c.y) < 0.6 * Math.max(c.size, 4) && last.artifact === c.artifact) {
    last.text += c.c;
    last.colours.add(c.color);
  } else lines.push({ page: c.page, y: c.y, x: c.x, text: c.c, colours: new Set([c.color]), artifact: c.artifact, first: index });
});
const show = (line: Line): string => `p${line.page} y${line.y.toFixed(0).padStart(4)} x${line.x.toFixed(0).padStart(4)} ${line.artifact ? 'A' : ' '} ${[...line.colours].filter((c) => c !== '000000').join(',').padEnd(6)} ${line.text.slice(0, 110)}`;
if (!needle) for (const line of lines) console.log(show(line));
else {
  const context = Number(contextArg ?? 6);
  const all = pdf.chars.map((c) => c.c).join('');
  for (let at = all.indexOf(needle); at >= 0; at = all.indexOf(needle, at + 1)) {
    const index = lines.findIndex((line, i) => line.first <= at && (lines[i + 1]?.first ?? Infinity) > at);
    console.log(`--- "${needle}" at character ${at}`);
    for (const line of lines.slice(Math.max(0, index - context), index + context + 1)) console.log(show(line));
  }
}
console.log(`${pdf.pages} pages, ${pdf.chars.length} characters, ${pdf.rules.length} thin lines`);
