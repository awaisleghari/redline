/**
 * Early checks, step 3: does the engine's DOCX-to-HTML conversion show (a)
 * tracked changes and (b) the static redline's formatting (see the style spec
 * in PLAN.md) faithfully enough for the in-page viewer?
 *
 * Writes HTML files to the folder and prints what markup each change received.
 *
 * Usage: npx tsx scripts/early-checks/viewer.ts <folder used by prepare.ts>
 */
import { Document, Packer, Paragraph, TextRun, UnderlineType } from 'docx';
import { convertDocxToHtml, initialize } from 'docxodus/core';
import { readFileSync, writeFileSync } from 'node:fs';

const folder = process.argv[2];
if (!folder) throw new Error('usage: viewer.ts <folder>');
await initialize();

// (a) Tracked changes.
for (const id of ['one-word-change', 'moved-paragraph', 'inserted-paragraph', 'deleted-paragraph', 'header-change', 'footnote-change', 'numbered-list-change', 'table-cell-change', 'x-format-change']) {
  // Tracked-change rendering is off by default; headers, footers and notes are too.
  const html = await convertDocxToHtml(readFileSync(`${folder}/${id}-npm.docx`), {
    renderTrackedChanges: true,
    renderHeadersAndFooters: true,
    renderFootnotesAndEndnotes: true,
  });
  writeFileSync(`${folder}/${id}-tracked.html`, html);
  const tags = [...html.matchAll(/<(ins|del)\b[^>]*>[^<]*/g)].map((m) => m[0]);
  console.log(`\n[tracked] ${id}: ${tags.length} <ins>/<del> elements; format-change markers: ${(html.match(/class="[^"]*rev-format[^"]*"|data-format-change|format-change/g) ?? []).length}`);
  for (const tag of tags.slice(0, 4)) console.log('   ', tag.slice(0, 220));
}

// (b) Static redline formatting, exactly as the style spec defines it.
const staticRedline = new Document({
  sections: [
    {
      children: [
        new Paragraph({
          children: [
            new TextRun('Unchanged text, then '),
            new TextRun({ text: 'inserted text', color: '0000FF', underline: { type: UnderlineType.DOUBLE } }),
            new TextRun(', then '),
            new TextRun({ text: 'deleted text', color: 'FF0000', strike: true }),
            new TextRun(', then '),
            new TextRun({ text: 'moved-from text', color: '008000', doubleStrike: true }),
            new TextRun(', then '),
            new TextRun({ text: 'moved-to text', color: '008000', underline: { type: UnderlineType.DOUBLE } }),
            new TextRun('.'),
          ],
        }),
      ],
    },
  ],
});
const html = await convertDocxToHtml(new Uint8Array(await Packer.toBuffer(staticRedline)));
writeFileSync(`${folder}/static-redline.html`, html);
console.log('\n[static] spans and the CSS rules they use:');
for (const label of ['inserted text', 'deleted text', 'moved-from text', 'moved-to text']) {
  const span = html.match(new RegExp(`<span[^>]*>${label}</span>`))?.[0] ?? '(not found as a single span)';
  const cls = span.match(/class="([^"]+)"/)?.[1];
  const rule = cls ? cls.split(/\s+/).map((c) => html.match(new RegExp(`[^}]*\\.${c}\\s*\\{[^}]*\\}`))?.[0]?.trim().replace(/\s+/g, ' ')).join(' ') : '';
  console.log(`  ${label}: ${span}\n      ${rule}`);
}
