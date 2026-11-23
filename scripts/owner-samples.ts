/**
 * Writes sample files for the owner to open in real Word ("owner checks" in
 * CLAUDE.md). For each chosen generated pair it writes the original, the
 * modified, and each output that the independent Checker passed, exactly as
 * the product would release them.
 *
 * Usage: npx tsx scripts/owner-samples.ts <output folder> [pair id ...]
 * With PDF_CONVERTER_URL set (e.g. http://127.0.0.1:8791, the converter container), the PDF is written too.
 * The folder is for looking at, not for committing (owner-checks/ is git-ignored).
 */
import { mkdirSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { initialize } from 'docxodus/core';
import { runComparison, runPdf } from '../src/app/flow';
import { formatComparedAt } from '../src/app/format';
import { checkPdf, checkStaticRedline, checkTrackedChanges } from '../src/checker';
import { makeStaticRedline } from '../src/producer/make-static-redline';
import { produceTrackedChanges } from '../src/producer/produce';
import { ALL_PAIR_SPECS, buildPair, pairById } from '../tests/fixtures/pairs';

const [folder, ...ids] = process.argv.slice(2);
if (!folder) throw new Error('usage: owner-samples.ts <output folder> [pair id ...]');
mkdirSync(folder, { recursive: true });
await initialize();

const specs = ids.length > 0 ? ids.map(pairById) : ALL_PAIR_SPECS;
for (const spec of specs) {
  const pair = await buildPair(spec);
  const details = { originalName: `${spec.id} - 1 original.docx`, modifiedName: `${spec.id} - 2 modified.docx`, comparedAt: formatComparedAt(new Date()) };
  const result = await runComparison(pair.original, pair.modified, details, {
    produce: async (original, modified) => {
      const started = performance.now();
      const trackedChanges = await produceTrackedChanges(original, modified);
      return { trackedChanges, engineMs: performance.now() - started };
    },
    makeStatic: async (tracked, original, info) => {
      const made = await makeStaticRedline(tracked, original, info);
      return { staticRedline: made.docx, counts: made.counts };
    },
    checkTracked: checkTrackedChanges,
    checkStatic: checkStaticRedline,
  });
  writeFileSync(join(folder, details.originalName), pair.original);
  writeFileSync(join(folder, details.modifiedName), pair.modified);
  for (const [reviewed, name] of [[result.tracked, 'redline (tracked changes)'], [result.staticRedline, 'redline (static)']] as const) {
    if (reviewed.file) {
      writeFileSync(join(folder, `${spec.id} - ${name}.docx`), reviewed.file);
      console.log(`${spec.id}: ${name} written (every check passed)`);
    } else {
      console.log(`${spec.id}: ${name} WITHHELD: ${reviewed.report.checks.find((c) => !c.passed)?.failure?.message}`);
      process.exitCode = 1;
    }
  }
  const converter = process.env.PDF_CONVERTER_URL;
  if (converter && result.staticRedline.file) {
    const pdf = await runPdf(result.staticRedline.file, {
      convert: async (docx) => {
        const response = await fetch(`${converter}/convert`, { method: 'POST', headers: { 'Content-Type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' }, body: new Blob([docx.slice().buffer]) });
        if (!response.ok) throw new Error(await response.text());
        return new Uint8Array(await response.arrayBuffer());
      },
      check: checkPdf,
    });
    if (pdf.file) {
      writeFileSync(join(folder, `${spec.id} - redline.pdf`), pdf.file);
      console.log(`${spec.id}: PDF written (both checks passed)`);
    } else {
      console.log(`${spec.id}: PDF WITHHELD: ${pdf.report.checks.find((c) => !c.passed)?.failure?.message}`);
      process.exitCode = 1;
    }
  }
}
