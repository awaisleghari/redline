/**
 * Writes sample files for the owner to open in real Word ("owner checks" in
 * CLAUDE.md). For each chosen generated pair it writes the original, the
 * modified and the tracked-changes output, and it writes an output only if the
 * independent Checker passed it, exactly as the product would.
 *
 * Usage: npx tsx scripts/owner-samples.ts <output folder> [pair id ...]
 * The folder is for looking at, not for committing (owner-checks/ is git-ignored).
 */
import { mkdirSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { compareDocuments, initialize } from 'docxodus/core';
import { runComparison } from '../src/app/flow';
import { checkTrackedChanges } from '../src/checker';
import { REVISION_AUTHOR } from '../src/producer/protocol';
import { ALL_PAIR_SPECS, buildPair, pairById } from '../tests/fixtures/pairs';

const [folder, ...ids] = process.argv.slice(2);
if (!folder) throw new Error('usage: owner-samples.ts <output folder> [pair id ...]');
mkdirSync(folder, { recursive: true });
await initialize();

const specs = ids.length > 0 ? ids.map(pairById) : ALL_PAIR_SPECS;
for (const spec of specs) {
  const pair = await buildPair(spec);
  const result = await runComparison(pair.original, pair.modified, {
    produce: async (original, modified) => {
      const started = performance.now();
      const trackedChanges = await compareDocuments(original, modified, { authorName: REVISION_AUTHOR });
      return { trackedChanges, engineMs: performance.now() - started };
    },
    check: checkTrackedChanges,
  });
  writeFileSync(join(folder, `${spec.id} - 1 original.docx`), pair.original);
  writeFileSync(join(folder, `${spec.id} - 2 modified.docx`), pair.modified);
  if (result.trackedChanges) {
    writeFileSync(join(folder, `${spec.id} - 3 redline (tracked changes).docx`), result.trackedChanges);
    console.log(`${spec.id}: written (every check passed)`);
  } else {
    console.log(`${spec.id}: output WITHHELD: ${result.report.checks.find((c) => !c.passed)?.failure?.message}`);
    process.exitCode = 1;
  }
}
