/**
 * Schema validation of every generated output (check 7 in PLAN.md), using
 * Microsoft's Open XML SDK validator through the small .NET tool in
 * scripts/validate-docx/. Needs the .NET 10 SDK (`dotnet`).
 *
 * Writes, for every generated pair and for the 100-page contract, the two
 * uploads, the tracked-changes output and the static redline to a temporary folder, validates them
 * all, and fails if any OUTPUT is invalid. Uploads are validated too, so that
 * a problem inherited from an upload can be told apart from one the engine
 * introduced; an invalid upload is reported but is not our failure.
 *
 * Usage: npm run test:schema
 */
import { spawnSync } from 'node:child_process';
import { mkdtempSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { compareDocuments, initialize } from 'docxodus/core';
import { unzip, zip } from '../src/common/zip';
import { makeStaticRedline } from '../src/producer/make-static-redline';
import { REVISION_AUTHOR } from '../src/producer/protocol';
import { buildLargeContract } from '../tests/fixtures/large';
import { ALL_PAIR_SPECS, buildPair } from '../tests/fixtures/pairs';

const root = join(dirname(fileURLToPath(import.meta.url)), '..');
const folder = mkdtempSync(join(tmpdir(), 'redline-schema-'));

try {
  await initialize();
  const pairs: { id: string; original: Uint8Array; modified: Uint8Array }[] = [];
  for (const spec of ALL_PAIR_SPECS) pairs.push({ id: spec.id, ...(await buildPair(spec)) });
  pairs.push({ id: 'large-contract', ...(await buildLargeContract()) });
  for (const pair of pairs) {
    writeFileSync(join(folder, `${pair.id}.upload-original.docx`), pair.original);
    writeFileSync(join(folder, `${pair.id}.upload-modified.docx`), pair.modified);
    const tracked = await compareDocuments(pair.original, pair.modified, { authorName: REVISION_AUTHOR });
    writeFileSync(join(folder, `${pair.id}.output-tracked.docx`), tracked);
    const staticRedline = await makeStaticRedline(tracked, { originalName: 'original.docx', modifiedName: 'modified.docx', comparedAt: '23 November 2026 at 14:05 (UTC+00:00)' });
    writeFileSync(join(folder, `${pair.id}.output-static.docx`), staticRedline.docx);
  }

  // The validator must have teeth: one deliberately schema-invalid file (an
  // element Word does not define, inside a run) has to be reported as invalid.
  const sample = await unzip(await compareDocuments(pairs[0]!.original, pairs[0]!.modified, { authorName: REVISION_AUTHOR }));
  const documentXml = new TextDecoder().decode(sample.get('word/document.xml')!);
  if (!documentXml.includes('<w:r><w:t')) throw new Error('schema-check: could not build the invalid control file');
  sample.set('word/document.xml', new TextEncoder().encode(documentXml.replace('<w:r><w:t', '<w:r><w:notARealElement/><w:t')));
  writeFileSync(join(folder, 'zz-control.deliberately-invalid.docx'), await zip(sample));

  const run = spawnSync('dotnet', ['run', '-c', 'Release', '--project', join(root, 'scripts', 'validate-docx'), '--', folder], {
    encoding: 'utf8',
    env: { ...process.env, DOTNET_CLI_TELEMETRY_OPTOUT: '1', DOTNET_NOLOGO: '1' },
  });
  if (run.error) {
    console.error(`schema-check: could not run "dotnet" (${run.error.message}). Install the .NET 10 SDK: apt-get install -y dotnet-sdk-10.0`);
    process.exit(2);
  }
  const lines = run.stdout.split('\n');
  const verdicts = lines.filter((line) => /^(VALID  |INVALID) /.test(line));
  const control = verdicts.filter((line) => line.includes('zz-control.'));
  if (verdicts.length !== pairs.length * 4 + 1 || control.length !== 1) {
    console.error(run.stdout, run.stderr);
    console.error(`schema-check: expected ${pairs.length * 4 + 1} verdicts, got ${verdicts.length}`);
    process.exit(2);
  }
  if (!control[0]!.startsWith('INVALID')) {
    console.error('schema-check: FAILED: the validator accepted a deliberately invalid file, so its verdicts cannot be trusted.');
    process.exit(1);
  }
  verdicts.splice(verdicts.indexOf(control[0]!), 1);
  const invalidOutputs = verdicts.filter((line) => line.startsWith('INVALID') && line.includes('.output-'));
  const invalidUploads = verdicts.filter((line) => line.startsWith('INVALID') && line.includes('.upload-'));
  if (invalidOutputs.length > 0 || invalidUploads.length > 0) console.log(run.stdout);
  console.log(`schema-check: ${verdicts.length} files validated with the Open XML SDK: ${pairs.length * 2} outputs (tracked changes and static redline for each pair), ${pairs.length * 2} uploads.`);
  if (invalidUploads.length > 0) console.log(`schema-check: note: ${invalidUploads.length} generated upload(s) are themselves invalid (a test-fixture matter, not an output failure).`);
  if (invalidOutputs.length > 0) {
    console.error(`schema-check: FAILED: ${invalidOutputs.length} output(s) do not validate.`);
    process.exit(1);
  }
  console.log('schema-check: every output validates against the Office Open XML schema (and the deliberately invalid control file was rejected).');
} finally {
  rmSync(folder, { recursive: true, force: true });
}
