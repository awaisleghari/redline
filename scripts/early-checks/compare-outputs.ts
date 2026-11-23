/**
 * Early checks, step 2: compare, part by part, the redline written by the npm
 * (WebAssembly) engine with the one written by the .NET engine for the same
 * pair, and summarise the tracked changes in each.
 *
 * Usage: npx tsx scripts/early-checks/compare-outputs.ts <folder used by prepare.ts>
 */
import JSZip from 'jszip';
import { readFileSync, readdirSync } from 'node:fs';
import { readTrackedChanges } from '../../tests/helpers/tracked-changes';

const folder = process.argv[2];
if (!folder) throw new Error('usage: compare-outputs.ts <folder>');

async function parts(path: string): Promise<Map<string, string>> {
  const zip = await JSZip.loadAsync(readFileSync(path));
  const result = new Map<string, string>();
  for (const name of Object.keys(zip.files).sort()) {
    if (!zip.files[name]!.dir) result.set(name, await zip.file(name)!.async('string'));
  }
  return result;
}

let allSame = true;
for (const file of readdirSync(folder).filter((f) => f.endsWith('-npm.docx')).sort()) {
  const id = file.replace('-npm.docx', '');
  const npm = await parts(`${folder}/${id}-npm.docx`);
  const dotnet = await parts(`${folder}/${id}-dotnet.docx`);
  const names = [...new Set([...npm.keys(), ...dotnet.keys()])].sort();
  const different = names.filter((name) => npm.get(name) !== dotnet.get(name));
  const changes = await readTrackedChanges(readFileSync(`${folder}/${id}-npm.docx`));
  const count = (key: 'insertions' | 'deletions' | 'movedFrom' | 'movedTo'): number =>
    Object.values(changes.parts).reduce((sum, part) => sum + part[key].length, 0);
  console.log(
    `${id}: ${names.length} parts, ${different.length === 0 ? 'IDENTICAL' : `DIFFERENT in ${different.join(', ')}`}` +
      ` | ins ${count('insertions')} del ${count('deletions')} moveFrom ${count('movedFrom')} moveTo ${count('movedTo')} format ${changes.formatChanges}`,
  );
  if (different.length) allSame = false;
}
console.log(allSame ? '\nEvery part of every output is byte-for-byte identical between the two builds.' : '\nSome outputs differ.');
