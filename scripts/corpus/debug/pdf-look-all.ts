// pdf-look-all.ts <folder of pdf-one.ts out-dirs>: the PDF check on every saved pair, one line each, then a count by reason.
import { existsSync, readdirSync, readFileSync } from 'node:fs';
import { checkPdf } from '../../../src/checker';
const parent = process.argv[2]!;
const only = process.argv[3];
const reasons = new Map<string, number>();
let passed = 0;
let total = 0;
for (const name of readdirSync(parent).sort()) {
  const dir = `${parent}/${name}`;
  if (!existsSync(`${dir}/redline.pdf`) || (only && !name.includes(only))) continue;
  total++;
  const started = Date.now();
  const report = await checkPdf({ staticRedline: new Uint8Array(readFileSync(`${dir}/static.docx`)), pdf: new Uint8Array(readFileSync(`${dir}/redline.pdf`)) });
  const failure = report.checks.find((c) => !c.passed)?.failure;
  if (report.passed) passed++;
  else reasons.set(failure?.message ?? '?', (reasons.get(failure?.message ?? '?') ?? 0) + 1);
  console.log(`${report.passed ? 'PASS' : 'FAIL'} ${name.slice(0, 40).padEnd(40)} ${((Date.now() - started) / 1000).toFixed(0)}s ${report.passed ? '' : `${failure?.message?.slice(0, 70)} | ${failure?.where} | exp ${String(failure?.expected).slice(0, 90)} | got ${String(failure?.found).slice(0, 60)} ${String(failure?.found).replace(/^.*\(page/, '(page').slice(0, 110)}`}`);
}
console.log(`\n${passed} of ${total} passed`);
for (const [reason, count] of [...reasons].sort((a, b) => b[1] - a[1])) console.log(`${count}  ${reason}`);
