// pdf-look.ts <out-dir of pdf-one.ts>: runs the PDF check on the saved pair and prints what failed.
import { readFileSync } from 'node:fs';
import { checkPdf } from '../../../src/checker';
const dir = process.argv[2]!;
const report = await checkPdf({ staticRedline: new Uint8Array(readFileSync(`${dir}/static.docx`)), pdf: new Uint8Array(readFileSync(`${dir}/redline.pdf`)) });
for (const check of report.checks) {
  console.log(`${check.id}: ${check.passed ? 'passed' : 'FAILED'}${check.note ? ` (${check.note})` : ''}`);
  if (check.failure) console.log(`  ${check.failure.message}\n  where: ${check.failure.where}\n  expected: ${check.failure.expected}\n  found: ${check.failure.found}`);
}
