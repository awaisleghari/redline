// pdf-run.ts <folder> <results.jsonl>: for each contract, ordinary edits -> both Word files -> PDF (local converter) -> check 9.
import { appendFileSync, existsSync, readFileSync, readdirSync } from 'node:fs';
import { join } from 'node:path';
import { initialize } from 'docxodus/core';
import { checkPdf, checkStaticRedline } from '../../../src/checker';
import { acceptTrackedChanges } from '../../../src/producer/accept-changes';
import { makeStaticRedline } from '../../../src/producer/make-static-redline';
import { produceTrackedChanges } from '../../../src/producer/produce';
import { mutate } from '../mutate';
const [folder, out] = process.argv.slice(2);
await initialize();
const done = new Set(existsSync(out!) ? readFileSync(out!, 'utf8').split('\n').filter(Boolean).map((l) => JSON.parse(l).file) : []);
const COVER = { originalName: 'original.docx', modifiedName: 'modified.docx', comparedAt: '26 November 2026 at 12:00 (UTC+00:00)' };
for (const file of readdirSync(folder!).filter((n) => /\.docx$/i.test(n)).sort()) {
  if (done.has(file)) continue;
  const row: Record<string, unknown> = { file };
  try {
    const seed = [...file].reduce((h, c) => (Math.imul(h, 31) + c.charCodeAt(0)) >>> 0, 7);
    const original = new Uint8Array(readFileSync(join(folder!, file)));
    const modified = (await mutate(await acceptTrackedChanges(original), seed)).docx;
    const tracked = await produceTrackedChanges(original, modified);
    const hadChanges = /<w:(ins|del|moveFrom|moveTo) /.test(new TextDecoder().decode(original.slice(0, 0))) ;
    void hadChanges;
    const made = await makeStaticRedline(tracked, original, COVER as never);
    const staticReport = await checkStaticRedline({ original, modified, output: made.docx, ...COVER });
    row.staticFailures = staticReport.checks.filter((c) => !c.passed).map((c) => c.id);
    row.staticBytes = made.docx.byteLength;
    const t0 = Date.now();
    const reply = await fetch('http://127.0.0.1:8099/convert', { method: 'POST', headers: { 'Content-Type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' }, body: new Blob([made.docx.slice().buffer]) });
    row.convertStatus = reply.status; row.convertSeconds = (Date.now() - t0) / 1000;
    if (!reply.ok) { row.convertMessage = (await reply.text()).slice(0, 200); }
    else {
      const pdf = new Uint8Array(await reply.arrayBuffer());
      row.pdfBytes = pdf.byteLength;
      const t1 = Date.now();
      const report = await checkPdf({ staticRedline: made.docx, pdf });
      row.checkSeconds = (Date.now() - t1) / 1000;
      row.pdfPassed = report.passed;
      row.pdfFailures = report.checks.filter((c) => !c.passed).map((c) => ({ id: c.id, message: c.failure?.message, where: c.failure?.where, expected: String(c.failure?.expected ?? '').slice(0, 80), found: String(c.failure?.found ?? '').slice(0, 80) }));
    }
  } catch (error) { row.error = String(error instanceof Error ? error.message : error).slice(0, 300); }
  appendFileSync(out!, JSON.stringify(row) + '\n');
  console.log(file.slice(0, 50), row.convertStatus, row.pdfPassed, row.error ?? '');
}
