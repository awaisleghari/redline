// pdf-one.ts <contract.docx> <out-dir>: ordinary edits -> both Word files -> PDF (converter on port 8099), all saved
// in <out-dir> (static.docx, redline.pdf), so that the PDF check can be studied again and again without redoing the rest.
import { mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { basename } from 'node:path';
import { initialize } from 'docxodus/core';
import { acceptTrackedChanges } from '../../../src/producer/accept-changes';
import { makeStaticRedline } from '../../../src/producer/make-static-redline';
import { produceTrackedChanges } from '../../../src/producer/produce';
import { mutate } from '../mutate';
const [path, out] = process.argv.slice(2);
await initialize();
const file = basename(path!);
const seed = [...file].reduce((h, c) => (Math.imul(h, 31) + c.charCodeAt(0)) >>> 0, 7);
const original = new Uint8Array(readFileSync(path!));
const modified = (await mutate(await acceptTrackedChanges(original), seed)).docx;
const tracked = await produceTrackedChanges(original, modified);
const made = await makeStaticRedline(tracked, original, { originalName: 'original.docx', modifiedName: 'modified.docx', comparedAt: '26 November 2026 at 12:00 (UTC+00:00)' });
mkdirSync(out!, { recursive: true });
writeFileSync(`${out}/static.docx`, made.docx);
const reply = await fetch(process.env.PDF_URL ?? 'http://127.0.0.1:8099/convert', { method: 'POST', headers: { 'Content-Type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' }, body: new Blob([made.docx.slice().buffer]) });
if (!reply.ok) throw new Error(`converter: status ${reply.status}`);
writeFileSync(`${out}/redline.pdf`, new Uint8Array(await reply.arrayBuffer()));
console.log(file, 'saved to', out);
