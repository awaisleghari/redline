/**
 * Stage 5: the PDF. These tests run the real converter container (Docker) and
 * the Checker's check 9 on its output.
 */
import { execFileSync } from 'node:child_process';
import { initialize } from 'docxodus/core';
import { beforeAll, describe, expect, it } from 'vitest';
import { checkPdf, CHECKS_FOR } from '../../src/checker';
import { unzip, zip } from '../../src/common/zip';
import { makeStaticRedline } from '../../src/producer/make-static-redline';
import { produceTrackedChanges } from '../../src/producer/produce';
import { BROKEN_PDF_CASES } from '../fixtures/broken-pdf';
import { PDF_DOCUMENT_PAIRS, pdfDocumentById } from '../fixtures/pdf-documents';
import { buildLargeContract } from '../fixtures/large';
import { ALL_PAIR_SPECS, CHECK_ONLY_PAIRS, buildDocx, buildPair, pairById } from '../fixtures/pairs';
import { containerLog, containerOfficeProcesses, containerScratchFiles, convertToPdf, CONVERT_URL, DOCX_TYPE, removeContainer, runInContainer, startExtraContainer } from './container';

const COVER = { originalName: 'original.docx', modifiedName: 'modified.docx', comparedAt: '23 November 2026 at 14:05 (UTC+00:00)' };

async function staticRedline(original: Uint8Array, modified: Uint8Array): Promise<Uint8Array> {
  return (await makeStaticRedline(await produceTrackedChanges(original, modified), original, COVER)).docx;
}

beforeAll(async () => {
  await initialize();
});

describe('the PDF of every generated pair passes check 9', () => {
  const pairs = [
    ...ALL_PAIR_SPECS.map((spec) => ({ id: spec.id, build: async () => buildPair(spec) })),
    ...CHECK_ONLY_PAIRS.map((spec) => ({ id: spec.id, build: async () => ({ original: await buildDocx(spec.original), modified: await buildDocx(spec.modified) }) })),
    ...PDF_DOCUMENT_PAIRS.map((spec) => ({ id: spec.id, build: spec.build })),
  ];
  it.each(pairs)('$id', async ({ build }) => {
    const pair = await build();
    const redline = await staticRedline(pair.original, pair.modified);
    const pdf = await convertToPdf(redline);
    const report = await checkPdf({ staticRedline: redline, pdf });
    expect(report.checks.filter((c) => !c.passed)).toEqual([]);
    expect(report.passed).toBe(true);
    expect(report.checks.map((c) => c.id)).toEqual(CHECKS_FOR.pdf);
  });

  it('the 100-page contract', async () => {
    const contract = await buildLargeContract();
    const redline = await staticRedline(contract.original, contract.modified);
    const pdf = await convertToPdf(redline);
    const report = await checkPdf({ staticRedline: redline, pdf });
    expect(report.checks.filter((c) => !c.passed)).toEqual([]);
    expect(report.checks[0]!.note).toMatch(/Compared \d{6} characters on \d+ pages\./);
  });
});

describe('deliberately wrong PDFs are caught', () => {
  it('covers words, numbers, notes, headers, clause numbers and each colour', () => {
    const ids = BROKEN_PDF_CASES.map((c) => c.id);
    for (const id of ['pdf-dropped-word', 'pdf-extra-word', 'pdf-extra-digit', 'pdf-footnote-altered', 'pdf-header-altered', 'pdf-clause-letters-wrong', 'pdf-deletion-in-black', 'pdf-insertion-in-red', 'pdf-move-in-blue']) expect(ids).toContain(id);
    expect(new Set(ids).size).toBe(ids.length);
  });

  it.each(BROKEN_PDF_CASES)('$id ($description)', async (brokenCase) => {
    const pair = await (pdfDocumentById(brokenCase.pairId)?.build() ?? buildPair(pairById(brokenCase.pairId)));
    const genuine = await staticRedline(pair.original, pair.modified);
    // The genuine PDF of the genuine redline passes, so any failure below is caused by the damage alone.
    expect((await checkPdf({ staticRedline: genuine, pdf: await convertToPdf(genuine) })).passed).toBe(true);
    const files = await unzip(genuine);
    brokenCase.damage(files);
    const wrongPdf = await convertToPdf(await zip(files));
    const report = await checkPdf({ staticRedline: genuine, pdf: wrongPdf });
    expect(report.passed).toBe(false);
    const failed = report.checks.filter((c) => !c.passed);
    expect(failed.map((c) => c.id).sort()).toEqual([...brokenCase.caughtBy].sort());
    for (const check of failed) expect(check.failure?.message ?? '').not.toBe('');
  });

  it('a PDF of a different document, a file that is not a PDF, and a damaged PDF', async () => {
    const pair = await buildPair(pairById('one-word-change'));
    const other = await buildPair(pairById('moved-paragraph'));
    const genuine = await staticRedline(pair.original, pair.modified);
    const genuinePdf = await convertToPdf(genuine);
    const otherPdf = await convertToPdf(await staticRedline(other.original, other.modified));
    expect((await checkPdf({ staticRedline: genuine, pdf: otherPdf })).passed).toBe(false);
    expect((await checkPdf({ staticRedline: genuine, pdf: new TextEncoder().encode('not a pdf at all') })).passed).toBe(false);
    expect((await checkPdf({ staticRedline: genuine, pdf: genuinePdf.slice(0, Math.floor(genuinePdf.length / 3)) })).passed).toBe(false);
    expect((await checkPdf({ staticRedline: new TextEncoder().encode('not a docx'), pdf: genuinePdf })).passed).toBe(false);
    // And the genuine PDF of the genuine redline does pass.
    expect((await checkPdf({ staticRedline: genuine, pdf: genuinePdf })).passed).toBe(true);
  });
});

describe('the converter keeps nothing and says nothing', () => {
  const post = (body: BodyInit | null, headers: Record<string, string>, path = '/convert') => fetch(CONVERT_URL.replace('/convert', path), { method: 'POST', headers, body });

  it('refuses what is not a Word document, plainly and without converting', async () => {
    expect((await post('plain text', { 'Content-Type': 'text/plain' })).status).toBe(415);
    expect((await post('not a zip', { 'Content-Type': DOCX_TYPE })).status).toBe(400);
    expect((await post('x', { 'Content-Type': DOCX_TYPE }, '/elsewhere')).status).toBe(404);
    const tooBig = await post(new Blob([new Uint8Array(31 * 1024 * 1024)]), { 'Content-Type': DOCX_TYPE });
    expect(tooBig.status).toBe(413);
    // A zip that is not a Word document: conversion fails, with a message that says nothing about the content.
    const broken = await post(new Blob([new TextEncoder().encode('PK\u0003\u0004 this is not really a document')]), { 'Content-Type': DOCX_TYPE });
    expect(broken.status).toBe(422);
    expect(await broken.text()).toBe('The document could not be converted to PDF.');
  });

  it('answers with a PDF that browsers will not cache or reinterpret', async () => {
    const pair = await buildPair(pairById('one-word-change'));
    const response = await post(new Blob([(await staticRedline(pair.original, pair.modified)).slice().buffer]), { 'Content-Type': DOCX_TYPE });
    expect(response.status).toBe(200);
    expect(response.headers.get('content-type')).toBe('application/pdf');
    expect(response.headers.get('cache-control')).toBe('no-store');
    expect(new TextDecoder().decode(new Uint8Array(await response.arrayBuffer()).slice(0, 5))).toBe('%PDF-');
  });

  it('a caller that hangs up part-way through sending, or before the reply, leaves no trace', async () => {
    // Done from inside the container, straight to the server: Docker's port forwarding would hide a dropped connection.
    const hangUp = `
import socket
for body in (b"PK" + b"x" * 1000, b""):
    s = socket.create_connection(("127.0.0.1", 8080))
    s.sendall(b"POST /convert HTTP/1.1\\r\\nHost: x\\r\\nContent-Type: ${DOCX_TYPE}\\r\\nContent-Length: 500000\\r\\n\\r\\n" + body)
    s.close()
s = socket.create_connection(("127.0.0.1", 8080))
s.sendall(b"GET /health HTTP/1.1\\r\\nHost: x\\r\\n\\r\\n")
s.close()
print("done")`;
    expect(runInContainer(hangUp).trim()).toBe('done');
    await new Promise((resolve) => setTimeout(resolve, 500));
    expect(containerLog()).toBe('');
    // And the server is still answering.
    expect((await fetch(CONVERT_URL.replace('/convert', '/health'))).status).toBe(200);
  });

  it('three conversions at once all succeed and leave nothing', async () => {
    const pair = await buildPair(pairById('several-changes'));
    const redline = await staticRedline(pair.original, pair.modified);
    const pdfs = await Promise.all([convertToPdf(redline), convertToPdf(redline), convertToPdf(redline)]);
    for (const pdf of pdfs) expect((await checkPdf({ staticRedline: redline, pdf })).passed).toBe(true);
  });

  it('a conversion that runs out of time is stopped completely: no process, no file, no log', async () => {
    const name = 'redline-pdf-timeout-test';
    startExtraContainer(name, 8792, { CONVERT_TIMEOUT_SECONDS: '0.3' });
    try {
      const url = 'http://127.0.0.1:8792';
      for (let attempt = 0; attempt < 60; attempt++) {
        try {
          if ((await fetch(`${url}/health`)).ok) break;
        } catch {
          await new Promise((resolve) => setTimeout(resolve, 250));
        }
      }
      const pair = await buildPair(pairById('several-changes'));
      const redline = await staticRedline(pair.original, pair.modified);
      for (let i = 0; i < 3; i++) {
        const response = await fetch(`${url}/convert`, { method: 'POST', headers: { 'Content-Type': DOCX_TYPE }, body: new Blob([redline.slice().buffer]) });
        expect(response.status).toBe(422);
        expect(await response.text()).toBe('The document could not be converted to PDF.');
      }
      await new Promise((resolve) => setTimeout(resolve, 1500));
      expect(containerOfficeProcesses(name)).toEqual([]);
      expect(containerScratchFiles(name)).toEqual([]);
      expect(containerLog(name)).toBe('');
    } finally {
      removeContainer(name);
    }
  });

  it('its own state report gives three counts and nothing else, and agrees with the search below', async () => {
    const report = await (await fetch(CONVERT_URL.replace('/convert', '/state'))).text();
    const counts = /^conversions=(\d+) files=(\d+) office_processes=(\d+)$/.exec(report);
    expect(counts, report).not.toBeNull();
    expect(Number(counts![1])).toBeGreaterThan(50); // every conversion this suite asked for was counted
    expect(Number(counts![2])).toBe(0);
    expect(Number(counts![3])).toBe(0);
  });

  it('the state report does count a file that is left behind (so a clean report means something)', async () => {
    // A converter with an ordinary writable disk, as on Cloudflare (the other tests use a read-only one).
    const name = 'redline-pdf-state';
    execFileSync('docker', ['rm', '-f', name], { stdio: 'ignore' });
    execFileSync('docker', ['run', '-d', '--name', name, '-p', '127.0.0.1:8793:8080', 'redline-pdf'], { stdio: 'ignore' });
    try {
      const state = async (): Promise<string> => (await fetch('http://127.0.0.1:8793/state')).text();
      for (let attempt = 0; attempt < 60; attempt++) {
        try {
          await state();
          break;
        } catch {
          await new Promise((resolve) => setTimeout(resolve, 250));
        }
      }
      expect(await state()).toBe('conversions=0 files=0 office_processes=0');
      // Every folder the converter's account can write to anywhere on the disk, found independently of the server.
      const writable = execFileSync('docker', ['exec', name, 'sh', '-c', 'find / -path /proc -prune -o -path /sys -prune -o -type d -writable -print 2>/dev/null; true'], { encoding: 'utf8' }).split('\n').filter(Boolean);
      expect(writable).toEqual(expect.arrayContaining(['/tmp', '/home/converter', '/run/lock']));
      // A file planted in each one, as the converter's own account, is counted.
      let planted = 0;
      for (const folder of writable) {
        execFileSync('docker', ['exec', name, 'sh', '-c', `touch "${folder}/left.docx"`]);
        planted++;
        expect(await state(), folder).toBe(`conversions=0 files=${planted} office_processes=0`);
      }
      // And a real conversion on a writable disk leaves the count where it was.
      const pair = await buildPair(pairById('one-word-change'));
      const reply = await fetch('http://127.0.0.1:8793/convert', { method: 'POST', headers: { 'Content-Type': DOCX_TYPE }, body: new Blob([pair.modified.slice().buffer]) });
      expect(reply.status).toBe(200);
      await reply.arrayBuffer();
      expect(await state()).toBe(`conversions=1 files=${planted} office_processes=0`);
    } finally {
      removeContainer(name);
    }
  });

  it('a second list of running processes mounted by the host (as Cloudflare does) is not mistaken for files', async () => {
    // Found on the live site: Cloudflare mounts a process list at /run/cloudchamber/proc, parts of which the
    // converter's account can "write" to (its own open connections). They are not files and must not be counted.
    const name = 'redline-pdf-cloudlike';
    execFileSync('docker', ['rm', '-f', name], { stdio: 'ignore' });
    execFileSync('docker', ['run', '-d', '--name', name, '--privileged', '--user', '0', '--entrypoint', 'sh', '-p', '127.0.0.1:8794:8080', 'redline-pdf', '-c',
      'mkdir -p /run/cloudchamber/proc /run/host/bpf && mount -t proc proc /run/cloudchamber/proc && mount -t bpf bpf /run/host/bpf && chmod 1777 /run/host/bpf && mount -t tmpfs tmpfs /sys/fs/bpf && chmod 1777 /sys/fs/bpf && echo old > /tmp/existing && chown 999 /tmp/existing && exec setpriv --reuid 999 --regid 999 --clear-groups python3 -u /opt/converter/server.py'], { stdio: 'ignore' });
    try {
      const ask = async (path: string): Promise<string> => (await fetch(`http://127.0.0.1:8794${path}`)).text();
      for (let attempt = 0; attempt < 60; attempt++) {
        try {
          await ask('/state');
          break;
        } catch {
          await new Promise((resolve) => setTimeout(resolve, 250));
        }
      }
      // The layout really is there, and really is writable in places by the converter's account.
      const writable = execFileSync('docker', ['exec', '-u', '999', name, 'sh', '-c', 'find /run/cloudchamber/proc -maxdepth 3 -type d -writable 2>/dev/null | head -5; true'], { encoding: 'utf8' });
      expect(writable).toContain('/run/cloudchamber/proc/');
      const pair = await buildPair(pairById('one-word-change'));
      const reply = await fetch('http://127.0.0.1:8794/convert', { method: 'POST', headers: { 'Content-Type': DOCX_TYPE }, body: new Blob([pair.modified.slice().buffer]) });
      expect(reply.status).toBe(200);
      await reply.arrayBuffer();
      expect(await ask('/state')).toBe('conversions=1 files=0 office_processes=0');
      const detail = await ask('/state-detail');
      expect(detail).toMatch(/^before_first=0 since_first=0 where=/);
      expect(detail).not.toContain('cloudchamber');
      expect(detail).toContain('/tmp:0');
      // A real file planted afterwards is still counted.
      execFileSync('docker', ['exec', '-u', '999', name, 'sh', '-c', 'touch /tmp/left.docx']);
      expect(await ask('/state')).toBe('conversions=1 files=1 office_processes=0');
      expect(await ask('/state-detail')).toContain('since_first=1');
      // Found by the Stage 6 review: other places an unprivileged account can keep data. Each is counted:
      // a writable special file system; a file system mounted inside /sys; a file that was already there
      // and has been written to; a block of shared memory (which is not a file at all).
      const plant = async (command: string, expected: number): Promise<void> => {
        execFileSync('docker', ['exec', '-u', '999', name, 'sh', '-c', command]);
        expect(await ask('/state'), command).toBe(`conversions=1 files=${expected} office_processes=0`);
      };
      await plant('mkdir /run/host/bpf/kept', 2);
      await plant('touch /sys/fs/bpf/kept', 3);
      await plant('echo more >> /tmp/existing', 4);
      await plant('python3 -c "import ctypes; ctypes.CDLL(None).shmget(0, 4096, 0o1600)"', 5);
      const where = await ask('/state-detail');
      expect(where).toContain('/run/host/bpf:1');
      expect(where).toContain('/sys/fs/bpf:1');
    } finally {
      removeContainer(name);
    }
  });

  it('after every conversion above, successful or not: no file is left behind, nothing is still running, and nothing was logged', () => {
    expect(containerScratchFiles()).toEqual([]);
    expect(containerOfficeProcesses()).toEqual([]);
    expect(containerLog()).toBe('');
  });
});

describe('a redline that shows its deletions in dark orange (the document itself uses red strikethrough)', () => {
  it('its PDF passes check 9: deletions are printed in dark orange, and the document\'s own red text is not taken for a change', async () => {
    const { withRedStrikethrough } = await import('../fixtures/real-world');
    const pair = await buildPair(pairById('several-changes'));
    const modified = await withRedStrikethrough(pair.modified);
    const cover = { originalName: 'original.docx', modifiedName: 'modified.docx', comparedAt: '24 November 2026 at 12:00 (UTC+00:00)' };
    const made = await makeStaticRedline(await produceTrackedChanges(pair.original, modified), pair.original, cover);
    expect(new TextDecoder().decode((await unzip(made.docx)).get('word/document.xml')!)).toContain('shown in dark orange');
    const report = await checkPdf({ staticRedline: made.docx, pdf: await convertToPdf(made.docx) });
    expect(report.checks.filter((check) => !check.passed)).toEqual([]);
    // The same PDF checked against a red redline of the plain pair must not pass: the colours are really compared.
    const red = await makeStaticRedline(await produceTrackedChanges(pair.original, pair.modified), pair.original, cover);
    expect((await checkPdf({ staticRedline: red.docx, pdf: await convertToPdf(made.docx) })).passed).toBe(false);
  });
});

describe('a redline that shows its insertions in magenta (the document itself uses blue double underline)', () => {
  it('its PDF passes check 9: insertions are printed in magenta, and the document\'s own blue text is not taken for a change', async () => {
    const { withBlueDoubleUnderline } = await import('../fixtures/real-world');
    const pair = await buildPair(pairById('several-changes'));
    const modified = await withBlueDoubleUnderline(pair.modified);
    const cover = { originalName: 'original.docx', modifiedName: 'modified.docx', comparedAt: '24 November 2026 at 12:00 (UTC+00:00)' };
    const made = await makeStaticRedline(await produceTrackedChanges(pair.original, modified), pair.original, cover);
    expect(new TextDecoder().decode((await unzip(made.docx)).get('word/document.xml')!)).toContain('shown in magenta');
    const report = await checkPdf({ staticRedline: made.docx, pdf: await convertToPdf(made.docx) });
    expect(report.checks.filter((check) => !check.passed)).toEqual([]);
    // The same PDF checked against a blue redline of the plain pair must not pass: the colours are really compared.
    const blue = await makeStaticRedline(await produceTrackedChanges(pair.original, pair.modified), pair.original, cover);
    expect((await checkPdf({ staticRedline: blue.docx, pdf: await convertToPdf(made.docx) })).passed).toBe(false);
  });
});
