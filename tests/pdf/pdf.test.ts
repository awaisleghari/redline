/**
 * Stage 5: the PDF. These tests run the real converter container (Docker) and
 * the Checker's check 9 on its output.
 */
import { initialize } from 'docxodus/core';
import { beforeAll, describe, expect, it } from 'vitest';
import { checkPdf, CHECKS_FOR } from '../../src/checker';
import { unzip, zip } from '../../src/common/zip';
import { makeStaticRedline } from '../../src/producer/make-static-redline';
import { produceTrackedChanges } from '../../src/producer/produce';
import { BROKEN_PDF_CASES } from '../fixtures/broken-pdf';
import { buildLargeContract } from '../fixtures/large';
import { ALL_PAIR_SPECS, CHECK_ONLY_PAIRS, buildDocx, buildPair, pairById } from '../fixtures/pairs';
import { containerLog, containerScratchFiles, convertToPdf, CONVERT_URL, DOCX_TYPE } from './container';

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
    const pair = await buildPair(pairById(brokenCase.pairId));
    const genuine = await staticRedline(pair.original, pair.modified);
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

  it('after every conversion above, successful or not: no file is left behind and nothing was logged', () => {
    expect(containerScratchFiles()).toEqual([]);
    expect(containerLog()).toBe('');
  });
});
