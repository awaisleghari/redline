/**
 * Words nobody changed are shown once in the static redline (owner's decisions, 2026-11-27):
 *   - a link added to, taken off or re-pointed on unchanged words is shown by shading;
 *   - unchanged wording whose container changed (a list taken out of a table) is shown once under a shaded note.
 * The tracked-changes Word file keeps Word's removed-and-added form in both cases.
 *
 * The second half proves the Checker's side: a note (or shading) cannot be used to hide a real change.
 * Every document here is generated; none comes from a real contract.
 */
import { initialize } from 'docxodus/core';
import { beforeAll, describe, expect, it } from 'vitest';
import { checkStaticRedline, checkTrackedChanges } from '../../src/checker';
import { unzip, zip } from '../../src/common/zip';
import { makeStaticRedline } from '../../src/producer/make-static-redline';
import { produceTrackedChanges } from '../../src/producer/produce';
import { buildDocx, PAYMENT_METHODS, type DocSpec } from '../fixtures/pairs';

const COVER = { originalName: 'original.docx', modifiedName: 'modified.docx', comparedAt: '27 November 2026 at 10:00 (UTC+00:00)' };
const decoder = new TextDecoder();
const encoder = new TextEncoder();
const OPENING = 'This Services Agreement is made between Alpha Holdings Limited and Beta Trading LLC.';
const CLOSING = 'This Agreement is governed by the laws of England and Wales.';
const ITEMS = PAYMENT_METHODS.flat();
const p = (text: string) => ({ kind: 'p' as const, text });
const numbered = (text: string) => ({ kind: 'numbered' as const, text });
const table = (rows: string[][]) => ({ kind: 'table' as const, rows });

beforeAll(async () => {
  await initialize();
});

async function partOf(docx: Uint8Array, name: string): Promise<string> {
  return decoder.decode((await unzip(docx)).get(name)!);
}

async function editPart(docx: Uint8Array, name: string, edit: (xml: string) => string): Promise<Uint8Array> {
  const files = await unzip(docx);
  const before = decoder.decode(files.get(name)!);
  const after = edit(before);
  expect(after, 'the edit was applied').not.toBe(before);
  files.set(name, encoder.encode(after));
  return zip(files);
}

/** Both Word files as the product makes them. The tracked-changes file must pass; the static redline's failed checks are returned. */
async function compare(original: Uint8Array, modified: Uint8Array): Promise<{ tracked: Uint8Array; staticRedline: Uint8Array; counts: { insertions: number; deletions: number; moves: number; formattingChanges: number }; failed: string[] }> {
  const tracked = await produceTrackedChanges(original, modified);
  const trackedReport = await checkTrackedChanges({ original, modified, output: tracked });
  expect(trackedReport.checks.filter((c) => !c.passed)).toEqual([]);
  const made = await makeStaticRedline(tracked, original, COVER);
  const report = await checkStaticRedline({ original, modified, output: made.docx, ...COVER });
  return { tracked, staticRedline: made.docx, counts: made.counts, failed: report.checks.filter((c) => !c.passed).map((c) => `${c.id}: ${c.failure?.message}`) };
}

const failedChecks = async (original: Uint8Array, modified: Uint8Array, output: Uint8Array): Promise<string[]> =>
  (await checkStaticRedline({ original, modified, output, ...COVER })).checks.filter((c) => !c.passed).map((c) => c.id);

/** Turn one word of a part's plain text into a link to a bookmark. */
const linkWord = (docx: Uint8Array, part: string, word: string, bookmark: string): Promise<Uint8Array> =>
  editPart(docx, part, (xml) =>
    xml.replace(new RegExp(`<w:r>(<w:rPr>.*?</w:rPr>)?<w:t xml:space="preserve">([^<]*?)${word}([^<]*?)</w:t></w:r>`), (_all, rPr = '', before: string, after: string) => {
      const run = (text: string): string => (text === '' ? '' : `<w:r>${rPr}<w:t xml:space="preserve">${text}</w:t></w:r>`);
      return `${run(before)}<w:hyperlink w:anchor="${bookmark}">${run(word)}</w:hyperlink>${run(after)}`;
    }),
  );

describe('a link added to, taken off or re-pointed on words that did not change', () => {
  it('a word in the page header that became a link is shown once, shaded; the tracked-changes file keeps the removed-and-added form', async () => {
    const spec: DocSpec = { blocks: [p(OPENING), p(CLOSING)], header: 'Prepared for the Client in confidence' };
    const original = await buildDocx(spec);
    const modified = await linkWord(original, 'word/header1.xml', 'Client', 'Parties');
    const { tracked, staticRedline, counts, failed } = await compare(original, modified);
    expect(failed).toEqual([]);
    expect(counts).toEqual({ insertions: 0, deletions: 0, moves: 0, formattingChanges: 1 });
    const trackedHeader = await partOf(tracked, 'word/header1.xml');
    expect(trackedHeader).toMatch(/<w:del\b[^>]*><w:r>.*?<w:delText[^>]*>Client<\/w:delText>/);
    expect(trackedHeader).toMatch(/<w:ins\b[^>]*><w:r>.*?<w:t[^>]*>Client<\/w:t>/);
    const header = await partOf(staticRedline, 'word/header1.xml');
    expect(header.match(/Client/g)).toHaveLength(1);
    expect(header).toMatch(/<w:hyperlink w:anchor="Parties"[^>]*><w:r><w:rPr>.*?<w:shd [^>]*w:fill="FFF2CC"\/>.*?<\/w:rPr><w:t[^>]*>Client<\/w:t><\/w:r><\/w:hyperlink>/);
    expect(header).not.toMatch(/<w:strike\/>|w:val="double"/);
  });

  it('a link that points to the same place in both documents is not shaded, and words changed inside it are marked as usual', async () => {
    const original = await linkWord(await buildDocx({ blocks: [p(OPENING), p('Notices go to the Client at its registered office.'), p(CLOSING)] }), 'word/document.xml', 'the Client', 'Parties');
    const modified = await linkWord(await buildDocx({ blocks: [p(OPENING), p('Notices go to the Client at its head office.'), p(CLOSING)] }), 'word/document.xml', 'the Client', 'Parties');
    const { counts, failed } = await compare(original, modified);
    expect(failed).toEqual([]);
    expect(counts).toEqual({ insertions: 1, deletions: 1, moves: 0, formattingChanges: 0 });
  });

  it('the Checker refuses a redline that shows the new link on a word without shading it, and one that keeps the old link', async () => {
    const original = await buildDocx({ blocks: [p(OPENING), p('Notices go to the Client at its registered office.'), p(CLOSING)] });
    const modified = await linkWord(original, 'word/document.xml', 'Client', 'Parties');
    const { staticRedline, failed } = await compare(original, modified);
    expect(failed).toEqual([]);
    // Shading taken off, and the count on the cover page lowered to match: only the link rule can catch it.
    const unshaded = await editPart(staticRedline, 'word/document.xml', (xml) =>
      xml.replace(/(<w:hyperlink w:anchor="Parties"[^>]*><w:r>)<w:rPr>.*?<\/w:rPr>/, '$1').replace('Formatting changes: 1', 'Formatting changes: 0'),
    );
    expect(await failedChecks(original, modified, unshaded)).toEqual(['format-marks']);
    // The other way round: the modified document's link is quietly left off.
    const unlinked = await editPart(staticRedline, 'word/document.xml', (xml) => xml.replace(/<w:hyperlink w:anchor="Parties"[^>]*>(.*?)<\/w:hyperlink>/, '$1'));
    expect(await failedChecks(original, modified, unlinked)).toContain('static-new');
  });
});

describe('unchanged wording whose container changed', () => {
  const asTable: DocSpec = { blocks: [p(OPENING), table(PAYMENT_METHODS), p(CLOSING)] };
  const asList: DocSpec = { blocks: [p(OPENING), ...ITEMS.map(p), p(CLOSING)] };
  const NOTE = /<w:p><w:pPr><w:pStyle w:val="RedlineLayoutNote"\/><\/w:pPr><w:r><w:rPr><w:shd [^>]*w:fill="FFF2CC"\/><\/w:rPr><w:t[^>]*>([^<]*)<\/w:t><\/w:r><\/w:p>/g;
  const notesIn = (xml: string): string[] => [...xml.matchAll(NOTE)].map((match) => match[1]!);
  const marksIn = (xml: string): number => (xml.split('</w:sectPr></w:pPr></w:p>').slice(1).join('').match(/<w:strike\/>|<w:dstrike\/>|<w:u w:val="double"\/>/g) ?? []).length;

  it('a list taken out of a table is shown once, unmarked, under one shaded note; the tracked-changes file shows the table removed and the list added', async () => {
    const [original, modified] = [await buildDocx(asTable), await buildDocx(asList)];
    const { tracked, staticRedline, counts, failed } = await compare(original, modified);
    expect(failed).toEqual([]);
    expect(counts).toEqual({ insertions: 0, deletions: 0, moves: 0, formattingChanges: 1 });
    const trackedXml = await partOf(tracked, 'word/document.xml');
    expect(trackedXml.match(/<w:trPr><w:del\b/g)).toHaveLength(2);
    expect(trackedXml.match(/<w:ins\b[^>]*><w:r>/g)).toHaveLength(4);
    const xml = await partOf(staticRedline, 'word/document.xml');
    expect(notesIn(xml)).toEqual(['Layout changed: the next 4 paragraphs were previously in a table.']);
    expect(marksIn(xml)).toBe(0);
    expect(xml).not.toContain('<w:tbl>');
    for (const item of ITEMS) expect(xml.split(`>${item}<`)).toHaveLength(2);
  });

  it('plain paragraphs set out as a table are shown once, as the table, under one shaded note', async () => {
    const [original, modified] = [await buildDocx(asList), await buildDocx(asTable)];
    const { staticRedline, counts, failed } = await compare(original, modified);
    expect(failed).toEqual([]);
    expect(counts).toEqual({ insertions: 0, deletions: 0, moves: 0, formattingChanges: 1 });
    const xml = await partOf(staticRedline, 'word/document.xml');
    expect(notesIn(xml)).toEqual(['Layout changed: the table below was previously 4 separate paragraphs.']);
    expect(marksIn(xml)).toBe(0);
    for (const item of ITEMS) expect(xml.split(`>${item}<`)).toHaveLength(2);
  });

  it('empty lines in the table do not count as wording', async () => {
    const spaced: DocSpec = { blocks: [p(OPENING), table([[`${ITEMS[0]}\n`, ITEMS[1]!], [ITEMS[2]!, `\n${ITEMS[3]}`]]), p(CLOSING)] };
    const { staticRedline, failed } = await compare(await buildDocx(spaced), await buildDocx(asList));
    expect(failed).toEqual([]);
    expect(notesIn(await partOf(staticRedline, 'word/document.xml'))).toHaveLength(1);
  });

  it('where one item was reworded there is no note: the change is marked in the ordinary way', async () => {
    const reworded: DocSpec = { blocks: [p(OPENING), ...ITEMS.map((item) => p(item.replace('delivery', 'collection'))), p(CLOSING)] };
    const { staticRedline, counts, failed } = await compare(await buildDocx(asTable), await buildDocx(reworded));
    expect(failed).toEqual([]);
    const xml = await partOf(staticRedline, 'word/document.xml');
    expect(notesIn(xml)).toEqual([]);
    expect(counts.insertions + counts.deletions).toBeGreaterThan(0);
    expect(counts.formattingChanges).toBe(0);
  });

  it('numbered clauses put into a table get no note: their old numbers must stay visible', async () => {
    const numberedList: DocSpec = { blocks: [p(OPENING), ...ITEMS.map(numbered), p(CLOSING)] };
    const { staticRedline, failed } = await compare(await buildDocx(numberedList), await buildDocx(asTable));
    expect(failed).toEqual([]);
    const xml = await partOf(staticRedline, 'word/document.xml');
    expect(notesIn(xml)).toEqual([]);
    expect(xml).toContain('RedlineOldNumber');
  });

  describe('a note cannot be used to hide a real change (the Checker, against redlines made for other documents)', () => {
    it('a reworded item under a note is refused', async () => {
      // A genuine redline of table -> list, with one word then altered under the note; checked against a
      // modified document that really has that word. The new side is right; only the note's claim is false.
      const original = await buildDocx(asTable);
      const { staticRedline } = await compare(original, await buildDocx(asList));
      const reworded = await buildDocx({ blocks: [p(OPENING), ...ITEMS.map((item) => p(item.replace('delivery', 'collection'))), p(CLOSING)] });
      const forged = await editPart(staticRedline, 'word/document.xml', (xml) => xml.replace('Cash on delivery', 'Cash on collection'));
      const failed = await failedChecks(original, reworded, forged);
      expect(failed).toContain('static-old');
      expect(failed).not.toContain('static-new');
    });

    it('an item dropped under a note is refused', async () => {
      const original = await buildDocx(asTable);
      const { staticRedline } = await compare(original, await buildDocx(asList));
      const shorter = await buildDocx({ blocks: [p(OPENING), ...ITEMS.slice(0, 3).map(p), p(CLOSING)] });
      const forged = await editPart(staticRedline, 'word/document.xml', (xml) =>
        xml.replace(/<w:p>(?:(?!<w:p>).)*?Cash on delivery.*?<\/w:p>/, '').replace('the next 4 paragraphs were', 'the next 3 paragraphs were'),
      );
      const failed = await failedChecks(original, shorter, forged);
      expect(failed).toContain('static-old');
      expect(failed).not.toContain('static-new');
    });

    it('items put in another order under a note are refused', async () => {
      const original = await buildDocx(asTable);
      const { staticRedline } = await compare(original, await buildDocx(asList));
      const swapped = [ITEMS[1]!, ITEMS[0]!, ITEMS[2]!, ITEMS[3]!];
      const reordered = await buildDocx({ blocks: [p(OPENING), ...swapped.map(p), p(CLOSING)] });
      const forged = await editPart(staticRedline, 'word/document.xml', (xml) => xml.replace(`>${ITEMS[0]}<`, '>\u0000<').replace(`>${ITEMS[1]}<`, `>${ITEMS[0]}<`).replace('>\u0000<', `>${ITEMS[1]}<`));
      const failed = await failedChecks(original, reordered, forged);
      expect(failed).toContain('static-old');
      expect(failed).not.toContain('static-new');
    });

    it('a note over clauses that were numbered in the original is refused', async () => {
      // The redline is made for an unnumbered list; the original it is then checked against has the same words, numbered.
      const modified = await buildDocx(asTable);
      const { staticRedline } = await compare(await buildDocx(asList), modified);
      const numberedOriginal = await buildDocx({ blocks: [p(OPENING), ...ITEMS.map(numbered), p(CLOSING)] });
      expect(await failedChecks(numberedOriginal, modified, staticRedline)).toContain('static-old');
    });

    it('a note inside a table cell is refused', async () => {
      const original = await buildDocx(asList);
      const modified = await buildDocx(asTable);
      const { staticRedline } = await compare(original, modified);
      const note = (await partOf(staticRedline, 'word/document.xml')).match(NOTE)![0]!;
      const moved = await editPart(staticRedline, 'word/document.xml', (xml) => xml.replace(note, '').replace(/(<w:tc>)(<w:p>)/, `$1${note}$2`));
      expect(await failedChecks(original, modified, moved)).toContain('completeness');
    });
  });
});
