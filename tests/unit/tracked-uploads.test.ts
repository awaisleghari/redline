/**
 * Uploads that already contain tracked changes are compared as if those
 * changes had been accepted (owner's decision, 2026-11-24).
 *
 * The documents with tracked changes used here are the engine's own
 * tracked-changes files for the generated pairs: for a pair (A, B), that file
 * reads as B when its changes are accepted and as A when they are rejected, and
 * between them the pairs cover insertions, deletions, moves, tables, rows,
 * headers, footers, notes, numbered lists and formatting changes.
 */
import { initialize } from 'docxodus/core';
import { beforeAll, describe, expect, it } from 'vitest';
import { checkStaticRedline, checkTrackedChanges } from '../../src/checker';
import { unzip, zip } from '../../src/common/zip';
import { containsTrackedChanges } from '../../src/app/validate';
import { acceptChangesInPart, acceptTrackedChanges, hasTrackedChanges } from '../../src/producer/accept-changes';
import { makeStaticRedline } from '../../src/producer/make-static-redline';
import { produceTrackedChanges } from '../../src/producer/produce';
import { ALL_PAIR_SPECS, CHECK_ONLY_PAIRS, buildDocx, buildPair, pairById } from '../fixtures/pairs';
import { openPackage, readStories } from './real-world-reading';

const COVER = { originalName: 'original.docx', modifiedName: 'modified.docx', comparedAt: '24 November 2026 at 10:00 (UTC+00:00)' };
const decoder = new TextDecoder();
const encoder = new TextEncoder();

beforeAll(async () => {
  await initialize();
});

const PAIRS: { id: string; build: () => Promise<{ original: Uint8Array; modified: Uint8Array }> }[] = [
  ...ALL_PAIR_SPECS.map((spec) => ({ id: spec.id, build: () => buildPair(spec) })),
  ...CHECK_ONLY_PAIRS.map((spec) => ({ id: spec.id, build: async () => ({ original: await buildDocx(spec.original), modified: await buildDocx(spec.modified) }) })),
];

async function bothPass(original: Uint8Array, modified: Uint8Array, label: string): Promise<{ tracked: Uint8Array; staticRedline: Uint8Array; counts: { insertions: number; deletions: number; moves: number; formattingChanges: number } }> {
  const tracked = await produceTrackedChanges(original, modified);
  const trackedReport = await checkTrackedChanges({ original, modified, output: tracked });
  expect(trackedReport.checks.filter((c) => !c.passed), `${label}: tracked-changes file`).toEqual([]);
  // As the page does: it says which uploads came with tracked changes, and the Producer puts the note on the cover page.
  const made = await makeStaticRedline(tracked, original, { ...COVER, trackedChangesIn: await whichHaveChanges(original, modified) });
  const staticReport = await checkStaticRedline({ original, modified, output: made.docx, ...COVER });
  expect(staticReport.checks.filter((c) => !c.passed), `${label}: static redline`).toEqual([]);
  return { tracked, staticRedline: made.docx, counts: made.counts };
}

async function whichHaveChanges(original: Uint8Array, modified: Uint8Array): Promise<'original' | 'modified' | 'both' | undefined> {
  const [inOriginal, inModified] = [await containsTrackedChanges(original), await containsTrackedChanges(modified)];
  return inOriginal && inModified ? 'both' : inOriginal ? 'original' : inModified ? 'modified' : undefined;
}
const coverOf = async (staticRedline: Uint8Array): Promise<string> => decoder.decode((await unzip(staticRedline)).get('word/document.xml')!).replace(/<[^>]+>/g, '|');

describe("the Producer's accepted copy of an upload", () => {
  it.each(PAIRS)('$id: reads, to the Checker, exactly as the Checker itself reads the upload with its changes accepted; and holds no tracked change', async ({ build }) => {
    const pair = await build();
    const withChanges = await produceTrackedChanges(pair.original, pair.modified);
    expect(await hasTrackedChanges(withChanges)).toBe(true);
    const accepted = await acceptTrackedChanges(withChanges);
    expect(await hasTrackedChanges(accepted)).toBe(false);
    // Two separate implementations of "accepted" (the Producer's here, the Checker's in its reading) must agree,
    // and both must agree with the document the changes lead to.
    const reading = readStories(await openPackage(accepted));
    expect(reading).toEqual(readStories(await openPackage(withChanges)));
    expect(reading).toEqual(readStories(await openPackage(pair.modified)));
  });

  it('an upload without tracked changes is passed on as the very same bytes', async () => {
    const pair = await buildPair(pairById('several-changes'));
    expect(await hasTrackedChanges(pair.original)).toBe(false);
    expect(await acceptTrackedChanges(pair.original)).toBe(pair.original);
  });

  it('accepts each kind of tracked change as Word does', () => {
    const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
    const accept = (body: string): string => acceptChangesInPart(`<w:document ${W}><w:body>${body}</w:body></w:document>`).replace(/^.*<w:body>|<\/w:body>.*$/g, '');
    const A = 'w:id="1" w:author="a" w:date="2026-01-01T00:00:00Z"';
    // Inserted text stays, deleted text goes.
    expect(accept(`<w:p><w:ins ${A}><w:r><w:t>new</w:t></w:r></w:ins><w:del ${A}><w:r><w:delText>old</w:delText></w:r></w:del></w:p>`)).toBe('<w:p><w:r><w:t>new</w:t></w:r></w:p>');
    // A deleted paragraph mark joins the paragraph to the next, which keeps its own properties.
    expect(accept(`<w:p><w:pPr><w:jc w:val="center"/><w:rPr><w:del ${A}/></w:rPr></w:pPr><w:r><w:t>one </w:t></w:r></w:p><w:p><w:pPr><w:jc w:val="right"/></w:pPr><w:r><w:t>two</w:t></w:r></w:p>`)).toBe(
      '<w:p><w:pPr><w:jc w:val="right"/></w:pPr><w:r><w:t>one </w:t></w:r><w:r><w:t>two</w:t></w:r></w:p>',
    );
    // A wholly deleted paragraph disappears; an inserted paragraph mark simply stays.
    expect(accept(`<w:p><w:r><w:t>a</w:t></w:r></w:p><w:p><w:pPr><w:rPr><w:del ${A}/></w:rPr></w:pPr><w:del ${A}><w:r><w:delText>gone</w:delText></w:r></w:del></w:p><w:p><w:pPr><w:rPr><w:ins ${A}/></w:rPr></w:pPr><w:r><w:t>b</w:t></w:r></w:p>`)).toBe(
      '<w:p><w:r><w:t>a</w:t></w:r></w:p><w:p><w:pPr><w:rPr/></w:pPr><w:r><w:t>b</w:t></w:r></w:p>',
    );
    // Moved text: gone from where it was, kept where it went; the range markers go.
    expect(accept(`<w:p><w:moveFromRangeStart w:id="2" w:name="m"/><w:moveFrom ${A}><w:r><w:t>moved</w:t></w:r></w:moveFrom><w:moveFromRangeEnd w:id="2"/><w:r><w:t>stay</w:t></w:r><w:moveToRangeStart w:id="3" w:name="m"/><w:moveTo ${A}><w:r><w:t>moved</w:t></w:r></w:moveTo><w:moveToRangeEnd w:id="3"/></w:p>`)).toBe(
      '<w:p><w:r><w:t>stay</w:t></w:r><w:r><w:t>moved</w:t></w:r></w:p>',
    );
    // A deleted table row goes; an inserted one stays; a record of earlier formatting goes and the present formatting stays.
    expect(accept(`<w:tbl><w:tr><w:trPr><w:del ${A}/></w:trPr><w:tc><w:p/></w:tc></w:tr><w:tr><w:trPr><w:ins ${A}/></w:trPr><w:tc><w:p><w:r><w:rPr><w:b/><w:rPrChange ${A}><w:rPr/></w:rPrChange></w:rPr><w:t>x</w:t></w:r></w:p></w:tc></w:tr></w:tbl>`)).toBe(
      '<w:tbl><w:tr><w:trPr/><w:tc><w:p><w:r><w:rPr><w:b/></w:rPr><w:t>x</w:t></w:r></w:p></w:tc></w:tr></w:tbl>',
    );
  });
});

describe('comparing when an upload contains tracked changes', () => {
  // For each pair (A, B): T is a document with tracked changes that reads as B.
  it.each(PAIRS)('$id: as the original, as the modified document, and as both', async ({ id, build }) => {
    const { original: a, modified: b } = await build();
    const t = await produceTrackedChanges(a, b);

    // T against B: the same document once its changes are accepted, so nothing is marked.
    const same = await bothPass(t, b, `${id}: tracked original against what it becomes`);
    expect(same.counts).toEqual({ insertions: 0, deletions: 0, moves: 0, formattingChanges: 0 });

    // A against T: the same redline as A against B.
    const plain = await bothPass(a, b, `${id}: plain pair`);
    const asModified = await bothPass(a, t, `${id}: tracked modified`);
    expect(asModified.counts).toEqual(plain.counts);
    // ...and it passes when checked against the clean B as well: the outputs do not depend on which form was uploaded.
    expect((await checkTrackedChanges({ original: a, modified: b, output: asModified.tracked })).passed).toBe(true);

    // T against A: T reads as B, so this is B against A.
    const reversed = await bothPass(b, a, `${id}: reversed plain pair`);
    const asOriginal = await bothPass(t, a, `${id}: tracked original`);
    expect(asOriginal.counts).toEqual(reversed.counts);
    // Against the clean B the static redline differs in one respect only: its cover page carries the note, which a clean upload must not have.
    expect((await checkStaticRedline({ original: b, modified: a, output: asOriginal.staticRedline, ...COVER })).checks.filter((check) => !check.passed).map((check) => check.id)).toEqual(['cover-page']);
  });

  it('both uploads contain tracked changes', async () => {
    const one = await buildPair(pairById('several-changes'));
    const two = await buildPair(pairById('numbered-clause-deleted'));
    const t1 = await produceTrackedChanges(one.original, one.modified);
    const t2 = await produceTrackedChanges(two.original, two.modified);
    await bothPass(t1, t2, 'both tracked');
    await bothPass(t2, t1, 'both tracked, other way');
  });

  it('the cover page says which upload came with tracked changes, the Checker requires exactly that, and its report says how the upload was read', async () => {
    const { original: a, modified: b } = await buildPair(pairById('several-changes'));
    const t = await produceTrackedChanges(a, b);
    const NOTE = {
      original: 'Note: the original contained tracked changes. It was compared as if they had all been accepted.',
      modified: 'Note: the modified document contained tracked changes. It was compared as if they had all been accepted.',
      both: 'Note: both documents contained tracked changes. They were compared as if those had all been accepted.',
    };
    for (const [original, modified, which] of [[t, a, 'original'], [a, t, 'modified'], [t, t, 'both']] as const) {
      const { staticRedline } = await bothPass(original, modified, which);
      expect(await coverOf(staticRedline)).toContain(NOTE[which]);
      const report = await checkStaticRedline({ original, modified, output: staticRedline, ...COVER });
      const notes = report.checks.map((check) => check.note ?? '').join(' ');
      if (which !== 'modified') expect(notes).toContain('The original contained tracked changes of its own; it was read as if they had all been accepted.');
      if (which !== 'original') expect(notes).toContain('The modified document contained tracked changes of its own; it was read as if they had all been accepted.');
      // A cover page without the note, or with the wrong one, is refused: the Checker decides from the uploads themselves.
      const tracked = await produceTrackedChanges(original, modified);
      for (const wrong of [undefined, which === 'both' ? 'original' : 'both'] as const) {
        const made = await makeStaticRedline(tracked, a, { ...COVER, trackedChangesIn: wrong ?? ('none' as never) });
        const refused = await checkStaticRedline({ original, modified, output: made.docx, ...COVER });
        expect(refused.checks.filter((check) => !check.passed).map((check) => check.id), `${which} with ${wrong}`).toContain('cover-page');
      }
    }
    // And a note on the cover page of an ordinary comparison is refused too.
    const plainTracked = await produceTrackedChanges(a, b);
    const noted = await makeStaticRedline(plainTracked, a, { ...COVER, trackedChangesIn: 'original' });
    expect((await checkStaticRedline({ original: a, modified: b, output: noted.docx, ...COVER })).checks.filter((check) => !check.passed).map((check) => check.id)).toEqual(['cover-page']);
    const plain = await bothPass(a, b, 'plain');
    expect(await coverOf(plain.staticRedline)).not.toContain('Note:');
  });

  it('an output made as if the changes in the upload had been rejected is caught', async () => {
    const { original: a, modified: b } = await buildPair(pairById('several-changes'));
    const c = (await buildPair(pairById('one-word-change'))).modified;
    const t = await produceTrackedChanges(a, b); // reads as B
    // The right output compares B with C. An output comparing A with C (the changes rejected, not accepted) must fail against the upload T.
    const wrong = await produceTrackedChanges(a, c);
    const report = await checkTrackedChanges({ original: t, modified: c, output: wrong });
    expect(report.checks.filter((check) => !check.passed).map((check) => check.id)).toEqual(expect.arrayContaining(['reject-all']));
    const wrongStatic = await makeStaticRedline(wrong, a, COVER);
    const staticReport = await checkStaticRedline({ original: t, modified: c, output: wrongStatic.docx, ...COVER });
    expect(staticReport.checks.filter((check) => !check.passed).map((check) => check.id)).toEqual(expect.arrayContaining(['static-old']));
  });

  it("a tracked change of the upload's own that survives into the output is caught", async () => {
    const { original: a, modified: b } = await buildPair(pairById('several-changes'));
    const t = await produceTrackedChanges(a, b);
    const { staticRedline } = await bothPass(t, b, 'baseline');
    const files = await unzip(staticRedline);
    const xml = decoder.decode(files.get('word/document.xml')!);
    const left = xml.replace(/<w:r>(?=(?:<w:rPr>(?:(?!<\/w:rPr>).)*<\/w:rPr>)?<w:t[ >])/, '<w:ins w:id="77" w:author="someone" w:date="2026-01-01T00:00:00Z"><w:r><w:t>left over</w:t></w:r></w:ins><w:r>');
    expect(left).not.toBe(xml);
    files.set('word/document.xml', encoder.encode(left));
    const report = await checkStaticRedline({ original: t, modified: b, output: await zip(files), ...COVER });
    expect(report.checks.filter((check) => !check.passed).map((check) => check.id)).toEqual(expect.arrayContaining(['no-tracked-changes']));
  });
});
