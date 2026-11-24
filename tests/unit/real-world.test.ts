/**
 * Regression tests for a pair of real documents the owner compared on the live
 * site (2026-11-24), which was wrongly withheld. Three separate causes; each is
 * reproduced here on generated documents.
 */
import { initialize } from 'docxodus/core';
import { beforeAll, describe, expect, it } from 'vitest';
import { checkStaticRedline, checkTrackedChanges } from '../../src/checker';
import { parseXml } from '../../src/checker/xml';
import { unzip, zip } from '../../src/common/zip';
import { makeSpacesExplicit, makeSpacesExplicitInPart } from '../../src/producer/explicit-spaces';
import { makeStaticRedline } from '../../src/producer/make-static-redline';
import { produceTrackedChanges } from '../../src/producer/produce';
import { buildPair, pairById } from '../fixtures/pairs';
import { compareDocuments } from 'docxodus/core';
import { repairNotes } from '../../src/producer/repair-notes';
import { withBlueDoubleUnderline, withClauseStyle, withNumberedClausesInCell, withDocumentWideSpaces, withNoteInTable, withSections } from '../fixtures/real-world';

const COVER = { originalName: 'original.docx', modifiedName: 'modified.docx', comparedAt: '24 November 2026 at 10:00 (UTC+00:00)' };
const decoder = new TextDecoder();
const encoder = new TextEncoder();

beforeAll(async () => {
  await initialize();
});

interface Outputs {
  tracked: Uint8Array;
  staticRedline: Uint8Array;
  counts: { insertions: number; deletions: number; moves: number; formattingChanges: number };
}

/** Both Word files, as the product makes them, each required to pass every one of its checks. */
async function bothPass(original: Uint8Array, modified: Uint8Array): Promise<Outputs> {
  const tracked = await produceTrackedChanges(original, modified);
  const trackedReport = await checkTrackedChanges({ original, modified, output: tracked });
  expect(trackedReport.checks.filter((c) => !c.passed)).toEqual([]);
  const made = await makeStaticRedline(tracked, original, COVER);
  const staticReport = await checkStaticRedline({ original, modified, output: made.docx, ...COVER });
  expect(staticReport.checks.filter((c) => !c.passed)).toEqual([]);
  return { tracked, staticRedline: made.docx, counts: made.counts };
}

async function editPart(docx: Uint8Array, part: RegExp, edit: (xml: string) => string): Promise<Uint8Array> {
  const files = await unzip(docx);
  let edited = 0;
  for (const [name, bytes] of files) {
    if (!part.test(name)) continue;
    const xml = decoder.decode(bytes);
    const out = edit(xml);
    if (out !== xml) edited++;
    files.set(name, encoder.encode(out));
  }
  expect(edited, 'the damage was applied').toBeGreaterThan(0);
  return zip(files);
}
const failed = async (report: Promise<{ checks: { id: string; passed: boolean }[] }>): Promise<string[]> => (await report).checks.filter((c) => !c.passed).map((c) => c.id);

describe('a document that says "keep white space" once, for the whole document', () => {
  it('the XML reader applies the setting to everything inside, until an inner element says otherwise', () => {
    const root = parseXml('<a xml:space="preserve"><b><t> x </t></b><c xml:space="default"><t> y </t><d xml:space="preserve"><t> z </t></d></c></a>');
    const [b, c] = root.children;
    expect(root.preserveSpace).toBe(true);
    expect(b!.children[0]!.preserveSpace).toBe(true);
    expect(c!.children[0]!.preserveSpace).toBe(false);
    expect(c!.children[1]!.children[0]!.preserveSpace).toBe(true);
    expect(parseXml('<a><t> x </t></a>').children[0]!.preserveSpace).toBe(false);
  });

  it('before the engine sees such a document, the setting is written onto each piece of text and taken off everything else; nothing else changes', async () => {
    const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
    const part = `<?xml version="1.0"?><!-- xml:space="x" --><w:document ${W} xmlns:a="urn:drawing" xml:space="preserve"><w:body><w:p><w:r><w:t>TERM</w:t></w:r><w:r><w:t> SHEET</w:t></w:r><w:r><w:t xml:space="default"> x </w:t></w:r><w:r xml:space='default'><w:t> y </w:t><w:instrText> PAGE </w:instrText></w:r><a:t> drawing </a:t><w:r><w:t>a&#13;b &amp; c</w:t><w:tab/></w:r></w:p></w:body></w:document>`;
    const explicit = makeSpacesExplicitInPart(part);
    expect(explicit).toBe(
      `<?xml version="1.0"?><!-- xml:space="x" --><w:document ${W} xmlns:a="urn:drawing"><w:body><w:p><w:r><w:t xml:space="preserve">TERM</w:t></w:r><w:r><w:t xml:space="preserve"> SHEET</w:t></w:r><w:r><w:t xml:space="default"> x </w:t></w:r><w:r><w:t> y </w:t><w:instrText> PAGE </w:instrText></w:r><a:t> drawing </a:t><w:r><w:t xml:space="preserve">a&#13;b &amp; c</w:t><w:tab/></w:r></w:p></w:body></w:document>`,
    );
    // A part that is not Word markup is left alone, whatever it says.
    const other = '<a:theme xmlns:a="urn:drawing" xml:space="preserve"><a:t> x </a:t></a:theme>';
    expect(makeSpacesExplicitInPart(other)).toBe(other);
    // An ordinary document is passed on as the very same bytes.
    const ordinary = (await buildPair(pairById('several-changes'))).modified;
    expect(await makeSpacesExplicit(ordinary)).toBe(ordinary);
  });

  it('the prepared document reads, to the Checker, exactly as the upload does, for every generated pair', async () => {
    const { ALL_PAIR_SPECS } = await import('../fixtures/pairs');
    const { openPackage, readStories } = await import('./real-world-reading');
    for (const spec of ALL_PAIR_SPECS) {
      const pair = await buildPair(spec);
      for (const docx of [pair.original, pair.modified]) {
        const upload = await withDocumentWideSpaces(docx);
        const prepared = await makeSpacesExplicit(upload);
        expect(prepared).not.toBe(upload);
        // No outer setting is left in what the engine is given.
        for (const [name, bytes] of await unzip(prepared)) if (/^word\/[^/]+\.xml$/.test(name)) expect(decoder.decode(bytes), `${spec.id} ${name}`).not.toMatch(/<w:(document|hdr|ftr|footnotes|endnotes|body|p|r)\b[^>]*xml:space/);
        expect(readStories(await openPackage(prepared)), spec.id).toEqual(readStories(await openPackage(upload)));
        expect(readStories(await openPackage(upload)), spec.id).toEqual(readStories(await openPackage(docx)));
      }
    }
  });

  for (const [which, reshape] of [
    ['the modified document only', { original: false, modified: true }],
    ['the original document only', { original: true, modified: false }],
    ['both documents', { original: true, modified: true }],
  ] as const) {
    it(`${which}: both Word files pass, and the spaces are really there`, async () => {
      for (const id of ['several-changes', 'footnote-change', 'header-change', 'moved-paragraph']) {
        const pair = await buildPair(pairById(id));
        const original = reshape.original ? await withDocumentWideSpaces(pair.original) : pair.original;
        const modified = reshape.modified ? await withDocumentWideSpaces(pair.modified) : pair.modified;
        if (reshape.modified) {
          const xml = decoder.decode((await unzip(modified)).get('word/document.xml')!);
          expect(xml).toMatch(/<w:document[^>]* xml:space="preserve"/);
          expect(xml).toContain('<w:t> </w:t>');
        }
        const { tracked, staticRedline } = await bothPass(original, modified);
        // The same documents written the ordinary way give the same verdict from the Checker: it reads the two forms alike.
        expect(await failed(checkTrackedChanges({ original: pair.original, modified: pair.modified, output: tracked }))).toEqual([]);
        expect(await failed(checkStaticRedline({ original: pair.original, modified: pair.modified, output: staticRedline, ...COVER }))).toEqual([]);
      }
    });
  }

  it('an output in which such a space is lost is caught (words run together), whichever of the two documents is written this way', async () => {
    const pair = await buildPair(pairById('several-changes'));
    for (const reshape of [{ original: false, modified: true }, { original: true, modified: false }, { original: true, modified: true }]) {
      const original = reshape.original ? await withDocumentWideSpaces(pair.original) : pair.original;
      const modified = reshape.modified ? await withDocumentWideSpaces(pair.modified) : pair.modified;
      const { tracked, staticRedline } = await bothPass(original, modified);
      // No output leaves the matter to an outer element.
      for (const output of [tracked, staticRedline]) {
        for (const [name, bytes] of await unzip(output)) if (/^word\/[^/]+\.xml$/.test(name)) expect(decoder.decode(bytes), name).not.toMatch(/<w:(document|hdr|ftr|footnotes|endnotes|body|p|r)\b[^>]*xml:space/);
      }
      // One piece of text loses its own setting.
      const lose = (xml: string): string => xml.replace(/<w:t xml:space="preserve">( [^<]*)<\/w:t>/, '<w:t>$1</w:t>');
      expect(await failed(checkTrackedChanges({ original, modified, output: await editPart(tracked, /^word\/document\.xml$/, lose) }))).toEqual(expect.arrayContaining(['accept-all']));
      expect(await failed(checkStaticRedline({ original, modified, output: await editPart(staticRedline, /^word\/document\.xml$/, lose), ...COVER }))).toEqual(expect.arrayContaining(['static-new']));
      // An output that states the setting on its outermost element instead is refused outright, even though it would read the same.
      const outer = (xml: string): string => xml.replace(/<w:document\b/, '<w:document xml:space="preserve"');
      expect(await failed(checkTrackedChanges({ original, modified, output: await editPart(tracked, /^word\/document\.xml$/, outer) }))).toEqual(['structure']);
      expect(await failed(checkStaticRedline({ original, modified, output: await editPart(staticRedline, /^word\/document\.xml$/, outer), ...COVER }))).toEqual(['structure']);
    }
  });
});

describe('a style of the same name defined differently in the two documents', () => {
  const OLD = '<w:pPr><w:spacing w:before="120"/></w:pPr><w:rPr><w:b/></w:rPr>';
  const NEW = '<w:pPr><w:spacing w:before="240"/><w:jc w:val="both"/></w:pPr><w:rPr><w:i/></w:rPr>';

  it('the engine records the change inside the styles part; the static redline carries no such record, and both files pass', async () => {
    const pair = await buildPair(pairById('one-word-change'));
    const original = await withClauseStyle(pair.original, OLD);
    const modified = await withClauseStyle(pair.modified, NEW);
    const { tracked, staticRedline, counts } = await bothPass(original, modified);
    const stylesOf = async (docx: Uint8Array): Promise<string> => decoder.decode((await unzip(docx)).get('word/styles.xml')!);
    // The situation is real: the tracked-changes file does hold change records in its styles part.
    expect(await stylesOf(tracked)).toMatch(/<w:[pr]PrChange\b/);
    expect(await stylesOf(staticRedline)).not.toMatch(/<w:[pr]PrChange\b/);
    // The new definition is the one kept.
    expect(await stylesOf(staticRedline)).toContain('<w:jc w:val="both"');
    // The paragraph's unchanged words went from bold to italic through the style alone: that is a formatting change.
    expect(counts.formattingChanges).toBeGreaterThan(0);
  });

  it('a static redline that still holds such a record is caught', async () => {
    const pair = await buildPair(pairById('one-word-change'));
    const original = await withClauseStyle(pair.original, OLD);
    const modified = await withClauseStyle(pair.modified, NEW);
    const { staticRedline } = await bothPass(original, modified);
    const left = await editPart(staticRedline, /^word\/styles\.xml$/, (xml) => xml.replace('<w:jc w:val="both"/>', '<w:jc w:val="both"/><w:pPrChange w:id="900" w:author="x"><w:pPr/></w:pPrChange>'));
    expect(await failed(checkStaticRedline({ original, modified, output: left, ...COVER }))).toEqual(expect.arrayContaining(['no-tracked-changes']));
  });
});

describe('a header and footer carried over from one section into the next', () => {
  const SPEC = 'footer-change';

  for (const [which, shape] of [
    ['later sections of the modified document carry them over', { original: undefined, modified: false }],
    ['both documents carry them over', { original: false, modified: false }],
    ['the original names them in every section, the modified carries them over', { original: true, modified: false }],
    ['the original carries them over, the modified names them in every section', { original: false, modified: true }],
  ] as const) {
    it(`${which}: both Word files pass`, async () => {
      const pair = await buildPair(pairById(SPEC));
      const original = shape.original === undefined ? pair.original : await withSections(pair.original, shape.original);
      const modified = await withSections(pair.modified, shape.modified);
      await bothPass(original, modified);
    });
  }

  it('an output whose later sections show a different footer, or none, is caught', async () => {
    const pair = await buildPair(pairById(SPEC));
    const modified = await withSections(pair.modified);
    const { tracked, staticRedline } = await bothPass(pair.original, modified);
    // No footer anywhere.
    const none = (xml: string): string => xml.replace(/<w:footerReference\b[^>]*\/>/g, '');
    expect(await failed(checkTrackedChanges({ original: pair.original, modified, output: await editPart(tracked, /^word\/document\.xml$/, none) }))).toEqual(expect.arrayContaining(['accept-all']));
    expect(await failed(checkStaticRedline({ original: pair.original, modified, output: await editPart(staticRedline, /^word\/document\.xml$/, none), ...COVER }))).toEqual(expect.arrayContaining(['static-new']));
    // The footer's words changed.
    const reword = (xml: string): string => xml.replace(/Holdings/g, 'Holdingz');
    expect(await failed(checkTrackedChanges({ original: pair.original, modified, output: await editPart(tracked, /^word\/footer\d*\.xml$/, reword) }))).toEqual(expect.arrayContaining(['accept-all']));
  });

  it('a later section that stops carrying the footer over (names a different one) is caught, where before the fix a carried-over footer was not read at all', async () => {
    const pair = await buildPair(pairById(SPEC));
    const modified = await withSections(pair.modified);
    const { tracked } = await bothPass(pair.original, modified);
    // In the modified document itself, give the last section a footer of its own with other words: the same output must now fail.
    const files = await unzip(modified);
    const rels = decoder.decode(files.get('word/_rels/document.xml.rels')!);
    const footerRel = /<Relationship\b[^>]*footer[^>]*>/.exec(rels)![0];
    const target = /Target="([^"]+)"/.exec(footerRel)![1]!;
    const footerPart = `word/${target}`;
    files.set('word/footerOther.xml', encoder.encode(decoder.decode(files.get(footerPart)!).replace(/Holdings/g, 'Other')));
    files.set('word/_rels/document.xml.rels', encoder.encode(rels.replace('</Relationships>', footerRel.replace(target, 'footerOther.xml').replace(/Id="[^"]+"/, 'Id="rIdOther"') + '</Relationships>')));
    files.set('[Content_Types].xml', encoder.encode(decoder.decode(files.get('[Content_Types].xml')!).replace('</Types>', '<Override PartName="/word/footerOther.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/></Types>')));
    const body = decoder.decode(files.get('word/document.xml')!);
    const at = body.lastIndexOf('<w:sectPr>');
    files.set('word/document.xml', encoder.encode(`${body.slice(0, at)}<w:sectPr><w:footerReference w:type="default" r:id="rIdOther"/>${body.slice(at + '<w:sectPr>'.length)}`));
    const different = await zip(files);
    expect(await failed(checkTrackedChanges({ original: pair.original, modified: different, output: tracked }))).toEqual(expect.arrayContaining(['accept-all']));
  });
});

/**
 * Found on real contracts (Stage 7, 39 of 161 documents): a change anywhere in a table makes the engine
 * re-issue every footnote referred to from that table as deleted and re-inserted, leaving the deleted
 * copy with nothing referring to it.
 */
describe('a footnote hanging off a table in which something else changes', () => {
  const notesOf = async (docx: Uint8Array): Promise<string> => decoder.decode((await unzip(docx)).get('word/footnotes.xml')!);
  const noteIds = (xml: string): string[] => [...xml.matchAll(/<w:footnote\b(?![^>]*w:type=)[^>]*w:id="(\d+)"/g)].map((m) => m[1]!);

  it('the engine does go wrong here; the repair puts it right, and both Word files pass with only the real change marked', async () => {
    const base = (await buildPair(pairById('footnote-change'))).original;
    const original = await withNoteInTable(base, 'The fee is payable monthly in arrear.');
    const modified = await withNoteInTable(base, 'The fee is payable quarterly in arrear.');
    // The fault is real: straight from the engine, the unchanged footnote is marked up and a stray copy exists.
    const raw = await compareDocuments(original, modified, { authorName: 'Redline' });
    const rawNotes = await notesOf(raw);
    expect(noteIds(rawNotes).length).toBeGreaterThan(noteIds(await notesOf(original)).length);
    expect(rawNotes).toMatch(/<w:ins /);
    // The repair alone removes it...
    const parts = new Map<string, string>();
    for (const [name, bytes] of await unzip(raw)) if (/^word\/[^/]+\.xml$/.test(name)) parts.set(name, decoder.decode(bytes));
    const repaired = repairNotes(parts).get('word/footnotes.xml')!;
    expect(noteIds(repaired)).toEqual(noteIds(await notesOf(original)));
    expect(repaired).not.toMatch(/<w:(ins|del) /);
    // ...and the product's own output passes every check, marking one word replaced and nothing in the footnote.
    const { tracked, counts } = await bothPass(original, modified);
    expect(await notesOf(tracked)).not.toMatch(/<w:(ins|del) /);
    expect(counts).toEqual({ insertions: 1, deletions: 1, moves: 0, formattingChanges: 0 });
  });

  it('when the footnote itself was edited as well, the note shows its old and new text, and both Word files pass', async () => {
    const pair = await buildPair(pairById('footnote-change'));
    const original = await withNoteInTable(pair.original, 'The fee is payable monthly in arrear.');
    const modified = await withNoteInTable(pair.modified, 'The fee is payable quarterly in arrear.');
    const { tracked } = await bothPass(original, modified);
    const notes = await notesOf(tracked);
    expect(noteIds(notes)).toEqual(noteIds(await notesOf(original)));
    expect(notes).toMatch(/<w:ins /);
    expect(notes).toMatch(/<w:del /);
  });

  it('a note that the engine rightly left with nothing referring to it is not touched, and an output with a stray note is still caught', async () => {
    const base = (await buildPair(pairById('footnote-change'))).original;
    const original = await withNoteInTable(base, 'The fee is payable monthly in arrear.');
    const modified = await withNoteInTable(base, 'The fee is payable quarterly in arrear.');
    const raw = await compareDocuments(original, modified, { authorName: 'Redline' });
    // The engine's unrepaired output must still fail the Checker: the repair is what makes it pass, not a looser check.
    const report = await checkTrackedChanges({ original, modified, output: raw });
    expect(report.checks.filter((c) => !c.passed).map((c) => c.id)).toEqual(expect.arrayContaining(['reject-all']));
    // Parts with no stray note come back unchanged.
    const clean = await produceTrackedChanges(base, (await buildPair(pairById('footnote-change'))).modified);
    const parts = new Map<string, string>();
    for (const [name, bytes] of await unzip(clean)) if (/^word\/[^/]+\.xml$/.test(name)) parts.set(name, decoder.decode(bytes));
    expect(repairNotes(parts).size).toBe(0);
  });
});

/**
 * Found on real contracts (Stage 7: 19 model leases): a document whose own text, or just a paragraph mark,
 * is formatted blue with a double underline, which is how the redline shows insertions. Owner's decision:
 * such a redline shows its insertions in magenta and says so in the legend; any other redline stays blue.
 */
describe('a document that itself uses blue double underline', () => {
  const ALT_LEGEND = 'Inserted text (shown in magenta, because this document itself contains blue double-underlined text)';
  const bodyOf = async (docx: Uint8Array): Promise<string> => decoder.decode((await unzip(docx)).get('word/document.xml')!);
  const insertedIn = (xml: string, colour: string): number => (xml.match(new RegExp(`<w:color w:val="${colour}"/><w:u w:val="double"/>`, 'g')) ?? []).length;

  for (const [which, shape] of [
    ['a stretch of text in the modified document', { original: undefined, modified: 'text' }],
    ['a stretch of text in the original only', { original: 'text', modified: undefined }],
    ['only a paragraph mark, in both documents (as in the leases)', { original: 'paragraph mark', modified: 'paragraph mark' }],
  ] as const) {
    it(`${which}: insertions are magenta, the legend says why, and both Word files pass`, async () => {
      const pair = await buildPair(pairById('several-changes'));
      const original = shape.original ? await withBlueDoubleUnderline(pair.original, shape.original) : pair.original;
      const modified = shape.modified ? await withBlueDoubleUnderline(pair.modified, shape.modified) : pair.modified;
      const { staticRedline, counts } = await bothPass(original, modified);
      const xml = await bodyOf(staticRedline);
      expect(xml).toContain(ALT_LEGEND);
      expect(insertedIn(xml, 'AA00AA')).toBeGreaterThan(1); // the legend sample and the insertions
      expect(counts.insertions).toBeGreaterThan(0);
      // The document's own blue text is still there, untouched, where the modified document has it.
      if (shape.modified) expect(insertedIn(xml, '0000FF')).toBeGreaterThan(0);
    });
  }

  it('an ordinary pair stays blue, with the ordinary legend', async () => {
    const pair = await buildPair(pairById('several-changes'));
    const { staticRedline } = await bothPass(pair.original, pair.modified);
    const xml = await bodyOf(staticRedline);
    expect(xml).not.toContain('magenta');
    expect(insertedIn(xml, 'AA00AA')).toBe(0);
    expect(insertedIn(xml, '0000FF')).toBeGreaterThan(1);
  });

  it('the wrong colour either way, or a legend that does not match, is caught', async () => {
    const pair = await buildPair(pairById('several-changes'));
    const blueModified = await withBlueDoubleUnderline(pair.modified);
    // A redline that should be magenta, made as if nothing were unusual (blue): the document's own blue text reads as an insertion.
    const tracked = await produceTrackedChanges(pair.original, blueModified);
    const madeBlue = await makeStaticRedline(tracked, pair.original, { ...COVER, style: { insertedColor: '0000FF' } });
    expect(await failed(checkStaticRedline({ original: pair.original, modified: blueModified, output: madeBlue.docx, ...COVER }))).not.toEqual([]);
    // A redline in magenta where nothing calls for it.
    const plainTracked = await produceTrackedChanges(pair.original, pair.modified);
    const madeMagenta = await makeStaticRedline(plainTracked, pair.original, { ...COVER, style: { insertedColor: 'AA00AA' } });
    expect(await failed(checkStaticRedline({ original: pair.original, modified: pair.modified, output: madeMagenta.docx, ...COVER }))).not.toEqual([]);
    // The right colour but the ordinary legend text.
    const { staticRedline } = await bothPass(pair.original, blueModified);
    const wrongLegend = await editPart(staticRedline, /^word\/document\.xml$/, (xml) => xml.replace(ALT_LEGEND, 'Inserted text'));
    expect(await failed(checkStaticRedline({ original: pair.original, modified: blueModified, output: wrongLegend, ...COVER }))).toContain('cover-page');
  });
});

/**
 * Found on real contracts (Stage 7, about a dozen documents): the last numbered clause in a table cell is deleted.
 * A cell's last paragraph mark cannot go, so the removal is shown on the mark of the paragraph before. The clause
 * that gives up its number must still be the deleted one; otherwise the surviving clause loses its number, or
 * both keep one and every later clause is numbered one too high.
 */
describe('numbered clauses at the end of a table cell', () => {
  it('the last clause deleted: it loses its number and shows its old one, the clause before keeps its own, and later clauses are numbered as in the modified document', async () => {
    for (const id of ['numbered-clause-deleted', 'numbered-clause-moved']) {
      const base = (await buildPair(pairById(id))).original;
      const original = await withNumberedClausesInCell(base);
      const modified = await withNumberedClausesInCell(base, 'last deleted');
      const { staticRedline, counts } = await bothPass(original, modified);
      expect(counts).toEqual({ insertions: 0, deletions: 1, moves: 0, formattingChanges: 0 });
      // In the cell: three paragraphs; only the last (the deleted clause) has had its automatic number taken off.
      const xml = decoder.decode((await unzip(staticRedline)).get('word/document.xml')!);
      const cell = /<w:tc>.*?<\/w:tc>/s.exec(xml)![0];
      const inCell = cell.match(/<w:p>.*?<\/w:p>/gs)!;
      expect(inCell).toHaveLength(3);
      expect(inCell.map((paragraph) => paragraph.includes('<w:numId w:val="0"/>'))).toEqual([false, false, true]);
      // And a redline in which the surviving clause loses its number instead is refused.
      const wrong = await editPart(staticRedline, /^word\/document\.xml$/, (doc) => doc.replace(inCell[1]!, inCell[1]!.replace(/<w:numPr>.*?<\/w:numPr>/s, '<w:numPr><w:numId w:val="0"/></w:numPr>')));
      expect(await failed(checkStaticRedline({ original, modified, output: wrong, ...COVER }))).toContain('numbering');
    }
  });

  it('the last two clauses run together, and a clause added at the end: both Word files pass', async () => {
    for (const id of ['numbered-clause-deleted', 'numbered-clause-moved']) {
      const base = (await buildPair(pairById(id))).original;
      await bothPass(await withNumberedClausesInCell(base), await withNumberedClausesInCell(base, 'joined'));
      await bothPass(await withNumberedClausesInCell(base, 'last deleted'), await withNumberedClausesInCell(base));
    }
  });
});
