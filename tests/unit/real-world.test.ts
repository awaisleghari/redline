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
import { repairParagraphProperties } from '../../src/producer/repair-paragraphs';
import { unwrapRowControls, unwrapRowControlsInPart } from '../../src/producer/unwrap-row-controls';
import { withBlueDoubleUnderline, withCellsInContentControls, withGlossary, withInk, withSvgPicture, SVG_PICTURE, withParagraphsAtCellEnd, withRedStrikethrough, withClauseStyle, withRowsInContentControls, withNumberedClausesInCell, withDocumentWideSpaces, withNoteInTable, withSections } from '../fixtures/real-world';

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

  it('an empty footnote re-issued the same way (no text, only its paragraph mark flagged) is restored too', () => {
    const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
    const A = 'w:id="1" w:author="x" w:date="2000-01-01T00:00:00Z"';
    const parts = new Map<string, string>([
      ['word/document.xml', `<w:document ${W}><w:body><w:p><w:r><w:footnoteReference w:id="5"/></w:r></w:p></w:body></w:document>`],
      ['word/footnotes.xml', `<w:footnotes ${W}><w:footnote w:id="5"><w:p><w:pPr><w:rPr><w:ins ${A}/></w:rPr></w:pPr></w:p></w:footnote><w:footnote w:id="9"><w:p><w:pPr><w:rPr><w:del ${A}/></w:rPr></w:pPr></w:p></w:footnote></w:footnotes>`],
    ]);
    const repaired = repairNotes(parts).get('word/footnotes.xml')!;
    expect(repaired).not.toContain('w:id="9"');
    expect(repaired).not.toMatch(/<w:(ins|del) /);
    expect(repaired).toContain('<w:footnote w:id="5">');
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

/**
 * Found on real contracts (Stage 7) and in an earlier review: the engine sometimes gives a paragraph that survives
 * the wrong properties (another paragraph's numbering or style, or none where the modified document has a style).
 */
describe('a surviving paragraph must carry the modified document\'s properties', () => {
  const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
  const doc = (body: string): string => `<w:document ${W}><w:body>${body}<w:sectPr/></w:body></w:document>`;
  const A = 'w:id="1" w:author="x" w:date="2000-01-01T00:00:00Z"';

  it('the repair gives each surviving paragraph the modified document\'s properties, keeps the mark\'s own flags, and leaves a file alone when the paragraphs do not line up', () => {
    const modified = doc('<w:p><w:pPr><w:pStyle w:val="Clause"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="7"/></w:numPr></w:pPr><w:r><w:t>One</w:t></w:r></w:p><w:p><w:pPr><w:pStyle w:val="Plain"/></w:pPr></w:p><w:p><w:r><w:t>Three</w:t></w:r></w:p>');
    // As the engine wrote it: a deleted paragraph (not part of the accepted text), then the empty paragraph wrongly numbered, an inserted mark, and a style where there should be none.
    const tracked = doc(
      `<w:p><w:pPr><w:pStyle w:val="Clause"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="7"/></w:numPr></w:pPr><w:r><w:t>One</w:t></w:r></w:p>` +
        `<w:p><w:pPr><w:pStyle w:val="Clause"/><w:rPr><w:del ${A}/></w:rPr></w:pPr><w:del ${A}><w:r><w:delText>gone</w:delText></w:r></w:del></w:p>` +
        `<w:p><w:pPr><w:pStyle w:val="Clause"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="7"/></w:numPr><w:rPr><w:ins ${A}/><w:b/></w:rPr></w:pPr></w:p>` +
        `<w:p><w:pPr><w:pStyle w:val="Clause"/></w:pPr><w:r><w:t>Three</w:t></w:r></w:p>`,
    );
    const repaired = repairParagraphProperties(tracked, modified)!;
    expect(repaired).toContain(`<w:p><w:pPr><w:pStyle w:val="Plain"/><w:rPr><w:ins ${A}/><w:b/></w:rPr></w:pPr></w:p>`);
    expect(repaired).toContain('<w:p><w:pPr/><w:r><w:t>Three</w:t></w:r></w:p>');
    // The deleted paragraph and the paragraph that was already right are untouched.
    expect(repaired).toContain(`<w:pPr><w:pStyle w:val="Clause"/><w:rPr><w:del ${A}/></w:rPr></w:pPr>`);
    expect(repaired).toContain('<w:p><w:pPr><w:pStyle w:val="Clause"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="7"/></w:numPr></w:pPr><w:r><w:t>One</w:t></w:r></w:p>');
    // Nothing to do, and nothing done when the two do not line up.
    expect(repairParagraphProperties(modified, modified)).toBeUndefined();
    expect(repairParagraphProperties(tracked, doc('<w:p><w:r><w:t>One</w:t></w:r></w:p>'))).toBeUndefined();
  });

  it('a paragraph that gains a named style where the original had none (the engine drops the style): both Word files now pass', async () => {
    const pair = await buildPair(pairById('one-word-change'));
    const modified = await withClauseStyle(pair.modified, '<w:pPr><w:spacing w:before="240"/><w:jc w:val="both"/></w:pPr><w:rPr><w:i/></w:rPr>');
    // The fault is real: straight from the engine, the tracked-changes file is refused.
    const raw = await compareDocuments(pair.original, modified, { authorName: 'Redline' });
    expect(await failed(checkTrackedChanges({ original: pair.original, modified, output: raw }))).not.toEqual([]);
    const { tracked } = await bothPass(pair.original, modified);
    expect(decoder.decode((await unzip(tracked)).get('word/document.xml')!)).toContain('<w:pStyle w:val="Clause"/>');
  });
});

/** Found on real contracts (Stage 7): things inside a Word file that Word itself accepts and the Checker refused. */
describe('files as other programs write them', () => {
  it('a link that names a part with different capitals (as a document-management system writes it) is accepted; a link to a part that is not there is still refused', async () => {
    const pair = await buildPair(pairById('several-changes'));
    const { tracked } = await bothPass(pair.original, pair.modified);
    // A custom data part, linked from the main text with different capitals in the folder name.
    const withData = async (target: string): Promise<Uint8Array> => {
      const files = await unzip(tracked);
      files.set('customXml/item1.xml', encoder.encode('<data/>'));
      const rels = decoder.decode(files.get('word/_rels/document.xml.rels')!);
      files.set('word/_rels/document.xml.rels', encoder.encode(rels.replace('</Relationships>', `<Relationship Id="rIdData" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml" Target="${target}"/></Relationships>`)));
      return zip(files);
    };
    expect(await failed(checkTrackedChanges({ original: pair.original, modified: pair.modified, output: await withData('/customXML/item1.xml') }))).toEqual([]);
    expect(await failed(checkTrackedChanges({ original: pair.original, modified: pair.modified, output: await withData('/customXML/item9.xml') }))).toContain('structure');
  });

  it("an item left by Word in its \"[trash]\" folder is not a part of the document; any other undeclared part is still refused", async () => {
    const pair = await buildPair(pairById('several-changes'));
    const { tracked } = await bothPass(pair.original, pair.modified);
    const withExtra = async (name: string): Promise<Uint8Array> => {
      const files = await unzip(tracked);
      files.set(name, new Uint8Array([1, 2, 3]));
      return zip(files);
    };
    expect(await failed(checkTrackedChanges({ original: pair.original, modified: pair.modified, output: await withExtra('[trash]/0000.dat') }))).toEqual([]);
    expect(await failed(checkTrackedChanges({ original: pair.original, modified: pair.modified, output: await withExtra('word/0000.dat') }))).toContain('structure');
  });
});

/**
 * Found on real contracts (Stage 7): a table whose cells sit in content controls (tick-box cells in forms and
 * schedules). When such a row was deleted or added, the static redline left that cell's own mark unmarked,
 * so the row did not read as removed.
 */
describe('table rows whose cells sit in content controls', () => {
  it.each(['table-row-deleted', 'table-row-added', 'table-cell-change'])('%s: both Word files pass', async (id) => {
    const pair = await buildPair(pairById(id));
    const { counts } = await bothPass(await withCellsInContentControls(pair.original), await withCellsInContentControls(pair.modified));
    expect(counts.insertions + counts.deletions).toBeGreaterThan(0);
  });
});

/**
 * Found on real contracts (Stage 7): a table whose rows sit in "repeating section" content controls. The engine
 * loses the table's ordinary rows and does not mark changes in the wrapped ones.
 */
describe('table rows wrapped in content controls', () => {
  it('the wrappers are taken off the copy given to the engine, and nothing else is touched', async () => {
    const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
    const row = (text: string): string => `<w:tr><w:tc><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc></w:tr>`;
    const control = (inner: string, id: number): string => `<w:sdt><w:sdtPr><w:id w:val="${id}"/></w:sdtPr><w:sdtContent>${inner}</w:sdtContent></w:sdt>`;
    const inCell = `<w:tr><w:tc>${control('<w:p><w:r><w:t>kept control</w:t></w:r></w:p>', 3)}</w:tc></w:tr>`;
    const part = `<w:document ${W}><w:body><w:tbl>${row('Heading')}${control(control(row('One'), 2), 1)}${inCell}</w:tbl>${control('<w:p><w:r><w:t>outside</w:t></w:r></w:p>', 4)}</w:body></w:document>`;
    const out = unwrapRowControlsInPart(part);
    expect(out).toContain(`<w:tbl>${row('Heading')}${row('One')}${inCell}</w:tbl>`);
    expect(out).toContain('w:val="3"'); // a control inside a cell stays
    expect(out).toContain('w:val="4"'); // and one round a paragraph
    const plain = (await buildPair(pairById('table-row-added'))).original;
    expect(await unwrapRowControls(plain)).toBe(plain);
  });

  it.each(['table-row-added', 'table-row-deleted', 'table-cell-change'])('%s: the engine goes wrong on such a table; the product passes both Word files', async (id) => {
    const pair = await buildPair(pairById(id));
    const original = await withRowsInContentControls(pair.original);
    const modified = await withRowsInContentControls(pair.modified);
    const { counts } = await bothPass(original, modified);
    expect(counts.insertions + counts.deletions).toBeGreaterThan(0);
    // The same changes are counted as for the same tables without the wrappers.
    expect(counts).toEqual((await bothPass(pair.original, pair.modified)).counts);
  });

  it('the fault is real: straight from the engine, at least one of these tables is refused by the Checker', async () => {
    let refused = 0;
    for (const id of ['table-row-added', 'table-row-deleted', 'table-cell-change']) {
      const pair = await buildPair(pairById(id));
      const original = await withRowsInContentControls(pair.original);
      const modified = await withRowsInContentControls(pair.modified);
      const raw = await compareDocuments(original, modified, { authorName: 'Redline' });
      if ((await failed(checkTrackedChanges({ original, modified, output: raw }))).length > 0) refused++;
    }
    expect(refused).toBeGreaterThan(0);
  });
});

/**
 * Found on a real contract (Stage 7): a document whose own text (there, just a paragraph mark) is formatted red
 * and struck through, which is how the redline shows deletions. The owner's rule for blue, applied to red: such a
 * redline shows its deletions in dark orange and says so in the legend; any other redline stays red.
 */
describe('a document that itself uses red strikethrough', () => {
  const ALT_LEGEND = 'Deleted text (shown in dark orange, because this document itself contains red struck-through text)';
  const bodyOf = async (docx: Uint8Array): Promise<string> => decoder.decode((await unzip(docx)).get('word/document.xml')!);
  const deletedIn = (xml: string, colour: string): number => (xml.match(new RegExp(`<w:strike/><w:color w:val="${colour}"/>`, 'g')) ?? []).length;

  for (const [which, shape] of [
    ['a stretch of text in the modified document', { original: undefined, modified: 'text' }],
    ['a stretch of text in the original only', { original: 'text', modified: undefined }],
    ['only a paragraph mark, in both documents (as in the real contract)', { original: 'paragraph mark', modified: 'paragraph mark' }],
  ] as const) {
    it(`${which}: deletions are dark orange, the legend says why, and both Word files pass`, async () => {
      const pair = await buildPair(pairById('several-changes'));
      const original = shape.original ? await withRedStrikethrough(pair.original, shape.original) : pair.original;
      const modified = shape.modified ? await withRedStrikethrough(pair.modified, shape.modified) : pair.modified;
      const { staticRedline, counts } = await bothPass(original, modified);
      const xml = await bodyOf(staticRedline);
      expect(xml).toContain(ALT_LEGEND);
      expect(deletedIn(xml, 'C55A11')).toBeGreaterThan(1); // the legend sample and the deletions
      expect(counts.deletions).toBeGreaterThan(0);
      // The document's own red struck-through text is still there, untouched, where the modified document has it.
      if (shape.modified) expect(deletedIn(xml, 'FF0000')).toBeGreaterThan(0);
    });
  }

  it('an ordinary pair stays red, with the ordinary legend', async () => {
    const pair = await buildPair(pairById('several-changes'));
    const { staticRedline } = await bothPass(pair.original, pair.modified);
    const xml = await bodyOf(staticRedline);
    expect(xml).not.toContain('dark orange');
    expect(xml).not.toContain('C55A11');
    expect(deletedIn(xml, 'FF0000')).toBeGreaterThan(1);
  });

  it('text that the upload\'s own tracked changes delete, and a footer nothing refers to, do not count (review finding): the redline stays red and passes', async () => {
    const pair = await buildPair(pairById('several-changes'));
    const A = 'w:id="900" w:author="x" w:date="2000-01-01T00:00:00Z"';
    const look = '<w:rPr><w:strike/><w:color w:val="FF0000"/></w:rPr>';
    // Red struck-through text inside a tracked deletion of the upload's own.
    const withDeleted = (docx: Uint8Array): Promise<Uint8Array> => editPart(docx, /^word\/document\.xml$/, (xml) => xml.replace('<w:p>', `<w:p><w:del ${A}><w:r>${look}<w:delText>struck before</w:delText></w:r></w:del>`));
    // Red struck-through text in a footer that no section refers to.
    const withFooter = async (docx: Uint8Array): Promise<Uint8Array> => {
      const files = await unzip(docx);
      files.set('word/footer9.xml', encoder.encode(`<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r>${look}<w:t>Left over</w:t></w:r></w:p></w:ftr>`));
      files.set('[Content_Types].xml', encoder.encode(decoder.decode(files.get('[Content_Types].xml')!).replace('</Types>', '<Override PartName="/word/footer9.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/></Types>')));
      files.set('word/_rels/document.xml.rels', encoder.encode(decoder.decode(files.get('word/_rels/document.xml.rels')!).replace('</Relationships>', '<Relationship Id="rIdLeft" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer9.xml"/></Relationships>')));
      return zip(files);
    };
    for (const [original, modified] of [
      [await withDeleted(pair.original), pair.modified],
      [await withFooter(pair.original), pair.modified],
      [pair.original, await withFooter(pair.modified)],
    ] as const) {
      const xml = await bodyOf((await bothPass(original, modified)).staticRedline);
      expect(xml).not.toContain('dark orange');
    }
    // The same in the modified upload (the redline is told, as the page tells it, that this upload had tracked changes).
    const modified = await withDeleted(pair.modified);
    const made = await makeStaticRedline(await produceTrackedChanges(pair.original, modified), pair.original, { ...COVER, trackedChangesIn: 'modified' });
    expect(await failed(checkStaticRedline({ original: pair.original, modified, output: made.docx, ...COVER }))).toEqual([]);
    expect(await bodyOf(made.docx)).not.toContain('dark orange');
  });

  it('a document with both looks of its own gets both alternative colours, and passes', async () => {
    const pair = await buildPair(pairById('several-changes'));
    const modified = await withBlueDoubleUnderline(await withRedStrikethrough(pair.modified, 'paragraph mark'));
    const xml = await bodyOf((await bothPass(pair.original, modified)).staticRedline);
    expect(xml).toContain(ALT_LEGEND);
    expect(xml).toContain('shown in magenta');
  });

  it('the wrong colour either way, or a legend that does not match, is caught', async () => {
    const pair = await buildPair(pairById('several-changes'));
    const redModified = await withRedStrikethrough(pair.modified);
    // A redline that should be dark orange, made as if nothing were unusual (red): the document's own red text reads as a deletion.
    const tracked = await produceTrackedChanges(pair.original, redModified);
    const madeRed = await makeStaticRedline(tracked, pair.original, { ...COVER, style: { deletedColor: 'FF0000' } });
    expect(await failed(checkStaticRedline({ original: pair.original, modified: redModified, output: madeRed.docx, ...COVER }))).not.toEqual([]);
    // A redline in dark orange where nothing calls for it.
    const plainTracked = await produceTrackedChanges(pair.original, pair.modified);
    const madeOrange = await makeStaticRedline(plainTracked, pair.original, { ...COVER, style: { deletedColor: 'C55A11' } });
    expect(await failed(checkStaticRedline({ original: pair.original, modified: pair.modified, output: madeOrange.docx, ...COVER }))).not.toEqual([]);
    // The right colour but the ordinary legend text.
    const { staticRedline } = await bothPass(pair.original, redModified);
    const wrongLegend = await editPart(staticRedline, /^word\/document\.xml$/, (xml) => xml.replace(ALT_LEGEND, 'Deleted text'));
    expect(await failed(checkStaticRedline({ original: pair.original, modified: redModified, output: wrongLegend, ...COVER }))).toContain('cover-page');
  });
});

/**
 * Found on real contracts (Stage 7): the last clause in a table cell is deleted and words are deleted from the
 * middle of the clause before it. The engine then records the removed paragraph break on the clause before (a cell's last mark cannot
 * be removed). The clause that went is still the last one: it shows its old number, and the one before keeps its own.
 */
describe('the last clause in a cell deleted, the one before it shortened', () => {
  it('both Word files pass; the surviving clause keeps its number and the deleted one shows its old number', async () => {
    for (const id of ['numbered-clause-deleted', 'numbered-clause-moved']) {
      const base = (await buildPair(pairById(id))).original;
      const original = await withNumberedClausesInCell(base);
      const modified = await withNumberedClausesInCell(base, 'last deleted, second shortened');
      const { staticRedline } = await bothPass(original, modified);
      const xml = decoder.decode((await unzip(staticRedline)).get('word/document.xml')!);
      const inCell = /<w:tc>.*?<\/w:tc>/s.exec(xml)![0].match(/<w:p>.*?<\/w:p>/gs)!;
      expect(inCell).toHaveLength(3);
      expect(inCell.map((paragraph) => paragraph.includes('<w:numId w:val="0"/>'))).toEqual([false, false, true]);
      expect(inCell.map((paragraph) => paragraph.includes('RedlineOldNumber'))).toEqual([false, false, true]);
    }
  });
});

/** Found on a real contract (Stage 7): a template whose glossary of building blocks has a style sheet of its own. */
describe('a document whose glossary has its own style sheet', () => {
  it('both Word files pass, old clause numbers are shown, and the redline\'s styles go into the document\'s own style sheet', async () => {
    const pair = await buildPair(pairById('numbered-clause-deleted'));
    const { staticRedline, counts } = await bothPass(await withGlossary(pair.original), await withGlossary(pair.modified));
    expect(counts).toEqual((await bothPass(pair.original, pair.modified)).counts);
    const files = await unzip(staticRedline);
    expect(decoder.decode(files.get('word/document.xml')!)).toContain('RedlineOldNumber');
    expect(decoder.decode(files.get('word/styles.xml')!)).toContain('w:styleId="RedlineOldNumber"');
    expect(decoder.decode(files.get('word/glossary/styles.xml')!)).not.toContain('Redline');
  });
});

/**
 * Found on a real contract (Stage 7): a stray pen mark. Word stores handwriting as pen strokes and, alongside, as a
 * picture of them; the engine keeps the picture. The strokes hold no text, so the picture is what is compared.
 */
describe('a pen mark (handwriting) in the document', () => {
  it('both Word files pass, and the pen mark\'s picture is in both', async () => {
    const pair = await buildPair(pairById('one-word-change'));
    const { tracked, staticRedline, counts } = await bothPass(await withInk(pair.original), await withInk(pair.modified));
    expect(counts).toEqual((await bothPass(pair.original, pair.modified)).counts);
    for (const output of [tracked, staticRedline]) expect(decoder.decode((await unzip(output)).get('word/document.xml')!)).toMatch(/<(a:)?blip /);
  });

  it('the pen mark is really checked: a different picture in the output, or a pen mark only one document has, is refused', async () => {
    const pair = await buildPair(pairById('one-word-change'));
    const original = await withInk(pair.original);
    const modified = await withInk(pair.modified);
    const tracked = await produceTrackedChanges(original, modified);
    // Another picture in place of the pen mark's.
    const files = await unzip(tracked);
    const pictures = [...files.keys()].filter((name) => /^word\/media\/.*\.png$/.test(name));
    expect(pictures.length).toBeGreaterThan(0);
    for (const picture of pictures) files.set(picture, Uint8Array.from([...files.get(picture)!, 0]));
    const failures = await failed(checkTrackedChanges({ original, modified, output: await zip(files) }));
    expect(failures).toContain('accept-all');
    expect(failures).toContain('reject-all');
    // The pen mark missing from the output's reading of one side.
    expect(await failed(checkTrackedChanges({ original, modified: pair.modified, output: tracked }))).toContain('accept-all');
    expect(await failed(checkTrackedChanges({ original: pair.original, modified, output: tracked }))).toContain('reject-all');
  });

  it('an output that keeps the pen strokes must keep them unchanged (review finding): changed strokes are refused, kept ones pass', async () => {
    const pair = await buildPair(pairById('one-word-change'));
    const original = await withInk(pair.original);
    // An output that still holds the strokes: the upload itself, compared with itself.
    expect(await failed(checkTrackedChanges({ original, modified: original, output: original }))).toEqual([]);
    const changed = await editPart(original, /^word\/ink\/ink1\.xml$/, (xml) => xml.replace("0 1,'0'0", "5 9,'3'7"));
    const failures = await failed(checkTrackedChanges({ original, modified: original, output: changed }));
    expect(failures).toContain('accept-all');
    expect(failures).toContain('reject-all');
  });

  it('the alternative forms of ordinary content must still say the same: text that only one form has is refused', async () => {
    const pair = await buildPair(pairById('one-word-change'));
    const odd = await editPart(await withInk(pair.original), /^word\/document\.xml$/, (xml) => xml.replace('</mc:Fallback>', '<w:t>only in one form</w:t></mc:Fallback>'));
    expect(await failed(checkTrackedChanges({ original: odd, modified: pair.modified, output: await produceTrackedChanges(pair.original, pair.modified) }))).toContain('completeness');
  });
});

/**
 * Found on a real contract (Stage 7): the first word of a bold heading is replaced, and the engine writes the new
 * word in a paragraph of its own whose mark is deleted. On "accept all" that word joins the heading and is bold,
 * because a joined paragraph has the properties of the mark that stays.
 */
describe('text that joins a paragraph of another style when its paragraph mark is removed', () => {
  const A = 'w:author="Redline" w:date="2000-01-01T00:00:00Z"';
  const heading = (inner: string): string => `<w:p><w:pPr><w:pStyle w:val="BoldHead"/></w:pPr>${inner}</w:p>`;
  const t = (text: string): string => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
  const intro = `<w:p>${t('The parties agree as follows.')}</w:p>`;
  const bodies = {
    original: `${intro}<w:p/>${heading(t('Termination of contract'))}`,
    modified: `${intro}<w:p/>${heading(t('without of contract'))}`,
    // As the engine wrote it on the real contract.
    tracked:
      `${intro}<w:p><w:pPr><w:rPr><w:ins w:id="1" ${A}/></w:rPr></w:pPr></w:p><w:p><w:pPr><w:rPr><w:del w:id="2" ${A}/></w:rPr></w:pPr><w:ins w:id="3" ${A}>${t('without')}</w:ins></w:p>` +
      heading(`<w:del w:id="4" ${A}><w:r><w:delText>Termination</w:delText></w:r></w:del>${t(' of contract')}`),
  };
  const build = async (body: string): Promise<Uint8Array> => {
    const base = (await buildPair(pairById('one-word-change'))).original;
    const styled = await editPart(base, /^word\/styles\.xml$/, (xml) => xml.replace('</w:styles>', '<w:style w:type="paragraph" w:customStyle="1" w:styleId="BoldHead"><w:name w:val="Bold Head"/><w:rPr><w:b/></w:rPr></w:style></w:styles>'));
    return editPart(styled, /^word\/document\.xml$/, (xml) => xml.replace(/(<w:body>).*?(<w:sectPr\b)/s, `$1${body}$2`));
  };

  it('the tracked-changes file passes: the new word is read as bold, as it is once the changes are accepted', async () => {
    const [original, modified, tracked] = await Promise.all([build(bodies.original), build(bodies.modified), build(bodies.tracked)]);
    expect(await failed(checkTrackedChanges({ original, modified, output: tracked }))).toEqual([]);
    // And it is really checked: if the heading it joins is not bold, the same file is refused.
    const plainHeading = await editPart(tracked, /^word\/document\.xml$/, (xml) => xml.replace('<w:pPr><w:pStyle w:val="BoldHead"/></w:pPr>', ''));
    expect(await failed(checkTrackedChanges({ original, modified, output: plainHeading }))).toContain('formatting');
  });

  it('a bookmark between the two paragraphs changes nothing (review finding): both files still pass', async () => {
    const between = (body: string): string => body.replace('</w:p><w:p><w:pPr><w:pStyle w:val="BoldHead"/>', '</w:p><w:bookmarkStart w:id="77" w:name="between"/><w:bookmarkEnd w:id="77"/><w:p><w:pPr><w:pStyle w:val="BoldHead"/>');
    expect(between(bodies.tracked)).not.toBe(bodies.tracked);
    const [original, modified, tracked] = await Promise.all([build(bodies.original), build(bodies.modified), build(between(bodies.tracked))]);
    expect(await failed(checkTrackedChanges({ original, modified, output: tracked }))).toEqual([]);
    const made = await makeStaticRedline(tracked, original, COVER);
    expect(await failed(checkStaticRedline({ original, modified, output: made.docx, ...COVER }))).toEqual([]);
  });

  it('the static redline shows the new word bold, as the modified document has it, and passes', async () => {
    const [original, modified, tracked] = await Promise.all([build(bodies.original), build(bodies.modified), build(bodies.tracked)]);
    const made = await makeStaticRedline(tracked, original, COVER);
    expect(await failed(checkStaticRedline({ original, modified, output: made.docx, ...COVER }))).toEqual([]);
    const xml = decoder.decode((await unzip(made.docx)).get('word/document.xml')!);
    expect(/<w:r><w:rPr>(?:(?!<\/w:rPr>).)*<w:b\/>(?:(?!<\/w:rPr>).)*<\/w:rPr><w:t[^>]*>without<\/w:t>/.test(xml)).toBe(true);
    // A static redline that shows the word plain is refused.
    const plain = await editPart(made.docx, /^word\/document\.xml$/, (doc) => doc.replace(/(<w:r><w:rPr>(?:(?!<\/w:rPr>).)*)<w:b\/>((?:(?!<\/w:rPr>).)*<\/w:rPr><w:t[^>]*>without<\/w:t>)/, '$1$2'));
    expect(await failed(checkStaticRedline({ original, modified, output: plain, ...COVER }))).toContain('formatting');
  });
});

/**
 * Found on a real contract (Stage 7): several paragraphs added after the only paragraph of a table cell, the last
 * of them empty. A cell's final paragraph mark cannot be shown as added or removed, so the whole run of marks is
 * shown one paragraph up; otherwise the original would read as having an extra empty paragraph in the cell.
 */
describe('several paragraphs added or removed at the end of a table cell', () => {
  const LINES = [
    ['', 'Must meet the requirement', ''],
    ['Must meet the requirement', 'Evidence to be attached'],
    ['', ''],
    ['', '', 'Must meet the requirement', '', '', 'Must meet the requirement', ''],
  ];

  it.each(LINES.map((lines) => [JSON.stringify(lines), lines] as const))('added %s: both Word files pass', async (_, lines) => {
    const base = (await buildPair(pairById('one-word-change'))).original;
    const { counts } = await bothPass(await withParagraphsAtCellEnd(base, []), await withParagraphsAtCellEnd(base, lines));
    expect(counts.insertions).toBeGreaterThan(0);
    expect(counts.deletions).toBe(0);
  });

  it('several numbered clauses removed from, or added at, the end of a cell (review finding): both Word files pass', async () => {
    const base = (await buildPair(pairById('numbered-clause-deleted'))).original;
    const all = await withNumberedClausesInCell(base);
    const firstOnly = await editPart(all, /^word\/document\.xml$/, (xml) => {
      const cell = /<w:tc>.*?<\/w:tc>/s.exec(xml)![0];
      const clauses = cell.match(/<w:p>.*?<\/w:p>/gs)!;
      expect(clauses).toHaveLength(3);
      return xml.replace(cell, cell.replace(clauses[1]! + clauses[2]!, ''));
    });
    expect((await bothPass(all, firstOnly)).counts.deletions).toBeGreaterThan(0);
    expect((await bothPass(firstOnly, all)).counts.insertions).toBeGreaterThan(0);
  });

  it.each(LINES.map((lines) => [JSON.stringify(lines), lines] as const))('removed %s: both Word files pass', async (_, lines) => {
    const base = (await buildPair(pairById('one-word-change'))).original;
    const { counts } = await bothPass(await withParagraphsAtCellEnd(base, lines), await withParagraphsAtCellEnd(base, []));
    expect(counts.deletions).toBeGreaterThan(0);
    expect(counts.insertions).toBe(0);
  });
});

/** Found on a real document (Stage 7): headers and footers left in the file that no section refers to any more. */
describe('a footer that no section refers to', () => {
  const withLeftover = async (docx: Uint8Array): Promise<Uint8Array> => {
    const files = await unzip(docx);
    files.set('word/footer9.xml', encoder.encode('<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>Left over from a deleted section</w:t></w:r></w:p></w:ftr>'));
    files.set('[Content_Types].xml', encoder.encode(decoder.decode(files.get('[Content_Types].xml')!).replace('</Types>', '<Override PartName="/word/footer9.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/></Types>')));
    const rels = decoder.decode(files.get('word/_rels/document.xml.rels')!);
    files.set('word/_rels/document.xml.rels', encoder.encode(rels.replace('</Relationships>', '<Relationship Id="rIdLeft" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer9.xml"/></Relationships>')));
    return zip(files);
  };

  it('in an upload it is not document content: both Word files pass, and neither carries the leftover', async () => {
    const pair = await buildPair(pairById('footer-change'));
    for (const [original, modified] of [[await withLeftover(pair.original), pair.modified], [pair.original, await withLeftover(pair.modified)], [await withLeftover(pair.original), await withLeftover(pair.modified)]] as const) {
      const { tracked, staticRedline } = await bothPass(original, modified);
      for (const output of [tracked, staticRedline]) {
        const files = await unzip(output);
        expect([...files.keys()].filter((name) => /footer9/.test(name))).toEqual([]);
        expect(decoder.decode(files.get('word/_rels/document.xml.rels')!)).not.toContain('footer9');
        expect(decoder.decode(files.get('[Content_Types].xml')!)).not.toContain('footer9');
      }
    }
  });

  it('in an output it is still refused, and so is an output that loses a footer that is in use', async () => {
    const pair = await buildPair(pairById('footer-change'));
    const { tracked } = await bothPass(pair.original, pair.modified);
    expect(await failed(checkTrackedChanges({ original: pair.original, modified: pair.modified, output: await withLeftover(tracked) }))).toContain('completeness');
    const dropped = await editPart(tracked, /^word\/document\.xml$/, (xml) => xml.replace(/<w:footerReference\b[^>]*\/>/g, ''));
    expect(await failed(checkTrackedChanges({ original: pair.original, modified: pair.modified, output: dropped }))).toContain('accept-all');
  });
});

/**
 * Found on real contracts (held-out run 1, about 17 documents of one family): the engine gathers a shape type (a
 * template for legacy shapes, which draws nothing) into a picture element of its own. A reader sees no difference.
 */
describe('a legacy picture element that holds only a shape template', () => {
  const V = 'xmlns:v="urn:schemas-microsoft-com:vml"';
  const template = `<v:shapetype ${V} id="_x0000_t202" coordsize="21600,21600" path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/></v:shapetype>`;
  const add = (docx: Uint8Array, inside: string): Promise<Uint8Array> => editPart(docx, /^word\/document\.xml$/, (xml) => xml.replace('<w:r>', `<w:r><w:pict>${inside}</w:pict></w:r><w:r>`));

  it('draws nothing, so an output that has one more of them than the uploads is passed', async () => {
    const pair = await buildPair(pairById('one-word-change'));
    const tracked = await produceTrackedChanges(pair.original, pair.modified);
    expect(await failed(checkTrackedChanges({ original: pair.original, modified: pair.modified, output: await add(tracked, template) }))).toEqual([]);
  });

  it('an actual shape is still an object: an output with one the uploads do not have is refused', async () => {
    const pair = await buildPair(pairById('one-word-change'));
    const tracked = await produceTrackedChanges(pair.original, pair.modified);
    const shape = `${template}<v:shape ${V} type="#_x0000_t202" style="width:10pt;height:10pt"/>`;
    const failures = await failed(checkTrackedChanges({ original: pair.original, modified: pair.modified, output: await add(tracked, shape) }));
    expect(failures).toContain('accept-all');
  });
});

/**
 * Found on real contracts (held-out run 1, about 10 documents of one family): an icon stored as text (SVG). The
 * engine reads it as XML and writes it out again laid out differently, so the picture file is no longer the uploaded one.
 */
describe('a picture stored as text (SVG)', () => {
  const svgsIn = async (docx: Uint8Array): Promise<Uint8Array[]> => [...(await unzip(docx))].filter(([name]) => name.endsWith('.svg')).map(([, bytes]) => bytes);

  it('both Word files pass, and the picture in them is the uploaded file, byte for byte', async () => {
    const pair = await buildPair(pairById('one-word-change'));
    const { tracked, staticRedline, counts } = await bothPass(await withSvgPicture(pair.original), await withSvgPicture(pair.modified));
    expect(counts).toEqual((await bothPass(pair.original, pair.modified)).counts);
    for (const output of [tracked, staticRedline]) {
      const svgs = await svgsIn(output);
      expect(svgs.length).toBeGreaterThan(0);
      for (const svg of svgs) expect(decoder.decode(svg)).toBe(decoder.decode(SVG_PICTURE));
    }
  });

  it('the picture is really checked: an output whose picture draws something else is refused', async () => {
    const pair = await buildPair(pairById('one-word-change'));
    const original = await withSvgPicture(pair.original);
    const modified = await withSvgPicture(pair.modified);
    const files = await unzip(await produceTrackedChanges(original, modified));
    for (const [name, bytes] of files) if (name.endsWith('.svg')) files.set(name, new TextEncoder().encode(decoder.decode(bytes).replace('#558ED5', '#FF0000')));
    const failures = await failed(checkTrackedChanges({ original, modified, output: await zip(files) }));
    expect(failures).toContain('accept-all');
    expect(failures).toContain('reject-all');
  });

  it('a picture that draws something else is not mistaken for the uploaded one: each document keeps its own', async () => {
    const pair = await buildPair(pairById('one-word-change'));
    const other = new TextEncoder().encode(decoder.decode(SVG_PICTURE).replace(/#558ED5/g, '#00AA00'));
    const original = await withSvgPicture(pair.original);
    const modified = await withSvgPicture(pair.modified, other);
    const tracked = await produceTrackedChanges(original, modified);
    expect(await failed(checkTrackedChanges({ original, modified, output: tracked }))).toEqual([]);
    expect((await svgsIn(tracked)).map((svg) => decoder.decode(svg)).sort()).toEqual([decoder.decode(SVG_PICTURE), decoder.decode(other)].sort());
  });
});

/**
 * Found on a real contract (held-out run 1): a cell holding two alternatives with "or" between them, both
 * alternatives removed. The static redline shows the removal of the cell's last paragraph on the mark of "or",
 * next to the removed mark of the first alternative; "or" itself stays, so the run of removed paragraphs stops there.
 */
describe('a kept line between two removed paragraphs at the end of a cell', () => {
  for (const [name, before, after] of [
    ['two alternatives removed, "or" kept', ['Clause eighteen of the general conditions applies here.', 'or', 'Clause eighteen of the general conditions has no application.'], ['or']],
    ['the same with a line kept above', ['Choose one.', 'Clause eighteen of the general conditions applies here.', 'or', 'Clause eighteen of the general conditions has no application.'], ['Choose one.', 'or']],
    ['two alternatives added around "or"', ['or'], ['Clause eighteen of the general conditions applies here.', 'or', 'Clause eighteen of the general conditions has no application.']],
  ] as const) {
    it(`both Word files pass: ${name}`, async () => {
      const pair = await buildPair(pairById('one-word-change'));
      await bothPass(await withParagraphsAtCellEnd(pair.original, [...before]), await withParagraphsAtCellEnd(pair.original, [...after]));
    });
  }
});
