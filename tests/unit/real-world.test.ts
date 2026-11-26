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
import { buildDocx, buildPair, pairById } from '../fixtures/pairs';
import type { DocSpec } from '../fixtures/pairs';
import { freshRelationshipIds, renameInPart } from '../../src/producer/fresh-relationship-ids';
import { repairSectionReferences } from '../../src/producer/repair-section-references';
import { compareDocuments } from 'docxodus/core';
import { repairNotes } from '../../src/producer/repair-notes';
import { repairParagraphProperties } from '../../src/producer/repair-paragraphs';
import { unwrapRowControls, unwrapRowControlsInPart } from '../../src/producer/unwrap-row-controls';
import { withBlueDoubleUnderline, withCellsInContentControls, withGlossary, withInk, withNestedTable, withSvgPicture, SVG_PICTURE, withParagraphsAtCellEnd, withRedStrikethrough, withClauseStyle, withRowsInContentControls, withNumberedClausesInCell, withDocumentWideSpaces, withNoteInTable, withSections } from '../fixtures/real-world';

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
async function bothPass(original: Uint8Array, modified: Uint8Array, trackedChangesIn?: 'original' | 'modified' | 'both'): Promise<Outputs> {
  const tracked = await produceTrackedChanges(original, modified);
  const trackedReport = await checkTrackedChanges({ original, modified, output: tracked });
  expect(trackedReport.checks.filter((c) => !c.passed)).toEqual([]);
  // (An upload that carried tracked changes is named on the cover page.)
  const made = await makeStaticRedline(tracked, original, trackedChangesIn ? { ...COVER, trackedChangesIn } : COVER);
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

  // (Until engine version 17 this test asserted the fault: the engine lost ordinary rows of such tables. Version 17
  // tracks changes inside row-level content controls, and the same three tables now come straight from the engine
  // in a form the Checker accepts. The unwrapping step is kept until removing it has been measured on real files.)
  it('since engine version 17 the fault is gone: straight from the engine, none of these tables is refused', async () => {
    let refused = 0;
    for (const id of ['table-row-added', 'table-row-deleted', 'table-cell-change']) {
      const pair = await buildPair(pairById(id));
      const original = await withRowsInContentControls(pair.original);
      const modified = await withRowsInContentControls(pair.modified);
      const raw = await compareDocuments(original, modified, { authorName: 'Redline' });
      if ((await failed(checkTrackedChanges({ original, modified, output: raw }))).length > 0) refused++;
    }
    expect(refused).toBe(0);
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

  it('nor does one that holds nothing at all (what the engine leaves where it took a template from)', async () => {
    const pair = await buildPair(pairById('one-word-change'));
    const tracked = await produceTrackedChanges(pair.original, pair.modified);
    const emptied = await editPart(tracked, /^word\/document\.xml$/, (xml) => xml.replace('<w:r>', '<w:r><w:pict/></w:r><w:r>'));
    expect(await failed(checkTrackedChanges({ original: pair.original, modified: pair.modified, output: emptied }))).toEqual([]);
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

/** Review findings on the first batch of held-out fixes. */
describe('held-out fixes, review findings', () => {
  it('an SVG picture that cannot be read as XML does not stop the comparison: it is left as it is', async () => {
    const pair = await buildPair(pairById('one-word-change'));
    const broken = new TextEncoder().encode('<svg xmlns="http://www.w3.org/2000/svg"><path></svg>');
    await expect(produceTrackedChanges(await withSvgPicture(pair.original, broken), await withSvgPicture(pair.modified, broken))).resolves.toBeInstanceOf(Uint8Array);
    await expect(produceTrackedChanges(await withSvgPicture(pair.original, new Uint8Array()), await withSvgPicture(pair.modified, new Uint8Array()))).resolves.toBeInstanceOf(Uint8Array);
  });

  it('a legacy shape must keep the template it is drawn from: an output that loses the template is refused', async () => {
    const V = 'xmlns:v="urn:schemas-microsoft-com:vml"';
    const template = `<w:r><w:pict><v:shapetype ${V} id="_x0000_t202" coordsize="21600,21600" path="m,l,21600r21600,l21600,xe"/></w:pict></w:r>`;
    const shape = `<w:r><w:pict><v:shape ${V} type="#_x0000_t202" style="width:10pt;height:10pt"/></w:pict></w:r>`;
    const pair = await buildPair(pairById('one-word-change'));
    const add = (docx: Uint8Array): Promise<Uint8Array> => editPart(docx, /^word\/document\.xml$/, (xml) => xml.replace('<w:r>', `${template}${shape}<w:r>`));
    const original = await add(pair.original);
    const modified = await add(pair.modified);
    const whole = await add(await produceTrackedChanges(pair.original, pair.modified));
    expect(await failed(checkTrackedChanges({ original, modified, output: whole }))).toEqual([]);
    const lost = await editPart(whole, /^word\/document\.xml$/, (xml) => xml.replace(template, ''));
    const failures = await failed(checkTrackedChanges({ original, modified, output: lost }));
    expect(failures).toContain('accept-all');
    expect(failures).toContain('reject-all');
  });
});

/**
 * Found on real contracts (held-out run 1, seven leases): the front page is a table whose last row holds a table.
 * When a row is added above it, the engine removes the inner table from that row (leaving the new row's text
 * there) and puts it back in a new row of its own. The inner table's rows then carry no flag of their own, and the
 * empty paragraph after the removed table has no neighbouring paragraph mark to show its removal on.
 */
describe('a table inside the last row of a table, with rows added or removed around it', () => {
  const OUTER = ['Dated', '', 'The landlord and the tenant and the guarantor', '', 'Lease relating to premises known as the address', '', 'Whole office open market rent'];
  const INNER = ['Term of years', 'Initial rent payable', 'Rent free period', 'Rent review pattern', 'Rent review basis'];
  const INNER_MORE = [INNER[0]!, 'Foregoing stated that notwithstanding law', ...INNER.slice(1)];
  const NEW_ROW = 'Liability in schedule without notice writing';
  for (const [name, before, after] of [
    ['a row added above it and a row added inside it', [OUTER, INNER], [[...OUTER, NEW_ROW], INNER_MORE]],
    ['the reverse: both rows removed', [[...OUTER, NEW_ROW], INNER_MORE], [OUTER, INNER]],
    ['rows added, removed and emptied above it', [OUTER, INNER], [[OUTER[0]!, OUTER[1]!, OUTER[2]!, 'To provided writing', '', '', OUTER[6]!, 'Liability in schedule'], INNER_MORE]],
    ['two rows added above it', [OUTER, INNER], [[...OUTER, NEW_ROW, 'Second new row here'], INNER_MORE]],
    ['the inner table rewritten', [OUTER, INNER], [[...OUTER, 'Liability'], ['Alpha', 'Beta', 'Gamma', 'Delta', 'Epsilon', 'Zeta']]],
  ] as const) {
    it(`both Word files pass: ${name}`, async () => {
      const pair = await buildPair(pairById('one-word-change'));
      await bothPass(await withNestedTable(pair.original, [...before[0]], [...before[1]]), await withNestedTable(pair.original, [...after[0]], [...after[1]]));
    });
  }
});

/**
 * Found on a real contract (held-out run 1, a draft carrying tracked changes): a cell's one line is replaced, and
 * the engine writes the old line removed together with its mark, then the new line added with the cell's last mark.
 * The cell has one paragraph in each version; shown as two paragraphs, one version always read as two.
 */
describe('the one line of a cell replaced, written as a removed paragraph and an added one', () => {
  const OLD = 'Total exercise price of the option';
  const NEW = 'Type of option granted to the holder';
  const run = (text: string): string => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
  const change = (tag: 'ins' | 'del', id: number): string => `<w:${tag} w:id="${9000 + id}" w:author="Redline" w:date="2026-11-25T00:00:00Z"`;
  const removed = (text: string): string => `<w:p><w:pPr><w:rPr>${change('del', 1)}/></w:rPr></w:pPr>${change('del', 2)}><w:r><w:delText xml:space="preserve">${text}</w:delText></w:r></w:del></w:p>`;
  const added = (text: string): string => `<w:p><w:pPr><w:rPr>${change('ins', 3)}/></w:rPr></w:pPr>${change('ins', 4)}>${run(text)}</w:ins></w:p>`;

  for (const [name, first, second, before, after] of [
    ['old line removed, new line added', removed(OLD), added(NEW), OLD, NEW],
    ['new line added, old line removed', added(NEW), removed(OLD), OLD, NEW],
  ] as const) {
    it(`the static redline passes and shows one paragraph: ${name}`, async () => {
      const pair = await buildPair(pairById('one-word-change'));
      const original = await withParagraphsAtCellEnd(pair.original, [before]);
      const modified = await withParagraphsAtCellEnd(pair.original, [after]);
      const tracked = await editPart(await withParagraphsAtCellEnd(pair.original, [OLD, NEW]), /^word\/document\.xml$/, (xml) => xml.replace(`<w:p>${run(OLD)}</w:p><w:p>${run(NEW)}</w:p>`, first + second));
      expect(await failed(checkTrackedChanges({ original, modified, output: tracked }))).toEqual([]);
      const made = await makeStaticRedline(tracked, original, COVER);
      const report = await checkStaticRedline({ original, modified, output: made.docx, ...COVER });
      expect(report.checks.filter((c) => !c.passed)).toEqual([]);
      const xml = decoder.decode((await unzip(made.docx)).get('word/document.xml')!);
      const cell = /<w:tc>.*?<\/w:tc>/s.exec(xml)![0];
      expect(cell.match(/<w:p[ >]/g)).toHaveLength(2); // the cell's first paragraph, and the replaced line as one
      expect(cell.indexOf(OLD)).toBeGreaterThan(-1);
      expect(cell.indexOf(NEW)).toBeGreaterThan(-1);
    });
  }
});

/**
 * Found on real contracts (held-out run 2, Irish and Islamic Development Bank forms, on drafts carrying tracked
 * changes): a cell's one paragraph split, with a new line typed in between, so the cell ends in an added paragraph
 * and a moved-in one. The run of additions was not seen as one run, the mark above it was left unmarked, and the
 * original read with a paragraph too many.
 */
describe('a cell that gains a typed paragraph and a moved-in paragraph at its end', () => {
  it('the static redline passes', async () => {
    const HEAD = 'Application of the foreign subsidies rule';
    const TAIL = ' of the tenth of July on distortive subsidies';
    const NEW = 'The parties shall review this annually';
    const run = (text: string): string => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
    const change = (tag: string, id: number, extra = ''): string => `<w:${tag} w:id="${9600 + id}" w:author="Redline" w:date="2026-11-26T00:00:00Z"${extra}`;
    const cellOf = (paragraphs: string): string => `<w:tbl><w:tblPr><w:tblW w:w="9000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="9000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="9000" w:type="dxa"/></w:tcPr>${paragraphs}</w:tc></w:tr></w:tbl>`;
    const base = (await buildPair(pairById('one-word-change'))).original;
    const withCell = (paragraphs: string): Promise<Uint8Array> => editPart(base, /^word\/document\.xml$/, (xml) => xml.replace(/<w:p>.*?<\/w:p>/s, (first) => first + cellOf(paragraphs)));
    const original = await withCell(`<w:p>${run(HEAD)}${run(TAIL)}</w:p>`);
    const modified = await withCell(`<w:p>${run(HEAD)}</w:p><w:p>${run(NEW)}</w:p><w:p>${run(TAIL)}</w:p>`);
    const tracked = await withCell(
      `<w:p>${run(HEAD)}${change('moveFromRangeStart', 1, ' w:name="move1"')}/>${change('moveFrom', 2)}><w:r><w:t xml:space="preserve">${TAIL}</w:t></w:r></w:moveFrom><w:moveFromRangeEnd w:id="9601"/></w:p>` +
        `<w:p><w:pPr><w:rPr>${change('ins', 3)}/></w:rPr></w:pPr>${change('ins', 4)}>${run(NEW)}</w:ins></w:p>` +
        `<w:p><w:pPr><w:rPr>${change('ins', 5)}/></w:rPr></w:pPr>${change('moveToRangeStart', 6, ' w:name="move1"')}/>${change('moveTo', 7)}>${run(TAIL)}</w:moveTo><w:moveToRangeEnd w:id="9606"/></w:p>`,
    );
    expect(await failed(checkTrackedChanges({ original, modified, output: tracked }))).toEqual([]);
    const made = await makeStaticRedline(tracked, original, COVER);
    expect((await checkStaticRedline({ original, modified, output: made.docx, ...COVER })).checks.filter((c) => !c.passed)).toEqual([]);
  });
});

/**
 * Found on real contracts (held-out run 2, Irish and Islamic Development Bank forms): the one line of a cell (a
 * form field) removed and an empty line left in its place, written as the old line removed with its mark and an
 * empty paragraph added with the cell's last mark. A rule added in the third batch wrongly left this case out.
 */
describe('the one line of a cell replaced by an empty line', () => {
  for (const mirrored of [false, true]) it(`the static redline passes and the cell holds one paragraph${mirrored ? ' (an empty line replaced by a line of text)' : ''}`, async () => {
    const OLD = 'Applicant entry goes on this line';
    const change = (tag: string, id: number): string => `<w:${tag} w:id="${9500 + id}" w:author="Redline" w:date="2026-11-26T00:00:00Z"`;
    const cellOf = (paragraphs: string): string => `<w:tbl><w:tblPr><w:tblW w:w="9000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="9000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="9000" w:type="dxa"/></w:tcPr>${paragraphs}</w:tc></w:tr></w:tbl>`;
    const base = (await buildPair(pairById('one-word-change'))).original;
    const withCell = (paragraphs: string): Promise<Uint8Array> => editPart(base, /^word\/document\.xml$/, (xml) => xml.replace(/<w:p>.*?<\/w:p>/s, (first) => first + cellOf(paragraphs)));
    const full = await withCell(`<w:p><w:r><w:t xml:space="preserve">${OLD}</w:t></w:r></w:p>`);
    const empty = await withCell('<w:p/>');
    const [original, modified] = mirrored ? [empty, full] : [full, empty];
    const tracked = await withCell(
      mirrored
        ? `<w:p><w:pPr><w:rPr>${change('ins', 1)}/></w:rPr></w:pPr>${change('ins', 2)}><w:r><w:t xml:space="preserve">${OLD}</w:t></w:r></w:ins></w:p>` +
            `<w:p><w:pPr><w:rPr>${change('del', 3)}/></w:rPr></w:pPr></w:p>`
        : `<w:p><w:pPr><w:rPr>${change('del', 1)}/></w:rPr></w:pPr>${change('del', 2)}><w:r><w:delText xml:space="preserve">${OLD}</w:delText></w:r></w:del></w:p>` +
            `<w:p><w:pPr><w:rPr>${change('ins', 3)}/></w:rPr></w:pPr></w:p>`,
    );
    expect(await failed(checkTrackedChanges({ original, modified, output: tracked }))).toEqual([]);
    const made = await makeStaticRedline(tracked, original, COVER);
    expect((await checkStaticRedline({ original, modified, output: made.docx, ...COVER })).checks.filter((c) => !c.passed)).toEqual([]);
    const cell = /<w:tc>.*?<\/w:tc>/s.exec(decoder.decode((await unzip(made.docx)).get('word/document.xml')!))![0];
    expect(cell.match(/<w:p[ >]/g)).toHaveLength(1);
  });
});

/**
 * Found on a real contract (held-out run 1, an ADB document carrying tracked changes): a cell's one line replaced
 * by twenty paragraphs. The engine writes the old line removed with its mark, then every new paragraph added with
 * its mark, the cell's last included. Left as written, the original read as two paragraphs (the old line, and an
 * empty one where the cell's last mark stayed), and the redline was withheld.
 */
describe('the one line of a cell replaced by several paragraphs', () => {
  const OLD = 'The scope of requirements is defined elsewhere';
  const NEW = ['Details of shipping documents follow', 'Three copies of the invoice', 'One certificate of origin'];
  const run = (text: string): string => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
  const change = (tag: string, id: number): string => `<w:${tag} w:id="${9300 + id}" w:author="Redline" w:date="2026-11-25T00:00:00Z"`;
  const removed = (text: string): string => `<w:p><w:pPr><w:rPr>${change('del', 1)}/></w:rPr></w:pPr>${change('del', 2)}><w:r><w:delText xml:space="preserve">${text}</w:delText></w:r></w:del></w:p>`;
  const added = (text: string, id: number, wrapper = 'ins'): string => `<w:p><w:pPr><w:rPr>${change('ins', id)}/></w:rPr></w:pPr>${change(wrapper, id + 1)}>${run(text)}</w:${wrapper}></w:p>`;
  const P = (text: string): string => `<w:p>${run(text)}</w:p>`;
  const cellOf = (paragraphs: string): string => `<w:tbl><w:tblPr><w:tblW w:w="9000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="9000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="9000" w:type="dxa"/></w:tcPr>${paragraphs}</w:tc></w:tr></w:tbl>`;

  for (const [name, lead] of [
    ['the line is all the cell held', ''],
    ['another line stands above it', P('Payment terms')],
  ] as const) {
    it(`the static redline passes and the old line shares a paragraph with the first new one: ${name}`, async () => {
      const base = (await buildPair(pairById('one-word-change'))).original;
      const withCell = (paragraphs: string): Promise<Uint8Array> => editPart(base, /^word\/document\.xml$/, (xml) => xml.replace(/<w:p>.*?<\/w:p>/s, (first) => first + cellOf(lead + paragraphs)));
      const original = await withCell(P(OLD));
      const modified = await withCell(NEW.map(P).join(''));
      const tracked = await withCell(removed(OLD) + added(NEW[0]!, 10) + added(NEW[1]!, 20) + added(NEW[2]!, 30));
      expect(await failed(checkTrackedChanges({ original, modified, output: tracked }))).toEqual([]);
      const made = await makeStaticRedline(tracked, original, COVER);
      const report = await checkStaticRedline({ original, modified, output: made.docx, ...COVER });
      expect(report.checks.filter((c) => !c.passed)).toEqual([]);
      const cell = /<w:tc>.*?<\/w:tc>/s.exec(decoder.decode((await unzip(made.docx)).get('word/document.xml')!))![0];
      expect(cell.match(/<w:p[ >]/g)).toHaveLength(NEW.length + (lead === '' ? 0 : 1));
      const joined = cell.split(/<w:p[ >]/).find((p) => p.includes(OLD))!;
      expect(joined.indexOf(OLD)).toBeLessThan(joined.indexOf(NEW[0]!));
    });
  }
});

/** Found on a real contract (held-out run 2): as above, with a smart tag in the middle of the last new paragraph. */
describe('the one line of a cell replaced by paragraphs, the last holding a smart tag', () => {
  it('the static redline passes', async () => {
    const OLD = 'Insert total price per line item';
    const run = (text: string): string => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
    const change = (tag: string, id: number): string => `<w:${tag} w:id="${9700 + id}" w:author="Redline" w:date="2026-11-26T00:00:00Z"`;
    const tag = (inner: string): string => `<w:smartTag w:uri="urn:schemas-microsoft-com:office:smarttags" w:element="place">${inner}</w:smartTag>`;
    const cellOf = (paragraphs: string): string => `<w:tbl><w:tblPr><w:tblW w:w="9000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="9000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="9000" w:type="dxa"/></w:tcPr>${paragraphs}</w:tc></w:tr></w:tbl>`;
    const base = (await buildPair(pairById('one-word-change'))).original;
    const withCell = (paragraphs: string): Promise<Uint8Array> => editPart(base, /^word\/document\.xml$/, (xml) => xml.replace(/<w:p>.*?<\/w:p>/s, (first) => first + cellOf(paragraphs)));
    const original = await withCell(`<w:p>${run(OLD)}</w:p>`);
    const modified = await withCell(`<w:p>${run('Total price per line item')}</w:p><w:p>${run('(')}${tag(run('Col.'))}${run(' 4 x 5)')}</w:p>`);
    const added = (id: number, text: string): string => `${change('ins', id)}>${run(text)}</w:ins>`;
    const tracked = await withCell(
      `<w:p><w:pPr><w:rPr>${change('del', 1)}/></w:rPr></w:pPr>${change('del', 2)}><w:r><w:delText xml:space="preserve">${OLD}</w:delText></w:r></w:del></w:p>` +
        `<w:p><w:pPr><w:rPr>${change('ins', 3)}/></w:rPr></w:pPr>${added(4, 'Total price per line item')}</w:p>` +
        `<w:p><w:pPr><w:rPr>${change('ins', 5)}/></w:rPr></w:pPr>${added(6, '(')}${tag(`<w:proofErr w:type="spellStart"/>${added(7, 'Col.')}<w:proofErr w:type="spellEnd"/>`)}${added(8, ' 4 x 5)')}</w:p>`,
    );
    expect(await failed(checkTrackedChanges({ original, modified, output: tracked }))).toEqual([]);
    const made = await makeStaticRedline(tracked, original, COVER);
    expect((await checkStaticRedline({ original, modified, output: made.docx, ...COVER })).checks.filter((c) => !c.passed)).toEqual([]);
  });
});

/** Review finding on the third batch: the replaced line and the new paragraphs numbered. The joined paragraph stays, so it shows no old number. */
describe('the one numbered line of a cell replaced by several numbered paragraphs', () => {
  it('the static redline passes, and the joined paragraph carries no old clause number', async () => {
    const base = (await buildPair(pairById('numbered-clause-deleted'))).original;
    const numPr = /<w:numPr>.*?<\/w:numPr>/s.exec(decoder.decode((await unzip(base)).get('word/document.xml')!))![0];
    const OLD = 'The scope of requirements is defined elsewhere';
    const NEW = ['Details of shipping documents follow', 'Three copies of the invoice', 'One certificate of origin'];
    const run = (text: string): string => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
    const change = (tag: string, id: number): string => `<w:${tag} w:id="${9400 + id}" w:author="Redline" w:date="2026-11-25T00:00:00Z"`;
    const P = (text: string): string => `<w:p><w:pPr>${numPr}</w:pPr>${run(text)}</w:p>`;
    const removed = `<w:p><w:pPr>${numPr}<w:rPr>${change('del', 1)}/></w:rPr></w:pPr>${change('del', 2)}><w:r><w:delText xml:space="preserve">${OLD}</w:delText></w:r></w:del></w:p>`;
    const added = (text: string, id: number): string => `<w:p><w:pPr>${numPr}<w:rPr>${change('ins', id)}/></w:rPr></w:pPr>${change('ins', id + 1)}>${run(text)}</w:ins></w:p>`;
    const withCell = (paragraphs: string): Promise<Uint8Array> =>
      editPart(base, /^word\/document\.xml$/, (xml) => xml.replace(/<w:p>.*?<\/w:p>/s, (first) => `${first}<w:tbl><w:tblPr><w:tblW w:w="9000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="9000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="9000" w:type="dxa"/></w:tcPr>${paragraphs}</w:tc></w:tr></w:tbl>`));
    const original = await withCell(P(OLD));
    const modified = await withCell(NEW.map(P).join(''));
    const tracked = await withCell(removed + added(NEW[0]!, 10) + added(NEW[1]!, 20) + added(NEW[2]!, 30));
    expect(await failed(checkTrackedChanges({ original, modified, output: tracked }))).toEqual([]);
    const made = await makeStaticRedline(tracked, original, COVER);
    expect((await checkStaticRedline({ original, modified, output: made.docx, ...COVER })).checks.filter((c) => !c.passed)).toEqual([]);
    const cell = /<w:tc>.*?<\/w:tc>/s.exec(decoder.decode((await unzip(made.docx)).get('word/document.xml')!))![0];
    expect(cell).not.toContain('RedlineOldNumber');
  });
});

/** Review findings on the second batch of held-out fixes: a table in a kept cell that comes or goes, with what follows it. */
describe('a table that comes or goes inside a kept cell, review findings', () => {
  const P = (text: string): string => (text === '' ? '<w:p/>' : `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`);
  const table = (width: number, rows: string[][]): string =>
    `<w:tbl><w:tblPr><w:tblW w:w="${width * rows[0]!.length}" w:type="dxa"/></w:tblPr><w:tblGrid>${rows[0]!.map(() => `<w:gridCol w:w="${width}"/>`).join('')}</w:tblGrid>${rows
      .map((row) => `<w:tr>${row.map((cell) => `<w:tc><w:tcPr><w:tcW w:w="${width}" w:type="dxa"/></w:tcPr>${cell}</w:tc>`).join('')}</w:tr>`)
      .join('')}</w:tbl>`;
  const inner = table(4000, ['Term of years', 'Initial rent payable', 'Rent free period', 'Rent review pattern'].map((label) => [P(label), P(`[${label} value]`)]));
  const outer = (cell: string): string => table(9000, [[P('Dated today here')], [cell]]);
  const withBlock = (docx: Uint8Array, block: string): Promise<Uint8Array> => editPart(docx, /^word\/document\.xml$/, (xml) => xml.replace(/<w:p>.*?<\/w:p>/s, (first) => first + block));
  const LEAD = P('Heads of terms follow');
  const NOTES = P('Notes follow the table');

  for (const [name, before, after] of [
    ['a table and the empty paragraph after it added after a line', LEAD, LEAD + inner + P('')],
    ['a table and the line after it replaced by another line', inner + NOTES, P('Completely different wording now')],
    ['a table and the line after it removed after a kept line', LEAD + inner + NOTES, LEAD],
    ['a table and a line after it added after a kept line', LEAD, LEAD + inner + NOTES],
  ] as const) {
    it(`both Word files pass: ${name}`, async () => {
      const pair = await buildPair(pairById('one-word-change'));
      await bothPass(await withBlock(pair.original, outer(before)), await withBlock(pair.original, outer(after)));
    });
  }
});

/**
 * Found on real contracts (held-out run 1, ADB procurement documents carrying tracked changes): at the end of a
 * cell an unnumbered line is replaced by a numbered clause, and the line after it removed. The engine writes the
 * new clause into the old line's paragraph (old properties, mark removed) and states the clause's numbering on
 * the cell's last paragraph, whose text is all removed. The clause must show its number where its text is.
 */
describe('a numbered clause that replaces the last two lines of a cell', () => {
  it('the static redline passes and the clause carries its number', async () => {
    const base = (await buildPair(pairById('numbered-clause-deleted'))).original;
    const numPr = /<w:numPr>.*?<\/w:numPr>/s.exec(decoder.decode((await unzip(base)).get('word/document.xml')!))![0];
    const OLD1 = 'Authorised representative of the partner (name and address, and';
    const OLD2 = 'e-mail address)';
    const NEW = 'Organisational chart and a list of the board of directors.';
    const run = (text: string): string => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
    const change = (tag: string, id: number): string => `<w:${tag} w:id="${9100 + id}" w:author="Redline" w:date="2026-11-25T00:00:00Z"`;
    const gone = (text: string, id: number): string => `${change('del', id)}><w:r><w:delText xml:space="preserve">${text}</w:delText></w:r></w:del>`;
    const inCell = (docx: Uint8Array, paragraphs: string): Promise<Uint8Array> =>
      editPart(docx, /^word\/document\.xml$/, (xml) => xml.replace(/<w:p>.*?<\/w:p>/s, (first) => `<w:tbl><w:tblPr><w:tblW w:w="9000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="9000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="9000" w:type="dxa"/></w:tcPr>${first}${paragraphs}</w:tc></w:tr></w:tbl>`));
    const original = await inCell(base, `<w:p>${run(OLD1)}</w:p><w:p>${run(OLD2)}</w:p>`);
    const modified = await inCell(base, `<w:p><w:pPr>${numPr}</w:pPr>${run(NEW)}</w:p>`);
    const tracked = await inCell(
      base,
      `<w:p><w:pPr><w:rPr>${change('del', 1)}/></w:rPr></w:pPr>${change('ins', 2)}>${run(NEW)}</w:ins>${gone(OLD1, 3)}</w:p>` +
        `<w:p><w:pPr>${numPr}${change('pPrChange', 4)}><w:pPr/></w:pPrChange></w:pPr>${gone(OLD2, 5)}</w:p>`,
    );
    expect(await failed(checkTrackedChanges({ original, modified, output: tracked }))).toEqual([]);
    const made = await makeStaticRedline(tracked, original, COVER);
    const report = await checkStaticRedline({ original, modified, output: made.docx, ...COVER });
    expect(report.checks.filter((c) => !c.passed)).toEqual([]);
    const xml = decoder.decode((await unzip(made.docx)).get('word/document.xml')!);
    const clause = xml.split('<w:p>').find((p) => p.includes(NEW))!;
    expect(clause).toContain(numPr);
  });
});

/**
 * Found on a real contract (held-out run 1, an NHS contract carrying tracked changes): every clause in a table
 * cell removed. The engine writes that as a removal on every paragraph mark but the cell's last, whose text is
 * removed. Read by the cell-end rule, the first paragraph of that run is the one that stays (emptied) and the
 * ones below it went; the redline typed an old clause number on the first as well, and was withheld.
 */
describe('a run of numbered clauses removed at the end of a cell', () => {
  const build = async (kept: string): Promise<{ original: Uint8Array; modified: Uint8Array; tracked: Uint8Array }> => {
    const base = (await buildPair(pairById('numbered-clause-deleted'))).original;
    const numPr = /<w:numPr>.*?<\/w:numPr>/s.exec(decoder.decode((await unzip(base)).get('word/document.xml')!))![0];
    const run = (text: string): string => (text === '' ? '' : `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`);
    const change = (tag: string, id: number): string => `<w:${tag} w:id="${9200 + id}" w:author="Redline" w:date="2026-11-25T00:00:00Z"`;
    const gone = (text: string, id: number): string => `${change('del', id)}><w:r><w:delText xml:space="preserve">${text}</w:delText></w:r></w:del>`;
    const markGone = (id: number): string => `<w:rPr>${change('del', id)}/></w:rPr>`;
    const withCell = (paragraphs: string): Promise<Uint8Array> =>
      editPart(base, /^word\/document\.xml$/, (xml) => xml.replace(/<w:p>.*?<\/w:p>/s, (first) => `${first}<w:tbl><w:tblPr><w:tblW w:w="9000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="9000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="9000" w:type="dxa"/></w:tcPr>${paragraphs}</w:tc></w:tr></w:tbl>`));
    const numbered = (text: string): string => `<w:p><w:pPr>${numPr}</w:pPr>${run(text)}</w:p>`;
    const A = 'give notice in writing of the change';
    const B = 'implement the change on the date mandated';
    const C = 'bear its own costs of the change';
    return {
      original: await withCell(numbered(kept === '' ? A : kept) + numbered(B) + numbered(C)),
      modified: await withCell(numbered(kept)),
      tracked: await withCell(
        `<w:p><w:pPr>${numPr}${markGone(1)}</w:pPr>${kept === '' ? gone(A, 2) : run(kept)}</w:p>` +
          `<w:p><w:pPr>${numPr}${markGone(3)}</w:pPr>${gone(B, 4)}</w:p>` +
          `<w:p><w:pPr>${numPr}</w:pPr>${gone(C, 5)}</w:p>`,
      ),
    };
  };
  const oldNumbersIn = (xml: string): number => xml.split('RedlineOldNumber"/>').length - 1;

  it('with the clause above them kept: the two that went show their old numbers, the kept one none', async () => {
    const { original, modified, tracked } = await build('The supplier must pay the fee');
    expect(await failed(checkTrackedChanges({ original, modified, output: tracked }))).toEqual([]);
    const made = await makeStaticRedline(tracked, original, COVER);
    expect((await checkStaticRedline({ original, modified, output: made.docx, ...COVER })).checks.filter((c) => !c.passed)).toEqual([]);
    expect(oldNumbersIn(decoder.decode((await unzip(made.docx)).get('word/document.xml')!))).toBe(2);
  });

  it('with every clause in the cell removed: the first is the one emptied, the two below it went', async () => {
    const { original, modified, tracked } = await build('');
    expect(await failed(checkTrackedChanges({ original, modified, output: tracked }))).toEqual([]);
    const made = await makeStaticRedline(tracked, original, COVER);
    expect((await checkStaticRedline({ original, modified, output: made.docx, ...COVER })).checks.filter((c) => !c.passed)).toEqual([]);
    expect(oldNumbersIn(decoder.decode((await unzip(made.docx)).get('word/document.xml')!))).toBe(2);
  });
});

/**
 * Found on real contracts (held-out run 1, ADB procurement documents): a text box holding a numbered note. The
 * engine replaces a changed text box whole, so the note is shown removed and put back. Old clause numbers belong
 * to the main text only; the removed copy inside the box was given one, and the redline was withheld.
 */
describe('a numbered paragraph inside a text box that is replaced', () => {
  it('both Word files pass, and no old number is typed inside the box', async () => {
    const base = (await buildPair(pairById('numbered-clause-deleted'))).original;
    const numPr = /<w:numPr>.*?<\/w:numPr>/s.exec(decoder.decode((await unzip(base)).get('word/document.xml')!))![0];
    const boxed = await buildPair(pairById('text-box-change'));
    const withBox = async (from: Uint8Array): Promise<Uint8Array> => {
      const xml = decoder.decode((await unzip(from)).get('word/document.xml')!);
      const holder = xml.match(/<w:p>(?:(?!<w:p>|<\/w:p>).)*<w:pict>.*?<\/w:pict>.*?<\/w:p>/s)![0];
      const inside = /(<w:txbxContent>)(.*?)(<\/w:txbxContent>)/gs;
      const numbered = holder.replace(inside, (_, open: string, lines: string, close: string) => open + lines.replace(/<w:p>/g, `<w:p><w:pPr>${numPr}</w:pPr>`) + close);
      expect(numbered).not.toBe(holder);
      return editPart(base, /^word\/document\.xml$/, (main) => main.replace(/<w:p>.*?<\/w:p>/s, (first) => first + numbered));
    };
    const outputs = await bothPass(await withBox(boxed.original), await withBox(boxed.modified));
    const xml = decoder.decode((await unzip(outputs.staticRedline)).get('word/document.xml')!);
    for (const box of xml.match(/<w:txbxContent>.*?<\/w:txbxContent>/gs)!) expect(box).not.toContain('RedlineOldNumber');
  });
});

/**
 * Found on real contracts (held-out run 1, ADB procurement documents, on drafts that were themselves made by the
 * compare engine): inside a Word file each link is tied to its address by a short internal name. The engine
 * names what it adds R1, R2, ...; a file it produced already holds such names, the new ones collide with them,
 * and a link came out tied to a picture, or to nothing. The names are replaced before the engine sees them.
 */
describe('a document whose internal reference names look like the engine\'s own', () => {
  const link = (n: number, extra = ''): unknown => ({ kind: 'linked', text: '', before: `Clause ${n}${extra} is published at `, linkText: `site ${n}`, url: `https://example.com/page-${n}`, after: '.' });
  const line = (text: string): unknown => ({ kind: 'paragraph', text });
  const build = (blocks: unknown[]): Promise<Uint8Array> => buildDocx({ blocks } as DocSpec);
  /** Rename every relationship of the main document R1, R2, ... as the engine would have. */
  const engineNamed = async (docx: Uint8Array): Promise<Uint8Array> => {
    const files = await unzip(docx);
    const rels = decoder.decode(files.get('word/_rels/document.xml.rels')!);
    const names = new Map([...rels.matchAll(/Id="([^"]+)"/g)].map((m, i) => [m[1]!, `R${i + 1}`]));
    files.set('word/_rels/document.xml.rels', encoder.encode(rels.replace(/Id="([^"]+)"/g, (_, id: string) => `Id="${names.get(id)}"`)));
    files.set('word/document.xml', encoder.encode(decoder.decode(files.get('word/document.xml')!).replace(/r:(id|embed)="([^"]+)"/g, (_, a: string, id: string) => `r:${a}="${names.get(id)}"`)));
    return zip(files);
  };

  it('both Word files pass, with every link pointing where it did', async () => {
    const original = await build([line('Heading'), link(1), link(2), link(3), link(4), link(5), line('End')]);
    const modified = await engineNamed(await build([line('Heading'), link(1, ' as amended'), link(2), link(6), link(3, ' new'), link(4), link(5), line('End')]));
    expect(decoder.decode((await unzip(modified)).get('word/document.xml')!)).toMatch(/<w:hyperlink [^>]*r:id="R\d+"/);
    await bothPass(original, modified);
    await bothPass(modified, original);
  });

  it('a document with no such name is passed on byte for byte', async () => {
    const docx = await build([line('Heading'), link(1)]);
    expect(await freshRelationshipIds(docx)).toBe(docx);
  });

  it('only the names are replaced, wherever a part uses them', () => {
    const names = new Map([['R1', 'rIdU1'], ['R2a', 'rIdU2']]);
    const part =
      `<w:document xmlns:w="w" xmlns:rel='http://schemas.openxmlformats.org/officeDocument/2006/relationships' xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:v="v">` +
      `<w:hyperlink rel:id="R1" w:anchor="R1"><w:t>R1 and "R2a" and the words rel:id="R1" typed as text</w:t></w:hyperlink><v:imagedata o:relid='R2a' o:title="R1"/><a:blip rel:embed="rId7" rel:link = "R2a"/></w:document>`;
    expect(renameInPart(part, names)).toBe(
      `<w:document xmlns:w="w" xmlns:rel='http://schemas.openxmlformats.org/officeDocument/2006/relationships' xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:v="v">` +
        `<w:hyperlink rel:id="rIdU1" w:anchor="R1"><w:t>R1 and "R2a" and the words rel:id="R1" typed as text</w:t></w:hyperlink><v:imagedata o:relid="rIdU2" o:title="R1"/><a:blip rel:embed="rId7" rel:link = "rIdU2"/></w:document>`,
    );
  });
});

/**
 * Found on a real contract (held-out run 1, an ADB document carrying tracked changes): the paragraph holding a
 * section break changed so much that the engine showed it removed and added again. Its output then held one
 * section break more than either document, and each later section named the header of the section after it.
 */
describe('a section break shown removed and added again', () => {
  const WNS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
  const ref = (type: string, id: string): string => `<w:headerReference w:type="${type}" r:id="${id}"/>`;
  const PAGE = '<w:pgSz w:w="12240" w:h="15840"/>';
  const breakWith = (refs: string, mark = '', margin = 1260): string => `<w:p><w:pPr>${mark ? `<w:rPr><w:${mark} w:id="1" w:author="x"/></w:rPr>` : ''}<w:sectPr>${refs}${PAGE}<w:pgMar w:top="${margin}"/></w:sectPr></w:pPr></w:p>`;
  const text = (t: string): string => `<w:p><w:r><w:t>${t}</w:t></w:r></w:p>`;
  const main = (blocks: string[], last: string): string => `<w:document ${WNS}><w:body>${blocks.join('')}<w:sectPr>${last}${PAGE}</w:sectPr></w:body></w:document>`;
  const RELS = `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${['A', 'B', 'C', 'D', 'E'].map((n) => `<Relationship Id="rId${n}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header${n}.xml"/>`).join('')}</Relationships>`;
  const AB = ref('even', 'rIdA') + ref('default', 'rIdB');
  const C = ref('default', 'rIdC');
  const DE = ref('even', 'rIdD') + ref('default', 'rIdE');
  // Both uploads: four sections. The second carries on with the first's headers; the third names one of its own.
  const upload = { main: main([text('one'), breakWith(AB), text('two'), breakWith(''), text('three'), breakWith(C, '', 990)], DE), rels: RELS };
  // The engine's output on the real contract: the first break twice, and the two after it naming the next section's headers.
  const engine = main([text('one'), breakWith(AB, 'del'), breakWith(AB, 'ins'), text('two'), breakWith(C), text('three'), breakWith(DE, '', 990)], DE);
  const namedAt = (xml: string): string[] => [...xml.matchAll(/<w:sectPr\b.*?<\/w:sectPr>/gs)].map((m) => [...m[0].matchAll(/w:type="(\w+)" r:id="rId(\w)"/g)].map((x) => `${x[1]}:${x[2]}`).join(','));

  it('each section break that stays names the headers the modified document shows in that section', () => {
    const repaired = repairSectionReferences(engine, RELS, upload, upload)!;
    // (A section that names none carries on with the previous section's: written out in full, as the engine does.)
    expect(namedAt(repaired)).toEqual(['even:A,default:B', 'even:A,default:B', 'even:A,default:B', 'even:A,default:C', 'even:D,default:E']);
    // Nothing else in a section break moves: the headers stay in front of the page settings, which are untouched.
    expect(repaired).toContain(`<w:sectPr>${ref('even', 'rIdA')}${C}${PAGE}<w:pgMar w:top="990"/></w:sectPr>`);
    expect(repaired.split(`<w:sectPr>${AB}${PAGE}<w:pgMar w:top="1260"/></w:sectPr>`)).toHaveLength(4);
  });

  it('nothing is touched where the output has as many section breaks as the uploads', () => {
    expect(repairSectionReferences(upload.main.replace(C, DE), RELS, upload, upload)).toBeUndefined();
  });

  it('nothing is touched where the uploads arrange their headers differently (review finding: the engine then points sections at parts of its own)', () => {
    const rearranged = { main: upload.main.replace(C, ref('default', 'rIdD')), rels: RELS };
    expect(repairSectionReferences(engine, RELS, upload, rearranged)).toBeUndefined();
    expect(repairSectionReferences(engine, RELS, rearranged, upload)).toBeUndefined();
    const extra = RELS.replace('</Relationships>', '<Relationship Id="rIdHeader" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header99.xml"/></Relationships>');
    expect(repairSectionReferences(engine, extra, upload, upload)).toBeUndefined();
  });

  it('nothing is touched where the breaks cannot be paired off, or a header is not in the output under the same name', () => {
    const unpaired = main([text('one'), breakWith(AB, 'del'), breakWith(AB, 'ins'), text('two'), breakWith(C), breakWith(C), text('three'), breakWith(DE)], DE);
    expect(repairSectionReferences(unpaired, RELS, upload, upload)).toBeUndefined();
    expect(repairSectionReferences(engine, RELS.replace('Target="headerC.xml"', 'Target="header9.xml"'), upload, upload)).toBeUndefined();
    expect(repairSectionReferences(engine, RELS.replace(/<Relationship Id="rIdC"[^>]*>/, ''), upload, upload)).toBeUndefined();
  });

  it('a document in which the engine shows a section break removed and added: both Word files pass', async () => {
    const base = (await buildPair(pairById('footer-change'))).original;
    const make = async (breakText: string): Promise<Uint8Array> => {
      const files = await unzip(base);
      const xml = decoder.decode(files.get('word/document.xml')!);
      const named = /<w:sectPr\b[^>]*>((?:(?!<\/w:sectPr>).)*)<\/w:sectPr>(?=<\/w:body>)/.exec(xml)![1]!;
      const bare = named.replace(/<w:(header|footer)Reference\b[^>]*\/>/g, '');
      const headerId = /<w:headerReference\b[^>]*r:id="([^"]+)"/.exec(named)![1]!;
      let rels = decoder.decode(files.get('word/_rels/document.xml.rels')!);
      const target = (new RegExp(`Id="${headerId}"[^>]*Target="([^"]+)"`).exec(rels) ?? new RegExp(`Target="([^"]+)"[^>]*Id="${headerId}"`).exec(rels))![1]!;
      let types = decoder.decode(files.get('[Content_Types].xml')!);
      for (let n = 1; n <= 5; n++) {
        files.set(`word/header9${n}.xml`, encoder.encode(decoder.decode(files.get(`word/${target}`)!).replace(/(<w:t[^>]*>)[^<]*/, `$1Header number ${n} of the set`)));
        rels = rels.replace('</Relationships>', `<Relationship Id="rIdH${n}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header9${n}.xml"/></Relationships>`);
        types = types.replace('</Types>', `<Override PartName="/word/header9${n}.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/></Types>`);
      }
      files.set('word/_rels/document.xml.rels', encoder.encode(rels));
      files.set('[Content_Types].xml', encoder.encode(types));
      const h = (type: string, n: number): string => `<w:headerReference w:type="${type}" r:id="rIdH${n}"/>`;
      const P = (t: string, sect?: string): string => `<w:p>${sect === undefined ? '' : `<w:pPr><w:sectPr>${sect}</w:sectPr></w:pPr>`}${t ? `<w:r><w:t xml:space="preserve">${t}</w:t></w:r>` : ''}</w:p>`;
      const body = [P('Part one of the agreement'), P(breakText, h('even', 1) + h('default', 2) + bare), P('Part two of the agreement'), P('', bare), P('Part three of the agreement'), P('', h('default', 3) + bare), P('Part four of the agreement'), `<w:sectPr>${h('even', 4) + h('default', 5) + bare}</w:sectPr>`];
      files.set('word/document.xml', encoder.encode(xml.replace(/<w:body>.*<\/w:body>/s, `<w:body>${body.join('')}</w:body>`)));
      return zip(files);
    };
    const original = await make('The first schedule ends on this line');
    const modified = await make('Completely different wording closes that chapter now');
    const { tracked } = await bothPass(original, modified);
    // The situation the repair is for: one section break more than either document.
    expect(decoder.decode((await unzip(tracked)).get('word/document.xml')!).split('<w:sectPr').length - 1).toBe(5);
  });
});

/**
 * Found on a real contract (held-out run 1, a tender schedule): words typed into the footnote separator (the
 * rule between the text and the footnotes), used as a second footer. The engine does not compare separators, so
 * a change of wording there was missing from the redline, which was withheld.
 */
describe('wording typed into the footnote separator', () => {
  const withSeparatorText = (docx: Uint8Array, wording: string): Promise<Uint8Array> =>
    editPart(docx, /^word\/footnotes\.xml$/, (xml) => {
      const typed = `<w:p><w:r><w:t xml:space="preserve">${wording}</w:t></w:r><w:r><w:tab/></w:r></w:p>`;
      const out = xml.replace(/(<w:footnote\b[^>]*w:type="separator"[^>]*>)/, `$1${typed}`);
      if (out === xml) throw new Error('the fixture has no footnote separator');
      return out;
    });

  it('a change of wording there is shown, and both Word files pass', async () => {
    const pair = await buildPair(pairById('footnote-change'));
    const original = await withSeparatorText(pair.original, 'Contract No: 1 Revision date: 1 May');
    const modified = await withSeparatorText(pair.modified, 'Contract No: 1 Revision date: 9 June');
    const { tracked, staticRedline } = await bothPass(original, modified);
    const notes = decoder.decode((await unzip(tracked)).get('word/footnotes.xml')!);
    expect(notes).toMatch(/<w:del\b[^>]*>(?:(?!<\/w:del>).)*Revision date: 1 May/s);
    expect(notes).toMatch(/<w:ins\b[^>]*>(?:(?!<\/w:ins>).)*Revision date: 9 June/s);
    expect(decoder.decode((await unzip(staticRedline)).get('word/footnotes.xml')!)).toContain('Revision date: 1 May');
  });

  it('the same wording in both is left exactly as the engine wrote it', async () => {
    const pair = await buildPair(pairById('footnote-change'));
    const original = await withSeparatorText(pair.original, 'Contract No: 1');
    const modified = await withSeparatorText(pair.modified, 'Contract No: 1');
    const { tracked } = await bothPass(original, modified);
    const separator = /<w:footnote\b[^>]*w:type="separator"[^>]*>.*?<\/w:footnote>/s.exec(decoder.decode((await unzip(tracked)).get('word/footnotes.xml')!))![0];
    expect(separator).not.toMatch(/<w:(ins|del)\b/);
  });
});

/**
 * Found on a real contract (held-out run 1, an ADB document): one numbered list running through a text box and
 * on into the main text. Word stores a text box twice (a modern form and a legacy one) and shows it once. The
 * redline counted both stored forms when working out a removed clause's old number, and the Checker, comparing
 * the two forms, counted them three times; neither agreed with Word, or with each other.
 */
describe('a numbered list that runs through a text box and on into the main text', () => {
  const NS_BOX = 'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:v="urn:schemas-microsoft-com:vml"';
  const withBoxedItem = async (docx: Uint8Array): Promise<Uint8Array> => {
    const numPr = /<w:numPr>.*?<\/w:numPr>/s.exec(decoder.decode((await unzip(docx)).get('word/document.xml')!))![0];
    const box = `<w:txbxContent><w:p><w:pPr>${numPr}</w:pPr><w:r><w:t xml:space="preserve">A note that opens the list</w:t></w:r></w:p></w:txbxContent>`;
    const holder = `<w:p><w:r><mc:AlternateContent ${NS_BOX}><mc:Choice Requires="wps"><w:drawing><wps:txbx>${box}</wps:txbx></w:drawing></mc:Choice><mc:Fallback><w:pict><v:textbox>${box}</v:textbox></w:pict></mc:Fallback></mc:AlternateContent></w:r></w:p>`;
    return editPart(docx, /^word\/document\.xml$/, (xml) => xml.replace(/<w:p>.*?<\/w:p>/s, (first) => first + holder));
  };
  const oldNumberIn = async (staticRedline: Uint8Array): Promise<string> => {
    const xml = decoder.decode((await unzip(staticRedline)).get('word/document.xml')!);
    return /RedlineOldNumber"\/>.*?<w:t[^>]*>([^<]*)</s.exec(xml)![1]!;
  };

  it('a removed clause shows the number Word gave it: the text box counted once', async () => {
    const pair = await buildPair(pairById('numbered-clause-deleted'));
    const plain = await oldNumberIn((await bothPass(pair.original, pair.modified)).staticRedline);
    const boxed = await oldNumberIn((await bothPass(await withBoxedItem(pair.original), await withBoxedItem(pair.modified))).staticRedline);
    expect(plain).toBe('(b)');
    expect(boxed).toBe('(c)');
  });

  it('an output whose legacy form of the box has lost its numbering is caught (review finding)', async () => {
    const pair = await buildPair(pairById('numbered-clause-deleted'));
    const original = await withBoxedItem(pair.original);
    const modified = await withBoxedItem(pair.modified);
    const { tracked, staticRedline } = await bothPass(original, modified);
    const damage = (docx: Uint8Array): Promise<Uint8Array> =>
      editPart(docx, /^word\/document\.xml$/, (xml) => {
        const out = xml.replace(/(<mc:Fallback\b[^>]*>(?:(?!<\/mc:Fallback>).)*?)<w:numPr>.*?<\/w:numPr>/s, '$1');
        if (out === xml) throw new Error('nothing was damaged');
        return out;
      });
    expect(await failed(checkTrackedChanges({ original, modified, output: await damage(tracked) }))).not.toEqual([]);
    const report = await checkStaticRedline({ original, modified, output: await damage(staticRedline), ...COVER });
    expect(report.checks.filter((c) => !c.passed)).not.toEqual([]);
  });
});

/**
 * Found on a real contract (held-out run 1, an NHS agreement published with tracked changes in it): a footnote
 * emptied in the draft, its one paragraph mark flagged as removed. Word keeps that paragraph (a note cannot lose
 * its last one). The Checker read the note as holding no paragraph at all, so no output could ever match it.
 */
describe('an upload whose emptied footnote has its only paragraph mark flagged as removed', () => {
  it('both Word files pass', async () => {
    const pair = await buildPair(pairById('footnote-change'));
    const emptied = '<w:footnote w:id="1"><w:p><w:pPr><w:pStyle w:val="FootnoteText"/><w:rPr><w:del w:id="900" w:author="A" w:date="2026-01-01T00:00:00Z"/></w:rPr></w:pPr></w:p></w:footnote>';
    const original = await editPart(pair.original, /^word\/footnotes\.xml$/, (xml) => {
      const out = xml.replace(/<w:footnote w:id="1">.*?<\/w:footnote>/s, emptied);
      if (out === xml) throw new Error('the fixture has no footnote 1');
      return out;
    });
    await bothPass(original, pair.modified, 'original');
    await bothPass(original, original, 'both');
  });

  it('an output whose footnote has lost its paragraph altogether is still caught (review finding)', async () => {
    const pair = await buildPair(pairById('footnote-change'));
    const blank = (docx: Uint8Array): Promise<Uint8Array> => editPart(docx, /^word\/footnotes\.xml$/, (xml) => xml.replace(/<w:footnote w:id="1">.*?<\/w:footnote>/s, '<w:footnote w:id="1"><w:p><w:pPr><w:pStyle w:val="FootnoteText"/></w:pPr></w:p></w:footnote>'));
    const original = await blank(pair.original);
    const modified = await blank(pair.modified);
    const { tracked } = await bothPass(original, modified);
    const damaged = await editPart(tracked, /^word\/footnotes\.xml$/, (xml) => {
      const out = xml.replace(/(<w:footnote\b(?![^>]*w:type=)[^>]*>).*?(<\/w:footnote>)/s, '$1$2');
      if (out === xml) throw new Error('nothing was damaged');
      return out;
    });
    expect(await failed(checkTrackedChanges({ original, modified, output: damaged }))).not.toEqual([]);
  });
});

/**
 * Found on real contracts (held-out run 2, Islamic Development Bank documents): a text box whose modern form
 * holds a web address as plain text and whose legacy form holds the same words as a link. A reader sees the same
 * words either way; the Checker refused the upload itself because the two forms did not read identically.
 */
describe('a text box whose two stored forms show the same words, one of them through a link field', () => {
  const NS_BOX = 'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:v="urn:schemas-microsoft-com:vml"';
  const plain = '<w:r><w:t>www.example.org</w:t></w:r>';
  const linked = (address: string): string =>
    `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> HYPERLINK "${address}" </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>${plain}<w:r><w:fldChar w:fldCharType="end"/></w:r>`;
  const box = (inner: string): string => `<w:txbxContent><w:p>${inner}</w:p></w:txbxContent>`;
  const withBox = (docx: Uint8Array): Promise<Uint8Array> =>
    editPart(docx, /^word\/document\.xml$/, (xml) =>
      xml.replace(/<w:p>.*?<\/w:p>/s, (first) => `${first}<w:p><w:r><mc:AlternateContent ${NS_BOX}><mc:Choice Requires="wps"><w:drawing><wps:txbx>${box(plain)}</wps:txbx></w:drawing></mc:Choice><mc:Fallback><w:pict><v:textbox>${box(linked('http://www.example.org'))}</v:textbox></w:pict></mc:Fallback></mc:AlternateContent></w:r></w:p>`),
    );

  it('both Word files pass; an output whose hidden form points the link elsewhere is caught', async () => {
    const pair = await buildPair(pairById('one-word-change'));
    const original = await withBox(pair.original);
    const modified = await withBox(pair.modified);
    const { tracked, staticRedline } = await bothPass(original, modified);
    const damage = (docx: Uint8Array): Promise<Uint8Array> =>
      editPart(docx, /^word\/document\.xml$/, (xml) => {
        const out = xml.replace('http://www.example.org', 'http://www.example.net');
        if (out === xml) throw new Error('nothing was damaged');
        return out;
      });
    expect(await failed(checkTrackedChanges({ original, modified, output: await damage(tracked) }))).not.toEqual([]);
    const report = await checkStaticRedline({ original, modified, output: await damage(staticRedline), ...COVER });
    expect(report.checks.filter((c) => !c.passed)).not.toEqual([]);
  });

  it('where the hidden link changes between the versions, the removed copy is held to the original (review finding)', async () => {
    const pair = await buildPair(pairById('one-word-change'));
    const original = await withBox(pair.original);
    const modified = await editPart(await withBox(pair.modified), /^word\/document\.xml$/, (xml) => xml.replace('http://www.example.org', 'http://www.example.net'));
    const { tracked, staticRedline } = await bothPass(original, modified);
    const damage = (docx: Uint8Array): Promise<Uint8Array> =>
      editPart(docx, /^word\/document\.xml$/, (xml) => {
        const out = xml.replace('http://www.example.org', 'http://elsewhere.example');
        if (out === xml) throw new Error('nothing was damaged');
        return out;
      });
    expect(await failed(checkTrackedChanges({ original, modified, output: await damage(tracked) }))).not.toEqual([]);
    const report = await checkStaticRedline({ original, modified, output: await damage(staticRedline), ...COVER });
    expect(report.checks.filter((c) => !c.passed)).not.toEqual([]);
  });

  it('a field that is not a whole, simple one is compared in full (review finding)', async () => {
    const pair = await buildPair(pairById('one-word-change'));
    const odd = `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> XE "x" </w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r><w:r><w:t>EXTRA VISIBLE WORDS</w:t></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>${plain}`;
    const withOddBox = (docx: Uint8Array): Promise<Uint8Array> =>
      editPart(docx, /^word\/document\.xml$/, (xml) =>
        xml.replace(/<w:p>.*?<\/w:p>/s, (first) => `${first}<w:p><w:r><mc:AlternateContent ${NS_BOX}><mc:Choice Requires="wps"><w:drawing><wps:txbx>${box(plain)}</wps:txbx></w:drawing></mc:Choice><mc:Fallback><w:pict><v:textbox>${box(odd)}</v:textbox></w:pict></mc:Fallback></mc:AlternateContent></w:r></w:p>`),
      );
    const original = await withOddBox(pair.original);
    const modified = await withOddBox(pair.modified);
    const tracked = await produceTrackedChanges(original, modified);
    expect(await failed(checkTrackedChanges({ original, modified, output: tracked }))).toContain('completeness');
  });

  it('forms that show different words are still refused', async () => {
    const pair = await buildPair(pairById('one-word-change'));
    const original = await withBox(pair.original);
    const differing = await editPart(await withBox(pair.modified), /^word\/document\.xml$/, (xml) => xml.replace(/(<mc:Fallback>.*?)www\.example\.org<\/w:t>/s, '$1www.example.net</w:t>'));
    const tracked = await produceTrackedChanges(original, differing);
    expect(await failed(checkTrackedChanges({ original, modified: differing, output: tracked }))).toContain('completeness');
  });
});
