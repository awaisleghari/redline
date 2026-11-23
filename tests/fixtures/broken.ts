/**
 * Deliberately broken outputs ("checking the Checker").
 *
 * Each case takes the genuine tracked-changes file the engine produced for one
 * of the generated pairs and damages it in one specific way. The Checker must
 * refuse every one of them. `caughtBy` lists exactly which checks must fail,
 * so the test also proves each check catches what it is meant to catch and
 * does not fire by accident.
 *
 * Add a new case here whenever a real bug is found.
 */
import type { CheckId } from '../../src/checker';
import { unzip, zip, type ZipEntries } from '../../src/common/zip';

export interface BrokenCase {
  id: string;
  /** Plain-English description of the damage. */
  description: string;
  /** Which generated pair's genuine output is damaged. */
  pairId: string;
  /** Damage the unzipped output in place, or return entirely new bytes. */
  damage: (files: ZipEntries) => void | Uint8Array;
  /** Exactly the checks that must fail. */
  caughtBy: CheckId[];
}

const decoder = new TextDecoder();
const encoder = new TextEncoder();

/** Rewrite one XML part. Throws if the part is missing or the edit changed nothing, so a case can never silently do nothing. */
function edit(files: ZipEntries, part: string, change: (xml: string) => string): void {
  const bytes = files.get(part);
  if (!bytes) throw new Error(`broken case: part ${part} not found`);
  const before = decoder.decode(bytes);
  const after = change(before);
  if (after === before) throw new Error(`broken case: the edit to ${part} changed nothing`);
  files.set(part, encoder.encode(after));
}

/** Replace exactly one occurrence of `from`. */
function swap(xml: string, from: string | RegExp, to: string): string {
  const matches = typeof from === 'string' ? xml.split(from).length - 1 : (xml.match(new RegExp(from.source, from.flags.includes('g') ? from.flags : from.flags + 'g')) ?? []).length;
  if (matches !== 1) throw new Error(`broken case: expected exactly one match for ${String(from)}, found ${matches}`);
  return xml.replace(from, to);
}

const DOC = 'word/document.xml';
const PARA = (text: string): string => `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
const SERVICES = 'The Supplier shall provide the Services with reasonable skill and care.';
const PAYMENT = 'The Customer shall pay the Fees within thirty days of the date of each invoice.';
const PARTIES = 'This Services Agreement is made between Alpha Holdings Limited (the Supplier) and Beta Trading LLC (the Customer).';

const TEXT: CheckId[] = ['accept-all', 'reject-all'];

export const BROKEN_CASES: BrokenCase[] = [
  // ----- text that did not change must survive untouched -----
  {
    id: 'dropped-word',
    description: 'A word of unchanged text is missing.',
    pairId: 'one-word-change',
    damage: (f) => edit(f, DOC, (x) => swap(x, ' skill and care.', ' and care.')),
    caughtBy: TEXT,
  },
  {
    id: 'added-word',
    description: 'A word was slipped into unchanged text without being tracked.',
    pairId: 'one-word-change',
    damage: (f) => edit(f, DOC, (x) => swap(x, 'The Customer shall pay', 'The Customer shall not pay')),
    caughtBy: TEXT,
  },
  {
    id: 'dropped-space',
    description: 'A single space is missing from unchanged text.',
    pairId: 'one-word-change',
    damage: (f) => edit(f, DOC, (x) => swap(x, 'thirty days of', 'thirty daysof')),
    caughtBy: TEXT,
  },
  {
    id: 'dropped-full-stop',
    description: 'A full stop is missing from unchanged text.',
    pairId: 'one-word-change',
    damage: (f) => edit(f, DOC, (x) => swap(x, 'England and Wales.', 'England and Wales')),
    caughtBy: TEXT,
  },
  {
    id: 'changed-digit-unchanged-text',
    description: 'A digit in unchanged text is different ($300,000 became $800,000).',
    pairId: 'one-word-change',
    damage: (f) => edit(f, DOC, (x) => swap(x, '$300,000', '$800,000')),
    caughtBy: TEXT,
  },
  {
    id: 'changed-case',
    description: 'A capital letter in unchanged text became lower case.',
    pairId: 'one-word-change',
    damage: (f) => edit(f, DOC, (x) => swap(x, 'pay the Fees', 'pay the fees')),
    caughtBy: TEXT,
  },
  // ----- the changes themselves must be exactly right -----
  {
    id: 'changed-digit-in-insertion',
    description: 'The inserted amount is wrong ($30,000 shown as $80,000).',
    pairId: 'number-change',
    damage: (f) => edit(f, DOC, (x) => swap(x, '<w:t>$30,000</w:t>', '<w:t>$80,000</w:t>')),
    caughtBy: ['accept-all'],
  },
  {
    id: 'changed-digit-in-deletion',
    description: 'The deleted amount is wrong ($300,000 shown as $900,000).',
    pairId: 'number-change',
    damage: (f) => edit(f, DOC, (x) => swap(x, '<w:delText>$300,000</w:delText>', '<w:delText>$900,000</w:delText>')),
    caughtBy: ['reject-all'],
  },
  {
    id: 'deletion-styled-as-insertion',
    description: 'A deletion is marked as an insertion.',
    pairId: 'one-word-change',
    damage: (f) =>
      edit(f, DOC, (x) =>
        swap(x, /<w:del ([^>]*)><w:r><w:delText>reasonable<\/w:delText><\/w:r><\/w:del>/, '<w:ins $1><w:r><w:t>reasonable</w:t></w:r></w:ins>'),
      ),
    caughtBy: TEXT,
  },
  {
    id: 'insertion-styled-as-deletion',
    description: 'An insertion is marked as a deletion.',
    pairId: 'one-word-change',
    damage: (f) =>
      edit(f, DOC, (x) =>
        swap(x, /<w:ins ([^>]*)><w:r><w:t>exceptional<\/w:t><\/w:r><\/w:ins>/, '<w:del $1><w:r><w:delText>exceptional</w:delText></w:r></w:del>'),
      ),
    caughtBy: TEXT,
  },
  {
    id: 'insertion-not-tracked',
    description: 'Inserted text appears as ordinary text, with no tracked change.',
    pairId: 'one-word-change',
    damage: (f) => edit(f, DOC, (x) => swap(x, /<w:ins [^>]*><w:r><w:t>exceptional<\/w:t><\/w:r><\/w:ins>/, '<w:r><w:t>exceptional</w:t></w:r>')),
    caughtBy: ['reject-all'],
  },
  {
    id: 'deletion-not-tracked',
    description: 'Deleted text appears as ordinary text, with no tracked change.',
    pairId: 'one-word-change',
    damage: (f) => edit(f, DOC, (x) => swap(x, /<w:del [^>]*><w:r><w:delText>reasonable<\/w:delText><\/w:r><\/w:del>/, '<w:r><w:t>reasonable</w:t></w:r>')),
    caughtBy: ['accept-all'],
  },
  {
    id: 'deletion-dropped',
    description: 'A deletion is missing altogether: the old word is simply gone.',
    pairId: 'one-word-change',
    damage: (f) => edit(f, DOC, (x) => swap(x, /<w:del [^>]*><w:r><w:delText>reasonable<\/w:delText><\/w:r><\/w:del>/, '')),
    caughtBy: ['reject-all'],
  },
  {
    id: 'inserted-paragraph-mark-not-tracked',
    description: 'A new paragraph has its text tracked but not the paragraph itself, so rejecting leaves an empty paragraph behind.',
    pairId: 'inserted-paragraph',
    damage: (f) => edit(f, DOC, (x) => swap(x, /<w:pPr><w:rPr><w:ins [^>]*\/><\/w:rPr><\/w:pPr>/, '')),
    caughtBy: ['reject-all'],
  },
  {
    id: 'deleted-paragraph-mark-not-tracked',
    description: 'A removed paragraph has its text tracked but not the paragraph itself, so accepting leaves an empty paragraph behind.',
    pairId: 'deleted-paragraph',
    damage: (f) => edit(f, DOC, (x) => swap(x, /<w:pPr><w:rPr><w:del [^>]*\/><\/w:rPr><\/w:pPr>/, '')),
    caughtBy: ['accept-all'],
  },
  // ----- paragraphs -----
  {
    id: 'dropped-paragraph',
    description: 'A whole unchanged paragraph is missing.',
    pairId: 'inserted-paragraph',
    damage: (f) => edit(f, DOC, (x) => swap(x, PARA(PAYMENT), '')),
    caughtBy: TEXT,
  },
  {
    id: 'swapped-paragraphs',
    description: 'Two unchanged paragraphs are in the wrong order.',
    pairId: 'inserted-paragraph',
    damage: (f) => edit(f, DOC, (x) => swap(x, PARA(SERVICES) + PARA(PAYMENT), PARA(PAYMENT) + PARA(SERVICES))),
    caughtBy: TEXT,
  },
  {
    id: 'merged-paragraphs',
    description: 'Two paragraphs were run together into one (every character is still there).',
    pairId: 'inserted-paragraph',
    damage: (f) => edit(f, DOC, (x) => swap(x, PARA(SERVICES) + PARA(PAYMENT), PARA(SERVICES + PAYMENT))),
    caughtBy: TEXT,
  },
  {
    id: 'duplicated-paragraph',
    description: 'An unchanged paragraph appears twice.',
    pairId: 'inserted-paragraph',
    damage: (f) => edit(f, DOC, (x) => swap(x, PARA(PARTIES), PARA(PARTIES) + PARA(PARTIES))),
    caughtBy: TEXT,
  },
  // ----- moves -----
  {
    id: 'move-destination-dropped',
    description: 'A moved clause was removed from its old place but never arrives at the new one.',
    pairId: 'moved-paragraph',
    damage: (f) => edit(f, DOC, (x) => swap(x, /<w:p><w:pPr><w:rPr><w:ins [^>]*\/><\/w:rPr><\/w:pPr><w:moveToRangeStart.*?<w:moveToRangeEnd [^>]*\/><\/w:p>/, '')),
    caughtBy: ['accept-all', 'structure'],
  },
  {
    id: 'move-text-altered',
    description: 'The moved clause reads differently at its destination ("secret" became "safe").',
    pairId: 'moved-paragraph',
    damage: (f) => edit(f, DOC, (x) => swap(x, /(<w:moveTo [^>]*><w:r><w:t xml:space="preserve">[^<]*?) secret /, '$1 safe ')),
    caughtBy: ['accept-all'],
  },
  {
    id: 'move-unpaired',
    description: "A move's origin and destination markers no longer match up.",
    pairId: 'moved-paragraph',
    damage: (f) => edit(f, DOC, (x) => swap(x, /(<w:moveToRangeStart [^>]*w:name=")move1"/, '$1move9"')),
    caughtBy: ['structure'],
  },
  {
    id: 'block-move-one-destination-unpaired',
    description: 'In a three-paragraph move, one destination no longer matches its origin.',
    pairId: 'moved-block',
    damage: (f) =>
      edit(f, DOC, (x) => {
        const names = [...x.matchAll(/<w:moveToRangeStart [^>]*w:name="([^"]+)"/g)];
        if (names.length !== 3) throw new Error(`broken case: expected 3 move destinations, found ${names.length}`);
        const last = names[2]!;
        return x.slice(0, last.index) + last[0].replace(`w:name="${last[1]}"`, 'w:name="elsewhere"') + x.slice(last.index + last[0].length);
      }),
    caughtBy: ['structure'],
  },
  {
    id: 'block-move-one-destination-dropped',
    description: 'In a three-paragraph move, one paragraph never arrives at the destination.',
    pairId: 'moved-block',
    damage: (f) =>
      edit(f, DOC, (x) => {
        const paragraphs = [...x.matchAll(/<w:p><w:pPr><w:rPr><w:ins [^>]*\/><\/w:rPr><\/w:pPr><w:moveToRangeStart.*?<w:moveToRangeEnd [^>]*\/><\/w:p>/g)];
        if (paragraphs.length !== 3) throw new Error(`broken case: expected 3 destination paragraphs, found ${paragraphs.length}`);
        const middle = paragraphs[1]!;
        return x.slice(0, middle.index) + x.slice(middle.index + middle[0].length);
      }),
    caughtBy: ['accept-all', 'structure'],
  },
  // ----- tables -----
  {
    id: 'table-cell-altered',
    description: 'Unchanged text in a table cell is different.',
    pairId: 'table-cell-change',
    damage: (f) => edit(f, DOC, (x) => swap(x, '>Hosting<', '>Housing<')),
    caughtBy: TEXT,
  },
  {
    id: 'table-cell-dropped',
    description: 'A table cell is missing from an unchanged row.',
    pairId: 'table-cell-change',
    damage: (f) => edit(f, DOC, (x) => swap(x, '<w:tc><w:p><w:r><w:t xml:space="preserve">USD 4,000</w:t></w:r></w:p></w:tc>', '')),
    caughtBy: TEXT,
  },
  {
    id: 'inserted-row-not-tracked',
    description: 'A new table row has its text tracked but not the row itself, so rejecting leaves an empty row behind.',
    pairId: 'table-row-added',
    damage: (f) => edit(f, DOC, (x) => swap(x, /<w:trPr><w:ins [^>]*\/><\/w:trPr>/, '')),
    caughtBy: ['reject-all'],
  },
  {
    id: 'deleted-row-not-tracked',
    description: 'A removed table row has its text tracked but not the row itself, so accepting leaves an empty row behind.',
    pairId: 'table-row-deleted',
    damage: (f) => edit(f, DOC, (x) => swap(x, /<w:trPr><w:del [^>]*\/><\/w:trPr>/, '')),
    caughtBy: ['accept-all'],
  },
  // ----- headers, footers, notes, text boxes -----
  {
    id: 'header-text-altered',
    description: 'Unchanged header text is different.',
    pairId: 'header-change',
    damage: (f) => edit(f, 'word/header1.xml', (x) => swap(x, ' - Confidential', ' - Public')),
    caughtBy: TEXT,
  },
  {
    id: 'footer-text-altered',
    description: 'The footer, which did not change, is different.',
    pairId: 'header-change',
    damage: (f) => edit(f, 'word/footer1.xml', (x) => swap(x, 'Alpha Holdings Limited', 'Alpha Holdings')),
    caughtBy: TEXT,
  },
  {
    id: 'header-part-missing',
    description: 'The header file inside the Word package is gone.',
    pairId: 'header-change',
    damage: (f) => {
      if (!f.delete('word/header1.xml')) throw new Error('broken case: header1.xml not found');
    },
    caughtBy: ['accept-all', 'reject-all', 'completeness', 'structure'],
  },
  {
    id: 'missing-footnote',
    description: 'A footnote is missing although the text still refers to it.',
    pairId: 'footnote-change',
    damage: (f) => edit(f, 'word/footnotes.xml', (x) => swap(x, /<w:footnote w:id="1">.*?<\/w:footnote>/, '')),
    caughtBy: ['accept-all', 'reject-all', 'completeness', 'structure'],
  },
  {
    id: 'footnote-reference-removed',
    description: 'The footnote mark is gone from the text, leaving the footnote orphaned.',
    pairId: 'footnote-change',
    damage: (f) => edit(f, DOC, (x) => swap(x, /<w:r>(<w:rPr>.*?<\/w:rPr>)?<w:footnoteReference w:id="1" ?\/><\/w:r>/, '')),
    caughtBy: ['accept-all', 'reject-all', 'completeness'],
  },
  {
    id: 'footnote-text-altered',
    description: 'Unchanged footnote text is different.',
    pairId: 'footnote-change',
    damage: (f) => edit(f, 'word/footnotes.xml', (x) => swap(x, ' of value added tax.', ' of sales tax.')),
    caughtBy: TEXT,
  },
  {
    id: 'endnote-text-altered',
    description: 'Unchanged endnote text is different.',
    pairId: 'endnote-change',
    damage: (f) => edit(f, 'word/endnotes.xml', (x) => swap(x, 'The courts of England', 'The courts of Scotland')),
    caughtBy: TEXT,
  },
  {
    id: 'text-box-altered',
    description: 'Text inside the new text box is different.',
    pairId: 'text-box-change',
    damage: (f) => edit(f, DOC, (x) => swap(x, '>Secretary<', '>Treasurer<')),
    caughtBy: ['accept-all'],
  },
  {
    id: 'unreferenced-part-with-text',
    description: 'The package holds an extra part with text that nothing refers to (hidden content).',
    pairId: 'one-word-change',
    damage: (f) => {
      f.set(
        'word/header9.xml',
        encoder.encode(
          '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>Hidden text</w:t></w:r></w:p></w:hdr>',
        ),
      );
      edit(f, '[Content_Types].xml', (x) =>
        swap(x, '</Types>', '<Override PartName="/word/header9.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/></Types>'),
      );
    },
    caughtBy: ['completeness'],
  },
  // ----- found by the Stage 2 independent review: each of these once passed -----
  {
    id: 'spaces-lost-preserve-flag-removed',
    description: 'A run lost its "preserve spaces" flag, so Word drops the space at its end and two words run together.',
    pairId: 'one-word-change',
    damage: (f) => edit(f, DOC, (x) => swap(x, '<w:t xml:space="preserve">The Supplier shall provide the Services with </w:t>', '<w:t>The Supplier shall provide the Services with </w:t>')),
    caughtBy: TEXT,
  },
  {
    id: 'table-dissolved-into-paragraphs',
    description: 'A table was replaced by its text as plain paragraphs.',
    pairId: 'table-cell-change',
    damage: (f) =>
      edit(f, DOC, (x) => {
        const table = x.match(/<w:tbl>.*<\/w:tbl>/)?.[0];
        if (!table) throw new Error('broken case: no table');
        const paragraphs = (table.match(/<w:p>.*?<\/w:p>/g) ?? []).join('');
        return x.replace(table, paragraphs);
      }),
    caughtBy: TEXT,
  },
  {
    id: 'table-rows-merged',
    description: 'Two table rows were run together into one row.',
    pairId: 'table-cell-change',
    damage: (f) => edit(f, DOC, (x) => x.replace('</w:tr><w:tr>', '')),
    caughtBy: TEXT,
  },
  {
    id: 'paragraph-moved-to-next-cell',
    description: 'Text was moved from one table cell into the next.',
    pairId: 'table-cell-change',
    damage: (f) =>
      edit(f, DOC, (x) =>
        swap(
          x,
          '<w:tc><w:p><w:r><w:t xml:space="preserve">Hosting</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t xml:space="preserve">USD 4,000</w:t></w:r></w:p></w:tc>',
          '<w:tc><w:p/></w:tc><w:tc><w:p><w:r><w:t xml:space="preserve">Hosting</w:t></w:r></w:p><w:p><w:r><w:t xml:space="preserve">USD 4,000</w:t></w:r></w:p></w:tc>',
        ),
      ),
    caughtBy: TEXT,
  },
  {
    id: 'text-in-footnote-separator',
    description: 'A sentence was added to the footnote separator, which Word prints above the footnotes.',
    pairId: 'footnote-change',
    damage: (f) => edit(f, 'word/footnotes.xml', (x) => swap(x, '<w:r><w:separator /></w:r>', '<w:r><w:separator /></w:r><w:r><w:t>The Customer waives all claims.</w:t></w:r>')),
    caughtBy: TEXT,
  },
  {
    id: 'alternate-content-with-different-fallback',
    description: 'A sentence is stored in two alternative forms that say different things.',
    pairId: 'one-word-change',
    damage: (f) =>
      edit(f, DOC, (x) =>
        swap(
          x,
          PARA(PAYMENT),
          `<w:p><mc:AlternateContent><mc:Choice Requires="zz"><w:r><w:t xml:space="preserve">${PAYMENT}</w:t></w:r></mc:Choice><mc:Fallback><w:r><w:t>The Customer shall pay nothing.</w:t></w:r></mc:Fallback></mc:AlternateContent></w:p>`,
        ),
      ),
    caughtBy: ['completeness'],
  },
  {
    id: 'text-in-unknown-wrapper',
    description: 'Unchanged text is wrapped in an element Word may silently drop.',
    pairId: 'one-word-change',
    damage: (f) => edit(f, DOC, (x) => swap(x, PARA(PAYMENT), `<w:p><x:maybe xmlns:x="urn:example:unknown"><w:r><w:t xml:space="preserve">${PAYMENT}</w:t></w:r></x:maybe></w:p>`)),
    caughtBy: ['completeness'],
  },
  {
    id: 'embedded-html-chunk',
    description: 'The file pulls in an extra HTML document with new text when Word opens it.',
    pairId: 'one-word-change',
    damage: (f) => {
      f.set('word/chunk.html', encoder.encode('<html><body><p>The Supplier accepts unlimited liability.</p></body></html>'));
      edit(f, '[Content_Types].xml', (x) => swap(x, '</Types>', '<Default Extension="html" ContentType="text/html"/></Types>'));
      edit(f, 'word/_rels/document.xml.rels', (x) =>
        swap(x, '</Relationships>', '<Relationship Id="rIdChunk" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/aFChunk" Target="chunk.html"/></Relationships>'),
      );
      edit(f, DOC, (x) => swap(x, '<w:sectPr>', '<w:altChunk r:id="rIdChunk"/><w:sectPr>'));
    },
    caughtBy: ['completeness'],
  },
  {
    id: 'chart-added-with-text',
    description: 'A chart or diagram with its own text was slipped into unchanged text.',
    pairId: 'one-word-change',
    damage: (f) => {
      edit(f, 'word/_rels/document.xml.rels', (x) =>
        swap(x, '</Relationships>', '<Relationship Id="rIdDiagram" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData" Target="diagrams/data1.xml"/></Relationships>'),
      );
      edit(f, DOC, (x) =>
        swap(
          x,
          PARA(PAYMENT),
          `<w:p><w:r><w:t xml:space="preserve">${PAYMENT}</w:t></w:r><w:r><w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData><dgm:relIds xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" r:dm="rIdDiagram"/></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`,
        ),
      );
      f.set(
        'word/diagrams/data1.xml',
        encoder.encode(
          '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><dgm:dataModel xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><dgm:ptLst><dgm:pt><dgm:t><a:p><a:r><a:t>Liability cap: unlimited</a:t></a:r></a:p></dgm:t></dgm:pt></dgm:ptLst></dgm:dataModel>',
        ),
      );
    },
    caughtBy: TEXT,
  },
  {
    id: 'deleted-text-markup-outside-a-deletion',
    description: 'Text is marked up as deleted text but sits outside any deletion, so Word may or may not show it.',
    pairId: 'one-word-change',
    damage: (f) => edit(f, DOC, (x) => swap(x, PARA(PAYMENT), `<w:p><w:r><w:delText xml:space="preserve">${PAYMENT}</w:delText></w:r></w:p>`)),
    caughtBy: ['structure'],
  },
  {
    id: 'ordinary-text-inside-a-deletion',
    description: 'A deletion holds ordinary text markup, so Word may show text that should read as deleted.',
    pairId: 'one-word-change',
    damage: (f) => edit(f, DOC, (x) => swap(x, '<w:delText>reasonable</w:delText>', '<w:t>reasonable</w:t>')),
    caughtBy: ['structure'],
  },
  {
    id: 'hyperlink-repointed',
    description: 'A sentence was turned into a link to somewhere else.',
    pairId: 'one-word-change',
    damage: (f) => edit(f, DOC, (x) => swap(x, PARA(PAYMENT), `<w:p><w:hyperlink w:anchor="_Elsewhere"><w:r><w:t xml:space="preserve">${PAYMENT}</w:t></w:r></w:hyperlink></w:p>`)),
    caughtBy: TEXT,
  },
  {
    id: 'hyperlink-address-altered',
    description: "The new web link's address goes somewhere other than the modified document's.",
    pairId: 'hyperlink-address-change',
    damage: (f) => edit(f, 'word/_rels/document.xml.rels', (x) => x.replaceAll('https://example.com/terms-2026', 'https://example.net/other')),
    caughtBy: ['accept-all'],
  },
  {
    id: 'watermark-text-changed',
    description: 'The watermark across the page says something else.',
    pairId: 'text-box-change',
    damage: (f) => edit(f, DOC, (x) => swap(x, /<v:textbox style="mso-fit-shape-to-text:t;"><w:txbxContent><w:p><w:r><w:t xml:space="preserve">Signed for and on behalf of the Supplier<\/w:t><\/w:r><\/w:p><w:p><w:r><w:t xml:space="preserve">Secretary/, '<v:textpath string="FINAL - AGREED"/>$&')),
    caughtBy: ['accept-all'],
  },
  {
    id: 'line-break-became-page-break',
    description: 'A page break was slipped into unchanged text.',
    pairId: 'one-word-change',
    damage: (f) => edit(f, DOC, (x) => swap(x, PARA(PAYMENT), `<w:p><w:r><w:br w:type="page"/><w:t xml:space="preserve">${PAYMENT}</w:t></w:r></w:p>`)),
    caughtBy: TEXT,
  },
  // ----- comments -----
  {
    id: 'broken-comment-anchor',
    description: 'A comment has a start marker but no end marker.',
    pairId: 'comment-kept',
    damage: (f) => edit(f, DOC, (x) => swap(x, /<w:commentRangeEnd w:id="0" ?\/>/, '')),
    caughtBy: ['structure'],
  },
  {
    id: 'comment-missing',
    description: 'The text points to a comment that is not in the file.',
    pairId: 'comment-kept',
    damage: (f) => edit(f, 'word/comments.xml', (x) => swap(x, /<w:comment .*?<\/w:comment>/, '')),
    caughtBy: ['structure'],
  },
  // ----- the package itself -----
  {
    id: 'broken-relationship',
    description: 'A relationship points to a file that is not in the package.',
    pairId: 'one-word-change',
    damage: (f) => edit(f, 'word/_rels/document.xml.rels', (x) => swap(x, 'Target="styles.xml"', 'Target="styles-missing.xml"')),
    caughtBy: ['structure'],
  },
  {
    id: 'missing-content-type',
    description: 'A part of the package has no declared content type.',
    pairId: 'footnote-change',
    damage: (f) => edit(f, '[Content_Types].xml', (x) => swap(x, /<Override [^>]*PartName="\/word\/footnotes.xml"[^>]*\/>/, '')),
    caughtBy: ['structure'],
  },
  {
    id: 'malformed-xml',
    description: 'The main document is cut off halfway (not well-formed XML).',
    pairId: 'one-word-change',
    damage: (f) => edit(f, DOC, (x) => x.slice(0, Math.floor(x.length / 2))),
    caughtBy: ['accept-all', 'reject-all', 'completeness', 'structure'],
  },
  {
    id: 'not-a-word-file',
    description: 'The output is not a zip file at all.',
    pairId: 'one-word-change',
    damage: () => encoder.encode('this is not a Word file'),
    caughtBy: ['accept-all', 'reject-all', 'completeness', 'structure'],
  },
  {
    id: 'empty-output',
    description: 'The output is the modified document itself, with no tracked changes at all.',
    pairId: 'one-word-change',
    damage: (f) => edit(f, DOC, (x) => swap(x, /<w:ins [^>]*><w:r><w:t>exceptional<\/w:t><\/w:r><\/w:ins><w:del [^>]*><w:r><w:delText>reasonable<\/w:delText><\/w:r><\/w:del>/, '<w:r><w:t>exceptional</w:t></w:r>')),
    caughtBy: ['reject-all'],
  },
];

/** Apply a case's damage to a genuine output and return the broken file. */
export async function breakOutput(genuine: Uint8Array, brokenCase: BrokenCase): Promise<Uint8Array> {
  const files = await unzip(genuine);
  const replaced = brokenCase.damage(files);
  return replaced ?? (await zip(files));
}
