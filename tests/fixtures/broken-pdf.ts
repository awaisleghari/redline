/**
 * Deliberately wrong PDFs ("checking the Checker", check 9).
 *
 * Each case damages the genuine static redline of a generated pair and has the
 * real converter turn the DAMAGED file into a PDF. That PDF is then checked
 * against the GENUINE static redline, so it stands for a conversion that went
 * wrong in one specific way. Every one must be refused.
 */
import type { CheckId } from '../../src/checker';
import type { ZipEntries } from '../../src/common/zip';

export interface BrokenPdfCase {
  id: string;
  description: string;
  pairId: string;
  /** Damage the static redline before it is converted. */
  damage: (files: ZipEntries) => void;
  /** 'pdf-text', or 'pdf-colours' alone when every character is right but a colour is not. */
  caughtBy: CheckId[];
}

const decoder = new TextDecoder();
const encoder = new TextEncoder();
const DOC = 'word/document.xml';

function edit(files: ZipEntries, part: string, change: (xml: string) => string): void {
  const bytes = files.get(part);
  if (!bytes) throw new Error(`broken case: part ${part} not found`);
  const before = decoder.decode(bytes);
  const after = change(before);
  if (after === before) throw new Error(`broken case: the edit to ${part} changed nothing`);
  files.set(part, encoder.encode(after));
}

function swap(xml: string, from: string | RegExp, to: string): string {
  const matches = typeof from === 'string' ? xml.split(from).length - 1 : (xml.match(new RegExp(from.source, 'g')) ?? []).length;
  if (matches !== 1) throw new Error(`broken case: expected exactly one match for ${String(from)}, found ${matches}`);
  return xml.replace(from, to);
}

const PARA = (text: string): string => `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
const SERVICES_START = 'The Supplier shall provide the Services with ';
const PAYMENT = 'The Customer shall pay the Fees within thirty days of the date of each invoice.';
const LAW = 'This Agreement is governed by the laws of England and Wales.';
const TEXT: CheckId[] = ['pdf-text', 'pdf-colours'];

export const BROKEN_PDF_CASES: BrokenPdfCase[] = [
  // ----- the words -----
  { id: 'pdf-dropped-word', description: 'A word is missing from the PDF.', pairId: 'one-word-change', damage: (f) => edit(f, DOC, (x) => swap(x, ' skill and care.', ' and care.')), caughtBy: TEXT },
  { id: 'pdf-extra-word', description: 'The PDF has an extra "not" in the middle of a sentence.', pairId: 'one-word-change', damage: (f) => edit(f, DOC, (x) => swap(x, 'shall pay the Fees', 'shall not pay the Fees')), caughtBy: TEXT },
  { id: 'pdf-extra-digit', description: 'The PDF shows $300,000 where the redline inserts $30,000.', pairId: 'number-change', damage: (f) => edit(f, DOC, (x) => swap(x, '<w:t>$30,000</w:t>', '<w:t>$300,000</w:t>')), caughtBy: TEXT },
  { id: 'pdf-changed-digit', description: 'One digit differs in unchanged text.', pairId: 'one-word-change', damage: (f) => edit(f, DOC, (x) => swap(x, '$300,000', '$800,000')), caughtBy: TEXT },
  { id: 'pdf-changed-letter-case', description: 'A capital letter became lower case.', pairId: 'one-word-change', damage: (f) => edit(f, DOC, (x) => swap(x, 'pay the Fees', 'pay the fees')), caughtBy: TEXT },
  { id: 'pdf-dropped-full-stop', description: 'A full stop is missing.', pairId: 'one-word-change', damage: (f) => edit(f, DOC, (x) => swap(x, 'England and Wales.', 'England and Wales')), caughtBy: TEXT },
  { id: 'pdf-dropped-paragraph', description: 'A whole paragraph is missing.', pairId: 'one-word-change', damage: (f) => edit(f, DOC, (x) => swap(x, PARA(PAYMENT), '')), caughtBy: TEXT },
  { id: 'pdf-swapped-paragraphs', description: 'Two paragraphs are in the wrong order.', pairId: 'one-word-change', damage: (f) => edit(f, DOC, (x) => swap(x, PARA(PAYMENT), '').replace(PARA(LAW), PARA(LAW) + PARA(PAYMENT))), caughtBy: TEXT },
  { id: 'pdf-extra-paragraph', description: 'The PDF ends with a sentence that is not in the redline.', pairId: 'one-word-change', damage: (f) => edit(f, DOC, (x) => swap(x, PARA(LAW), PARA(LAW) + PARA('The Supplier accepts unlimited liability.'))), caughtBy: TEXT },
  { id: 'pdf-deleted-text-missing', description: 'The struck-through (deleted) text is not printed at all.', pairId: 'one-word-change', damage: (f) => edit(f, DOC, (x) => swap(x, /<w:r><w:rPr><w:strike\/><w:color w:val="FF0000"\/><\/w:rPr><w:t>reasonable<\/w:t><\/w:r>/, '')), caughtBy: TEXT },
  { id: 'pdf-text-hidden', description: 'A sentence is formatted as hidden text and so is not printed.', pairId: 'one-word-change', damage: (f) => edit(f, DOC, (x) => swap(x, PARA(LAW), `<w:p><w:r><w:rPr><w:vanish/></w:rPr><w:t xml:space="preserve">${LAW}</w:t></w:r></w:p>`)), caughtBy: TEXT },
  // ----- everything that is not the main text -----
  { id: 'pdf-cover-count-wrong', description: 'The cover page in the PDF shows a different count.', pairId: 'one-word-change', damage: (f) => edit(f, DOC, (x) => swap(x, 'Insertions: 1<', 'Insertions: 2<')), caughtBy: TEXT },
  { id: 'pdf-header-altered', description: 'The header in the PDF says something else.', pairId: 'header-change', damage: (f) => edit(f, 'word/header1.xml', (x) => swap(x, ' - Confidential', ' - Public')), caughtBy: TEXT },
  { id: 'pdf-footer-missing', description: 'The footer is not printed.', pairId: 'header-change', damage: (f) => edit(f, 'word/footer1.xml', (x) => swap(x, 'Alpha Holdings Limited', '')), caughtBy: TEXT },
  { id: 'pdf-footnote-altered', description: 'A footnote in the PDF says something else.', pairId: 'footnote-change', damage: (f) => edit(f, 'word/footnotes.xml', (x) => swap(x, ' of value added tax.', ' of sales tax.')), caughtBy: TEXT },
  { id: 'pdf-endnote-altered', description: 'An endnote in the PDF says something else.', pairId: 'endnote-change', damage: (f) => edit(f, 'word/endnotes.xml', (x) => swap(x, 'The courts of England', 'The courts of Scotland')), caughtBy: TEXT },
  { id: 'pdf-text-box-altered', description: 'Text in a text box differs.', pairId: 'text-box-change', damage: (f) => edit(f, DOC, (x) => swap(x, '>Secretary<', '>Treasurer<')), caughtBy: TEXT },
  { id: 'pdf-table-cell-altered', description: 'A table cell differs.', pairId: 'table-cell-change', damage: (f) => edit(f, DOC, (x) => swap(x, '>Hosting<', '>Housing<')), caughtBy: TEXT },
  { id: 'pdf-clause-letters-wrong', description: 'The list is numbered 1, 2, 3 in the PDF where the redline has (a), (b), (c).', pairId: 'numbered-list-change', damage: (f) => edit(f, 'word/numbering.xml', (x) => x.replaceAll('w:val="lowerLetter"', 'w:val="decimal"')), caughtBy: TEXT },
  { id: 'pdf-clause-number-missing', description: 'A numbered clause has no number in the PDF.', pairId: 'numbered-list-change', damage: (f) => edit(f, DOC, (x) => x.replace('<w:numPr><w:ilvl w:val="0"/><w:numId w:val="2"/></w:numPr>', '')), caughtBy: TEXT },
  { id: 'pdf-old-number-wrong', description: "A deleted clause's old number is wrong in the PDF.", pairId: 'numbered-clause-deleted', damage: (f) => edit(f, DOC, (x) => swap(x, '>(b)</w:t><w:tab/>', '>(d)</w:t><w:tab/>')), caughtBy: TEXT },
  // ----- found by the Stage 5 independent review: each of these once passed -----
  { id: 'pdf-last-clause-number-missing', description: 'The last numbered clause has no number in the PDF.', pairId: 'numbered-list-change', damage: (f) => edit(f, DOC, (x) => { const at = x.lastIndexOf('<w:numPr><w:ilvl w:val="0"/><w:numId w:val="2"/></w:numPr>'); return at < 0 ? x : x.slice(0, at) + x.slice(at + '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="2"/></w:numPr>'.length); }), caughtBy: TEXT },
  { id: 'pdf-all-clause-numbers-missing', description: 'No clause is numbered in the PDF.', pairId: 'numbered-list-change', damage: (f) => edit(f, DOC, (x) => x.replaceAll('<w:numPr><w:ilvl w:val="0"/><w:numId w:val="2"/></w:numPr>', '')), caughtBy: TEXT },
  { id: 'pdf-header-digit-wrong', description: 'The header says "Draft 9" where the redline inserts "2".', pairId: 'header-change', damage: (f) => edit(f, 'word/header1.xml', (x) => swap(x, '<w:t>2</w:t>', '<w:t>9</w:t>')), caughtBy: TEXT },
  { id: 'pdf-header-change-in-black', description: 'The inserted digit in the header is printed in black.', pairId: 'header-change', damage: (f) => edit(f, 'word/header1.xml', (x) => swap(x, '<w:color w:val="0000FF"/>', '<w:color w:val="000000"/>')), caughtBy: ['pdf-colours'] },
  { id: 'pdf-footer-date-wrong', description: 'A date and reference number in the footer are different.', pairId: 'page-numbered-header', damage: (f) => edit(f, 'word/footer1.xml', (x) => swap(x, 'Draft of 12 March 2026 - Ref 4471', 'Draft of 31 March 2029 - Ref 9999')), caughtBy: TEXT },
  { id: 'pdf-digits-at-end-of-document', description: 'A figure is printed after the last line of the document.', pairId: 'one-word-change', damage: (f) => edit(f, DOC, (x) => swap(x, PARA(LAW), PARA(LAW) + PARA('-£9,999,999'))), caughtBy: TEXT },
  { id: 'pdf-digits-before-footnote-mark', description: 'Digits are slipped in just before a footnote mark, so "1" reads as "501".', pairId: 'footnote-change', damage: (f) => edit(f, DOC, (x) => swap(x, 'of each invoice.</w:t>', 'of each invoice.50</w:t>')), caughtBy: TEXT },
  { id: 'pdf-word-at-end-in-document-with-endnotes', description: 'A short word made of roman-numeral letters is added at the end of a document that has endnotes.', pairId: 'endnote-change', damage: (f) => edit(f, DOC, (x) => { const at = x.lastIndexOf('<w:sectPr>'); return x.slice(0, at) + PARA('civil') + x.slice(at); }), caughtBy: TEXT },
  { id: 'pdf-bullets-printed-as-numbers', description: 'A bulleted list is printed as 1, 2, 3.', pairId: 'bulleted-list', damage: (f) => edit(f, 'word/numbering.xml', (x) => x.replaceAll('w:val="bullet"', 'w:val="decimal"').replace(/<w:lvlText w:val="[^"%]{1,2}" ?\/>/g, '<w:lvlText w:val="%1."/>')), caughtBy: TEXT },
  { id: 'pdf-footnotes-swapped', description: "Two footnotes are attached to each other's marks.", pairId: 'two-footnotes', damage: (f) => edit(f, 'word/footnotes.xml', (x) => x.replace('<w:footnote w:id="1">', '<w:footnote w:id="TEMP">').replace('<w:footnote w:id="2">', '<w:footnote w:id="1">').replace('<w:footnote w:id="TEMP">', '<w:footnote w:id="2">')), caughtBy: TEXT },
  { id: 'pdf-page-number-header-word-changed', description: 'A word in a page-numbered header is different.', pairId: 'page-numbered-header', damage: (f) => edit(f, 'word/header1.xml', (x) => swap(x, 'Services Agreement - Page ', 'Supply Agreement - Page ')), caughtBy: TEXT },
  { id: 'pdf-tab-leader-line-has-extra-text', description: 'Text is added on a signature line.', pairId: 'tab-leaders', damage: (f) => edit(f, DOC, (x) => swap(x, 'Signed:', 'Signed: A. N. Other')), caughtBy: TEXT },
  { id: 'pdf-table-heading-changed', description: 'The repeated table heading says something else.', pairId: 'repeating-table-heading', damage: (f) => edit(f, DOC, (x) => swap(x, '>Description<', '>Details<')), caughtBy: TEXT },
  { id: 'pdf-sentence-printed-among-footnotes', description: 'The last sentence of the document is printed down in the footnote instead of in the text.', pairId: 'footnote-change', damage: (f) => { edit(f, DOC, (x) => swap(x, PARA(LAW), '')); edit(f, 'word/footnotes.xml', (x) => swap(x, ' of value added tax.</w:t></w:r></w:p>', ` of value added tax.</w:t></w:r></w:p>${PARA(LAW)}`)); }, caughtBy: TEXT },
  // ----- found by the second Stage 5 review: each of these once passed -----
  { id: 'pdf-dot-before-figure-after-tab', description: '"Fee: 5 per cent" is printed "Fee: .5 per cent".', pairId: 'tabs-before-figures-symbol-and-bullets', damage: (f) => edit(f, DOC, (x) => swap(x, 'Fee:\t5 per cent', 'Fee:\t.5 per cent')), caughtBy: TEXT },
  { id: 'pdf-minus-before-figure-after-tab', description: '"Total 500" is printed "Total -500".', pairId: 'tabs-before-figures-symbol-and-bullets', damage: (f) => edit(f, DOC, (x) => swap(x, 'Total\t500', 'Total\t-500')), caughtBy: TEXT },
  { id: 'pdf-symbol-changed', description: 'A tick is printed as a cross.', pairId: 'tabs-before-figures-symbol-and-bullets', damage: (f) => edit(f, DOC, (x) => swap(x, 'w:char="F0FC"', 'w:char="F0FB"')), caughtBy: TEXT },
  { id: 'pdf-symbol-printed-as-digit', description: 'A tick is printed as the digit 9.', pairId: 'tabs-before-figures-symbol-and-bullets', damage: (f) => edit(f, DOC, (x) => swap(x, '<w:sym w:char="F0FC" w:font="Wingdings"/>', '<w:t>9</w:t>')), caughtBy: TEXT },
  { id: 'pdf-minus-after-bullet', description: 'A bulleted "500 units" is printed "-500 units".', pairId: 'tabs-before-figures-symbol-and-bullets', damage: (f) => edit(f, DOC, (x) => swap(x, '>500 units of stock;', '>-500 units of stock;')), caughtBy: TEXT },
  { id: 'pdf-bullet-changed', description: 'A bullet is printed as a dash.', pairId: 'bulleted-list', damage: (f) => edit(f, 'word/numbering.xml', (x) => x.replace(/<w:lvlText w:val="[^"%]{1,2}" ?\/>/g, '<w:lvlText w:val="-"/>')), caughtBy: TEXT },
  { id: 'pdf-inserted-clause-number-in-black', description: 'The number of an inserted clause is printed in black.', pairId: 'numbered-clause-inserted', damage: (f) => edit(f, DOC, (x) => swap(x, '</w:numPr><w:rPr><w:color w:val="0000FF"/><w:u w:val="double"/></w:rPr></w:pPr>', '</w:numPr></w:pPr>')), caughtBy: ['pdf-colours'] },
  { id: 'pdf-unchanged-clause-number-in-red', description: 'The number of an unchanged clause is printed in red and struck through.', pairId: 'numbered-clause-inserted', damage: (f) => edit(f, DOC, (x) => swap(x, '</w:numPr></w:pPr><w:r><w:t xml:space="preserve">either party', '</w:numPr><w:rPr><w:strike/><w:color w:val="FF0000"/></w:rPr></w:pPr><w:r><w:t xml:space="preserve">either party')), caughtBy: ['pdf-colours'] },
  { id: 'pdf-digits-at-foot-of-page', description: '"pay USD 5" is printed "pay USD 5000" at the foot of a page, in a document with page numbers.', pairId: 'page-edges-and-short-table', damage: (f) => edit(f, DOC, (x) => swap(x, 'pay USD 5</w:t>', 'pay USD 5000</w:t>')), caughtBy: TEXT },
  { id: 'pdf-digit-at-top-of-page', description: '"50 days" is printed "150 days" at the top of a page, in a document with page numbers.', pairId: 'page-edges-and-short-table', damage: (f) => edit(f, DOC, (x) => swap(x, '>50 days after', '>150 days after')), caughtBy: TEXT },
  { id: 'pdf-word-at-top-of-page', description: 'The word "mild" is added at the top of a page.', pairId: 'page-edges-and-short-table', damage: (f) => edit(f, DOC, (x) => swap(x, '>50 days after', '>mild 50 days after')), caughtBy: TEXT },
  { id: 'pdf-table-heading-at-top-of-page-outside-table', description: "A table's heading-row text is printed at the top of a page where no table continues.", pairId: 'page-edges-and-short-table', damage: (f) => edit(f, DOC, (x) => swap(x, '>50 days after', '>Item Code Agreed 50 days after')), caughtBy: TEXT },
  { id: 'pdf-table-heading-at-end-of-document', description: "A table's heading-row text is printed after the last line.", pairId: 'page-edges-and-short-table', damage: (f) => edit(f, DOC, (x) => swap(x, PARA('End.'), PARA('End.') + PARA('Item Code Agreed'))), caughtBy: TEXT },
  { id: 'pdf-page-total-wrong', description: 'The header says "of 35" in a PDF of three pages.', pairId: 'page-edges-and-short-table', damage: (f) => edit(f, 'word/header1.xml', (x) => swap(x, '</w:p></w:hdr>', '<w:r><w:t>5</w:t></w:r></w:p></w:hdr>')), caughtBy: TEXT },
  { id: 'pdf-page-number-wrong', description: 'The header says "Page 91" on page 1.', pairId: 'page-edges-and-short-table', damage: (f) => edit(f, 'word/header1.xml', (x) => swap(x, 'Services Agreement - Page ', 'Services Agreement - Page 9')), caughtBy: TEXT },
  { id: 'pdf-footer-page-total-wrong', description: 'A "Page X of Y" footer gives the wrong total.', pairId: 'page-x-of-y-footer', damage: (f) => edit(f, 'word/footer1.xml', (x) => swap(x, '</w:p></w:ftr>', '<w:r><w:t>0</w:t></w:r></w:p></w:ftr>')), caughtBy: TEXT },
  { id: 'pdf-dash-footer-number-wrong', description: 'A "- X -" footer prints a different number.', pairId: 'dash-page-number-footer', damage: (f) => edit(f, 'word/footer1.xml', (x) => swap(x, '>- </w:t>', '>- 1</w:t>')), caughtBy: TEXT },
  { id: 'pdf-bare-page-number-wrong', description: 'A footer that is only the page number prints a different number.', pairId: 'page-number-only-footer', damage: (f) => edit(f, 'word/footer1.xml', (x) => swap(x, '</w:p></w:ftr>', '<w:r><w:t>7</w:t></w:r></w:p></w:ftr>')), caughtBy: TEXT },
  // ----- the colours -----
  { id: 'pdf-deletion-in-black', description: 'Deleted text is printed in black.', pairId: 'one-word-change', damage: (f) => edit(f, DOC, (x) => swap(x, '<w:rPr><w:strike/><w:color w:val="FF0000"/></w:rPr><w:t>reasonable', '<w:rPr><w:strike/><w:color w:val="000000"/></w:rPr><w:t>reasonable')), caughtBy: ['pdf-colours'] },
  { id: 'pdf-insertion-in-red', description: 'Inserted text is printed in red.', pairId: 'one-word-change', damage: (f) => edit(f, DOC, (x) => swap(x, '<w:rPr><w:color w:val="0000FF"/><w:u w:val="double"/></w:rPr><w:t>exceptional', '<w:rPr><w:color w:val="FF0000"/><w:u w:val="double"/></w:rPr><w:t>exceptional')), caughtBy: ['pdf-colours'] },
  { id: 'pdf-move-in-blue', description: 'Moved text is printed in blue.', pairId: 'moved-paragraph', damage: (f) => edit(f, DOC, (x) => x.replaceAll('<w:color w:val="008000"/><w:u w:val="double"/></w:rPr><w:t xml:space="preserve">Each party', '<w:color w:val="0000FF"/><w:u w:val="double"/></w:rPr><w:t xml:space="preserve">Each party')), caughtBy: ['pdf-colours'] },
  { id: 'pdf-footnote-change-in-black', description: 'A change inside a footnote is printed in black.', pairId: 'footnote-change', damage: (f) => edit(f, 'word/footnotes.xml', (x) => swap(x, '<w:color w:val="0000FF"/>', '<w:color w:val="000000"/>')), caughtBy: ['pdf-colours'] },
  { id: 'pdf-legend-colour-wrong', description: 'The legend on the cover page shows "Inserted text" in green.', pairId: 'one-word-change', damage: (f) => edit(f, DOC, (x) => swap(x, '<w:color w:val="0000FF"/><w:u w:val="double"/></w:rPr><w:t xml:space="preserve">Inserted text', '<w:color w:val="008000"/><w:u w:val="double"/></w:rPr><w:t xml:space="preserve">Inserted text')), caughtBy: ['pdf-colours'] },
];

// SERVICES_START is kept for readers of the cases above: the changed word follows it.
void SERVICES_START;
