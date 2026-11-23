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
  // ----- the colours -----
  { id: 'pdf-deletion-in-black', description: 'Deleted text is printed in black.', pairId: 'one-word-change', damage: (f) => edit(f, DOC, (x) => swap(x, '<w:rPr><w:strike/><w:color w:val="FF0000"/></w:rPr><w:t>reasonable', '<w:rPr><w:strike/><w:color w:val="000000"/></w:rPr><w:t>reasonable')), caughtBy: ['pdf-colours'] },
  { id: 'pdf-insertion-in-red', description: 'Inserted text is printed in red.', pairId: 'one-word-change', damage: (f) => edit(f, DOC, (x) => swap(x, '<w:rPr><w:color w:val="0000FF"/><w:u w:val="double"/></w:rPr><w:t>exceptional', '<w:rPr><w:color w:val="FF0000"/><w:u w:val="double"/></w:rPr><w:t>exceptional')), caughtBy: ['pdf-colours'] },
  { id: 'pdf-move-in-blue', description: 'Moved text is printed in blue.', pairId: 'moved-paragraph', damage: (f) => edit(f, DOC, (x) => x.replaceAll('<w:color w:val="008000"/><w:u w:val="double"/></w:rPr><w:t xml:space="preserve">Each party', '<w:color w:val="0000FF"/><w:u w:val="double"/></w:rPr><w:t xml:space="preserve">Each party')), caughtBy: ['pdf-colours'] },
  { id: 'pdf-footnote-change-in-black', description: 'A change inside a footnote is printed in black.', pairId: 'footnote-change', damage: (f) => edit(f, 'word/footnotes.xml', (x) => swap(x, '<w:color w:val="0000FF"/>', '<w:color w:val="000000"/>')), caughtBy: ['pdf-colours'] },
  { id: 'pdf-legend-colour-wrong', description: 'The legend on the cover page shows "Inserted text" in green.', pairId: 'one-word-change', damage: (f) => edit(f, DOC, (x) => swap(x, '<w:color w:val="0000FF"/><w:u w:val="double"/></w:rPr><w:t xml:space="preserve">Inserted text', '<w:color w:val="008000"/><w:u w:val="double"/></w:rPr><w:t xml:space="preserve">Inserted text')), caughtBy: ['pdf-colours'] },
];

// SERVICES_START is kept for readers of the cases above: the changed word follows it.
void SERVICES_START;
