/**
 * Document pairs for the PDF check, covering what a page prints that the
 * document does not spell out: tab leaders, a table heading row repeated on
 * each page, page numbers in a header, bullets, several footnotes, an empty
 * numbered paragraph, and characters that print differently from how they are
 * typed. The Stage 5 independent review found the first version of check 9
 * wrongly withheld most of these.
 */
import {
  AlignmentType,
  Document,
  Footer,
  FootnoteReferenceRun,
  Header,
  LeaderType,
  LevelFormat,
  Packer,
  PageNumber,
  Paragraph,
  SectionType,
  SymbolRun,
  Table,
  TableCell,
  TableRow,
  TabStopType,
  TextRun,
  WidthType,
} from 'docx';

export interface PdfDocumentPair {
  id: string;
  description: string;
  build: () => Promise<{ original: Uint8Array; modified: Uint8Array }>;
}

type Body = (Paragraph | Table)[];
interface Extras {
  header?: Paragraph[];
  footer?: Paragraph[];
  footnotes?: Record<number, string>;
}

const text = (value: string): Paragraph => new Paragraph({ children: [new TextRun(value)] });
const numbered = (value: string): Paragraph => new Paragraph({ children: value === '' ? [] : [new TextRun(value)], numbering: { reference: 'clauses', level: 0 } });
const bulleted = (value: string): Paragraph => new Paragraph({ children: [new TextRun(value)], bullet: { level: 0 } });

async function pack(children: Body, extras: Extras = {}): Promise<Uint8Array> {
  const footnotes: Record<number, { children: Paragraph[] }> = {};
  for (const [id, value] of Object.entries(extras.footnotes ?? {})) footnotes[Number(id)] = { children: [text(value)] };
  const document = new Document({
    numbering: { config: [{ reference: 'clauses', levels: [{ level: 0, format: LevelFormat.DECIMAL, text: '%1.', alignment: AlignmentType.START }] }] },
    footnotes,
    sections: [
      {
        headers: extras.header ? { default: new Header({ children: extras.header }) } : undefined,
        footers: extras.footer ? { default: new Footer({ children: extras.footer }) } : undefined,
        children,
      },
    ],
  });
  return new Uint8Array(await Packer.toBuffer(document));
}

const pair = (id: string, description: string, make: (changed: boolean) => { body: Body; extras?: Extras }): PdfDocumentPair => ({
  id,
  description,
  build: async () => {
    const before = make(false);
    const after = make(true);
    return { original: await pack(before.body, before.extras), modified: await pack(after.body, after.extras) };
  },
});

const CLAUSE = 'The Supplier shall provide the Services with reasonable skill and care, in accordance with good industry practice and all applicable laws.';
const word = (changed: boolean): string => (changed ? 'exceptional' : 'reasonable');
const sentence = (changed: boolean): Paragraph => text(`The Supplier shall provide the Services with ${word(changed)} skill and care.`);

/** A pair built from whole documents (several sections, first-page and even-page headers). */
const custom = (id: string, description: string, make: (changed: boolean) => Document): PdfDocumentPair => ({
  id,
  description,
  build: async () => ({ original: new Uint8Array(await Packer.toBuffer(make(false))), modified: new Uint8Array(await Packer.toBuffer(make(true))) }),
});
const newPage = (value: string): Paragraph => new Paragraph({ pageBreakBefore: true, children: [new TextRun(value)] });
const pageNumber = (): Paragraph => new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ children: [PageNumber.CURRENT] })] });

export const PDF_DOCUMENT_PAIRS: PdfDocumentPair[] = [
  pair('tab-leaders', 'A signature line and a contents-style line, each drawn with a tab leader.', (changed) => ({
    body: [
      sentence(changed),
      new Paragraph({ tabStops: [{ type: TabStopType.RIGHT, position: 9000, leader: LeaderType.DOT }], children: [new TextRun('Schedule 1: Fees\t12')] }),
      new Paragraph({ tabStops: [{ type: TabStopType.RIGHT, position: 9000, leader: LeaderType.UNDERSCORE }], children: [new TextRun('Signed:\t')] }),
      text('Director'),
    ],
  })),
  pair('repeating-table-heading', 'A long table whose heading row is repeated at the top of each page.', (changed) => ({
    body: [
      sentence(false),
      new Table({
        width: { size: 100, type: WidthType.PERCENTAGE },
        rows: [
          new TableRow({ tableHeader: true, children: ['Item', 'Description', 'Fee'].map((value) => new TableCell({ children: [text(value)] })) }),
          ...Array.from({ length: 90 }, (_, index) => new TableRow({ children: [`${index + 1}`, `Service line ${index + 1}`, index === 60 && changed ? 'USD 2,500' : 'USD 1,500'].map((value) => new TableCell({ children: [text(value)] })) })),
        ],
      }),
      text('End of schedule.'),
    ],
  })),
  pair('page-numbered-header', 'Several pages with "Page X of Y" in the header and a reference in the footer.', (changed) => ({
    body: Array.from({ length: 45 }, (_, index) => text(`${index + 1}. ${index === 20 && changed ? CLAUSE.replace('reasonable', 'exceptional') : CLAUSE}`)),
    extras: {
      header: [new Paragraph({ children: [new TextRun('Services Agreement - Page '), new TextRun({ children: [PageNumber.CURRENT] }), new TextRun(' of '), new TextRun({ children: [PageNumber.TOTAL_PAGES] })] })],
      footer: [text('Draft of 12 March 2026 - Ref 4471')],
    },
  })),
  pair('page-number-only-footer', 'A footer that holds nothing but the page number.', (changed) => ({
    body: Array.from({ length: 45 }, (_, index) => text(index === 30 && changed ? CLAUSE.replace('reasonable', 'exceptional') : CLAUSE)),
    extras: { footer: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ children: [PageNumber.CURRENT] })] })] },
  })),
  pair('bulleted-list', 'A bulleted list with an item added.', (changed) => ({
    body: [text('The Supplier shall deliver:'), bulleted('a project plan;'), ...(changed ? [bulleted('a risk register;')] : []), bulleted('a monthly report;'), bulleted('the final deliverables.'), text('End.')],
  })),
  pair('empty-numbered-paragraph', 'A numbered list with an empty item in the middle.', (changed) => ({
    body: [text('The parties agree:'), numbered('the Fees are payable monthly;'), numbered(''), numbered(`the term is ${changed ? 'three' : 'two'} years;`), numbered('either party may terminate on notice.'), text('End.')],
  })),
  pair('two-footnotes', 'Two footnotes, the second of which changes.', (changed) => ({
    body: [
      new Paragraph({ children: [new TextRun('The Customer shall pay the Fees within thirty days.'), new FootnoteReferenceRun(1)] }),
      sentence(false),
      new Paragraph({ children: [new TextRun('This Agreement is governed by the laws of England and Wales.'), new FootnoteReferenceRun(2)] }),
    ],
    extras: { footnotes: { 1: 'Fees are stated exclusive of value added tax.', 2: `The courts of England have ${changed ? 'non-exclusive' : 'exclusive'} jurisdiction.` } },
  })),
  pair('special-characters', 'A non-breaking hyphen, a sharp s in capitals formatting, accents, quotes and dashes.', (changed) => ({
    body: [
      text('The licence is non‑exclusive and non‑transferable – see clause 4 — “the Licence”.'),
      new Paragraph({ children: [new TextRun('Registered office: '), new TextRun({ text: 'Hauptstraße 12, München', allCaps: true }), new TextRun('.')] }),
      new Paragraph({ children: [new TextRun({ text: 'Société Générale', smallCaps: true }), new TextRun(` acts as ${changed ? 'security agent' : 'agent'}.`)] }),
    ],
  })),
  // ----- found by the second Stage 5 review -----
  pair('page-x-of-y-footer', 'A footer that says only "Page X of Y", over several pages.', (changed) => ({
    body: Array.from({ length: 45 }, (_, index) => text(index === 12 && changed ? CLAUSE.replace('reasonable', 'exceptional') : CLAUSE)),
    extras: { footer: [new Paragraph({ children: [new TextRun('Page '), new TextRun({ children: [PageNumber.CURRENT] }), new TextRun(' of '), new TextRun({ children: [PageNumber.TOTAL_PAGES] })] })] },
  })),
  pair('dash-page-number-footer', 'A footer of the form "- X -".', (changed) => ({
    body: Array.from({ length: 45 }, (_, index) => text(index === 40 && changed ? CLAUSE.replace('reasonable', 'exceptional') : CLAUSE)),
    extras: { footer: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun('- '), new TextRun({ children: [PageNumber.CURRENT] }), new TextRun(' -')] })] },
  })),
  pair('short-page-number-header-and-footer', 'A header "Page X" and a footer "p. X/Y" on a short document.', (changed) => ({
    body: [sentence(changed), text('End.')],
    extras: {
      header: [new Paragraph({ children: [new TextRun('Page '), new TextRun({ children: [PageNumber.CURRENT] })] })],
      footer: [new Paragraph({ children: [new TextRun('p. '), new TextRun({ children: [PageNumber.CURRENT] }), new TextRun('/'), new TextRun({ children: [PageNumber.TOTAL_PAGES] })] })],
    },
  })),
  pair('header-and-footer-share-first-letters', 'The footer says "Confidential" and a page begins "Confidentiality"; the header says "Supply Agreement" and a page begins "Services".', (changed) => ({
    body: [
      text('Services are provided as set out below.'),
      sentence(changed),
      new Paragraph({ pageBreakBefore: true, children: [new TextRun('Confidentiality')] }),
      text('Each party shall keep the terms of this Agreement confidential.'),
      new Paragraph({ pageBreakBefore: true, children: [new TextRun('Supply Agreement terms prevail over any purchase order.')] }),
    ],
    extras: { header: [text('Supply Agreement')], footer: [text('Confidential')] },
  })),
  pair('page-edges-and-short-table', 'A page-numbered header, a short table with a heading row, and a new page that begins with a figure.', (changed) => ({
    body: [
      sentence(changed),
      new Table({
        width: { size: 100, type: WidthType.PERCENTAGE },
        rows: [
          new TableRow({ tableHeader: true, children: ['Item', 'Code', 'Agreed'].map((value) => new TableCell({ children: [text(value)] })) }),
          ...[['1', 'QX', 'No'], ['2', 'AB', 'Yes']].map((cells) => new TableRow({ children: cells.map((value) => new TableCell({ children: [text(value)] })) })),
        ],
      }),
      text('The Customer shall pay USD 5'),
      new Paragraph({ pageBreakBefore: true, children: [new TextRun('50 days after the date of the invoice the Fees fall due.')] }),
      text('End.'),
    ],
    extras: { header: [new Paragraph({ children: [new TextRun('Services Agreement - Page '), new TextRun({ children: [PageNumber.CURRENT] }), new TextRun(' of '), new TextRun({ children: [PageNumber.TOTAL_PAGES] })] })] },
  })),
  pair('tabs-before-figures-symbol-and-bullets', 'Figures after plain tabs, a symbol-font tick, and bulleted figures.', (changed) => ({
    body: [
      sentence(changed),
      new Paragraph({ tabStops: [{ type: TabStopType.LEFT, position: 3000 }], children: [new TextRun('Fee:\t5 per cent')] }),
      new Paragraph({ tabStops: [{ type: TabStopType.LEFT, position: 3000 }], children: [new TextRun('Total\t500')] }),
      new Paragraph({ children: [new TextRun('Approved '), new SymbolRun({ char: 'F0FC', symbolfont: 'Wingdings' }), new TextRun(' by the board')] }),
      bulleted('500 units of stock;'),
      bulleted('20 days of support.'),
    ],
  })),
  pair('numbered-clause-inserted', 'A numbered clause is added between two others.', (changed) => ({
    body: [text('The parties agree:'), numbered('the Fees are payable monthly;'), ...(changed ? [numbered('interest accrues on late payment;')] : []), numbered('either party may terminate on notice.'), text('End.')],
  })),
  // ----- found by the third Stage 5 review: headers and footers that do not belong on every page -----
  custom('first-page-without-footer', 'A title page with no footer; every other page has a footer that is only the page number. Pages end in figures.', (changed) =>
    new Document({
      sections: [
        {
          properties: { titlePage: true },
          footers: { default: new Footer({ children: [pageNumber()] }) },
          children: [text('Share Purchase Agreement'), sentence(changed), text('Draft: Version 1'), newPage('The Customer shall pay USD 7'), newPage('The deposit is USD 3'), text('End.')],
        },
      ],
    }),
  ),
  custom('two-sections-with-their-own-footers', 'An agreement and a schedule, each a section with its own page-number footer. Pages end in figures.', (changed) =>
    new Document({
      sections: [
        { footers: { default: new Footer({ children: [pageNumber()] }) }, children: [sentence(changed), text('The fee is GBP 1')] },
        { properties: { type: SectionType.NEXT_PAGE }, footers: { default: new Footer({ children: [pageNumber()] }) }, children: [text('Schedule 1'), text('Number of Shares: 2')] },
      ],
    }),
  ),
  custom('first-page-and-default-headers', 'A first-page header and footer, and a different header on the other pages; the document begins with the words of that other header.', (changed) =>
    new Document({
      sections: [
        {
          properties: { titlePage: true },
          headers: { first: new Header({ children: [text('Subject to contract')] }), default: new Header({ children: [text('Share Purchase Agreement')] }) },
          footers: { first: new Footer({ children: [text('Not for release')] }) },
          children: [text('Share Purchase Agreement'), sentence(changed), newPage('The offer made by the Buyer is binding.'), text('This press announcement is approved.')],
        },
      ],
    }),
  ),
  custom('odd-and-even-headers', 'Different headers on odd and even pages, with a page-number footer.', (changed) =>
    new Document({
      evenAndOddHeaderAndFooters: true,
      sections: [
        {
          headers: { default: new Header({ children: [text('Alpha Holdings Limited')] }), even: new Header({ children: [text('Share Purchase Agreement')] }) },
          footers: { default: new Footer({ children: [pageNumber()] }), even: new Footer({ children: [pageNumber()] }) },
          children: [sentence(changed), newPage('Second page.'), newPage('Third page.')],
        },
      ],
    }),
  ),
  pair('footnote-begins-like-the-next-page', 'A footnote that begins with the same words as the paragraph at the top of the next page.', (changed) => ({
    body: [
      new Paragraph({ children: [new TextRun('The Supplier'), new FootnoteReferenceRun(1), new TextRun(` shall deliver the ${changed ? 'Goods and Services' : 'Goods'}.`)] }),
      newPage('The Supplier means to perform this Agreement in full and on time.'),
    ],
    extras: { footnotes: { 1: 'The Supplier means Alpha Holdings Limited.' } },
  })),
];

export function pdfDocumentById(id: string): PdfDocumentPair | undefined {
  return PDF_DOCUMENT_PAIRS.find((p) => p.id === id);
}
