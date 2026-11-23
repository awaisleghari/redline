/**
 * Generated test pairs: small, contract-like Word documents built in code.
 *
 * Each pair has an original and a modified document that differ in exactly one,
 * known way, together with the tracked changes a correct comparison must
 * produce. Nothing here comes from a real client document.
 */
import {
  AlignmentType,
  CommentRangeEnd,
  CommentRangeStart,
  CommentReference,
  Document,
  EndnoteReferenceRun,
  Footer,
  FootnoteReferenceRun,
  Header,
  LevelFormat,
  Packer,
  Paragraph,
  Table,
  TableCell,
  TableRow,
  Textbox,
  TextRun,
  UnderlineType,
  WidthType,
} from 'docx';

/** Which part of the Word file a change sits in. */
export type PartKind = 'body' | 'header' | 'footer' | 'footnotes' | 'endnotes';

export interface ExpectedChanges {
  /** Text of each insertion, in document order. */
  insertions: string[];
  /** Text of each deletion, in document order. */
  deletions: string[];
  /** Text of each block moved away from its original place. */
  movedFrom: string[];
  /** Text of each block moved to its new place. */
  movedTo: string[];
  /** The moved-from and moved-to marks in document order, so a move's direction is checked. */
  moveOrder: ('from' | 'to')[];
  /** Paragraphs whose paragraph mark is tracked as inserted (the paragraph itself is new). */
  insertedParagraphMarks: number;
  /** Paragraphs whose paragraph mark is tracked as deleted (the paragraph itself is gone). */
  deletedParagraphMarks: number;
  /** Table rows tracked as inserted. */
  insertedRows: number;
  /** Table rows tracked as deleted. */
  deletedRows: number;
}

export interface PairSpec {
  /** Short machine-friendly name, used in test titles and file names. */
  id: string;
  /** Plain-English description of the difference. */
  description: string;
  original: DocSpec;
  modified: DocSpec;
  /** Expected tracked changes, per part of the file. Parts not listed must have none. */
  expected: Partial<Record<PartKind, ExpectedChanges>>;
}

export interface BuiltPair {
  spec: PairSpec;
  original: Uint8Array;
  modified: Uint8Array;
}

/** A stretch of text with its own character formatting. */
export interface RunSpec {
  text: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  allCaps?: boolean;
}

type Block =
  | { kind: 'p'; text: string }
  | { kind: 'rich'; text: string; runs: RunSpec[] }
  | { kind: 'endnoted'; text: string; endnoteId: number }
  | { kind: 'commented'; text: string; before: string; anchored: string; after: string; commentId: number }
  | { kind: 'textbox'; text: string; lines: string[] }
  | { kind: 'numbered'; text: string }
  | { kind: 'footnoted'; text: string; footnoteId: number }
  | { kind: 'table'; rows: string[][] };

export interface DocSpec {
  blocks: Block[];
  header?: string;
  footer?: string;
  footnotes?: Record<number, string>;
  endnotes?: Record<number, string>;
  comments?: Record<number, string>;
}

const p = (text: string): Block => ({ kind: 'p', text });
/** A paragraph made of separately formatted runs. `text` is their concatenation. */
const rich = (...runs: RunSpec[]): Block => ({ kind: 'rich', text: runs.map((r) => r.text).join(''), runs });
const numbered = (text: string): Block => ({ kind: 'numbered', text });

// ---------------------------------------------------------------------------
// Shared clause text
// ---------------------------------------------------------------------------

const PARTIES =
  'This Services Agreement is made between Alpha Holdings Limited (the Supplier) and Beta Trading LLC (the Customer).';
const SERVICES = 'The Supplier shall provide the Services with reasonable skill and care.';
const PAYMENT = 'The Customer shall pay the Fees within thirty days of the date of each invoice.';
const LIABILITY = 'The total liability of the Supplier under this Agreement shall not exceed $300,000 in aggregate.';
const TERMINATION =
  'Either party may terminate this Agreement by giving ninety days written notice to the other party.';
const CONFIDENTIALITY =
  'Each party shall keep the Confidential Information of the other party secret and shall not disclose it to any third person.';
const GOVERNING_LAW = 'This Agreement is governed by the laws of England and Wales.';

const BASE: Block[] = [p(PARTIES), p(SERVICES), p(PAYMENT), p(LIABILITY), p(TERMINATION), p(GOVERNING_LAW)];

const none = (): ExpectedChanges => ({
  insertions: [],
  deletions: [],
  movedFrom: [],
  movedTo: [],
  moveOrder: [],
  insertedParagraphMarks: 0,
  deletedParagraphMarks: 0,
  insertedRows: 0,
  deletedRows: 0,
});
const changes = (partial: Partial<ExpectedChanges>): ExpectedChanges => ({ ...none(), ...partial });

/** Replace the one block whose text equals `from`. Throws if it is not there exactly once. */
export function replaceBlock(blocks: Block[], from: string, to: Block): Block[] {
  const hits = blocks.filter((b) => b.kind !== 'table' && b.text === from).length;
  if (hits !== 1) throw new Error(`expected exactly one block with text "${from}", found ${hits}`);
  return blocks.map((b) => (b.kind !== 'table' && b.text === from ? to : b));
}

// ---------------------------------------------------------------------------
// The pairs
// ---------------------------------------------------------------------------

const ASSIGNMENT = 'Neither party may assign this Agreement without the prior written consent of the other party.';

const FEE_TABLE = [
  ['Service', 'Monthly fee'],
  ['Hosting', 'USD 4,000'],
  ['Support', 'USD 1,500'],
];
const FEE_TABLE_MODIFIED = [
  ['Service', 'Monthly fee'],
  ['Hosting', 'USD 4,000'],
  ['Support', 'USD 2,500'],
];

const LIST_INTRO = 'The Supplier shall deliver the following:';
const LIST = [
  'a project plan within ten Business Days;',
  'a monthly progress report;',
  'the final deliverables by the Completion Date.',
];
const NEW_LIST_ITEM = 'a risk register updated every quarter;';

export const PAIR_SPECS: PairSpec[] = [
  {
    id: 'one-word-change',
    description: 'One word is replaced in a sentence.',
    original: { blocks: BASE },
    modified: {
      blocks: replaceBlock(BASE, SERVICES, p('The Supplier shall provide the Services with exceptional skill and care.')),
    },
    expected: { body: changes({ deletions: ['reasonable'], insertions: ['exceptional'] }) },
  },
  {
    id: 'number-change',
    description: 'A money amount changes from $300,000 to $30,000.',
    original: { blocks: BASE },
    modified: {
      blocks: replaceBlock(
        BASE,
        LIABILITY,
        p('The total liability of the Supplier under this Agreement shall not exceed $30,000 in aggregate.'),
      ),
    },
    expected: { body: changes({ deletions: ['$300,000'], insertions: ['$30,000'] }) },
  },
  {
    id: 'inserted-paragraph',
    description: 'A whole new paragraph is added.',
    original: { blocks: BASE },
    modified: { blocks: [...BASE.slice(0, 5), p(ASSIGNMENT), ...BASE.slice(5)] },
    expected: { body: changes({ insertions: [ASSIGNMENT], insertedParagraphMarks: 1 }) },
  },
  {
    id: 'deleted-paragraph',
    description: 'A whole paragraph is removed.',
    original: { blocks: BASE },
    modified: { blocks: BASE.filter((b) => b.kind === 'table' || b.text !== PAYMENT) },
    expected: { body: changes({ deletions: [PAYMENT], deletedParagraphMarks: 1 }) },
  },
  {
    id: 'moved-paragraph',
    description: 'A paragraph is moved, unchanged, from near the top to near the bottom.',
    original: { blocks: [p(PARTIES), p(CONFIDENTIALITY), p(SERVICES), p(PAYMENT), p(LIABILITY), p(TERMINATION), p(GOVERNING_LAW)] },
    modified: { blocks: [p(PARTIES), p(SERVICES), p(PAYMENT), p(LIABILITY), p(TERMINATION), p(CONFIDENTIALITY), p(GOVERNING_LAW)] },
    // The clause leaves its old place near the top ('from' comes first in the
    // document) and arrives near the bottom; both paragraph marks are tracked.
    expected: {
      body: changes({
        movedFrom: [CONFIDENTIALITY],
        movedTo: [CONFIDENTIALITY],
        moveOrder: ['from', 'to'],
        insertedParagraphMarks: 1,
        deletedParagraphMarks: 1,
      }),
    },
  },
  {
    id: 'table-cell-change',
    description: 'One cell of a table changes.',
    original: { blocks: [p(PARTIES), p(PAYMENT), { kind: 'table', rows: FEE_TABLE }, p(GOVERNING_LAW)] },
    modified: { blocks: [p(PARTIES), p(PAYMENT), { kind: 'table', rows: FEE_TABLE_MODIFIED }, p(GOVERNING_LAW)] },
    expected: { body: changes({ deletions: ['1,500'], insertions: ['2,500'] }) },
  },
  {
    id: 'header-change',
    description: 'The page header changes; the body is identical.',
    original: { blocks: BASE, header: 'Draft 1 - Confidential', footer: 'Alpha Holdings Limited' },
    modified: { blocks: BASE, header: 'Draft 2 - Confidential', footer: 'Alpha Holdings Limited' },
    expected: { header: changes({ deletions: ['1'], insertions: ['2'] }) },
  },
  {
    id: 'footnote-change',
    description: 'The text of a footnote changes; the body is identical.',
    original: {
      blocks: replaceBlock(BASE, PAYMENT, { kind: 'footnoted', text: PAYMENT, footnoteId: 1 }),
      footnotes: { 1: 'Fees are stated exclusive of value added tax.' },
    },
    modified: {
      blocks: replaceBlock(BASE, PAYMENT, { kind: 'footnoted', text: PAYMENT, footnoteId: 1 }),
      footnotes: { 1: 'Fees are stated inclusive of value added tax.' },
    },
    expected: { footnotes: changes({ deletions: ['exclusive'], insertions: ['inclusive'] }) },
  },
  {
    id: 'numbered-list-change',
    description: 'A new item is added in the middle of an automatically numbered list.',
    original: { blocks: [p(PARTIES), p(LIST_INTRO), ...LIST.map(numbered), p(GOVERNING_LAW)] },
    modified: {
      blocks: [p(PARTIES), p(LIST_INTRO), numbered(LIST[0]!), numbered(LIST[1]!), numbered(NEW_LIST_ITEM), numbered(LIST[2]!), p(GOVERNING_LAW)],
    },
    expected: { body: changes({ insertions: [NEW_LIST_ITEM], insertedParagraphMarks: 1 }) },
  },
];

/**
 * Further pairs, beyond the nine the plan lists, covering parts of a Word file
 * the Checker must also read: footers, endnotes, text boxes, comments, table
 * rows, character formatting, and several changes at once.
 */
const FORMATTED_SERVICES: RunSpec[] = [
  { text: 'The Supplier shall provide the ' },
  { text: 'Services', bold: true },
  { text: ' with ' },
  { text: 'reasonable skill and care', italic: true },
  { text: ', and ' },
  { text: 'time is of the essence', underline: true },
  { text: '. ' },
  { text: 'Important', allCaps: true },
  { text: ' terms follow.' },
];

export const EXTRA_PAIR_SPECS: PairSpec[] = [
  {
    id: 'footer-change',
    description: 'The page footer changes; the body and header are identical.',
    original: { blocks: BASE, header: 'Confidential', footer: 'Alpha Holdings Limited' },
    modified: { blocks: BASE, header: 'Confidential', footer: 'Alpha Holdings plc' },
    expected: { footer: changes({ deletions: ['Limited'], insertions: ['plc'] }) },
  },
  {
    id: 'endnote-change',
    description: 'The text of an endnote changes; the body is identical.',
    original: {
      blocks: replaceBlock(BASE, GOVERNING_LAW, { kind: 'endnoted', text: GOVERNING_LAW, endnoteId: 1 }),
      endnotes: { 1: 'The courts of England have exclusive jurisdiction.' },
    },
    modified: {
      blocks: replaceBlock(BASE, GOVERNING_LAW, { kind: 'endnoted', text: GOVERNING_LAW, endnoteId: 1 }),
      endnotes: { 1: 'The courts of England have non-exclusive jurisdiction.' },
    },
    // The engine treats the hyphen as a word boundary, so only "non-" is new.
    expected: { endnotes: changes({ insertions: ['non-'] }) },
  },
  {
    id: 'text-box-change',
    description: 'A word changes inside a text box.',
    original: { blocks: [p(PARTIES), { kind: 'textbox', text: '', lines: ['Signed for and on behalf of the Supplier', 'Director'] }, p(GOVERNING_LAW)] },
    modified: { blocks: [p(PARTIES), { kind: 'textbox', text: '', lines: ['Signed for and on behalf of the Supplier', 'Secretary'] }, p(GOVERNING_LAW)] },
    // The engine does not compare inside a text box: it replaces the whole box,
    // deleting the old one and inserting the new one (each in its own paragraph).
    expected: {
      body: changes({
        deletions: ['Signed for and on behalf of the SupplierDirector'],
        insertions: ['Signed for and on behalf of the SupplierSecretary'],
        insertedParagraphMarks: 1,
        deletedParagraphMarks: 1,
      }),
    },
  },
  {
    id: 'comment-kept',
    description: 'A word changes in a document that carries a reviewer comment on other text.',
    original: {
      blocks: [
        p(PARTIES),
        { kind: 'commented', text: PAYMENT, before: 'The Customer shall pay the Fees within ', anchored: 'thirty days', after: ' of the date of each invoice.', commentId: 0 },
        p(SERVICES),
        p(GOVERNING_LAW),
      ],
      comments: { 0: 'Check the payment period with the client.' },
    },
    modified: {
      blocks: [
        p(PARTIES),
        { kind: 'commented', text: PAYMENT, before: 'The Customer shall pay the Fees within ', anchored: 'thirty days', after: ' of the date of each invoice.', commentId: 0 },
        p('The Supplier shall provide the Services with exceptional skill and care.'),
        p(GOVERNING_LAW),
      ],
      comments: { 0: 'Check the payment period with the client.' },
    },
    expected: { body: changes({ deletions: ['reasonable'], insertions: ['exceptional'] }) },
  },
  {
    id: 'table-row-added',
    description: 'A row is added to a table.',
    original: { blocks: [p(PARTIES), { kind: 'table', rows: FEE_TABLE.slice(0, 2) }, p(GOVERNING_LAW)] },
    modified: { blocks: [p(PARTIES), { kind: 'table', rows: FEE_TABLE }, p(GOVERNING_LAW)] },
    expected: { body: changes({ insertions: ['Support', 'USD 1,500'], insertedRows: 1, insertedParagraphMarks: 2 }) },
  },
  {
    id: 'table-row-deleted',
    description: 'A row is removed from a table.',
    original: { blocks: [p(PARTIES), { kind: 'table', rows: FEE_TABLE }, p(GOVERNING_LAW)] },
    modified: { blocks: [p(PARTIES), { kind: 'table', rows: FEE_TABLE.slice(0, 2) }, p(GOVERNING_LAW)] },
    expected: { body: changes({ deletions: ['Support', 'USD 1,500'], deletedRows: 1, deletedParagraphMarks: 2 }) },
  },
  {
    id: 'moved-block',
    description: 'Three consecutive paragraphs are moved together, unchanged, from near the top to near the bottom.',
    original: { blocks: [p(PARTIES), p(CONFIDENTIALITY), p(ASSIGNMENT), p(TERMINATION), p(SERVICES), p(PAYMENT), p(LIABILITY), p(GOVERNING_LAW)] },
    modified: { blocks: [p(PARTIES), p(SERVICES), p(PAYMENT), p(LIABILITY), p(CONFIDENTIALITY), p(ASSIGNMENT), p(TERMINATION), p(GOVERNING_LAW)] },
    expected: {
      body: changes({
        movedFrom: [CONFIDENTIALITY, ASSIGNMENT, TERMINATION],
        movedTo: [CONFIDENTIALITY, ASSIGNMENT, TERMINATION],
        moveOrder: ['from', 'from', 'from', 'to', 'to', 'to'],
        insertedParagraphMarks: 3,
        deletedParagraphMarks: 3,
      }),
    },
  },
  {
    id: 'formatted-text-word-change',
    description: 'One word changes in a paragraph that has bold, italic, underlined and all-capitals text.',
    original: { blocks: [p(PARTIES), rich(...FORMATTED_SERVICES), p(GOVERNING_LAW)] },
    modified: {
      blocks: [
        p(PARTIES),
        rich(...FORMATTED_SERVICES.map((run) => (run.text === ' terms follow.' ? { ...run, text: ' provisions follow.' } : run))),
        p(GOVERNING_LAW),
      ],
    },
    expected: { body: changes({ deletions: ['terms'], insertions: ['provisions'] }) },
  },
  {
    id: 'several-changes',
    description: 'A word change, a number change and a rewritten clause in one document, with header and footnote changes.',
    original: {
      blocks: replaceBlock(BASE, PAYMENT, { kind: 'footnoted', text: PAYMENT, footnoteId: 1 }),
      header: 'Draft 1 - Confidential',
      footnotes: { 1: 'Fees are stated exclusive of value added tax.' },
    },
    modified: {
      blocks: [
        p(PARTIES),
        p('The Supplier shall provide the Services with exceptional skill and care.'),
        { kind: 'footnoted', text: PAYMENT, footnoteId: 1 },
        p('The total liability of the Supplier under this Agreement shall not exceed $30,000 in aggregate.'),
        p(ASSIGNMENT),
        p(GOVERNING_LAW),
      ],
      header: 'Draft 2 - Confidential',
      footnotes: { 1: 'Fees are stated inclusive of value added tax.' },
    },
    expected: {
// The termination clause is replaced by the assignment clause in the same
      // position; the two share enough words that the engine shows it as one
      // paragraph edited word by word, not as a delete plus an insert.
      body: changes({
        insertions: ['exceptional', '$30,000', 'Neither', 'assign', 'without the prior', 'consent of'],
        deletions: ['reasonable', '$300,000', 'Either', 'terminate', 'by giving ninety days', 'notice to'],
      }),
      header: changes({ deletions: ['1'], insertions: ['2'] }),
      footnotes: changes({ deletions: ['exclusive'], insertions: ['inclusive'] }),
    },
  },
];

/** Every generated pair: the plan's nine plus the extras. */
export const ALL_PAIR_SPECS: PairSpec[] = [...PAIR_SPECS, ...EXTRA_PAIR_SPECS];

// ---------------------------------------------------------------------------
// Building the .docx files
// ---------------------------------------------------------------------------

const NUMBERING_REFERENCE = 'clause-list';

function renderBlock(block: Block): Paragraph | Table | Textbox {
  switch (block.kind) {
    case 'p':
      return new Paragraph({ children: [new TextRun(block.text)] });
    case 'rich':
      return new Paragraph({
        children: block.runs.map(
          (run) =>
            new TextRun({
              text: run.text,
              bold: run.bold,
              italics: run.italic,
              underline: run.underline ? { type: UnderlineType.SINGLE } : undefined,
              allCaps: run.allCaps,
            }),
        ),
      });
    case 'endnoted':
      return new Paragraph({ children: [new TextRun(block.text), new EndnoteReferenceRun(block.endnoteId)] });
    case 'commented':
      return new Paragraph({
        children: [
          new TextRun(block.before),
          new CommentRangeStart(block.commentId),
          new TextRun(block.anchored),
          new CommentRangeEnd(block.commentId),
          new TextRun({ children: [new CommentReference(block.commentId)] }),
          new TextRun(block.after),
        ],
      });
    case 'textbox':
      return new Textbox({
        alignment: 'left',
        style: { width: '300pt', height: '60pt' },
        children: block.lines.map((line) => new Paragraph({ children: [new TextRun(line)] })),
      });
    case 'numbered':
      return new Paragraph({
        children: [new TextRun(block.text)],
        numbering: { reference: NUMBERING_REFERENCE, level: 0 },
      });
    case 'footnoted':
      return new Paragraph({ children: [new TextRun(block.text), new FootnoteReferenceRun(block.footnoteId)] });
    case 'table':
      return new Table({
        width: { size: 100, type: WidthType.PERCENTAGE },
        rows: block.rows.map(
          (cells) =>
            new TableRow({
              children: cells.map(
                (text) => new TableCell({ children: [new Paragraph({ children: [new TextRun(text)] })] }),
              ),
            }),
        ),
      });
  }
}

export async function buildDocx(spec: DocSpec): Promise<Uint8Array> {
  const footnotes: Record<number, { children: Paragraph[] }> = {};
  for (const [id, text] of Object.entries(spec.footnotes ?? {})) {
    footnotes[Number(id)] = { children: [new Paragraph({ children: [new TextRun(text)] })] };
  }
  const endnotes: Record<number, { children: Paragraph[] }> = {};
  for (const [id, text] of Object.entries(spec.endnotes ?? {})) {
    endnotes[Number(id)] = { children: [new Paragraph({ children: [new TextRun(text)] })] };
  }
  const doc = new Document({
    creator: 'Redline test fixtures',
    title: 'Generated test document',
    numbering: {
      config: [
        {
          reference: NUMBERING_REFERENCE,
          levels: [{ level: 0, format: LevelFormat.LOWER_LETTER, text: '(%1)', alignment: AlignmentType.START }],
        },
      ],
    },
    footnotes,
    endnotes,
    comments: spec.comments
      ? {
          children: Object.entries(spec.comments).map(([id, text]) => ({
            id: Number(id),
            author: 'Reviewer',
            date: new Date('2026-01-15T10:00:00Z'),
            children: [new Paragraph({ children: [new TextRun(text)] })],
          })),
        }
      : undefined,
    sections: [
      {
        headers: spec.header
          ? { default: new Header({ children: [new Paragraph({ children: [new TextRun(spec.header)] })] }) }
          : undefined,
        footers: spec.footer
          ? { default: new Footer({ children: [new Paragraph({ children: [new TextRun(spec.footer)] })] }) }
          : undefined,
        children: spec.blocks.map(renderBlock),
      },
    ],
  });
  return new Uint8Array(await Packer.toBuffer(doc));
}

export async function buildPair(spec: PairSpec): Promise<BuiltPair> {
  const [original, modified] = await Promise.all([buildDocx(spec.original), buildDocx(spec.modified)]);
  return { spec, original, modified };
}

export function pairById(id: string): PairSpec {
  const spec = ALL_PAIR_SPECS.find((s) => s.id === id);
  if (!spec) throw new Error(`no test pair called "${id}"`);
  return spec;
}
