/**
 * Generated test pairs: small, contract-like Word documents built in code.
 *
 * Each pair has an original and a modified document that differ in exactly one,
 * known way, together with the tracked changes a correct comparison must
 * produce. Nothing here comes from a real client document.
 */
import {
  AlignmentType,
  Document,
  Footer,
  FootnoteReferenceRun,
  Header,
  LevelFormat,
  Packer,
  Paragraph,
  Table,
  TableCell,
  TableRow,
  TextRun,
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

type Block =
  | { kind: 'p'; text: string }
  | { kind: 'numbered'; text: string }
  | { kind: 'footnoted'; text: string; footnoteId: number }
  | { kind: 'table'; rows: string[][] };

interface DocSpec {
  blocks: Block[];
  header?: string;
  footer?: string;
  footnotes?: Record<number, string>;
}

const p = (text: string): Block => ({ kind: 'p', text });
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

const none = (): ExpectedChanges => ({ insertions: [], deletions: [], movedFrom: [], movedTo: [] });
const changes = (partial: Partial<ExpectedChanges>): ExpectedChanges => ({ ...none(), ...partial });

/** Replace the one block whose text equals `from`. Throws if it is not there exactly once. */
function replaceBlock(blocks: Block[], from: string, to: Block): Block[] {
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
    expected: { body: changes({ insertions: [ASSIGNMENT] }) },
  },
  {
    id: 'deleted-paragraph',
    description: 'A whole paragraph is removed.',
    original: { blocks: BASE },
    modified: { blocks: BASE.filter((b) => b.kind === 'table' || b.text !== PAYMENT) },
    expected: { body: changes({ deletions: [PAYMENT] }) },
  },
  {
    id: 'moved-paragraph',
    description: 'A paragraph is moved, unchanged, from near the top to near the bottom.',
    original: { blocks: [p(PARTIES), p(CONFIDENTIALITY), p(SERVICES), p(PAYMENT), p(LIABILITY), p(TERMINATION), p(GOVERNING_LAW)] },
    modified: { blocks: [p(PARTIES), p(SERVICES), p(PAYMENT), p(LIABILITY), p(TERMINATION), p(CONFIDENTIALITY), p(GOVERNING_LAW)] },
    expected: { body: changes({ movedFrom: [CONFIDENTIALITY], movedTo: [CONFIDENTIALITY] }) },
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
    expected: { body: changes({ insertions: [NEW_LIST_ITEM] }) },
  },
];

// ---------------------------------------------------------------------------
// Building the .docx files
// ---------------------------------------------------------------------------

const NUMBERING_REFERENCE = 'clause-list';

function renderBlock(block: Block): Paragraph | Table {
  switch (block.kind) {
    case 'p':
      return new Paragraph({ children: [new TextRun(block.text)] });
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
  const spec = PAIR_SPECS.find((s) => s.id === id);
  if (!spec) throw new Error(`no test pair called "${id}"`);
  return spec;
}
