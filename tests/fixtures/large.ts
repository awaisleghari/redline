/**
 * A generated contract of roughly 100 pages, with a modified version carrying a
 * realistic spread of edits. Used to measure how the engine copes with a large
 * document. The text is built from a fixed word list with a seeded random
 * number generator, so every run produces the same two files.
 */
import { Document, HeadingLevel, Packer, Paragraph, TextRun } from 'docx';

const WORDS = (
  'the Supplier Customer shall may must not provide deliver pay indemnify notify terminate assign perform ' +
  'Services Fees Agreement Party Parties Confidential Information Intellectual Property Rights Business Day ' +
  'notice written consent prior reasonable material breach obligation liability warranty representation ' +
  'in accordance with subject to pursuant under within following upon after before during any each all such ' +
  'other applicable law regulation claim loss damage cost expense third person affiliate subcontractor ' +
  'term period date invoice payment amount interest rate schedule clause paragraph provision remedy right ' +
  'and or of for to by on at as is be been has have will would including without limitation whether howsoever arising'
).split(' ');

/** Small seeded pseudo-random generator (mulberry32). */
function seeded(seed: number): () => number {
  let state = seed >>> 0;
  return () => {
    state = (state + 0x6d2b79f5) >>> 0;
    let t = state;
    t = Math.imul(t ^ (t >>> 15), t | 1);
    t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

function sentence(random: () => number, words: number): string {
  const parts: string[] = [];
  for (let i = 0; i < words; i++) parts.push(WORDS[Math.floor(random() * WORDS.length)]!);
  const text = parts.join(' ');
  return text.charAt(0).toUpperCase() + text.slice(1) + '.';
}

export interface LargeContract {
  original: Uint8Array;
  modified: Uint8Array;
  /** Number of body paragraphs in the original. */
  paragraphs: number;
  /** Number of words in the original. */
  words: number;
  /** How many edits of each kind the modified version carries. */
  edits: { wordChanges: number; insertedParagraphs: number; deletedParagraphs: number; movedParagraphs: number };
}

type Clause = { heading: boolean; text: string };

async function pack(clauses: Clause[]): Promise<Uint8Array> {
  const doc = new Document({
    creator: 'Redline test fixtures',
    sections: [
      {
        children: clauses.map((clause) =>
          clause.heading
            ? new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun(clause.text)] })
            : new Paragraph({ children: [new TextRun(clause.text)] }),
        ),
      },
    ],
  });
  return new Uint8Array(await Packer.toBuffer(doc));
}

/**
 * About 45,000 words: at a typical 450 words to a page of contract text, about
 * 100 pages.
 */
export async function buildLargeContract(sections = 100, clausesPerSection = 9): Promise<LargeContract> {
  const random = seeded(20261123);
  const original: Clause[] = [];
  for (let s = 1; s <= sections; s++) {
    original.push({ heading: true, text: `${s}. ${sentence(random, 4).replace('.', '')}` });
    for (let c = 1; c <= clausesPerSection; c++) {
      original.push({ heading: false, text: `${s}.${c} ${sentence(random, 30 + Math.floor(random() * 40))}` });
    }
  }

  const modified: Clause[] = original.map((clause) => ({ ...clause }));
  const edits = { wordChanges: 0, insertedParagraphs: 0, deletedParagraphs: 0, movedParagraphs: 0 };
  const editRandom = seeded(7);
  const bodyIndexes = (): number[] => modified.map((clause, i) => (clause.heading ? -1 : i)).filter((i) => i >= 0);
  const pick = (): number => {
    const candidates = bodyIndexes();
    return candidates[Math.floor(editRandom() * candidates.length)]!;
  };

  // 150 single-word replacements scattered through the document.
  for (let i = 0; i < 150; i++) {
    const clause = modified[pick()]!;
    const words = clause.text.split(' ');
    const at = 2 + Math.floor(editRandom() * (words.length - 3));
    words[at] = `amended${i}`;
    clause.text = words.join(' ');
    edits.wordChanges++;
  }
  // 20 deleted paragraphs.
  for (let i = 0; i < 20; i++) {
    modified.splice(pick(), 1);
    edits.deletedParagraphs++;
  }
  // 20 inserted paragraphs.
  for (let i = 0; i < 20; i++) {
    modified.splice(pick(), 0, { heading: false, text: `New clause ${i}: ${sentence(editRandom, 40)}` });
    edits.insertedParagraphs++;
  }
  // 5 paragraphs moved a long way.
  for (let i = 0; i < 5; i++) {
    const [moved] = modified.splice(pick(), 1);
    modified.splice(pick(), 0, moved!);
    edits.movedParagraphs++;
  }

  const words = original.reduce((sum, clause) => sum + clause.text.split(' ').length, 0);
  return {
    original: await pack(original),
    modified: await pack(modified),
    paragraphs: original.length,
    words,
    edits,
  };
}
