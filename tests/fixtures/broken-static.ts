/**
 * Deliberately broken static redlines ("checking the Checker", Stage 3).
 * Each takes the genuine static redline of a generated pair and damages it in
 * one way. `caughtBy` lists the checks that catch the damage directly;
 * `expectedStaticFailures` adds the checks that cannot run once a text check
 * has failed.
 */
import type { CheckId } from '../../src/checker';
import type { ZipEntries } from '../../src/common/zip';

export interface BrokenStaticCase {
  id: string;
  description: string;
  pairId: string;
  damage: (files: ZipEntries) => void;
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

const DEL = '<w:rPr><w:strike/><w:color w:val="FF0000"/></w:rPr>';
const INS = '<w:rPr><w:color w:val="0000FF"/><w:u w:val="double"/></w:rPr>';
const SHADE = '<w:shd w:val="clear" w:color="auto" w:fill="FFF2CC"/>';

export const BROKEN_STATIC_CASES: BrokenStaticCase[] = [
  // ----- check 3: the two documents must be recoverable from the formatting -----
  {
    id: 'static-dropped-word',
    description: 'A word of unchanged text is missing.',
    pairId: 'one-word-change',
    damage: (f) => edit(f, DOC, (x) => swap(x, ' skill and care.', ' and care.')),
    caughtBy: ['static-new', 'static-old'],
  },
  {
    id: 'static-deletion-styled-as-insertion',
    description: 'Deleted text is formatted as an insertion.',
    pairId: 'one-word-change',
    damage: (f) => edit(f, DOC, (x) => swap(x, `<w:r>${DEL}<w:t>reasonable</w:t></w:r>`, `<w:r>${INS}<w:t>reasonable</w:t></w:r>`)),
    caughtBy: ['static-new', 'static-old', 'cover-page'],
  },
  {
    id: 'static-deletion-not-struck',
    description: 'Deleted text is red but not struck through, so it reads as ordinary text.',
    pairId: 'one-word-change',
    damage: (f) => edit(f, DOC, (x) => swap(x, `<w:r>${DEL}<w:t>reasonable</w:t></w:r>`, '<w:r><w:rPr><w:color w:val="FF0000"/></w:rPr><w:t>reasonable</w:t></w:r>')),
    caughtBy: ['static-new', 'cover-page'],
  },
  {
    id: 'static-insertion-wrong-colour',
    description: 'Inserted text is double underlined but not in the insertion colour.',
    pairId: 'one-word-change',
    damage: (f) => edit(f, DOC, (x) => swap(x, `<w:r>${INS}<w:t>exceptional</w:t></w:r>`, '<w:r><w:rPr><w:color w:val="1F3864"/><w:u w:val="double"/></w:rPr><w:t>exceptional</w:t></w:r>')),
    caughtBy: ['static-old', 'formatting', 'cover-page'],
  },
  {
    id: 'static-insertion-unmarked',
    description: 'Inserted text carries no marking at all.',
    pairId: 'one-word-change',
    damage: (f) => edit(f, DOC, (x) => swap(x, `<w:r>${INS}<w:t>exceptional</w:t></w:r>`, '<w:r><w:t>exceptional</w:t></w:r>')),
    caughtBy: ['static-old', 'cover-page'],
  },
  {
    id: 'static-wrong-digit-in-insertion',
    description: 'The inserted amount is wrong ($30,000 shown as $80,000).',
    pairId: 'number-change',
    damage: (f) => edit(f, DOC, (x) => swap(x, '<w:t>$30,000</w:t>', '<w:t>$80,000</w:t>')),
    caughtBy: ['static-new'],
  },
  {
    id: 'static-deleted-paragraph-mark-unmarked',
    description: "A deleted paragraph's text is struck but the paragraph itself is not marked as deleted.",
    pairId: 'deleted-paragraph',
    damage: (f) => edit(f, DOC, (x) => swap(x, `<w:pPr><w:numPr><w:numId w:val="0"/></w:numPr>${DEL}</w:pPr>`, '<w:pPr><w:numPr><w:numId w:val="0"/></w:numPr></w:pPr>')),
    // (The count of deletions is unaffected: the struck text is still one unbroken deletion.)
    caughtBy: ['static-new'],
  },
  {
    id: 'static-move-shown-as-deletion',
    description: 'A moved-from clause is formatted as a plain deletion.',
    pairId: 'moved-paragraph',
    damage: (f) => edit(f, DOC, (x) => x.replaceAll('<w:rPr><w:dstrike/><w:color w:val="008000"/></w:rPr>', DEL)),
    caughtBy: ['cover-page'],
  },
  {
    id: 'static-header-altered',
    description: 'Unchanged header text is different.',
    pairId: 'header-change',
    damage: (f) => edit(f, 'word/header1.xml', (x) => swap(x, ' - Confidential', ' - Public')),
    caughtBy: ['static-new', 'static-old'],
  },
  {
    id: 'static-deleted-row-left-standing',
    description: "A deleted table row's cells are not all marked as deleted, so an empty row would survive.",
    pairId: 'table-row-deleted',
    damage: (f) =>
      edit(f, DOC, (x) => {
        const marks = [...x.matchAll(/<w:tc><w:p><w:pPr>(<w:numPr><w:numId w:val="0"\/><\/w:numPr>)?<w:rPr><w:strike\/><w:color w:val="FF0000"\/><\/w:rPr><\/w:pPr>/g)];
        if (marks.length !== 2) throw new Error(`broken case: expected 2 deleted cells, found ${marks.length}`);
        const last = marks[1]!;
        return x.slice(0, last.index) + '<w:tc><w:p><w:pPr></w:pPr>' + x.slice(last.index + last[0].length);
      }),
    caughtBy: ['static-new'],
  },
  // ----- check 4: no tracked changes may remain -----
  {
    id: 'static-leftover-tracked-change',
    description: 'A tracked insertion is still in the file.',
    pairId: 'one-word-change',
    damage: (f) => edit(f, DOC, (x) => swap(x, `<w:r>${INS}<w:t>exceptional</w:t></w:r>`, `<w:ins w:id="900" w:author="x" w:date="2026-01-01T00:00:00Z"><w:r>${INS}<w:t>exceptional</w:t></w:r></w:ins>`)),
    caughtBy: ['no-tracked-changes'],
  },
  {
    id: 'static-leftover-formatting-record',
    description: 'A tracked formatting-change record is still in the file.',
    pairId: 'one-word-change',
    damage: (f) =>
      edit(f, DOC, (x) => swap(x, `${INS}<w:t>exceptional`, `${INS.replace('</w:rPr>', '<w:rPrChange w:id="901" w:author="x" w:date="2026-01-01T00:00:00Z"><w:rPr/></w:rPrChange></w:rPr>')}<w:t>exceptional`)),
    caughtBy: ['no-tracked-changes'],
  },
  {
    id: 'static-leftover-move-marker',
    description: 'A move range marker is still in the file.',
    pairId: 'moved-paragraph',
    damage: (f) =>
      edit(f, DOC, (x) =>
        swap(x, '<w:p><w:r><w:t xml:space="preserve">This Agreement is governed', '<w:p><w:moveFromRangeEnd w:id="5"/><w:r><w:t xml:space="preserve">This Agreement is governed'),
      ),
    caughtBy: ['no-tracked-changes', 'structure'],
  },
  // ----- check 5: the cover page -----
  {
    id: 'static-wrong-cover-count',
    description: 'The cover page claims two insertions where there is one.',
    pairId: 'one-word-change',
    damage: (f) => edit(f, DOC, (x) => swap(x, 'Insertions: 1<', 'Insertions: 2<')),
    caughtBy: ['cover-page'],
  },
  {
    id: 'static-wrong-cover-file-name',
    description: 'The cover page names a different file.',
    pairId: 'one-word-change',
    damage: (f) => edit(f, DOC, (x) => swap(x, 'Modified: modified.docx<', 'Modified: other.docx<')),
    caughtBy: ['cover-page'],
  },
  {
    id: 'static-cover-page-missing',
    description: 'The cover page is missing.',
    pairId: 'one-word-change',
    damage: (f) => edit(f, DOC, (x) => x.replace(/<w:body>.*?<w:sectPr>.*?<\/w:sectPr><\/w:pPr><\/w:p>/, '<w:body>')),
    caughtBy: ['cover-page'],
  },
  {
    id: 'static-text-hidden-in-cover-style',
    description: 'A sentence of the document was given the cover-page style, which would exempt it from the text checks.',
    pairId: 'one-word-change',
    damage: (f) =>
      edit(f, DOC, (x) =>
        swap(
          x,
          '<w:p><w:r><w:t xml:space="preserve">This Agreement is governed by the laws of England and Wales.</w:t></w:r></w:p>',
          '<w:p><w:pPr><w:pStyle w:val="RedlineCoverText"/></w:pPr><w:r><w:t xml:space="preserve">This Agreement is governed by the laws of England and Wales.</w:t></w:r></w:p>',
        ),
      ),
    caughtBy: ['completeness'],
  },
  // ----- found by the Stage 3 independent review: each of these once passed -----
  {
    id: 'static-deletion-shown-as-move-with-cover-adjusted',
    description: 'A deletion is formatted as moved-from text and the cover page is altered to agree, though nothing arrives anywhere.',
    pairId: 'one-word-change',
    damage: (f) =>
      edit(f, DOC, (x) =>
        swap(swap(swap(x, `<w:r>${DEL}<w:t>reasonable</w:t></w:r>`, '<w:r><w:rPr><w:dstrike/><w:color w:val="008000"/></w:rPr><w:t>reasonable</w:t></w:r>'), 'Deletions: 1<', 'Deletions: 0<'), 'Moves: 0<', 'Moves: 1<'),
      ),
    caughtBy: ['cover-page'],
  },
  {
    id: 'static-legend-sample-in-wrong-style',
    description: 'The legend shows "Inserted text" in the deletion style.',
    pairId: 'one-word-change',
    damage: (f) => edit(f, DOC, (x) => swap(x, `${INS}<w:t xml:space="preserve">Inserted text`, `${DEL}<w:t xml:space="preserve">Inserted text`)),
    caughtBy: ['cover-page'],
  },
  {
    id: 'static-cover-count-hidden',
    description: 'A count on the cover page is formatted as hidden text.',
    pairId: 'one-word-change',
    damage: (f) => edit(f, DOC, (x) => swap(x, '<w:r><w:t xml:space="preserve">Deletions: 1</w:t>', '<w:r><w:rPr><w:vanish/></w:rPr><w:t xml:space="preserve">Deletions: 1</w:t>')),
    caughtBy: ['cover-page'],
  },
  {
    id: 'static-cover-style-on-a-changed-table-cell',
    description: 'A changed paragraph inside a table is given the cover-page style, which would hide it from the counts.',
    pairId: 'table-cell-change',
    damage: (f) => edit(f, DOC, (x) => swap(x, '<w:tc><w:p><w:r><w:t xml:space="preserve">USD </w:t>', '<w:tc><w:p><w:pPr><w:pStyle w:val="RedlineCoverText"/></w:pPr><w:r><w:t xml:space="preserve">USD </w:t>')),
    // Refused outright, and the counts no longer agree either (cover-styled paragraphs are not counted).
    caughtBy: ['completeness', 'cover-page'],
  },
  {
    id: 'static-text-inside-deleted-box-unmarked',
    description: 'Text inside a deleted text box is not itself marked as deleted.',
    pairId: 'text-box-change',
    damage: (f) => edit(f, DOC, (x) => swap(x, `<w:r>${DEL}<w:t xml:space="preserve">Director</w:t></w:r>`, '<w:r><w:t xml:space="preserve">Director</w:t></w:r>')),
    caughtBy: ['completeness', 'cover-page'],
  },
  {
    id: 'static-insertion-also-struck',
    description: 'Inserted text is also struck through, so it reads as deleted.',
    pairId: 'one-word-change',
    damage: (f) => edit(f, DOC, (x) => swap(x, `<w:r>${INS}<w:t>exceptional</w:t></w:r>`, '<w:r><w:rPr><w:strike/><w:color w:val="0000FF"/><w:u w:val="double"/></w:rPr><w:t>exceptional</w:t></w:r>')),
    caughtBy: ['formatting'],
  },
  {
    id: 'static-unchanged-text-struck',
    description: 'An unchanged sentence is struck through in black.',
    pairId: 'one-word-change',
    damage: (f) =>
      edit(f, DOC, (x) =>
        swap(
          x,
          '<w:p><w:r><w:t xml:space="preserve">This Agreement is governed by the laws of England and Wales.</w:t></w:r></w:p>',
          '<w:p><w:r><w:rPr><w:strike/></w:rPr><w:t xml:space="preserve">This Agreement is governed by the laws of England and Wales.</w:t></w:r></w:p>',
        ),
      ),
    caughtBy: ['formatting', 'format-marks'],
  },
  {
    id: 'static-theme-colour-over-deletion-red',
    description: 'Deleted text carries a theme colour that Word shows in place of red.',
    pairId: 'one-word-change',
    damage: (f) => edit(f, DOC, (x) => swap(x, `<w:r>${DEL}<w:t>reasonable</w:t></w:r>`, '<w:r><w:rPr><w:strike/><w:color w:val="FF0000" w:themeColor="text1"/></w:rPr><w:t>reasonable</w:t></w:r>')),
    // It reads as ordinary struck-through text: wrong on the modified side, and unmarked strikethrough on the original side.
    caughtBy: ['static-new', 'cover-page', 'format-marks'],
  },
  {
    id: 'static-old-number-hidden',
    description: "A deleted clause's old number is formatted as hidden text.",
    pairId: 'numbered-clause-deleted',
    damage: (f) => edit(f, DOC, (x) => swap(x, '<w:rStyle w:val="RedlineOldNumber"/><w:strike/>', '<w:rStyle w:val="RedlineOldNumber"/><w:strike/><w:vanish/>')),
    caughtBy: ['completeness'],
  },
  // ----- check 8 and formatting marks -----
  {
    id: 'static-bold-lost',
    description: 'Unchanged bold text is no longer bold.',
    pairId: 'formatted-text-word-change',
    damage: (f) => edit(f, DOC, (x) => swap(x, /<w:r><w:rPr><w:b\/><w:bCs\/><\/w:rPr>(<w:t[^>]*>Services<\/w:t>)<\/w:r>/, '<w:r>$1</w:r>')),
    caughtBy: ['formatting', 'format-marks'],
  },
  {
    id: 'static-text-hidden',
    description: 'An unchanged sentence was made hidden text.',
    pairId: 'one-word-change',
    damage: (f) =>
      edit(f, DOC, (x) =>
        swap(
          x,
          '<w:p><w:r><w:t xml:space="preserve">This Agreement is governed by the laws of England and Wales.</w:t></w:r></w:p>',
          '<w:p><w:r><w:rPr><w:vanish/></w:rPr><w:t xml:space="preserve">This Agreement is governed by the laws of England and Wales.</w:t></w:r></w:p>',
        ),
      ),
    caughtBy: ['formatting', 'format-marks'],
  },
  {
    id: 'static-format-change-not-shaded',
    description: 'Text that became bold is not marked as a formatting change.',
    pairId: 'formatting-only-change',
    damage: (f) => edit(f, DOC, (x) => swap(x, `<w:rPr><w:b/><w:bCs/>${SHADE}</w:rPr>`, '<w:rPr><w:b/><w:bCs/></w:rPr>')),
    caughtBy: ['format-marks', 'cover-page'],
  },
  {
    id: 'static-shaded-without-format-change',
    description: 'Text whose formatting did not change is marked as a formatting change.',
    pairId: 'formatting-only-change',
    damage: (f) =>
      edit(f, DOC, (x) =>
        swap(
          x,
          '<w:r><w:t xml:space="preserve">This Agreement is governed by the laws of England and Wales.</w:t></w:r>',
          `<w:r><w:rPr>${SHADE}</w:rPr><w:t xml:space="preserve">This Agreement is governed by the laws of England and Wales.</w:t></w:r>`,
        ),
      ),
    caughtBy: ['format-marks', 'cover-page'],
  },
  // ----- clause numbers -----
  {
    id: 'static-old-number-wrong',
    description: 'A deleted clause shows the wrong old number.',
    pairId: 'numbered-clause-deleted',
    damage: (f) => edit(f, DOC, (x) => swap(x, '>(b)</w:t><w:tab/>', '>(d)</w:t><w:tab/>')),
    caughtBy: ['numbering'],
  },
  {
    id: 'static-old-number-missing',
    description: 'A deleted clause does not show its old number.',
    pairId: 'numbered-clause-deleted',
    damage: (f) => edit(f, DOC, (x) => swap(x, /<w:r><w:rPr><w:rStyle w:val="RedlineOldNumber"\/>.*?<\/w:r>/, '')),
    caughtBy: ['numbering'],
  },
  {
    id: 'static-deleted-clause-still-numbered',
    description: 'A deleted clause keeps its automatic number, so every later clause is numbered one too high.',
    pairId: 'numbered-clause-deleted',
    damage: (f) => edit(f, DOC, (x) => swap(x, '<w:numPr><w:numId w:val="0"/></w:numPr>', '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="2"/></w:numPr>')),
    caughtBy: ['numbering'],
  },
  {
    id: 'static-clause-unnumbered',
    description: 'A surviving clause lost its automatic number.',
    pairId: 'numbered-list-change',
    damage: (f) => edit(f, DOC, (x) => x.replace('<w:numPr><w:ilvl w:val="0"/><w:numId w:val="2"/></w:numPr>', '')),
    caughtBy: ['numbering'],
  },
  // ----- the package -----
  {
    id: 'static-broken-comment-anchor',
    description: 'A comment has a start marker but no end marker.',
    pairId: 'comment-kept',
    damage: (f) => edit(f, DOC, (x) => swap(x, /<w:commentRangeEnd w:id="0" ?\/>/, '')),
    caughtBy: ['structure'],
  },
];

/** Every check expected to fail: the direct ones, plus those that cannot run once a text check has failed. */
export function expectedStaticFailures(brokenCase: BrokenStaticCase): CheckId[] {
  const failing = new Set<CheckId>(brokenCase.caughtBy);
  if (failing.has('static-new')) {
    failing.add('formatting');
    failing.add('numbering');
  }
  if (failing.has('static-old')) {
    failing.add('format-marks');
    failing.add('numbering');
  }
  return [...failing].sort();
}

export function damageStatic(files: ZipEntries, brokenCase: BrokenStaticCase): void {
  brokenCase.damage(files);
}
