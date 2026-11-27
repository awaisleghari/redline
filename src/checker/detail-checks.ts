/**
 * Checks that go beyond plain text: formatting (check 8), the marking of
 * formatting changes, clause numbers, leftover tracked changes (check 4) and
 * the cover page (check 5). All work from the Checker's own readings.
 */
import { NS } from '../common/namespaces';
import { describeFormat, FORMAT, FORMAT_MASK, STANDARD_COLOURS, staticKindOf, type ChangeColours, type StaticKind } from './formatting';
import type { DocxPackage } from './package';
import type { CheckFailure } from './report';
import { CHAR, FORM_READING, isStructureMark, linkAt, STATIC_STYLE, type Para, type Story } from './stories';
import { attribute, child, descendants, is, type XmlElement } from './xml';

const W = NS.w;
const FORMAT_BITS = FORMAT_MASK;
/** Strikethrough is how deleted and moved-from text is shown, so it is not compared on such text. */
const STRIKES = FORMAT.strike | FORMAT.doubleStrike;

function excerptAt(text: string, at: number): string {
  const start = Math.max(0, at - 30);
  const end = Math.min(text.length, at + 40);
  return `${start > 0 ? '…' : ''}${text.slice(start, end)}${end < text.length ? '…' : ''}`;
}

/**
 * Walk two readings that have identical text, paragraph by paragraph, handing
 * each pair of paragraphs to `visit`. Returns the first failure `visit`
 * reports. If the texts are not identical the readings cannot be lined up, and
 * that is itself reported (the text checks will have failed too).
 */
function alongside(
  expected: Story[],
  actual: Story[],
  expectedName: string,
  visit: (want: Para, got: Para, where: string) => CheckFailure | undefined,
): CheckFailure | undefined {
  const actualByKey = new Map(actual.map((s) => [s.key, s]));
  for (const want of expected) {
    const got = actualByKey.get(want.key);
    const sameShape = got && got.details.length === want.details.length && got.details.every((p, i) => p.text === want.details[i]!.text);
    if (!got || !sameShape) return { message: `This could not be checked because the text does not match the ${expectedName} (see the text checks).`, where: want.label };
    let paragraphNumber = 0;
    for (let i = 0; i < want.details.length; i++) {
      if (isStructureMark(want.details[i]!.text)) continue;
      paragraphNumber++;
      const failure = visit(want.details[i]!, got.details[i]!, `${want.label}, paragraph ${paragraphNumber}`);
      if (failure) return failure;
    }
  }
  if (actual.some((s) => !expected.some((e) => e.key === s.key))) {
    return { message: `This could not be checked because the text does not match the ${expectedName} (see the text checks).` };
  }
  return undefined;
}

/**
 * Check 8: the output's text, as it stands once changes are accepted, has the
 * modified document's bold, italic, underline, capitals and hidden-text
 * settings. In a static redline, inserted text is underlined by design, so its
 * underline is not compared.
 */
export function formattingFailure(modified: Story[], accepted: Story[], isStatic: boolean): CheckFailure | undefined {
  return alongside(modified, accepted, 'modified document', (want, got, where) => {
    for (let i = 0; i < want.text.length; i++) {
      const wantCode = want.meta.charCodeAt(i);
      const gotCode = got.meta.charCodeAt(i);
      const ignore = isStatic && gotCode & CHAR.changed ? FORMAT.underline : 0;
      const a = wantCode & FORMAT_BITS & ~ignore;
      const b = gotCode & FORMAT_BITS & ~ignore;
      if (a !== b) {
        return {
          message: 'Text does not have the formatting it has in the modified document.',
          where,
          expected: `"${excerptAt(want.text, i)}" should be ${describeFormat(a)} at the point shown`,
          found: describeFormat(b),
        };
      }
    }
    return undefined;
  });
}

/**
 * Formatting changes are marked exactly where formatting changed: in the
 * static redline, a kept character is shaded if and only if its bold, italic,
 * underline, capitals or hidden-text setting differs from the original
 * document's, or it links somewhere else than it did in the original (a link
 * added, taken off or re-pointed); deleted and inserted text is never shaded.
 *
 * This is also where the original's links are verified in a static redline. The file shows each kept word
 * once, in the link it has now, so the original's link on a shaded word cannot be read from it. What is
 * verified instead: every word the original links is either in that same link still, or shaded; no word is
 * shaded unless its formatting or its link really differs; and deleted text links where it did.
 */
export function formatMarksFailure(original: Story[], staticRejected: Story[], staticAccepted: Story[]): CheckFailure | undefined {
  const fromOriginal = alongside(original, staticRejected, 'original document', (want, got, where) => {
    for (let i = 0; i < want.text.length; i++) {
      const gotCode = got.meta.charCodeAt(i);
      const shaded = (gotCode & CHAR.shaded) !== 0;
      const isText = (gotCode & CHAR.isText) !== 0;
      const ignore = gotCode & CHAR.changed ? STRIKES : 0;
      const formatDiffers = (want.meta.charCodeAt(i) & FORMAT_BITS & ~ignore) !== (gotCode & FORMAT_BITS & ~ignore);
      const [linkedBefore, linkedNow] = [linkAt(want, i), linkAt(got, i)];
      const linkDiffers = linkedBefore !== linkedNow;
      const differs = formatDiffers || linkDiffers;
      const linkWas = `linked to${linkedBefore || ' nothing (no link)'}`;
      const linkIs = `linked to${linkedNow || ' nothing (no link)'}`;
      if (gotCode & CHAR.changed) {
        if (shaded) return { message: 'Deleted text is marked as a formatting change.', where, found: excerptAt(got.text, i) };
        if (linkDiffers) return { message: 'Deleted text does not link where it did in the original document.', where, expected: linkWas, found: `"${excerptAt(got.text, i)}" is ${linkIs}` };
        if (formatDiffers && isText) {
          return { message: 'Deleted text does not have the formatting it had in the original document.', where, expected: describeFormat(want.meta.charCodeAt(i) & FORMAT_BITS), found: `"${excerptAt(got.text, i)}" is ${describeFormat(gotCode & FORMAT_BITS)}` };
        }
        continue;
      }
      if (linkDiffers && !shaded) {
        return { message: 'Text whose link changed is not marked as a formatting change.', where, expected: `"${excerptAt(got.text, i)}" was ${linkWas} and is now ${linkIs}: it should be shaded`, found: 'not shaded' };
      }
      if (!isText) continue; // markers (tabs, objects) follow the text around them
      if (differs && !shaded) {
        return { message: 'Text whose formatting changed is not marked as a formatting change.', where, expected: `"${excerptAt(got.text, i)}" was ${describeFormat(want.meta.charCodeAt(i) & FORMAT_BITS)} and is now ${describeFormat(gotCode & FORMAT_BITS)}: it should be shaded`, found: 'not shaded' };
      }
      if (!differs && shaded) {
        return { message: 'Text is marked as a formatting change although its formatting did not change.', where, expected: 'not shaded', found: `"${excerptAt(got.text, i)}" is shaded` };
      }
    }
    return undefined;
  });
  if (fromOriginal) return fromOriginal;
  for (const story of staticAccepted) {
    let paragraphNumber = 0;
    for (const paragraph of story.details) {
      if (isStructureMark(paragraph.text)) continue;
      paragraphNumber++;
      for (let i = 0; i < paragraph.text.length; i++) {
        const code = paragraph.meta.charCodeAt(i);
        if (code & CHAR.changed && code & CHAR.shaded) {
          return { message: 'Inserted text is marked as a formatting change.', where: `${story.label}, paragraph ${paragraphNumber}`, found: excerptAt(paragraph.text, i) };
        }
      }
    }
  }
  return undefined;
}

/** Clause numbers: every paragraph carries the automatic number it has in the modified document. */
export function numberingFailure(modified: Story[], accepted: Story[]): CheckFailure | undefined {
  return alongside(modified, accepted, 'modified document', (want, got, where) =>
    want.number === got.number
      ? undefined
      : {
          message: 'A paragraph does not carry the number it has in the modified document.',
          where: `${where} ("${excerptAt(want.text, 0)}")`,
          expected: want.number === '' ? '(no number)' : want.number,
          found: got.number === '' ? '(no number)' : got.number,
        },
  );
}

/** Which paragraphs are numbered in the stored forms of text boxes that are not the form read: the same as in the modified document. */
export function otherFormsFailure(upload: string[], output: string[], side: 'modified' | 'original' = 'modified'): CheckFailure | undefined {
  // On the old side only the readings are compared: a removed paragraph in a static redline gives up its numbering by design.
  const of = (list: string[]): string[] => (side === 'original' ? list.filter((entry) => entry.startsWith(FORM_READING)) : list);
  const [want, got] = [of(upload), of(output)];
  if (want.length === got.length && want.every((form, i) => form === got[i])) return undefined;
  return {
    message: `A text box is stored in a second form (for older programs) whose numbered paragraphs or field instructions differ from the ${side} document.`,
    where: 'text boxes and drawings',
  };
}

/**
 * Old numbers: in a static redline, a removed paragraph that was numbered in
 * the original shows that old number, struck through, in front of it; nothing
 * else carries one.
 */
export function oldNumberFailure(original: Story[], staticRejected: Story[]): CheckFailure | undefined {
  // Old numbers are shown in the main text only (PLAN.md); elsewhere none may appear.
  const inBody = new Set<Para>(original.filter((s) => s.key === 'body').flatMap((s) => s.details));
  return alongside(original, staticRejected, 'original document', (want, got, where) => {
    // A removed clause must not also keep a live automatic number.
    if (got.removed && got.number !== '') {
      return { message: 'A removed clause still carries a live automatic number.', where: `${where} ("${excerptAt(want.text, 0)}")`, expected: '(no number)', found: got.number };
    }
    const expected = got.removed && want.number !== '' && inBody.has(want) ? `${want.number}\t` : '';
    if (got.oldNumber === expected) return undefined;
    return {
      message: expected === '' ? 'An old clause number is shown where there should be none.' : 'A removed clause does not show the number it had in the original document.',
      where: `${where} ("${excerptAt(want.text, 0)}")`,
      expected: expected === '' ? '(no old number)' : expected.trimEnd(),
      found: got.oldNumber === '' ? '(no old number)' : got.oldNumber.trimEnd(),
    };
  });
}

const TRACKED_ELEMENTS = new Set(['ins', 'del', 'moveFrom', 'moveTo', 'delText', 'delInstrText', 'cellIns', 'cellDel', 'cellMerge', 'numberingChange']);

/** Check 4: names of any tracked-change markup left in the package, with the part it is in. */
export function leftoverTrackedChanges(pkg: DocxPackage): string[] {
  const found: string[] = [];
  for (const partName of pkg.partNames) {
    if (!partName.endsWith('.xml') || partName === '[Content_Types].xml') continue;
    const root = pkg.xml(partName);
    if (root.ns !== W) continue;
    const names = new Set<string>();
    for (const node of descendants(root)) {
      if (node.ns !== W) continue;
      if (TRACKED_ELEMENTS.has(node.name) || /Change$/.test(node.name) || /^(move|customXml)\w*Range(Start|End)$/.test(node.name)) names.add(node.name);
    }
    for (const name of names) found.push(`"${partName}" still contains tracked-change markup (${name}).`);
  }
  return found;
}

export interface StaticCounts {
  insertions: number;
  deletions: number;
  moves: number;
  formattingChanges: number;
  /** Stretches of text marked as moved to here (must be present exactly when moves are). */
  moveDestinations: number;
}

/**
 * The Checker's own count of the changes in a static redline, read from the
 * formatting alone. Definition (PLAN.md): reading each part in order, a change
 * is an unbroken stretch of text and paragraph marks of one kind; a move is
 * counted once, at its origin; a formatting change is an unbroken stretch of
 * shaded runs within one paragraph. The cover page itself is not counted.
 */
export function countStaticChanges(pkg: DocxPackage, colours: ChangeColours = STANDARD_COLOURS): StaticCounts {
  const counts: Record<StaticKind, number> = { inserted: 0, deleted: 0, movedFrom: 0, movedTo: 0 };
  let formattingChanges = 0;

  const isCover = (paragraph: XmlElement): boolean => {
    const pPr = child(paragraph, W, 'pPr');
    const style = pPr && child(pPr, W, 'pStyle');
    return !!style && (attribute(style, W, 'val') ?? '').startsWith(STATIC_STYLE.coverPrefix);
  };

  for (const partName of pkg.partNames) {
    if (!partName.endsWith('.xml') || partName === '[Content_Types].xml') continue;
    const root = pkg.xml(partName);
    if (root.ns !== W || !['document', 'hdr', 'ftr', 'footnotes', 'endnotes', 'comments'].includes(root.name)) continue;

    let previous: StaticKind | 'unchanged' | 'boundary' = 'boundary';
    const token = (kind: StaticKind | 'unchanged' | 'boundary'): void => {
      if (kind !== previous && kind !== 'unchanged' && kind !== 'boundary') counts[kind]++;
      previous = kind;
    };
    const walk = (element: XmlElement): void => {
      for (const node of element.children) {
        if (node.ns !== W) {
          walk(node);
          continue;
        }
        if (node.name === 'p') {
          if (isCover(node)) continue;
          let shadedBefore = false;
          const runs = (parent: XmlElement): void => {
            for (const c of parent.children) {
              if (is(c, W, 'r')) {
                const rPr = child(c, W, 'rPr');
                if (!c.children.some((x) => x !== rPr)) continue;
                const shading = rPr && child(rPr, W, 'shd');
                const shaded = !!shading && attribute(shading, W, 'fill')?.toUpperCase() === STATIC_STYLE.shadeFill;
                if (shaded && !shadedBefore) formattingChanges++;
                shadedBefore = shaded;
              } else if (c.ns === W && ['hyperlink', 'sdt', 'sdtContent', 'smartTag', 'fldSimple', 'customXml'].includes(c.name)) runs(c);
            }
          };
          runs(node);
          walk(node);
          const pPr = child(node, W, 'pPr');
          token(staticKindOf(pPr && child(pPr, W, 'rPr'), colours) ?? 'unchanged');
        } else if (node.name === 'r') {
          const rPr = child(node, W, 'rPr');
          const runStyle = rPr && child(rPr, W, 'rStyle');
          // The old number typed in front of a removed clause belongs to that clause's change; it is not counted on its own.
          if (runStyle && attribute(runStyle, W, 'val') === STATIC_STYLE.oldNumber) continue;
          if (node.children.some((x) => x !== rPr)) token(staticKindOf(rPr, colours) ?? 'unchanged');
          walk(node);
        } else if (node.name === 'footnote' || node.name === 'endnote' || node.name === 'comment') {
          token('boundary');
          walk(node);
          token('boundary');
        } else if (node.name !== 'pPr' && node.name !== 'rPr') {
          walk(node);
        }
      }
    };
    walk(root);
  }
  return { insertions: counts.inserted, deletions: counts.deleted, moves: counts.movedFrom, formattingChanges, moveDestinations: counts.movedTo };
}

export interface CoverExpectation {
  originalName: string;
  modifiedName: string;
  comparedAt: string;
  /** Which uploads the Checker itself found tracked changes in. */
  trackedChangesIn?: 'original' | 'modified' | 'both';
  /** True when the Checker itself found that a document already uses blue double underline, so insertions must be shown in the alternative colour. */
  alternativeInsertionColour?: boolean;
  alternativeDeletionColour?: boolean;
}

/** The legend's line for inserted text: plain, or saying why insertions are not in the usual blue. */
export const DELETED_LEGEND = { standard: 'Deleted text', alternative: 'Deleted text (shown in dark orange, because this document itself contains red struck-through text)' } as const;
export const INSERTED_LEGEND = { standard: 'Inserted text', alternative: 'Inserted text (shown in magenta, because this document itself contains blue double-underlined text)' } as const;

/** What the cover page must say when an upload came with tracked changes of its own (and must not say otherwise). */
export const ACCEPTED_NOTE_EXPECTED = {
  original: 'Note: the original contained tracked changes. It was compared as if they had all been accepted.',
  modified: 'Note: the modified document contained tracked changes. It was compared as if they had all been accepted.',
  both: 'Note: both documents contained tracked changes. They were compared as if those had all been accepted.',
} as const;

/** Check 5: the cover page says exactly what it should, with the Checker's own counts. */
export function coverPageFailure(cover: string[], coverLooks: string[], expectation: CoverExpectation, counts: StaticCounts): CheckFailure | undefined {
  // Text shown as moved away must have arrived somewhere, and the other way round.
  if ((counts.moves === 0) !== (counts.moveDestinations === 0)) {
    return {
      message: counts.moves > 0 ? 'Text is marked as moved from somewhere, but nothing is marked as moved to anywhere.' : 'Text is marked as moved to here, but nothing is marked as moved from anywhere.',
      expected: 'every move has an origin and a destination',
      found: `${counts.moves} moved-from, ${counts.moveDestinations} moved-to`,
    };
  }
  const insertedLegend = expectation.alternativeInsertionColour ? INSERTED_LEGEND.alternative : INSERTED_LEGEND.standard;
  const deletedLegend = expectation.alternativeDeletionColour ? DELETED_LEGEND.alternative : DELETED_LEGEND.standard;
  const expected = [
    'Comparison of documents',
    `Original: ${expectation.originalName}`,
    `Modified: ${expectation.modifiedName}`,
    `Compared: ${expectation.comparedAt}`,
    ...(expectation.trackedChangesIn ? [ACCEPTED_NOTE_EXPECTED[expectation.trackedChangesIn]] : []),
    'Legend',
    insertedLegend,
    deletedLegend,
    'Text moved from here',
    'Text moved to here',
    'Text whose formatting changed',
    'Changes',
    `Insertions: ${counts.insertions}`,
    `Deletions: ${counts.deletions}`,
    `Moves: ${counts.moves}`,
    `Formatting changes: ${counts.formattingChanges}`,
    '', // the empty paragraph that ends the cover page's own section
  ];
  if (cover.length === 0) return { message: 'The static redline has no cover page.' };
  for (let i = 0; i < Math.max(expected.length, cover.length); i++) {
    if (expected[i] === cover[i]) continue;
    const isCount = /^(Insertions|Deletions|Moves|Formatting changes): /.test(expected[i] ?? '');
    return {
      message: isCount ? 'A count on the cover page does not match the changes actually in the redline.' : 'The cover page does not say what it should.',
      where: `cover page, line ${i + 1}`,
      expected: expected[i] ?? '(no further lines)',
      found: cover[i] ?? '(line missing)',
    };
  }
  // The legend must show each style as it really is, and no line may be hidden or dressed as a change.
  const looks = expected.map((line) =>
    line === insertedLegend ? 'inserted' : line === deletedLegend ? 'deleted' : line === 'Text moved from here' ? 'movedFrom' : line === 'Text moved to here' ? 'movedTo' : line === 'Text whose formatting changed' ? 'shaded' : line === '' ? '' : 'plain',
  );
  for (let i = 0; i < looks.length; i++) {
    if (looks[i] === coverLooks[i]) continue;
    return { message: 'A line of the cover page is not formatted as it should be (the legend must show each style exactly, and nothing may be hidden).', where: `cover page, line ${i + 1} ("${expected[i]}")`, expected: looks[i] || '(empty)', found: coverLooks[i] || '(empty)' };
  }
  return undefined;
}
