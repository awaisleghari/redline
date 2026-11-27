/**
 * Layout notes in a static redline (owner's decision, 2026-11-27).
 *
 * Where a passage's wording is unchanged and only its container changed (a list taken out of a table), the
 * static redline shows the passage once, as it now stands, under a short shaded note. Read with the inserted
 * text taken out, such a passage is laid out as in the modified document, not as in the original: the
 * original's table (or its separate paragraphs) cannot be read back from the file.
 *
 * What the Checker verifies instead, and it is all a note can excuse:
 *   - the note says exactly how much it covers, and the Checker finds that passage itself;
 *   - nothing inside the passage is marked as a change;
 *   - at that very place the original has one table (or that many paragraphs), and its paragraphs that hold any
 *     words are, one for one and in the same order, the passage's paragraphs that hold any words;
 *   - no paragraph of the original's passage carried a clause number (a number the reader of the original saw
 *     would otherwise vanish unremarked).
 * Not verified inside a noted passage, and declared: how the original's table was divided into rows and cells,
 * its empty paragraphs, and bullets. Formatting marks and links are verified there as everywhere else, on the
 * paragraphs paired here.
 *
 * The passage as the modified document has it is verified in full by the new-side checks; nothing here touches them.
 */
import type { CheckFailure } from './report';
import { CHAR, isStructureMark, MARK, type Para, type Story } from './stories';

export interface Aligned {
  original: Story[];
  output: Story[];
}

/** The entry that closes the table opening at `start`, or -1. */
function tableEnd(details: Para[], start: number): number {
  let depth = 0;
  for (let i = start; i < details.length; i++) {
    if (details[i]!.meta !== '') continue;
    if (details[i]!.text === MARK.table) depth++;
    else if (details[i]!.text === MARK.tableEnd && --depth === 0) return i;
  }
  return -1;
}

const isParagraph = (entry: Para | undefined): entry is Para => !!entry && !isStructureMark(entry.text);
const marked = (entry: Para): boolean => entry.removed || entry.markKind !== undefined || [...entry.meta].some((c) => c.charCodeAt(0) & CHAR.changed);
const excerpt = (text: string): string => (text.length > 60 ? `${text.slice(0, 60)}…` : text);

/**
 * Line the original up with a static redline read with its inserted text taken out: each noted passage is
 * reduced, in both, to its paragraphs that hold words, once the claims of its note are verified. The two can
 * then be compared entry for entry like any other reading. Returns the failure if a note's claim does not hold.
 */
export function alignLayoutNotes(original: Story[], output: Story[]): Aligned | CheckFailure {
  if (!output.some((story) => story.layoutNotes)) return { original, output };
  const aligned: Aligned = { original: [...original], output: [...output] };
  for (const [index, story] of output.entries()) {
    if (!story.layoutNotes) continue;
    const where = story.label;
    const originalIndex = original.findIndex((candidate) => candidate.key === story.key);
    if (originalIndex < 0) return { message: 'A layout note stands in a part of the document that the original does not have.', where };
    const [now, before] = [story.details, original[originalIndex]!.details];
    const [outNow, outBefore]: [Para[], Para[]] = [[], []];
    let [n, b] = [0, 0];
    for (const note of story.layoutNotes) {
      // Up to the note, the two are compared entry for entry.
      const plain = note.at - n;
      if (plain < 0) return { message: 'Two layout notes cover the same passage.', where };
      outNow.push(...now.slice(n, note.at));
      outBefore.push(...before.slice(b, b + plain));
      [n, b] = [note.at, b + plain];
      // The passage as it stands now, and the passage of the original at the same place.
      const nowEnd = note.was === 'table' ? n + note.count : tableEnd(now, n) + 1;
      const beforeEnd = note.was === 'table' ? tableEnd(before, b) + 1 : b + note.count;
      const [passage, was] = [now.slice(n, nowEnd), before.slice(b, beforeEnd)];
      const [tableSide, textSide] = note.was === 'table' ? [was, passage] : [passage, was];
      const tableName = note.was === 'table' ? 'the original' : 'the redline';
      if (tableSide.length === 0 || tableSide[0]!.text !== MARK.table || tableSide[0]!.meta !== '') {
        return { message: `A layout note says a passage ${note.was === 'table' ? 'was previously in a table' : 'is now a table'}, but ${tableName} has no table at that place.`, where };
      }
      if (textSide.length !== note.count || !textSide.every(isParagraph)) {
        return { message: `A layout note covers ${note.count} paragraph(s), but ${note.was === 'table' ? 'the redline' : 'the original'} does not have that many plain paragraphs at that place.`, where };
      }
      const changed = passage.find((entry) => isParagraph(entry) && marked(entry));
      if (changed) return { message: 'A passage under a layout note (which says only the layout changed) contains text marked as a change.', where, found: excerpt(changed.text) };
      const words = (entries: Para[]): Para[] => entries.filter((entry) => isParagraph(entry) && entry.text !== '');
      const [wordsNow, wordsBefore] = [words(passage), words(was)];
      if (wordsNow.length === 0) return { message: 'A layout note stands over a passage that holds no words.', where };
      for (let i = 0; i < Math.max(wordsNow.length, wordsBefore.length); i++) {
        if (wordsNow[i]?.text === wordsBefore[i]?.text) continue;
        return {
          message: 'A passage under a layout note (which says only the layout changed) does not have the wording of the original document.',
          where,
          expected: wordsBefore[i] ? excerpt(wordsBefore[i]!.text) : '(no further paragraphs)',
          found: wordsNow[i] ? excerpt(wordsNow[i]!.text) : '(no further paragraphs)',
        };
      }
      // (Every paragraph of the original's passage, the empty ones too: an empty numbered paragraph prints its number. Review finding.)
      const numbered = was.find((entry) => isParagraph(entry) && entry.number !== '');
      if (numbered) return { message: 'A passage under a layout note had a clause number in the original document, which the redline does not show.', where, expected: numbered.number, found: excerpt(numbered.text) };
      outNow.push(...wordsNow);
      outBefore.push(...wordsBefore);
      [n, b] = [nowEnd, beforeEnd];
    }
    outNow.push(...now.slice(n));
    outBefore.push(...before.slice(b));
    aligned.output[index] = { ...story, details: outNow, paragraphs: outNow.map((entry) => entry.text) };
    aligned.original[originalIndex] = { ...original[originalIndex]!, details: outBefore, paragraphs: outBefore.map((entry) => entry.text) };
  }
  return aligned;
}
