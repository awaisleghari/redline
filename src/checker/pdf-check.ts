/**
 * Check 9: the PDF says what the static redline says, in order, and changed
 * text is in the right colours.
 *
 * The static redline is read exactly as it is printed (every visible
 * character, cover page included). The PDF is read as a stream of characters
 * with their colours. White space is not compared: a PDF records where
 * characters sit on the page, not the spaces between words. Every other
 * character is compared exactly.
 *
 * How the two are matched:
 *   1. The main text must appear in the PDF, complete and in order.
 *   2. What is left in the PDF must then account, in turn, for every text box,
 *      footnote and endnote (complete and in order), and for the headers and
 *      footers (which repeat on each page; their page numbers are not compared).
 *   3. Whatever is still left can only be what a page shows without the
 *      document spelling it out, and only in the place where a page shows it:
 *        - a clause number or bullet, immediately before a list item;
 *        - a footnote number, where the text has a footnote mark;
 *        - page numbers and note numbers, at the edge of a page (never letters).
 *      Anything else, anywhere else, is text the redline does not contain: one
 *      extra digit in the middle of a sentence fails the check.
 */
import { FORMAT, type StaticKind } from './formatting';
import { isBlank, type PdfChar, type PdfReading } from './pdf';
import type { CheckFailure } from './report';
import { CHAR, isStructureMark, kindOfCode, type Story, type StoryReading } from './stories';

/** The colour each kind of change must be printed in (the style spec in PLAN.md). */
const COLOUR: Record<StaticKind, string> = { inserted: '0000FF', deleted: 'FF0000', movedFrom: '008000', movedTo: '008000' };
const KIND_NAME: Record<StaticKind, string> = { inserted: 'inserted', deleted: 'deleted', movedFrom: 'moved-from', movedTo: 'moved-to' };

/** One character the PDF is expected to show. */
interface Expected {
  c: string;
  kind: StaticKind | undefined;
  /** Characters left in this paragraph after this one (0 for its last). */
  toParagraphEnd: number;
  /** Paragraph number within its story, for reports. */
  paragraph: number;
  /** True when a mark that prints something of its own (a footnote number, a symbol, a field) follows this character. */
  markAfter: boolean;
  /** For the first character of a paragraph: the clause number printed before it ('' if none), and whether it is a list item. */
  label?: { text: string; listed: boolean };
}

/** The characters a story prints, without white space, as the PDF would show them. */
function expectedChars(story: Story, withoutDigits = false): Expected[] {
  const out: Expected[] = [];
  let paragraphNumber = 0;
  for (const paragraph of story.details) {
    if (isStructureMark(paragraph.text)) continue;
    paragraphNumber++;
    const start = out.length;
    let markPending = false;
    for (let i = 0; i < paragraph.text.length; i++) {
      const code = paragraph.meta.charCodeAt(i);
      if (code & FORMAT.hidden) continue; // hidden text is not printed
      if (!(code & CHAR.isText)) {
        // A marker is not a character, but it may print something (a footnote number, a symbol).
        if (!isBlank(paragraph.text[i]!)) {
          if (out.length > start) out[out.length - 1]!.markAfter = true;
          else markPending = true;
        }
        continue;
      }
      let text = paragraph.text[i]!.normalize('NFKC');
      if (code & (FORMAT.caps | FORMAT.smallCaps)) text = text.toUpperCase(); // capitals formatting prints capitals
      for (const c of text) {
        if (isBlank(c) || (withoutDigits && /\d/.test(c))) continue;
        out.push({ c, kind: kindOfCode(code), toParagraphEnd: 0, paragraph: paragraphNumber, markAfter: false });
      }
    }
    for (let k = start; k < out.length; k++) out[k]!.toParagraphEnd = out.length - 1 - k;
    if (out.length > start) {
      const label = [...paragraph.number.normalize('NFKC')].filter((c) => !isBlank(c)).join('');
      out[start]!.label = { text: label, listed: paragraph.listed || markPending };
    }
  }
  return out;
}

interface Source {
  chars: PdfChar[];
  used: boolean[];
  /** For characters matched to the main text: which character of the main text. */
  bodyIndex: (number | undefined)[];
}

interface MatchResult {
  /** How many expected characters were found, in order. Equal to the stream's length on success. */
  matched: number;
  /** The first expected character that was found in the wrong colour. */
  wrongColour?: { expected: Expected; index: number; found: string };
  /** Position in the source just after the last match. */
  sourceAt: number;
}

/** How many characters in a row must agree before a match is believed after a gap. */
const RUN = 6;

/**
 * Find `stream` in the unused characters of `source`, complete and in order.
 * After any gap, a match is accepted only where several characters in a row
 * agree, so that a stray character elsewhere on the page (a header, a clause
 * number) is not mistaken for the next character of the text.
 */
function consume(stream: Expected[], source: Source, skipDigits = false, isBody = false): MatchResult {
  const { chars, used } = source;
  const taken: number[] = [];
  let wrongColour: MatchResult['wrongColour'];
  let i = 0;
  let inRun = false;
  let lastSource = 0;
  const agreesAhead = (from: number, at: number): boolean => {
    const need = Math.min(RUN, stream[from]!.toParagraphEnd + 1, stream.length - from);
    let j = at;
    for (let t = 0; t < need; t++) {
      while (j < chars.length && (used[j] || (skipDigits && /\d/.test(chars[j]!.c)))) j++;
      if (j >= chars.length || chars[j]!.c !== stream[from + t]!.c) return false;
      j++;
    }
    return true;
  };
  for (let j = 0; j < chars.length && i < stream.length; j++) {
    if (used[j]) continue;
    if (skipDigits && /\d/.test(chars[j]!.c)) continue;
    if (chars[j]!.c === stream[i]!.c && (inRun || agreesAhead(i, j))) {
      const expected = stream[i]!;
      if (expected.kind && chars[j]!.color !== COLOUR[expected.kind] && !wrongColour) wrongColour = { expected, index: i, found: chars[j]!.color };
      taken.push(j);
      lastSource = j + 1;
      i++;
      inRun = true;
    } else {
      inRun = false;
    }
  }
  if (i === stream.length) {
    taken.forEach((j, index) => {
      used[j] = true;
      if (isBody) source.bodyIndex[j] = index;
    });
    // Page numbers inside a header or footer belong to it.
    if (skipDigits && taken.length > 0) for (let j = taken[0]!; j <= taken[taken.length - 1]!; j++) if (/\d/.test(chars[j]!.c)) used[j] = true;
  }
  return { matched: i, wrongColour, sourceAt: lastSource };
}

const excerpt = (chars: { c: string }[], from: number, length = 50): string =>
  chars
    .slice(Math.max(0, from), from + length)
    .map((x) => x.c)
    .join('');

export interface PdfCheckResult {
  text?: CheckFailure;
  colours?: CheckFailure;
  note: string;
}

/**
 * Compare a PDF with the static redline it was made from. `printed` is the
 * static redline read with `readStories(pkg, 'all', 'static')`.
 */
export function comparePdf(printed: StoryReading, pdf: PdfReading): PdfCheckResult {
  const source: Source = { chars: pdf.chars, used: new Array<boolean>(pdf.chars.length).fill(false), bodyIndex: new Array<number | undefined>(pdf.chars.length).fill(undefined) };
  let bodyStream: Expected[] = [];
  let colours: CheckFailure | undefined;
  const noteColour = (story: Story, result: MatchResult): void => {
    if (colours || !result.wrongColour) return;
    const { expected, found } = result.wrongColour;
    colours = {
      message: `Text that is ${KIND_NAME[expected.kind!]} in the redline is not printed in its colour in the PDF.`,
      where: `${story.label}, paragraph ${expected.paragraph}`,
      expected: `colour ${COLOUR[expected.kind!]}`,
      found: `colour ${found}`,
    };
  };
  const note = `Compared ${pdf.chars.length} characters on ${pdf.pages} page${pdf.pages === 1 ? '' : 's'}.`;
  const missing = (story: Story, stream: Expected[], result: MatchResult): CheckFailure => ({
    message: `The PDF does not contain the ${story.label} of the redline, complete and in order.`,
    where: `${story.label}, paragraph ${stream[result.matched]?.paragraph ?? '?'}`,
    expected: `…${excerpt(stream, result.matched - 15, 65)}…`,
    found: `…${excerpt(pdf.chars, result.sourceAt - 15, 65)}… (white space is not compared)`,
  });

  if (printed.problems.length > 0) return { text: { message: printed.problems[0]! }, note };
  if (pdf.chars.length === 0) return { text: { message: 'The PDF contains no text.' }, note };

  // 1. The main text, complete and in order.
  const body = printed.stories.find((s) => s.key === 'body');
  if (body) {
    const stream = expectedChars(body);
    const result = consume(stream, source, false, true);
    if (result.matched < stream.length) return { text: missing(body, stream, result), note };
    noteColour(body, result);
    bodyStream = stream;
  }

  // 2. Text boxes and notes, each exactly once; then headers and footers, as often as they repeat.
  const repeating = (s: Story): boolean => s.kind === 'header' || s.kind === 'footer' || s.key.startsWith('header/') || s.key.startsWith('footer/');
  for (const story of printed.stories) {
    if (story.key === 'body' || repeating(story)) continue;
    const stream = expectedChars(story);
    if (stream.length === 0) continue;
    const result = consume(stream, source);
    if (result.matched < stream.length) return { text: missing(story, stream, result), note };
    noteColour(story, result);
  }
  for (const story of printed.stories) {
    if (!repeating(story)) continue;
    const stream = expectedChars(story, true);
    if (stream.length === 0) continue;
    let times = 0;
    for (;;) {
      const result = consume(stream, source, true);
      if (result.matched < stream.length) {
        // The ordinary header or footer of a section must be printed at least once.
        if (times === 0 && story.key.endsWith('/default')) return { text: missing(story, stream, result), note };
        break;
      }
      noteColour(story, result);
      if (++times > pdf.pages) break;
    }
  }

  // 3. What is left must be only what a page shows without the document
  // spelling it out, and only where a page shows it.
  const { chars, used, bodyIndex } = source;
  const isLetter = (c: string): boolean => /\p{L}/u.test(c);
  // Note numbers are digits, or (Word's default for endnotes) small roman numerals, or symbols such as *.
  const hasNotes = printed.stories.some((s) => s.kind === 'footnote' || s.kind === 'endnote');
  const isNoteNumber = (t: string): boolean => /^(\d{1,3}|[ivxlc]{1,6}|[^\p{L}\d]{1,3})$/u.test(t);
  const extra = (j: number, why: string): PdfCheckResult => ({
    text: {
      message: `The PDF contains text that is not in the static redline (${why}).`,
      where: `page ${chars[j]!.page}`,
      expected: 'nothing here',
      found: `…${excerpt(chars, j - 25, 25)}[${excerpt(chars, j, 12)}]…  (the bracketed text is the extra; white space is not compared)`,
    },
    colours,
    note,
  });
  // The nearest main-text character before and after each position.
  const before: (number | undefined)[] = [];
  let last: number | undefined;
  for (let j = 0; j < chars.length; j++) {
    before[j] = last;
    if (bodyIndex[j] !== undefined) last = j;
  }
  const after: (number | undefined)[] = new Array<number | undefined>(chars.length).fill(undefined);
  let next: number | undefined;
  for (let j = chars.length - 1; j >= 0; j--) {
    after[j] = next;
    if (bodyIndex[j] !== undefined) next = j;
  }

  for (let j = 0; j < chars.length; j++) {
    if (used[j]) continue;
    // An unbroken stretch of unexplained characters.
    let end = j;
    while (end + 1 < chars.length && !used[end + 1]) end++;
    const gap = chars.slice(j, end + 1);
    const text = gap.map((x) => x.c).join('');
    const p = before[j];
    const n = after[end];
    const samePage = p !== undefined && n !== undefined && chars[p]!.page === chars[n]!.page && gap.every((x) => x.page === chars[p]!.page);
    const prev = p !== undefined ? bodyStream[bodyIndex[p]!] : undefined;
    const following = n !== undefined ? bodyStream[bodyIndex[n]!] : undefined;
    const consecutive = p !== undefined && n !== undefined && bodyIndex[n]! === bodyIndex[p]! + 1;

    if (samePage && consecutive && prev!.toParagraphEnd > 0) {
      // Inside a paragraph, on one page: only a footnote number or symbol, where the text has such a mark.
      if (!prev!.markAfter) return extra(j, 'in the middle of a paragraph');
      if (!isNoteNumber(text)) return extra(j, 'beside a footnote mark or symbol');
    } else if (samePage && consecutive && following!.label) {
      // Between two paragraphs on one page: only the clause number or bullet of a list item (and a note number before it).
      const label = following!.label;
      // A note number may come first, when the paragraph before ends with a footnote mark.
      const allowed = prev!.markAfter ? text.replace(/^(\d{1,3}|[ivxlc]{1,6}(?![a-z])|[^\p{L}\d]{1,3})/u, '') : text;
      if (allowed === '') {
        // nothing but the note number
      } else if (label.text !== '') {
        if (allowed !== label.text) return extra(j, `where the clause number "${label.text}" belongs`);
      } else if (!label.listed || allowed.length > 2 || [...allowed].some(isLetter)) {
        return extra(j, 'between two paragraphs');
      }
    } else {
      // At the edge of a page, or beside a text box, note, header or footer: page and note numbers, a clause
      // number carried to the new page, bullets and rules. Never words.
      const label = following?.label?.text ?? '';
      const rest = label !== '' && text.endsWith(label) ? text.slice(0, text.length - label.length) : text;
      // Digits and symbols; and, in a document with notes, a short roman numeral (an endnote's number).
      const withoutRoman = hasNotes ? rest.replace(/[ivxlc]{1,6}/, '') : rest;
      if (rest.length > 12 || [...withoutRoman].some(isLetter)) return extra(j, 'at the edge of a page');
    }
    j = end;
  }
  return { colours, note };
}
