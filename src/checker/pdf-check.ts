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
 *   1. The main text must appear in the PDF, complete and in order. So must
 *      every text box, footnote, endnote, header and footer.
 *   2. A page prints a few things the document does not spell out. Each is
 *      required or allowed in exactly one place, and must be exactly right:
 *        - a clause number: required immediately before the paragraph it
 *          belongs to, and it must be the number the Checker works out itself;
 *        - a bullet: one or two symbols before a bulleted paragraph;
 *        - a note's number: required where the text has the note's mark, and
 *          again in front of the note itself;
 *        - a page number: digits where a header or footer has a page-number field;
 *        - the dots or line of a tab leader, where the text has a tab;
 *        - a table's heading row, repeated where the table runs onto a new page.
 *   3. Nothing else may be there. One stray character anywhere fails the check.
 */
import { FORMAT, type StaticKind } from './formatting';
import { isBlank, type PdfChar, type PdfReading } from './pdf';
import type { CheckFailure } from './report';
import { CHAR, isStructureMark, kindOfCode, MARK, type Story, type StoryReading } from './stories';

/** The colour each kind of change must be printed in (the style spec in PLAN.md). */
const COLOUR: Record<StaticKind, string> = { inserted: '0000FF', deleted: 'FF0000', movedFrom: '008000', movedTo: '008000' };
const KIND_NAME: Record<StaticKind, string> = { inserted: 'inserted', deleted: 'deleted', movedFrom: 'moved-from', movedTo: 'moved-to' };

/** Something a page prints at a point in the text without the document spelling it out. */
type Token = { t: 'note'; kind: 'footnote' | 'endnote'; n: number } | { t: 'symbol' } | { t: 'tab' } | { t: 'page-number' } | { t: 'label'; text: string } | { t: 'bullet' };

/** One character the PDF is expected to show. */
interface Expected {
  c: string;
  /** Compare without regard to case (text formatted as capitals is printed in capitals). */
  fold: boolean;
  kind: StaticKind | undefined;
  /** Characters left in this paragraph after this one (0 for its last). */
  toParagraphEnd: number;
  /** Paragraph number within its story, for reports. */
  paragraph: number;
  /** What a page prints immediately before this character. */
  before: Token[];
  /** What a page prints immediately after this character (used for the last character of a story). */
  after: Token[];
}

/** Characters as compared: compatibility forms spelled out, and every kind of hyphen as a plain one. */
const canon = (c: string): string => (c === '‐' || c === '‑' ? '-' : c);
const same = (found: string, expected: Expected): boolean => (expected.fold ? found.toLowerCase() === expected.c : found === expected.c);

const ROMAN: [number, string][] = [[1000, 'm'], [900, 'cm'], [500, 'd'], [400, 'cd'], [100, 'c'], [90, 'xc'], [50, 'l'], [40, 'xl'], [10, 'x'], [9, 'ix'], [5, 'v'], [4, 'iv'], [1, 'i']];
function roman(n: number): string {
  let out = '';
  for (const [size, symbol] of ROMAN) for (; n >= size; n -= size) out += symbol;
  return out;
}
const escape = (text: string): string => text.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');

/** The text a token may print, as a regular-expression fragment. */
function pattern(token: Token): string {
  switch (token.t) {
    case 'note':
      // A note is numbered in order: as a number, or (Word's default for endnotes) a roman numeral, or symbols.
      return `(?:${token.n}|${roman(token.n)}|${roman(token.n).toUpperCase()}|[^\\p{L}\\d]{1,3})`;
    case 'symbol':
      return '.';
    case 'tab':
      return '[._\\-–—·…]*';
    case 'page-number':
      return '(?:\\d{1,6}|[ivxlcdm]{1,8}|[IVXLCDM]{1,8})';
    case 'label':
      return escape(token.text);
    case 'bullet':
      return '[^\\p{L}\\d]{0,2}';
  }
}

interface NoteCounters {
  footnote: number;
  endnote: number;
}

/** The characters a story prints, without white space, with what a page adds around them. */
function expectedChars(story: Story, counters: NoteCounters): { stream: Expected[]; trailing: Token[] } {
  const out: Expected[] = [];
  let pending: Token[] = []; // tokens waiting for the next character
  let paragraphNumber = 0;
  for (const paragraph of story.details) {
    if (isStructureMark(paragraph.text)) continue;
    paragraphNumber++;
    // The clause number or bullet printed in front of the paragraph.
    const label = [...paragraph.number.normalize('NFKC')].filter((c) => !isBlank(c)).map(canon).join('');
    if (label !== '') pending.push({ t: 'label', text: label });
    else if (paragraph.listed) pending.push({ t: 'bullet' });

    const start = out.length;
    const text = paragraph.text;
    let field: { code: string; pageNumber: boolean; inResult: boolean } | undefined;
    for (let i = 0; i < text.length; ) {
      const code = paragraph.meta.charCodeAt(i);
      if (!(code & CHAR.isText)) {
        // A marker: not a character of the text, but it may stand for something a page prints.
        const at = (mark: string): boolean => text.startsWith(mark, i);
        if (at(MARK.footnote)) {
          pending.push({ t: 'note', kind: 'footnote', n: ++counters.footnote });
          i += MARK.footnote.length;
        } else if (at(MARK.endnote)) {
          pending.push({ t: 'note', kind: 'endnote', n: ++counters.endnote });
          i += MARK.endnote.length;
        } else if (at('⟦symbol ')) {
          pending.push({ t: 'symbol' });
          i = text.indexOf('⟧', i) + 1 || text.length;
        } else if (at(MARK.fieldBegin)) {
          field = { code: '', pageNumber: false, inResult: false };
          i += MARK.fieldBegin.length;
        } else if (at(MARK.fieldResult)) {
          if (field) {
            field.inResult = true;
            // A page-number field prints the number of the page it lands on, whatever the document had stored.
            field.pageNumber = /^\s*(PAGE|NUMPAGES|SECTIONPAGES)\b/i.test(field.code);
            if (field.pageNumber) pending.push({ t: 'page-number' });
          }
          i += MARK.fieldResult.length;
        } else if (at(MARK.fieldEnd)) {
          field = undefined;
          i += MARK.fieldEnd.length;
        } else {
          if (field && !field.inResult) field.code += text[i];
          else if (text[i] === '\t') pending.push({ t: 'tab' });
          else if (text[i] === '‑' && !(code & FORMAT.hidden)) out.push({ c: '-', fold: false, kind: kindOfCode(code), toParagraphEnd: 0, paragraph: paragraphNumber, before: pending.splice(0), after: [] });
          i++;
        }
        continue;
      }
      i++;
      if (code & FORMAT.hidden) continue; // hidden text is not printed
      if (field?.pageNumber && field.inResult) continue; // the stored page number is not what the page prints
      const fold = (code & (FORMAT.caps | FORMAT.smallCaps)) !== 0;
      for (const raw of text[i - 1]!.normalize('NFKC')) {
        if (isBlank(raw)) continue;
        const c = canon(raw);
        out.push({ c: fold ? c.toLowerCase() : c, fold, kind: kindOfCode(code), toParagraphEnd: 0, paragraph: paragraphNumber, before: pending.splice(0), after: [] });
      }
    }
    for (let k = start; k < out.length; k++) out[k]!.toParagraphEnd = out.length - 1 - k;
  }
  // Tokens after the last character of the story.
  const trailing = pending;
  pending = [];
  if (out.length > 0) out[out.length - 1]!.after = trailing;
  return { stream: out, trailing };
}

/** Who has accounted for a character of the PDF. */
type Owner = undefined | 'body' | 'other' | 'accepted';

interface Source {
  chars: PdfChar[];
  owner: Owner[];
}

interface Match {
  /** For each character of the stream, the position in the PDF it was found at. Shorter than the stream on failure. */
  at: number[];
  wrongColour?: { expected: Expected; found: string };
}

/** How many characters in a row must agree before a match is believed after a gap. */
const RUN = 6;

/**
 * Find `stream` among the characters of the PDF that nothing has accounted
 * for yet, complete and in order. After any gap, a match is accepted only
 * where several characters in a row agree, so that a stray character elsewhere
 * on the page is not mistaken for the next character of the text.
 */
function find(stream: Expected[], source: Source, as: 'body' | 'other'): Match {
  const { chars, owner } = source;
  const at: number[] = [];
  let wrongColour: Match['wrongColour'];
  let inRun = false;
  const agreesAhead = (from: number, position: number): boolean => {
    const need = Math.min(RUN, stream[from]!.toParagraphEnd + 1, stream.length - from);
    let j = position;
    for (let t = 0; t < need; t++) {
      while (j < chars.length && owner[j]) j++;
      if (j >= chars.length || !same(chars[j]!.c, stream[from + t]!)) return false;
      j++;
    }
    return true;
  };
  for (let j = 0; j < chars.length && at.length < stream.length; j++) {
    if (owner[j]) continue;
    const expected = stream[at.length]!;
    if (same(chars[j]!.c, expected) && (inRun || agreesAhead(at.length, j))) {
      if (expected.kind && chars[j]!.color !== COLOUR[expected.kind] && !wrongColour) wrongColour = { expected, found: chars[j]!.color };
      at.push(j);
      inRun = true;
    } else {
      inRun = false;
    }
  }
  if (at.length === stream.length) for (const j of at) owner[j] = as;
  return { at, wrongColour };
}

const excerpt = (chars: { c: string }[], from: number, length = 50): string =>
  chars
    .slice(Math.max(0, from), Math.max(0, from) + length)
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
  const chars = pdf.chars.map((x) => ({ ...x, c: canon(x.c) }));
  const source: Source = { chars, owner: new Array<Owner>(chars.length).fill(undefined) };
  const { owner } = source;
  let colours: CheckFailure | undefined;
  const note = `Compared ${chars.length} characters on ${pdf.pages} page${pdf.pages === 1 ? '' : 's'}.`;
  const noteColour = (story: Story, match: Match): void => {
    if (colours || !match.wrongColour) return;
    const { expected, found } = match.wrongColour;
    colours = {
      message: `Text that is ${KIND_NAME[expected.kind!]} in the redline is not printed in its colour in the PDF.`,
      where: `${story.label}, paragraph ${expected.paragraph}`,
      expected: `colour ${COLOUR[expected.kind!]}`,
      found: `colour ${found}`,
    };
  };
  const missing = (story: Story, stream: Expected[], match: Match): PdfCheckResult => ({
    text: {
      message: `The PDF does not contain the ${story.label} of the redline, complete and in order.`,
      where: `${story.label}, paragraph ${stream[match.at.length]?.paragraph ?? '?'}`,
      expected: `…${excerpt(stream, match.at.length - 15, 65)}…`,
      found: `…${excerpt(chars, (match.at[match.at.length - 1] ?? 0) - 14, 65)}… (white space is not compared)`,
    },
    note,
  });
  const extra = (j: number, why: string, expected = 'nothing here'): PdfCheckResult => ({
    text: {
      message: `The PDF does not match the static redline: ${why}.`,
      where: `page ${chars[Math.min(j, chars.length - 1)]?.page ?? '?'}`,
      expected,
      found: `…${excerpt(chars, j - 25, 25)}[${excerpt(chars, j, 14)}]…  (the bracketed text is where it differs; white space is not compared)`,
    },
    colours,
    note,
  });

  if (printed.problems.length > 0) return { text: { message: printed.problems[0]! }, note };
  if (chars.length === 0) return { text: { message: 'The PDF contains no text.' }, note };

  // ---- 1. Find every story in the PDF, complete and in order. ----
  const counters: NoteCounters = { footnote: 0, endnote: 0 };
  const repeating = (s: Story): boolean => s.key.startsWith('header/') || s.key.startsWith('footer/');
  const isNote = (s: Story): boolean => (s.kind === 'footnote' || s.kind === 'endnote') && /^(footnote|endnote)\/\d+$/.test(s.key);
  interface Found {
    story: Story;
    stream: Expected[];
    match: Match;
  }
  const found: Found[] = [];
  let pageNumbersPrinted = false;

  const body = printed.stories.find((s) => s.key === 'body');
  let bodyFound: Found | undefined;
  if (body) {
    const { stream } = expectedChars(body, counters);
    const match = find(stream, source, 'body');
    if (match.at.length < stream.length) return missing(body, stream, match);
    noteColour(body, match);
    bodyFound = { story: body, stream, match };
  }
  for (const story of printed.stories) {
    if (story.key === 'body' || repeating(story)) continue;
    const { stream } = expectedChars(story, counters);
    if (stream.length === 0) continue;
    const match = find(stream, source, 'other');
    if (match.at.length < stream.length) return missing(story, stream, match);
    noteColour(story, match);
    found.push({ story, stream, match });
  }
  for (const story of printed.stories) {
    if (!repeating(story)) continue;
    const { stream, trailing } = expectedChars(story, counters);
    if (stream.some((x) => [...x.before, ...x.after].some((t) => t.t === 'page-number')) || trailing.some((t) => t.t === 'page-number')) pageNumbersPrinted = true;
    if (stream.length === 0) continue;
    for (let times = 0; times <= pdf.pages; times++) {
      const match = find(stream, source, 'other');
      if (match.at.length < stream.length) {
        // The ordinary header or footer of a section must be printed at least once.
        if (times === 0 && story.key.endsWith('/default')) return missing(story, stream, match);
        break;
      }
      noteColour(story, match);
      found.push({ story, stream, match });
    }
  }

  // ---- 2. Everything else in the PDF must be exactly what a page adds, exactly where it adds it. ----
  /** The unaccounted-for characters strictly between two positions, and whether anything else lies between. */
  const between = (from: number, to: number): { free: number[]; foreign: boolean } => {
    const free: number[] = [];
    let foreign = false;
    for (let j = from + 1; j < to; j++) {
      if (owner[j] === undefined) free.push(j);
      else if (owner[j] !== 'accepted') foreign = true;
    }
    return { free, foreign };
  };
  const textOf = (positions: number[]): string => positions.map((j) => chars[j]!.c).join('');
  const accept = (positions: number[]): void => positions.forEach((j) => void (owner[j] = 'accepted'));
  const fits = (text: string, tokens: Token[], around = ''): boolean => new RegExp(`^${around}${tokens.map(pattern).join('')}${around}$`, 'u').test(text);
  const describe = (tokens: Token[]): string =>
    tokens.length === 0
      ? 'nothing here'
      : tokens.map((t) => (t.t === 'label' ? `the clause number "${t.text}"` : t.t === 'note' ? `the number of ${t.kind} ${t.n}` : t.t === 'bullet' ? 'a bullet' : t.t === 'tab' ? 'a tab leader' : t.t === 'page-number' ? 'a page number' : 'a symbol')).join(', then ');

  // 2a. Inside the notes, text boxes, headers and footers.
  for (const { story, stream, match } of found) {
    for (let i = 1; i < stream.length; i++) {
      const { free, foreign } = between(match.at[i - 1]!, match.at[i]!);
      const tokens = [...stream[i - 1]!.after, ...stream[i]!.before];
      if (foreign) continue; // something else sits between (a page break inside a note); the main text's pass deals with it
      if (!fits(textOf(free), tokens)) return extra(free[0] ?? match.at[i]!, `in the ${story.label}`, describe(tokens));
      accept(free);
    }
    // Immediately before the story: a note's own number (required), and whatever precedes its first character.
    const first = match.at[0]!;
    const lead: number[] = [];
    for (let j = first - 1; j >= 0 && owner[j] === undefined && lead.length < 12; j--) lead.unshift(j);
    const number = isNote(story) ? [{ t: 'note', kind: story.kind as 'footnote' | 'endnote', n: Number(story.key.split('/')[1]) } satisfies Token] : [];
    const tokens: Token[] = [...number, ...stream[0]!.before];
    // Take the longest stretch right before the story that fits (other things may precede it on the page).
    let taken = false;
    for (let from = 0; from <= lead.length; from++) {
      if (fits(textOf(lead.slice(from)), tokens)) {
        accept(lead.slice(from));
        taken = true;
        break;
      }
    }
    if (!taken) return extra(first, `in front of the ${story.label}`, describe(tokens));
    // Immediately after the story: a page number that ends a header or footer.
    const trail: number[] = [];
    const lastToken = stream[stream.length - 1]!.after;
    if (lastToken.length > 0) {
      for (let j = match.at[match.at.length - 1]! + 1; j < chars.length && owner[j] === undefined && trail.length < 8; j++) trail.push(j);
      for (let to = trail.length; to >= 0; to--) {
        if (fits(textOf(trail.slice(0, to)), lastToken)) {
          accept(trail.slice(0, to));
          break;
        }
      }
    }
  }

  // 2b. Along the main text: between each character and the next, and before the first and after the last.
  const repeatRows = printed.repeatRows.map((row) => [...row.normalize('NFKC')].filter((c) => !isBlank(c)).map(canon).join('')).filter((row) => row !== '');
  // At the edge of a page a page number may be printed (only if some header or footer has a page-number field).
  const edge = pageNumbersPrinted ? '(?:\\d{1,6}|[ivxlcdm]{1,8}|[IVXLCDM]{1,8}){0,4}' : '';
  if (bodyFound) {
    const { stream, match } = bodyFound;
    for (let i = 0; i <= stream.length; i++) {
      const from = i === 0 ? -1 : match.at[i - 1]!;
      const to = i === stream.length ? chars.length : match.at[i]!;
      const { free, foreign } = between(from, to);
      const tokens = [...(i > 0 ? stream[i - 1]!.after : []), ...(i < stream.length ? stream[i]!.before : [])];
      const samePage = i > 0 && i < stream.length && chars[from]!.page === chars[to]!.page;
      let text = textOf(free);
      if (samePage && !foreign) {
        if (!fits(text, tokens)) return extra(free[0] ?? to, i < stream.length && stream[i]!.before.length > 0 ? 'in front of a paragraph' : 'in the middle of the text', describe(tokens));
      } else {
        // A page break, the start or end of the document, or a note, text box, header or footer in between.
        for (const row of repeatRows) if (row !== '' && text.includes(row)) text = text.replace(row, '');
        if (!fits(text, tokens, edge)) return extra(free[0] ?? Math.min(to, chars.length - 1), 'at the edge of a page', describe(tokens));
      }
      accept(free);
    }
  }
  // 3. Nothing may be left over.
  const stray = owner.findIndex((o) => o === undefined);
  if (stray >= 0) return extra(stray, 'it contains text that is not in the redline');
  return { colours, note };
}
