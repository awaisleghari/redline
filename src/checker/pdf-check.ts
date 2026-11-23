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
 *   1. Page by page, the header is matched whole at the top of the page and
 *      the footer whole at the bottom, each with the page number that belongs
 *      on that page (and the number of pages the PDF really has).
 *   2. The main text must then appear, complete and in order. So must every
 *      text box, footnote and endnote.
 *   3. A page prints a few things the document does not spell out. Each is
 *      required or allowed in exactly one place, and must be exactly right:
 *        - a clause number: required immediately before the paragraph it
 *          belongs to; it must be the number the Checker works out itself,
 *          in the colour of the paragraph's own mark;
 *        - a bullet: the list's own bullet character, before a bulleted paragraph;
 *        - a note's number: required where the text has the note's mark, and
 *          again in front of the note itself;
 *        - a symbol-font character: exactly the character the document names;
 *        - the dots or line of a tab leader (three or more), where the text has a tab;
 *        - a table's heading row, repeated where a page turns inside that table.
 *   4. Nothing else may be there. One stray character anywhere fails the check.
 */
import { FORMAT, type StaticKind } from './formatting';
import { isBlank, type PdfChar, type PdfReading } from './pdf';
import type { CheckFailure } from './report';
import { CHAR, isStructureMark, kindOfCode, MARK, type Story, type StoryReading } from './stories';

/** The colour each kind of change must be printed in (the style spec in PLAN.md). */
const COLOUR: Record<StaticKind, string> = { inserted: '0000FF', deleted: 'FF0000', movedFrom: '008000', movedTo: '008000' };
const KIND_NAME: Record<StaticKind, string> = { inserted: 'inserted', deleted: 'deleted', movedFrom: 'moved-from', movedTo: 'moved-to' };

/** Something a page prints at a point in the text without the document spelling it out. */
type Token =
  | { t: 'note'; kind: 'footnote' | 'endnote'; n: number }
  | { t: 'symbol'; code: number }
  | { t: 'tab' }
  | { t: 'page-number'; field: 'PAGE' | 'NUMPAGES' | 'SECTIONPAGES' }
  | { t: 'label'; text: string; kind: StaticKind | undefined }
  | { t: 'bullet'; text: string; kind: StaticKind | undefined };

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
  /** Which table the character is in (0 when it is in none). */
  table: number;
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
const escape = (text: string): string => text.replace(/[.*+?^${}()|[\]\\/]/g, '\\$&');
const numerals = (numbers: number[]): string =>
  `(?:${[...new Set(numbers.filter((n) => n >= 1).flatMap((n) => [String(n), roman(n), roman(n).toUpperCase()]))].sort((a, b) => b.length - a.length).join('|') || '(?!)'})`;

/**
 * The characters a symbol-font glyph may be reported as. Symbol fonts keep
 * their glyphs in a private range (F000 to F0FF); a PDF reports the same
 * glyph by that code, by its low byte, or not at all in any other way.
 */
function glyph(code: number): string {
  if (!Number.isInteger(code) || code <= 0x20 || code > 0x10ffff) return '(?!)';
  const forms = new Set([code]);
  if (code >= 0xf000 && code <= 0xf0ff) forms.add(code & 0xff);
  if (code > 0x20 && code <= 0xff) forms.add(0xf000 | code);
  return `(?:${[...forms].map((form) => escape([...String.fromCodePoint(form).normalize('NFKC')].map(canon).join(''))).join('|')})`;
}

/** Where on the PDF a token is being looked for (a page number depends on the page). */
interface Where {
  page: number;
  pages: number;
}

/** The dots or line of a tab leader: nothing, or at least three of the same leader character. */
const LEADER = '(?:\\.{3,}|_{3,}|-{3,}|–{3,}|—{3,}|·{3,})?';

/** The text a token may print, as a regular-expression fragment. */
function pattern(token: Token, where: Where): string {
  switch (token.t) {
    case 'note':
      // A note is numbered in order: as a number, or (Word's default for endnotes) a roman numeral, or the usual symbols.
      return `(?:${token.n}|${roman(token.n)}|${roman(token.n).toUpperCase()}|[*†‡§]{1,4})`;
    case 'symbol':
      return glyph(token.code);
    case 'tab':
      return LEADER;
    case 'page-number':
      // The number of the page it is printed on (the cover page comes before page 1), or the number of pages.
      return token.field === 'PAGE' ? numerals([where.page, where.page - 1]) : token.field === 'NUMPAGES' ? numerals([where.pages]) : numerals([where.pages, where.pages - 1]);
    case 'label':
      return escape(token.text);
    case 'bullet':
      return [...token.text].map((c) => glyph(c.codePointAt(0)!)).join('');
  }
}
/** The most characters a token can print. */
const longest = (token: Token): number => (token.t === 'tab' ? 400 : token.t === 'label' ? [...token.text].length : token.t === 'bullet' ? [...token.text].length * 3 : token.t === 'symbol' ? 3 : 8);

interface NoteCounters {
  footnote: number;
  endnote: number;
}

const spelled = (text: string): string => [...text.normalize('NFKC')].filter((c) => !isBlank(c)).map(canon).join('');

/** The characters a story prints, without white space, with what a page adds around them. */
function expectedChars(story: Story, counters: NoteCounters): { stream: Expected[]; trailing: Token[] } {
  const out: Expected[] = [];
  let pending: Token[] = []; // tokens waiting for the next character
  let paragraphNumber = 0;
  let depth = 0;
  let tables = 0;
  for (const paragraph of story.details) {
    if (paragraph.text === MARK.table && depth++ === 0) tables++;
    if (paragraph.text === MARK.tableEnd) depth = Math.max(0, depth - 1);
    if (isStructureMark(paragraph.text)) continue;
    const table = depth > 0 ? tables : 0;
    paragraphNumber++;
    // The bullet or clause number printed in front of the paragraph, in the colour of the paragraph's own mark.
    const bullet = [...paragraph.bullet].filter((c) => !isBlank(c)).join('');
    const label = spelled(paragraph.number);
    if (bullet !== '') pending.push({ t: 'bullet', text: bullet, kind: paragraph.markKind });
    else if (label !== '') pending.push({ t: 'label', text: label, kind: paragraph.markKind });

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
          const close = text.indexOf('⟧', i);
          const inside = text.slice(i, close < 0 ? text.length : close).trim();
          pending.push({ t: 'symbol', code: parseInt(inside.slice(inside.lastIndexOf(' ') + 1), 16) });
          i = close + 1 || text.length;
        } else if (at(MARK.fieldBegin)) {
          field = { code: '', pageNumber: false, inResult: false };
          i += MARK.fieldBegin.length;
        } else if (at(MARK.fieldResult)) {
          if (field) {
            field.inResult = true;
            // A page-number field prints the number of the page it lands on, whatever the document had stored.
            const name = /^\s*(PAGE|NUMPAGES|SECTIONPAGES)\b/i.exec(field.code)?.[1]?.toUpperCase() as 'PAGE' | 'NUMPAGES' | 'SECTIONPAGES' | undefined;
            field.pageNumber = name !== undefined;
            if (name) pending.push({ t: 'page-number', field: name });
          }
          i += MARK.fieldResult.length;
        } else if (at(MARK.fieldEnd)) {
          field = undefined;
          i += MARK.fieldEnd.length;
        } else {
          if (field && !field.inResult) field.code += text[i];
          else if (text[i] === '\t') pending.push({ t: 'tab' });
          else if (text[i] === '‑' && !(code & FORMAT.hidden)) out.push({ c: '-', fold: false, kind: kindOfCode(code), toParagraphEnd: 0, paragraph: paragraphNumber, table, before: pending.splice(0), after: [] });
          i++;
        }
        continue;
      }
      i++;
      if (code & FORMAT.hidden) continue; // hidden text is not printed
      if (field?.pageNumber && field.inResult) continue; // the stored page number is not what the page prints
      const fold = (code & (FORMAT.caps | FORMAT.smallCaps)) !== 0;
      // A tab typed in the text prints its leader (dots, a line) if the paragraph's tab stop has one.
      if (text[i - 1] === '\t') pending.push({ t: 'tab' });
      for (const raw of text[i - 1]!.normalize('NFKC')) {
        if (isBlank(raw)) continue;
        const c = canon(raw);
        out.push({ c: fold ? c.toLowerCase() : c, fold, kind: kindOfCode(code), toParagraphEnd: 0, paragraph: paragraphNumber, table, before: pending.splice(0), after: [] });
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
/** How far ahead the two readings of an ambiguous character are compared. */
const SCAN = 20000;

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
  // For each character: how many characters follow on from it with nothing printed in between (itself included).
  const runway = new Array<number>(stream.length).fill(1);
  for (let i = stream.length - 2; i >= 0; i--) if (stream[i]!.toParagraphEnd > 0 && stream[i + 1]!.before.length === 0) runway[i] = runway[i + 1]! + 1;
  const agreesAhead = (from: number, position: number): boolean => {
    const need = Math.min(RUN, runway[from]!);
    let j = position;
    for (let t = 0; t < need; t++) {
      while (j < chars.length && owner[j]) j++;
      if (j >= chars.length || !same(chars[j]!.c, stream[from + t]!)) return false;
      j++;
    }
    return true;
  };
  const nearest = (from: number, start: number, end: number): number => {
    for (let j = start; j < end; j++) if (!owner[j] && agreesAhead(from, j)) return j;
    return -1;
  };
  for (let j = 0; j < chars.length && at.length < stream.length; j++) {
    if (owner[j]) continue;
    const k = at.length;
    const expected = stream[k]!;
    if (same(chars[j]!.c, expected) && (inRun || agreesAhead(k, j))) {
      if (inRun && runway[k]! >= RUN) {
        // The run breaks right here: the text carries on with the next character, but the PDF does not.
        // Either this character is the text's and something else follows it (the turn of a page), or this
        // character is something else's (the first letter of a note, say) and the text resumes further on.
        // It is the second only if the text, from this character, is found further on before its continuation is.
        let next = j + 1;
        while (next < chars.length && owner[next]) next++;
        if (next >= chars.length || !same(chars[next]!.c, stream[k + 1]!)) {
          const end = Math.min(chars.length, j + 1 + SCAN);
          const resumes = nearest(k, j + 1, end);
          if (resumes !== -1 && nearest(k + 1, j + 1, resumes) === -1) {
            inRun = false;
            continue;
          }
        }
      }
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

/** A header or footer as one page prints it: its characters, and what the page adds between them. */
type Piece = { lit: Expected } | { tok: Token };

const regexCache = new Map<string, RegExp>();
const whole = (fragment: string): RegExp => {
  let regex = regexCache.get(fragment);
  if (!regex) regexCache.set(fragment, (regex = new RegExp(`^(?:${fragment})$`, 'u')));
  return regex;
};

/**
 * Match a header or footer at one exact place: from `start`, character after
 * character with nothing skipped, ending exactly at `mustEnd` if given.
 * Returns the positions of its characters and the end, or undefined.
 */
function matchPieces(pieces: Piece[], source: Source, start: number, limit: number, mustEnd: number | undefined, where: Where): { lits: number[]; end: number } | undefined {
  const { chars, owner } = source;
  const lits: number[] = [];
  const attempt = (k: number, position: number): number | undefined => {
    const mark = lits.length;
    for (; k < pieces.length; k++) {
      const piece = pieces[k]!;
      if ('lit' in piece) {
        if (position >= limit || owner[position] || !same(chars[position]!.c, piece.lit)) break;
        lits.push(position++);
        continue;
      }
      // A token: try each length it can take here, longest first.
      const regex = whole(pattern(piece.tok, where));
      const most = Math.min(longest(piece.tok), limit - position);
      const texts: string[] = [''];
      for (let n = 1; n <= most && !owner[position + n - 1]; n++) texts.push(texts[n - 1]! + chars[position + n - 1]!.c);
      for (let n = texts.length - 1; n >= 0; n--) {
        if (!regex.test(texts[n]!)) continue;
        const end = attempt(k + 1, position + n);
        if (end !== undefined) return end;
      }
      break;
    }
    if (k === pieces.length && (mustEnd === undefined || position === mustEnd)) return position;
    lits.length = mark;
    return undefined;
  };
  const end = attempt(0, start);
  return end === undefined ? undefined : { lits, end };
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

const CHANGE_COLOURS = new Set(Object.values(COLOUR));

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
  const wrongColour = (story: Story, expected: Expected, found: string): void => {
    colours ??= {
      message: `Text that is ${KIND_NAME[expected.kind!]} in the redline is not printed in its colour in the PDF.`,
      where: `${story.label}, paragraph ${expected.paragraph}`,
      expected: `colour ${COLOUR[expected.kind!]}`,
      found: `colour ${found}`,
    };
  };
  const noteColour = (story: Story, match: Match): void => {
    if (match.wrongColour) wrongColour(story, match.wrongColour.expected, match.wrongColour.found);
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

  const counters: NoteCounters = { footnote: 0, endnote: 0 };
  const repeating = (s: Story): boolean => s.key.startsWith('header/') || s.key.startsWith('footer/');
  const isNote = (s: Story): boolean => (s.kind === 'footnote' || s.kind === 'endnote') && /^(footnote|endnote)\/\d+$/.test(s.key);
  const accept = (positions: number[]): void => positions.forEach((j) => void (owner[j] = 'accepted'));

  // The body is read first so that its notes are numbered first, as on the page.
  const body = printed.stories.find((s) => s.key === 'body');
  const bodyStream = body ? expectedChars(body, counters).stream : [];
  const others = printed.stories.filter((s) => s.key !== 'body' && !repeating(s)).map((story) => ({ story, stream: expectedChars(story, counters).stream }));

  // ---- 1. Headers and footers, page by page. ----
  // A page prints its header first and its footer last. Each is matched whole, at exactly that place, with the
  // page number that belongs on that page. (Matched before the main text, so that nothing at the edge of a page
  // is ever left to be explained away afterwards.)
  interface Repeating {
    story: Story;
    pieces: Piece[];
    literals: number;
    /** The same text and fields as another header or footer (a document often stores one header several times). */
    signature: string;
    span: number;
    printedOn: Set<number>;
  }
  const repeats: Repeating[] = [];
  for (const story of printed.stories) {
    if (!repeating(story)) continue;
    const { stream, trailing } = expectedChars(story, counters);
    const pieces: Piece[] = [];
    for (const expected of stream) pieces.push(...expected.before.map((tok) => ({ tok })), { lit: expected });
    pieces.push(...trailing.map((tok) => ({ tok })));
    const literals = stream.length;
    // Something that may print nothing at all (an empty header, a lone tab) has nothing to be matched.
    if (literals === 0 && !pieces.some((piece) => 'tok' in piece && piece.tok.t === 'page-number')) continue;
    const signature = pieces.map((piece) => ('lit' in piece ? piece.lit.c : `\u0000${piece.tok.t}${piece.tok.t === 'page-number' ? piece.tok.field : ''}\u0000`)).join('');
    const span = pieces.reduce((sum, piece) => sum + ('lit' in piece ? 1 : longest(piece.tok)), 0);
    repeats.push({ story, pieces, literals, signature, span, printedOn: new Set() });
  }
  repeats.sort((a, b) => b.pieces.length - a.pieces.length); // the fullest first, so that a bare page number does not take a digit of "Page 1 of 3"
  const pageStart: number[] = [];
  const pageEnd: number[] = [];
  chars.forEach((x, j) => {
    pageStart[x.page] ??= j;
    pageEnd[x.page] = j + 1;
  });
  const claim = (repeat: Repeating, page: number, found: { lits: number[]; end: number }, start: number): void => {
    for (let j = start; j < found.end; j++) owner[j] = 'accepted';
    let n = 0;
    for (const piece of repeat.pieces) {
      if (!('lit' in piece)) continue;
      const j = found.lits[n++]!;
      owner[j] = 'other';
      if (piece.lit.kind && chars[j]!.color !== COLOUR[piece.lit.kind]) wrongColour(repeat.story, piece.lit, chars[j]!.color);
    }
    repeat.printedOn.add(page);
  };
  // The cover page is a section of its own, with no header or footer.
  const firstPage = printed.cover.length > 0 ? 2 : 1;
  for (let page = firstPage; page <= pdf.pages; page++) {
    if (pageStart[page] === undefined) continue;
    const where: Where = { page, pages: pdf.pages };
    const limit = pageEnd[page]!;
    for (let progress = true; progress; ) {
      progress = false;
      for (const repeat of repeats) {
        if (repeat.printedOn.has(page)) continue;
        let first = pageStart[page]!;
        while (first < limit && owner[first]) first++;
        let last = limit;
        while (last > first && owner[last - 1]) last--;
        if (first >= last) break;
        // A header stands at the top of the page (after anything already matched there), a footer at the bottom.
        let start = first;
        let found: ReturnType<typeof matchPieces>;
        if (repeat.story.key.startsWith('header/')) {
          found = matchPieces(repeat.pieces, source, start, last, undefined, where);
        } else {
          for (start = Math.max(first, last - repeat.span); start < last; start++) {
            found = matchPieces(repeat.pieces, source, start, last, last, where);
            if (found) break;
          }
        }
        if (!found) continue;
        claim(repeat, page, found, start);
        progress = true;
      }
    }
  }

  // ---- 2. The main text, then every note and text box: complete and in order. ----
  interface Found {
    story: Story;
    stream: Expected[];
    match: Match;
  }
  const found: Found[] = [];
  let bodyFound: Found | undefined;
  if (body) {
    const match = find(bodyStream, source, 'body');
    if (match.at.length < bodyStream.length) return missing(body, bodyStream, match);
    noteColour(body, match);
    bodyFound = { story: body, stream: bodyStream, match };
  }
  for (const { story, stream } of others) {
    if (stream.length === 0) continue;
    const match = find(stream, source, 'other');
    if (match.at.length < stream.length) return missing(story, stream, match);
    noteColour(story, match);
    found.push({ story, stream, match });
  }

  // A header or footer that a page prints somewhere other than first or last (beside a text box, say): accepted
  // where it stands whole and unbroken, if it has enough text of its own to be told from anything else.
  for (let page = firstPage; page <= pdf.pages; page++) {
    if (pageStart[page] === undefined) continue;
    const where: Where = { page, pages: pdf.pages };
    for (const repeat of repeats) {
      if (repeat.printedOn.has(page) || repeat.literals < 4) continue;
      for (let start = pageStart[page]!; start < pageEnd[page]!; start++) {
        if (owner[start]) continue;
        const match = matchPieces(repeat.pieces, source, start, pageEnd[page]!, undefined, where);
        if (!match) continue;
        claim(repeat, page, match, start);
        break;
      }
    }
  }
  // The ordinary header or footer of a section must be printed at least once.
  for (const repeat of repeats) {
    if (!repeat.story.key.endsWith('/default') || repeat.literals === 0) continue;
    if (repeats.some((other) => other.signature === repeat.signature && other.printedOn.size > 0)) continue;
    const stream = repeat.pieces.flatMap((piece) => ('lit' in piece ? [piece.lit] : []));
    return missing(repeat.story, stream, { at: [] });
  }

  // On a page, the notes are printed below the main text. Main text that comes after a note on the same
  // page has been printed in the wrong place (down among the notes).
  if (bodyFound) {
    const firstNoteOnPage = new Map<number, number>();
    for (const { story, match } of found) {
      if (!isNote(story)) continue;
      for (const j of match.at) {
        const page = chars[j]!.page;
        if (!firstNoteOnPage.has(page) || j < firstNoteOnPage.get(page)!) firstNoteOnPage.set(page, j);
      }
    }
    for (const j of bodyFound.match.at) {
      const note = firstNoteOnPage.get(chars[j]!.page);
      if (note !== undefined && j > note) return extra(j, 'main text is printed among the footnotes', 'the main text above the notes');
    }
  }

  // ---- 3. Everything else in the PDF must be exactly what a page adds, exactly where it adds it. ----
  /** The unaccounted-for characters strictly between two positions. */
  const between = (from: number, to: number): number[] => {
    const free: number[] = [];
    for (let j = from + 1; j < to; j++) if (owner[j] === undefined) free.push(j);
    return free;
  };
  const textOf = (positions: number[]): string => positions.map((j) => chars[j]!.c).join('');
  const whereOf = (j: number): Where => ({ page: chars[Math.max(0, Math.min(j, chars.length - 1))]!.page, pages: pdf.pages });
  const fits = (text: string, tokens: Token[], where: Where, middle = ''): boolean => whole(tokens.map((token) => pattern(token, where)).join('') + middle).test(text);
  const describe = (tokens: Token[]): string =>
    tokens.length === 0
      ? 'nothing here'
      : tokens.map((t) => (t.t === 'label' ? `the clause number "${t.text}"` : t.t === 'note' ? `the number of ${t.kind} ${t.n}` : t.t === 'bullet' ? `the bullet "${t.text}"` : t.t === 'tab' ? 'nothing, or a tab leader of at least three dots or dashes' : t.t === 'page-number' ? 'the page number' : 'the symbol the document has here')).join(', then ');
  /** A clause number or bullet is printed like its paragraph's own mark: in the change's colour if the paragraph was inserted or moved, and otherwise not in any change colour. */
  const checkLabelColour = (story: Story, free: number[], tokens: Token[], next: Expected | undefined): void => {
    const last = tokens[tokens.length - 1];
    if (!last || (last.t !== 'label' && last.t !== 'bullet') || !next) return;
    const length = last.t === 'label' ? [...last.text].length : 1;
    for (const j of free.slice(-length)) {
      const colour = chars[j]!.color;
      if (last.kind ? colour === COLOUR[last.kind] : !CHANGE_COLOURS.has(colour)) continue;
      colours ??= {
        message: last.kind ? `The number or bullet of a paragraph that is ${KIND_NAME[last.kind]} in the redline is not printed in its colour in the PDF.` : 'The number or bullet of an unchanged paragraph is printed in the colour of a change in the PDF.',
        where: `${story.label}, paragraph ${next.paragraph}`,
        expected: last.kind ? `colour ${COLOUR[last.kind]}` : 'not the colour of a change',
        found: `colour ${colour}`,
      };
    }
  };

  // 3a. Inside the notes and text boxes.
  for (const { story, stream, match } of found) {
    for (let i = 1; i < stream.length; i++) {
      const free = between(match.at[i - 1]!, match.at[i]!);
      const tokens = [...stream[i - 1]!.after, ...stream[i]!.before];
      if (!fits(textOf(free), tokens, whereOf(match.at[i]!))) return extra(free[0] ?? match.at[i]!, `in the ${story.label}`, describe(tokens));
      checkLabelColour(story, free, tokens, stream[i]);
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
      if (fits(textOf(lead.slice(from)), tokens, whereOf(first))) {
        checkLabelColour(story, lead.slice(from), tokens, stream[0]);
        accept(lead.slice(from));
        taken = true;
        break;
      }
    }
    if (!taken) return extra(first, `in front of the ${story.label}`, describe(tokens));
    // Immediately after the story: what follows its last character.
    const trail: number[] = [];
    const lastToken = stream[stream.length - 1]!.after;
    if (lastToken.length > 0) {
      const end = match.at[match.at.length - 1]!;
      for (let j = end + 1; j < chars.length && owner[j] === undefined && trail.length < 8; j++) trail.push(j);
      for (let to = trail.length; to >= 0; to--) {
        if (fits(textOf(trail.slice(0, to)), lastToken, whereOf(end))) {
          accept(trail.slice(0, to));
          break;
        }
      }
    }
  }

  // 3b. Along the main text: between each character and the next, and before the first and after the last.
  // Where a table runs onto a new page, its heading rows are printed again at the top of that page.
  const rows = printed.repeatRows.map(spelled).filter((row) => row !== '');
  const repeatedRows = rows.map((row) => `(?:${escape(row)})?`).join('');
  if (bodyFound) {
    const { story, stream, match } = bodyFound;
    for (let i = 0; i <= stream.length; i++) {
      const from = i === 0 ? -1 : match.at[i - 1]!;
      const to = i === stream.length ? chars.length : match.at[i]!;
      const free = between(from, to);
      const after = i > 0 ? stream[i - 1]!.after : [];
      const before = i < stream.length ? stream[i]!.before : [];
      const where = whereOf(i < stream.length ? to : from);
      const turnOfPage = i > 0 && i < stream.length && chars[from]!.page !== chars[to]!.page;
      const text = textOf(free);
      // A heading row is repeated only where the page turns inside its table.
      const insideTable = turnOfPage && stream[i]!.table !== 0 && stream[i - 1]!.table === stream[i]!.table;
      const ok = turnOfPage ? whole(after.map((t) => pattern(t, whereOf(from))).join('') + (insideTable ? repeatedRows : '') + before.map((t) => pattern(t, where)).join('')).test(text) : fits(text, [...after, ...before], where);
      if (!ok) {
        const place = turnOfPage ? 'at the turn of a page' : i === 0 ? 'before the text begins' : i === stream.length ? 'after the text ends' : before.length > 0 ? 'in front of a paragraph' : 'in the middle of the text';
        return extra(free[0] ?? Math.min(to, chars.length - 1), place, describe([...after, ...before]));
      }
      checkLabelColour(story, free, [...after, ...before], stream[i]);
      accept(free);
    }
  }
  // 4. Nothing may be left over.
  const stray = owner.findIndex((o) => o === undefined);
  if (stray >= 0) return extra(stray, 'it contains text that is not in the redline');
  return { colours, note };
}
