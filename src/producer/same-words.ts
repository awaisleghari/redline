/**
 * Producer: words nobody changed are not shown struck out and typed again in the static redline
 * (owner's decisions, 2026-11-27).
 *
 * Word can record some changes only as "removed and added", although every word is the same:
 *   - a link added to, taken off, or re-pointed on unchanged words;
 *   - (see layout notes, below) unchanged wording whose container changed.
 * The tracked-changes Word file keeps that form, the only one Word can accept or reject. The static redline
 * and the PDF show the words once, as they now stand. A changed link is then shown by the formatting-change
 * shading (applied where the link differs from the original's: markFormattingChanges in static-redline.ts).
 *
 * These steps run on the tracked-changes markup, before it is converted to fixed formatting.
 */
import { NS } from '../common/namespaces';
import { isWord, kids, linkKeyOf, shownText, type LinkResolver } from './original-reader';

const W = NS.w;
type El = Element;

/** A wrapper holding nothing but runs of plain text (their properties and their characters). */
function plainRuns(wrapper: El): El[] | undefined {
  const runs = kids(wrapper);
  if (runs.length === 0) return undefined;
  for (const run of runs) {
    if (!isWord(run, 'r')) return undefined;
    if (!kids(run).every((c) => isWord(c, 'rPr') || isWord(c, 't') || isWord(c, 'delText'))) return undefined;
  }
  return runs;
}

interface Piece {
  side: 'added' | 'removed';
  /** The element standing in the paragraph: the wrapper itself, or the link around it. */
  element: El;
  /** The tracked wrappers (w:ins or w:del) it consists of. */
  wrappers: El[];
  link: string;
  text: string;
}

/** A child of a paragraph that is one added or removed stretch of plain text, in a link or not. */
function pieceOf(child: El, resolve: LinkResolver): Piece | undefined {
  const sideOf = (wrapper: El): Piece['side'] | undefined => (isWord(wrapper, 'ins') ? 'added' : isWord(wrapper, 'del') ? 'removed' : undefined);
  const textOf = (wrapper: El): string | undefined => {
    const runs = plainRuns(wrapper);
    return runs?.map((run) => kids(run).filter((c) => !isWord(c, 'rPr')).map(shownText).join('')).join('');
  };
  if (isWord(child, 'hyperlink')) {
    const wrappers = kids(child);
    const side = wrappers[0] && sideOf(wrappers[0]);
    if (!side || !wrappers.every((wrapper) => sideOf(wrapper) === side)) return undefined;
    const texts = wrappers.map(textOf);
    if (texts.some((text) => text === undefined)) return undefined;
    return { side, element: child, wrappers, link: linkKeyOf(child, resolve), text: texts.join('') };
  }
  const side = sideOf(child);
  const text = side && textOf(child);
  return side && text !== undefined ? { side, element: child, wrappers: [child], link: '', text } : undefined;
}

/**
 * Within a paragraph: a stretch removed and a stretch added, side by side, that read exactly the same and differ
 * in where they link (found on the owner's documents: a word in a page header that became a link). The removed
 * copy goes, and the added copy stays as ordinary, unchanged text in its link.
 */
export function mergeRelinkedWords(root: El, resolve: LinkResolver): void {
  const paragraphs = root.getElementsByTagNameNS(W, 'p');
  for (let i = 0; i < paragraphs.length; i++) {
    const children = kids(paragraphs.item(i) as El);
    let at = 0;
    while (at < children.length) {
      const group: Piece[] = [];
      for (let piece = pieceOf(children[at]!, resolve); piece; piece = at < children.length ? pieceOf(children[at]!, resolve) : undefined) {
        group.push(piece);
        at++;
      }
      if (group.length === 0) at++;
      else mergeGroup(group);
    }
  }
}

function mergeGroup(group: Piece[]): void {
  const side = (which: Piece['side']): { text: string; links: string[] } => {
    const pieces = group.filter((piece) => piece.side === which);
    return { text: pieces.map((piece) => piece.text).join(''), links: pieces.flatMap((piece) => Array.from({ length: piece.text.length }, () => piece.link)) };
  };
  const [removed, added] = [side('removed'), side('added')];
  if (removed.text === '' || removed.text !== added.text) return;
  // Only where a link really differs. (The same words removed and added for another reason are left as they are.)
  if (!removed.links.some((link, k) => link !== added.links[k])) return;
  for (const piece of group) {
    if (piece.side === 'removed') {
      piece.element.parentNode!.removeChild(piece.element);
      continue;
    }
    for (const wrapper of piece.wrappers) {
      const parent = wrapper.parentNode!;
      while (wrapper.firstChild) parent.insertBefore(wrapper.firstChild, wrapper);
      parent.removeChild(wrapper);
    }
  }
}

// ---------------------------------------------------------------------------
// Layout notes: unchanged wording whose container changed
// ---------------------------------------------------------------------------

/** Paragraph style of a layout note. The Checker recognises the note by it. */
export const LAYOUT_NOTE_STYLE = 'RedlineLayoutNote';

/** What a layout note says. `count` paragraphs follow the note (a table became text), or were there before (text became the table below). */
export function layoutNoteText(was: 'table' | 'paragraphs', count: number): string {
  if (was === 'table') return count === 1 ? 'Layout changed: the next paragraph was previously in a table.' : `Layout changed: the next ${count} paragraphs were previously in a table.`;
  return count === 1 ? 'Layout changed: the table below was previously a single paragraph.' : `Layout changed: the table below was previously ${count} separate paragraphs.`;
}

/** (Not a type guard: a failed test must not narrow the element's type away.) */
const is = (node: El, name: string): boolean => isWord(node, name);
type Side = 'added' | 'removed';
const WRAPPER: Record<Side, string> = { added: 'ins', removed: 'del' };
/** Markers that hold nothing a reader sees. */
const SILENT = new Set(['bookmarkStart', 'bookmarkEnd', 'proofErr']);

const flagged = (properties: El | undefined, name: string): boolean => !!properties && kids(properties).some((c) => is(c, name));
const childNamed = (parent: El | undefined, name: string): El | undefined => (parent ? kids(parent).find((c) => is(c, name)) : undefined);

/** The characters of a wrapper that holds only runs of text, tabs, line breaks and hyphens; undefined if it holds anything else. */
function wordsOf(wrapper: El): string | undefined {
  let text = '';
  for (const run of kids(wrapper)) {
    if (!is(run, 'r')) return undefined;
    for (const c of kids(run)) {
      if (is(c, 'rPr')) continue;
      if (is(c, 't') || is(c, 'delText')) text += shownText(c);
      else if (is(c, 'tab')) text += '\t';
      else if (is(c, 'br') && !c.getAttributeNS(W, 'type')) text += '\n';
      else if (is(c, 'noBreakHyphen')) text += '‑';
      else if (is(c, 'softHyphen')) text += '­';
      else return undefined;
    }
  }
  return text;
}

/**
 * A paragraph that was added (or removed) whole and holds nothing but words, in links or not: its text.
 * Undefined for any other paragraph.
 */
function wholeParagraph(paragraph: El, side: Side): string | undefined {
  if (!is(paragraph, 'p')) return undefined;
  const pPr = childNamed(paragraph, 'pPr');
  if (!flagged(childNamed(pPr, 'rPr'), WRAPPER[side]) || childNamed(pPr, 'sectPr')) return undefined;
  let text = '';
  const read = (wrapper: El): boolean => {
    const words = is(wrapper, WRAPPER[side]) ? wordsOf(wrapper) : undefined;
    if (words === undefined) return false;
    text += words;
    return true;
  };
  for (const c of kids(paragraph)) {
    if (c === pPr || (c.namespaceURI === W && SILENT.has(c.localName ?? ''))) continue;
    if (is(c, 'hyperlink') ? !kids(c).every(read) : !read(c)) return undefined;
  }
  return text;
}

/** A table that was added (or removed) whole, whose cells hold only such paragraphs: their texts, row by row. */
function wholeTable(table: El, side: Side): { paragraphs: El[]; texts: string[] } | undefined {
  const paragraphs: El[] = [];
  const texts: string[] = [];
  let rows = 0;
  for (const row of kids(table)) {
    if (is(row, 'tblPr') || is(row, 'tblGrid') || (row.namespaceURI === W && SILENT.has(row.localName ?? ''))) continue;
    if (!is(row, 'tr') || !flagged(childNamed(row, 'trPr'), WRAPPER[side])) return undefined;
    rows++;
    for (const cell of kids(row)) {
      if (is(cell, 'trPr') || is(cell, 'tblPrEx')) continue;
      if (!is(cell, 'tc')) return undefined;
      for (const block of kids(cell)) {
        if (is(block, 'tcPr') || (block.namespaceURI === W && SILENT.has(block.localName ?? ''))) continue;
        const text = wholeParagraph(block, side);
        if (text === undefined) return undefined;
        paragraphs.push(block);
        texts.push(text);
      }
    }
  }
  return rows > 0 ? { paragraphs, texts } : undefined;
}

/** Take the tracked-change markup off a paragraph that is to be shown as unchanged. */
function settle(paragraph: El, side: Side): void {
  const mark = childNamed(childNamed(paragraph, 'pPr'), 'rPr');
  for (const flag of mark ? kids(mark) : []) if (is(flag, WRAPPER[side])) mark!.removeChild(flag);
  const unwrap = (parent: El): void => {
    for (const c of kids(parent)) {
      if (is(c, 'hyperlink')) unwrap(c);
      if (!is(c, WRAPPER[side])) continue;
      while (c.firstChild) parent.insertBefore(c.firstChild, c);
      parent.removeChild(c);
    }
  };
  unwrap(paragraph);
}

function noteParagraph(before: El, text: string, fill: string): void {
  const doc = before.ownerDocument!;
  const prefix = before.prefix ?? 'w';
  const make = (name: string, attributes: Record<string, string> = {}): El => {
    const element = doc.createElementNS(W, `${prefix}:${name}`);
    for (const [key, value] of Object.entries(attributes)) element.setAttributeNS(W, `${prefix}:${key}`, value);
    return element;
  };
  const paragraph = make('p');
  const pPr = paragraph.appendChild(make('pPr'));
  pPr.appendChild(make('pStyle', { val: LAYOUT_NOTE_STYLE }));
  const run = paragraph.appendChild(make('r'));
  run.appendChild(make('rPr')).appendChild(make('shd', { val: 'clear', color: 'auto', fill }));
  const t = run.appendChild(make('t'));
  t.setAttribute('xml:space', 'preserve');
  t.appendChild(doc.createTextNode(text));
  before.parentNode!.insertBefore(paragraph, before);
}

/**
 * Where a table was removed and paragraphs added beside it (or the other way round) and both say exactly the same,
 * paragraph for paragraph (found on the owner's documents: a list taken out of a borderless table), nobody
 * changed the wording: only its container changed. The wording is then shown once, unmarked, as it now stands,
 * under one short shaded note saying that the layout changed.
 *
 * Narrow on purpose: directly in the main text, a header or a footer; plain words and links only; empty
 * paragraphs do not count as wording; and not where a paragraph that went carried a clause number (the note
 * would then hide a number that the reader of the original saw). `numberedBefore` says whether a removed
 * paragraph carried a clause number in the original (true when that is not known).
 */
export function noteLayoutChanges(root: El, fill: string, numberedBefore: (paragraph: El) => boolean): void {
  const container = root.localName === 'document' ? childNamed(root, 'body') : root.localName === 'hdr' || root.localName === 'ftr' ? root : undefined;
  if (!container) return;
  for (const table of kids(container)) {
    if (!is(table, 'tbl') || table.parentNode !== container) continue;
    for (const side of ['removed', 'added'] as const) {
      const whole = wholeTable(table, side);
      if (!whole) continue;
      const other: Side = side === 'removed' ? 'added' : 'removed';
      const wording = whole.texts.filter((text) => text !== '');
      if (wording.length === 0) continue;
      // The run of paragraphs of the other kind directly before the table, or directly after it.
      for (const step of ['previousSibling', 'nextSibling'] as const) {
        const run: El[] = [];
        for (let n = table[step] as El | null; n && (n.nodeType !== 1 || wholeParagraph(n, other) !== undefined); n = n[step] as El | null) if (n.nodeType === 1) run.push(n);
        if (step === 'previousSibling') run.reverse();
        const texts = run.map((paragraph) => wholeParagraph(paragraph, other)!);
        // The stretch of that run that says what the table says: from one paragraph with words to another.
        let found: El[] | undefined;
        for (let start = 0; start < run.length && !found; start++) {
          if (texts[start] === '') continue;
          let matched = 0;
          let end = start;
          for (; end < run.length && matched < wording.length; end++) {
            if (texts[end] === '') continue;
            if (texts[end] !== wording[matched]) break;
            matched++;
          }
          if (matched === wording.length) found = run.slice(start, end);
        }
        if (!found) continue;
        const removedParagraphs = side === 'removed' ? whole.paragraphs : found;
        if (removedParagraphs.some(numberedBefore)) continue;
        if (side === 'removed') {
          // A table became plain paragraphs: the paragraphs stay, the table goes.
          noteParagraph(found[0]!, layoutNoteText('table', found.length), fill);
          for (const paragraph of found) settle(paragraph, 'added');
          container.removeChild(table);
        } else {
          // Plain paragraphs became a table: the table stays, the paragraphs go. The table takes the place the
          // paragraphs had, so that removed paragraphs around them stay in the original's order (review finding;
          // only removed paragraphs stand between the two places, so the modified document reads as before).
          container.insertBefore(table, found[0]!);
          noteParagraph(table, layoutNoteText('paragraphs', found.length), fill);
          for (const paragraph of whole.paragraphs) settle(paragraph, 'added');
          for (const row of kids(table)) {
            const trPr = childNamed(row, 'trPr');
            for (const flag of trPr ? kids(trPr) : []) if (is(flag, 'ins')) trPr!.removeChild(flag);
          }
          for (const paragraph of found) container.removeChild(paragraph);
        }
        break;
      }
      break;
    }
  }
}
