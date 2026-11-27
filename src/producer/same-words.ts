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
