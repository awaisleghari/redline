/**
 * Producer: one model of paragraph marks, applied to the tracked-changes markup before it is turned into a
 * static redline.
 *
 * A run of paragraphs stands for two readings, the original's and the modified document's. In each reading a
 * paragraph ends at a mark that reading has; a mark flagged as removed belongs to the original only, a mark
 * flagged as added to the modified document only. With u unflagged, i added and d removed marks before the fixed
 * last mark of a table cell (or of a note, a text box, the text itself), the modified document has u + i + 1
 * paragraphs there and the original u + d + 1.
 *
 * The same two readings can be written in more than one way, and the compare engine uses several. A clause
 * removed from the end of a cell may be written with the removal on its own mark, or one mark early (on the
 * clause before, which in fact stays), because the last mark of a cell cannot be removed in Word. In a static
 * redline nothing is accepted or rejected: every line is printed, and a reader takes a line's number, and whether
 * the line is struck out, from the line itself. So before conversion each run is put into ONE form:
 *
 *     each reading's paragraph ends at its last line that holds anything of that reading.
 *
 * Lines after that one, holding nothing of the reading, are then lines the reading does not have at all: a
 * removed clause is the line that is struck through from its first word to its mark, and the clause before it is
 * untouched. Both readings stay exactly what they were (only marks that separate nothing move). Where the
 * modified document's paragraph now ends at an earlier line, that line takes over the paragraph's settings
 * (its style and its clause number), which the tracked-changes file states on the line the paragraph used to end at.
 *
 * This replaces a set of separate rules for the ends of table cells (a removed last clause, runs of added or
 * removed clauses, a line replaced by another, a clause recorded one line early). The static conversion itself
 * now treats a paragraph in a cell like any other.
 */
import { NS } from '../common/namespaces';
import { isWord, kids } from './original-reader';

const W = NS.w;
type El = Element;
type Reading = 'old' | 'new';

/** Not a type guard, so that a failed test does not narrow the element's type away. */
const is = (node: El, name: string): boolean => isWord(node, name);
const named = (node: El, names: ReadonlySet<string>): boolean => node.namespaceURI === W && names.has(node.localName ?? '');
const child = (parent: El | undefined, name: string): El | undefined => (parent ? kids(parent).find((c) => is(c, name)) : undefined);

/** Markers between or inside paragraphs that hold nothing a reader sees. */
const SILENT = new Set(['bookmarkStart', 'bookmarkEnd', 'proofErr', 'permStart', 'permEnd', 'commentRangeStart', 'commentRangeEnd', 'moveFromRangeStart', 'moveFromRangeEnd', 'moveToRangeStart', 'moveToRangeEnd']);
/** Inside a run: what prints nothing in the text (the rest of a run's content does). */
const PRINTS_NOTHING = new Set(['rPr', 'lastRenderedPageBreak', 'footnoteRef', 'endnoteRef', 'annotationRef', 'commentReference', 'separator', 'continuationSeparator']);
const TEXT = new Set(['t', 'delText', 'instrText', 'delInstrText']);
const OLD_ONLY = new Set(['del', 'moveFrom']);
const NEW_ONLY = new Set(['ins', 'moveTo']);
const FLAGS = new Set(['ins', 'del', 'moveFrom', 'moveTo']);
/** Inside a paragraph's settings: what must stand after the mark's own settings. */
const AFTER_MARK = new Set(['sectPr', 'pPrChange']);

/** Which readings hold anything of this paragraph's content. */
function holds(paragraph: El): Record<Reading, boolean> {
  const found = { old: false, new: false };
  const walk = (parent: El, inOld: boolean, inNew: boolean): void => {
    for (const c of kids(parent)) {
      if (found.old && found.new) return;
      if (c.namespaceURI !== W) {
        // A drawing, an equation, an alternative form: something is shown.
        if (inOld) found.old = true;
        if (inNew) found.new = true;
        continue;
      }
      const name = c.localName ?? '';
      if (name === 'pPr' || SILENT.has(name)) continue;
      if (OLD_ONLY.has(name)) walk(c, inOld, false);
      else if (NEW_ONLY.has(name)) walk(c, false, inNew);
      else if (name === 'r') {
        const prints = kids(c).some((x) => {
          if (x.namespaceURI !== W) return true;
          if (PRINTS_NOTHING.has(x.localName ?? '')) return false;
          if (TEXT.has(x.localName ?? '')) return (x.textContent ?? '') !== '';
          // A holder for legacy shapes that holds none draws nothing.
          if (x.localName === 'pict') return kids(x).some((shape) => shape.localName !== 'shapetype');
          return true;
        });
        if (prints && inOld) found.old = true;
        if (prints && inNew) found.new = true;
      } else walk(c, inOld, inNew);
    }
  };
  walk(paragraph, true, true);
  return found;
}

const markSettings = (paragraph: El): El | undefined => child(child(paragraph, 'pPr'), 'rPr');

/** Which readings have this paragraph's mark, and the flag that says so. */
function markOf(paragraph: El): { old: boolean; new: boolean; flag?: string } {
  const settings = markSettings(paragraph);
  const flag = (settings && kids(settings).find((c) => named(c, FLAGS))?.localName) ?? undefined;
  return { old: !flag || OLD_ONLY.has(flag), new: !flag || NEW_ONLY.has(flag), flag };
}

/** Write which readings have a paragraph's mark. */
function setMark(paragraph: El, has: Record<Reading, boolean>, flagBefore: string | undefined): void {
  const prefix = paragraph.prefix ?? 'w';
  const doc = paragraph.ownerDocument!;
  let settings = markSettings(paragraph);
  for (const flag of settings ? kids(settings) : []) if (named(flag, FLAGS)) settings!.removeChild(flag);
  if (has.old && has.new) return;
  // (A moved paragraph's mark keeps saying "moved" where it still belongs to the same reading alone.)
  const name = has.old ? (flagBefore === 'moveFrom' ? 'moveFrom' : 'del') : flagBefore === 'moveTo' ? 'moveTo' : 'ins';
  if (!settings) {
    let pPr = child(paragraph, 'pPr');
    if (!pPr) pPr = paragraph.insertBefore(doc.createElementNS(W, `${prefix}:pPr`), paragraph.firstChild) as El;
    settings = pPr.insertBefore(doc.createElementNS(W, `${prefix}:rPr`), kids(pPr).find((c) => named(c, AFTER_MARK)) ?? null) as El;
  }
  settings.insertBefore(doc.createElementNS(W, `${prefix}:${name}`), settings.firstChild);
}

/** Give `to` the paragraph settings of `from` (style, numbering, layout). A section break stays where it is. */
function takeSettings(to: El, from: El): void {
  const own = child(to, 'pPr');
  const ownBreak = child(own, 'sectPr');
  const source = child(from, 'pPr');
  const copy = source ? (source.cloneNode(true) as El) : (to.ownerDocument!.createElementNS(W, `${to.prefix ?? 'w'}:pPr`) as El);
  const copiedBreak = child(copy, 'sectPr');
  if (copiedBreak) copy.removeChild(copiedBreak);
  if (ownBreak) copy.insertBefore(ownBreak, child(copy, 'pPrChange') ?? null);
  if (own) to.replaceChild(copy, own);
  else to.insertBefore(copy, to.firstChild);
}

/** Put one run of paragraphs (side by side in one container) into the one form described at the top of this file. */
function settleRun(run: El[], closed: boolean): void {
  const before = run.map(markOf);
  const content = run.map(holds);
  const after = before.map((mark) => ({ old: mark.old, new: mark.new }));
  const settingsFrom = new Map<number, number>();
  for (const reading of ['old', 'new'] as const) {
    let start = 0;
    /** One paragraph of this reading: lines `start` to `end`; it ends at line `end`'s mark, or (`end` undefined) where the run ends. */
    const paragraph = (from: number, end: number | undefined): void => {
      let last: number | undefined;
      for (let i = end ?? run.length - 1; i >= from && last === undefined; i--) if (content[i]![reading]) last = i;
      if (last === undefined || last === end) return;
      after[last]![reading] = true;
      for (let i = last + 1; i <= (end ?? run.length - 1); i++) after[i]![reading] = false;
      // The modified document's paragraph has the settings stated on the line it ended at.
      if (reading === 'new' && end !== undefined) settingsFrom.set(last, end);
    };
    for (let i = 0; i < run.length; i++) {
      if (!before[i]![reading]) continue;
      paragraph(start, i);
      start = i + 1;
    }
    // What follows the last mark this reading has: a paragraph of its own if it holds anything (and nothing can follow it).
    if (start < run.length && closed) paragraph(start, undefined);
  }
  if (after.some((mark) => !mark.old && !mark.new)) return; // cannot happen; if it did, the run is left as the engine wrote it
  for (const [to, from] of settingsFrom) takeSettings(run[to]!, run[from]!);
  run.forEach((paragraph, i) => {
    if (settingsFrom.has(i) || after[i]!.old !== before[i]!.old || after[i]!.new !== before[i]!.new) setMark(paragraph, after[i]!, before[i]!.flag);
  });
}

/** Put every run of paragraphs under `root` (one part of the tracked-changes file) into the one form. */
export function settleParagraphMarks(root: El): void {
  const containers = new Set<El>();
  const paragraphs = root.getElementsByTagNameNS(W, 'p');
  for (let i = 0; i < paragraphs.length; i++) containers.add(paragraphs.item(i)!.parentNode as El);
  for (const container of containers) {
    let run: El[] = [];
    const end = (closed: boolean): void => {
      if (run.some((paragraph) => markOf(paragraph).flag)) settleRun(run, closed);
      run = [];
    };
    for (const block of kids(container)) {
      if (is(block, 'p')) run.push(block);
      else if (named(block, SILENT) || is(block, 'tcPr') || is(block, 'sectPr')) continue;
      // A table ends the paragraph before it. Anything else (a content control, say) may carry the paragraph on: its end is not judged.
      else end(is(block, 'tbl'));
    }
    end(true);
  }
}
