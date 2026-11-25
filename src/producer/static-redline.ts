/**
 * Producer: turn the tracked-changes Word file into a static redline, in which
 * every change is ordinary text carrying fixed formatting (the style spec in
 * PLAN.md) and no tracked changes remain.
 *
 *   inserted   blue (0000FF), double underline
 *   deleted    red (FF0000), strikethrough
 *   moved from green (008000), double strikethrough
 *   moved to   green (008000), double underline
 *
 * This file is Producer code. The Checker verifies its result independently
 * and shares nothing with it beyond `src/common/`.
 */
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { NS } from '../common/namespaces';
import { BIT, isWord, kid, kids, paragraphsInCountingOrder, shownText, unitsOf, type OriginalInfo, type StyleSheet } from './original-reader';

const W = NS.w;

/** Shading that marks unchanged text whose formatting changed. */
export const FORMAT_CHANGE_FILL = 'FFF2CC';
/** Character style of the old clause number typed in front of a removed paragraph. */
export const OLD_NUMBER_STYLE = 'RedlineOldNumber';

export type ChangeKind = 'inserted' | 'deleted' | 'movedFrom' | 'movedTo';

export interface ChangeCounts {
  insertions: number;
  deletions: number;
  /** Counted once per block moved away from its original place. */
  moves: number;
  /** Unbroken stretches of unchanged text whose formatting changed. */
  formattingChanges: number;
}

/** What the conversion needs from the original document and the tracked file's styles. */
export interface OriginalContext {
  info: OriginalInfo;
  /** Styles of the tracked-changes file (the formatting the text has now). */
  styles: StyleSheet;
  /** Whether the part being converted is the main document (old clause numbers are shown there). */
  isMain: boolean;
}

/** The formatting that marks each kind of change. */
export const CHANGE_STYLE: Record<ChangeKind, { color: string; underline?: 'double'; strike?: 'single' | 'double' }> = {
  inserted: { color: '0000FF', underline: 'double' },
  deleted: { color: 'FF0000', strike: 'single' },
  movedFrom: { color: '008000', strike: 'double' },
  movedTo: { color: '008000', underline: 'double' },
};

const WRAPPER_KIND: Record<string, ChangeKind> = { ins: 'inserted', del: 'deleted', moveFrom: 'movedFrom', moveTo: 'movedTo' };

/** Records of a change to properties (formatting, numbering, table shape). Removed: the new state is kept. */
export const PROPERTY_CHANGE_RECORDS = new Set(['rPrChange', 'pPrChange', 'sectPrChange', 'tblPrChange', 'tblPrExChange', 'tcPrChange', 'trPrChange', 'tblGridChange', 'numberingChange']);
/** Zero-width markers that only make sense alongside tracked changes. Removed. */
const RANGE_MARKERS = new Set(['moveFromRangeStart', 'moveFromRangeEnd', 'moveToRangeStart', 'moveToRangeEnd', 'customXmlInsRangeStart', 'customXmlInsRangeEnd', 'customXmlDelRangeStart', 'customXmlDelRangeEnd', 'customXmlMoveFromRangeStart', 'customXmlMoveFromRangeEnd', 'customXmlMoveToRangeStart', 'customXmlMoveToRangeEnd']);
/** Property containers in which w:ins / w:del etc. are flags on the owner, not wrappers round content. */
const FLAG_PARENTS = new Set(['rPr', 'trPr', 'tcPr', 'numPr', 'tblPrEx']);
const CELL_FLAGS = new Set(['cellIns', 'cellDel', 'cellMerge']);

/** Run properties in the order Word's schema lists them, so that added ones land in a valid position. */
const RPR_ORDER = ['rStyle', 'rFonts', 'b', 'bCs', 'i', 'iCs', 'caps', 'smallCaps', 'strike', 'dstrike', 'outline', 'shadow', 'emboss', 'imprint', 'noProof', 'snapToGrid', 'vanish', 'webHidden', 'color', 'spacing', 'w', 'kern', 'position', 'sz', 'szCs', 'highlight', 'u', 'effect', 'bdr', 'shd', 'fitText', 'vertAlign', 'rtl', 'cs', 'em', 'lang', 'eastAsianLayout', 'specVanish', 'oMath'];
/** Paragraph properties that must come after w:rPr inside w:pPr. */
const PPR_AFTER_RPR = new Set(['sectPr', 'pPrChange']);

const isW = (node: Node | null | undefined, name?: string): node is Element =>
  !!node && node.nodeType === 1 && (node as Element).namespaceURI === W && (name === undefined || (node as Element).localName === name);

function elementChildren(parent: Element): Element[] {
  const result: Element[] = [];
  for (let n = parent.firstChild; n; n = n.nextSibling) if (n.nodeType === 1) result.push(n as Element);
  return result;
}

function childW(parent: Element, name: string): Element | undefined {
  return elementChildren(parent).find((c) => isW(c, name));
}

/** Set one run property, replacing any existing one and keeping schema order. */
function setRunProperty(rPr: Element, name: string, attributes: Record<string, string> = {}): void {
  removeRunProperty(rPr, name);
  const prefix = rPr.prefix ?? 'w';
  const element = rPr.ownerDocument!.createElementNS(W, `${prefix}:${name}`);
  for (const [key, value] of Object.entries(attributes)) element.setAttributeNS(W, `${prefix}:${key}`, value);
  const rank = RPR_ORDER.indexOf(name);
  const before = elementChildren(rPr).find((c) => {
    if (!isW(c)) return false;
    const other = RPR_ORDER.indexOf(c.localName!);
    return other === -1 ? c.localName === 'rPrChange' : other > rank;
  });
  rPr.insertBefore(element, before ?? null);
}

function removeRunProperty(rPr: Element, name: string): void {
  for (const c of elementChildren(rPr)) if (isW(c, name)) rPr.removeChild(c);
}

/** Apply a change's formatting to a set of run properties (of a run, or of a paragraph mark). */
function applyChangeStyle(rPr: Element, kind: ChangeKind, colors: { insertedColor: string; deletedColor: string }): void {
  const style = kind === 'inserted' ? { ...CHANGE_STYLE.inserted, color: colors.insertedColor } : kind === 'deleted' ? { ...CHANGE_STYLE.deleted, color: colors.deletedColor } : CHANGE_STYLE[kind];
  setRunProperty(rPr, 'color', { val: style.color });
  if (style.underline) setRunProperty(rPr, 'u', { val: style.underline });
  if (style.strike) {
    // Single and double strikethrough exclude each other.
    removeRunProperty(rPr, 'strike');
    removeRunProperty(rPr, 'dstrike');
    setRunProperty(rPr, style.strike === 'double' ? 'dstrike' : 'strike');
  }
}

function runProperties(run: Element): Element {
  const existing = childW(run, 'rPr');
  if (existing) return existing;
  const rPr = run.ownerDocument!.createElementNS(W, `${run.prefix ?? 'w'}:rPr`);
  run.insertBefore(rPr, run.firstChild);
  return rPr;
}

/** The run properties of a paragraph's mark (w:pPr/w:rPr), created if absent. */
function markProperties(paragraph: Element): Element {
  const prefix = paragraph.prefix ?? 'w';
  let pPr = childW(paragraph, 'pPr');
  if (!pPr) {
    pPr = paragraph.ownerDocument!.createElementNS(W, `${prefix}:pPr`);
    paragraph.insertBefore(pPr, paragraph.firstChild);
  }
  let rPr = childW(pPr, 'rPr');
  if (!rPr) {
    rPr = paragraph.ownerDocument!.createElementNS(W, `${prefix}:rPr`);
    const before = elementChildren(pPr).find((c) => isW(c) && PPR_AFTER_RPR.has(c.localName!));
    pPr.insertBefore(rPr, before ?? null);
  }
  return rPr;
}

function rename(element: Element, localName: string): Element {
  const doc = element.ownerDocument!;
  const replacement = doc.createElementNS(W, `${element.prefix ?? 'w'}:${localName}`);
  for (let i = 0; i < element.attributes.length; i++) {
    const a = element.attributes.item(i)!;
    if (a.namespaceURI) replacement.setAttributeNS(a.namespaceURI, a.name, a.value);
    else replacement.setAttribute(a.name, a.value);
  }
  while (element.firstChild) replacement.appendChild(element.firstChild);
  element.parentNode!.replaceChild(replacement, element);
  return replacement;
}

/** One piece of the document in reading order, for counting changes. */
type Token = ChangeKind | 'unchanged' | 'boundary';

/** Choices about how the static redline is laid out (see the decisions log in PLAN.md). */
export interface StaticRedlineStyleOptions {
  /** Where a phrase is replaced, show the deleted text before the inserted text. */
  deletionBeforeInsertion: boolean;
  /**
   * Take automatic numbering off paragraphs that were deleted or moved away, so
   * that the remaining clause numbers are exactly those of the modified document.
   */
  unnumberRemovedParagraphs: boolean;
  /** Type the old clause number, struck through, in front of a removed numbered paragraph. */
  showOldNumbers: boolean;
  /** Shade unchanged text whose bold, italic, underline, capitals or hidden-text setting changed. */
  markFormattingChanges: boolean;
  /** The colour of inserted text: the standard blue, or the alternative where a document itself uses blue double underline. */
  insertedColor: string;
  /** The colour of deleted text (see DELETED_COLOR). */
  deletedColor: string;
}

/**
 * Inserted text is blue unless one of the two documents itself contains text (or a paragraph mark) formatted
 * blue with a double underline, which a reader could not tell from an insertion. Then insertions are shown in
 * a dark magenta, and the legend says so (owner's decision, 2026-11-24).
 */
export const INSERTED_COLOR = { standard: '0000FF', alternative: 'AA00AA' } as const;
/**
 * Deleted text, on the same rule: red, unless one of the two documents itself contains text (or a paragraph mark)
 * formatted red with a strikethrough (found on a real contract). Then deletions are shown in a dark orange, and the
 * legend says so.
 */
export const DELETED_COLOR = { standard: 'FF0000', alternative: 'C55A11' } as const;

/** Whether a part's own runs or paragraph marks are formatted red with a (single) strikethrough, stated on the run or mark itself. */
export function usesDeletionLook(xml: string): boolean {
  if (!/w:val="FF0000"/i.test(xml) || !/<w:strike\b/.test(xml)) return false;
  const root = new DOMParser().parseFromString(xml, 'application/xml').documentElement as unknown as Element;
  for (const name of ['r', 'pPr']) {
    const owners = root.getElementsByTagNameNS(W, name);
    for (let i = 0; i < owners.length; i++) {
      const rPr = elementChildren(owners.item(i) as unknown as Element).find((c) => isW(c, 'rPr'));
      if (!rPr) continue;
      const color = elementChildren(rPr).find((c) => isW(c, 'color'));
      const strike = elementChildren(rPr).find((c) => isW(c, 'strike'));
      const struck = !!strike && !['0', 'false', 'off'].includes(strike.getAttributeNS(W, 'val') ?? '');
      if (color && !color.getAttributeNS(W, 'themeColor') && color.getAttributeNS(W, 'val')?.toUpperCase() === DELETED_COLOR.standard && struck) return true;
    }
  }
  return false;
}

/** Whether a part's own runs or paragraph marks are formatted blue with a double underline (stated on the run or mark itself). */
export function usesInsertionLook(xml: string): boolean {
  if (!/w:val="0000FF"/i.test(xml) || !/<w:u\b[^>]*w:val="double"/.test(xml)) return false;
  const root = new DOMParser().parseFromString(xml, 'application/xml').documentElement as unknown as Element;
  for (const name of ['r', 'pPr']) {
    const owners = root.getElementsByTagNameNS(W, name);
    for (let i = 0; i < owners.length; i++) {
      const rPr = elementChildren(owners.item(i) as unknown as Element).find((c) => isW(c, 'rPr'));
      if (!rPr) continue;
      const color = elementChildren(rPr).find((c) => isW(c, 'color'));
      const underline = elementChildren(rPr).find((c) => isW(c, 'u'));
      if (color && !color.getAttributeNS(W, 'themeColor') && color.getAttributeNS(W, 'val')?.toUpperCase() === INSERTED_COLOR.standard && underline?.getAttributeNS(W, 'val') === 'double') return true;
    }
  }
  return false;
}

export const DEFAULT_STYLE_OPTIONS: StaticRedlineStyleOptions = { deletionBeforeInsertion: true, unnumberRemovedParagraphs: true, showOldNumbers: true, markFormattingChanges: true, insertedColor: INSERTED_COLOR.standard, deletedColor: DELETED_COLOR.standard };

const hasFlag = (properties: Element | undefined, names: string[]): boolean => !!properties && kids(properties).some((c) => isWord(c) && names.includes(c.localName!));

/** Is this element inside a tracked wrapper of one of these kinds? */
function insideWrapper(node: Element, names: string[]): boolean {
  for (let n: Node | null = node.parentNode; n; n = n.parentNode) {
    if (!isWord(n) || !names.includes(n.localName!)) continue;
    const parent = n.parentNode as Element | null;
    if (!(parent && isWord(parent) && FLAG_PARENTS.has(parent.localName!))) return true;
  }
  return false;
}

function nearestParagraph(node: Element): Element | undefined {
  for (let n: Node | null = node.parentNode; n; n = n.parentNode) if (isWord(n, 'p')) return n;
  return undefined;
}

/**
 * Shade a run, or just the characters of it whose formatting changed
 * (splitting the run where needed). `changed` has one entry per character of
 * the run's text, in order.
 */
function shadeRun(run: Element, changed: boolean[]): void {
  const shade = (target: Element): void => setRunProperty(runProperties(target), 'shd', { val: 'clear', color: 'auto', fill: FORMAT_CHANGE_FILL });
  if (changed.every((c) => c)) {
    shade(run);
    return;
  }
  // Mixed: rebuild the run as a sequence of runs, each wholly changed or wholly unchanged.
  const parent = run.parentNode!;
  const doc = run.ownerDocument!;
  const rPr = kid(run, 'rPr');
  let at = 0;
  let current: { element: Element; changed: boolean } | undefined;
  const piece = (isChanged: boolean): Element => {
    if (!current || current.changed !== isChanged) {
      const element = doc.createElementNS(W, run.nodeName);
      if (rPr) element.appendChild(rPr.cloneNode(true));
      parent.insertBefore(element, run);
      if (isChanged) shade(element);
      current = { element, changed: isChanged };
    }
    return current.element;
  };
  for (const child of kids(run)) {
    if (child === rPr) continue;
    if (!isWord(child, 't') && !isWord(child, 'delText')) {
      // Tabs, breaks and the like stay with the text before them (or after, at the start of the run).
      piece(current?.changed ?? changed[at] ?? false).appendChild(child);
      continue;
    }
    const text = shownText(child);
    let start = 0;
    while (start < text.length) {
      const flag = changed[at + start]!;
      let end = start;
      while (end < text.length && changed[at + end] === flag) end++;
      const part = doc.createElementNS(W, child.nodeName);
      part.setAttribute('xml:space', 'preserve');
      part.appendChild(doc.createTextNode(text.slice(start, end)));
      piece(flag).appendChild(part);
      start = end;
    }
    at += text.length;
  }
  parent.removeChild(run);
}

/**
 * Before conversion, while the tracked changes still say what is kept and what
 * is deleted: shade unchanged text whose formatting differs from the original's.
 */
function markFormattingChanges(root: Element, original: OriginalContext): void {
  for (const unit of unitsOf(root)) {
    // This unit as the original had it: unchanged and deleted text, in order.
    const pieces: { run: Element; kept: boolean; bits: number; length: number }[] = [];
    let text = '';
    const runs = unit.getElementsByTagNameNS(W, 'r');
    for (let i = 0; i < runs.length; i++) {
      const run = runs.item(i) as Element;
      if (insideWrapper(run, ['ins', 'moveTo'])) continue;
      const kept = !insideWrapper(run, ['del', 'moveFrom']);
      let length = 0;
      for (const child of kids(run)) {
        if (isWord(child, 't') || isWord(child, 'delText')) {
          const shown = shownText(child);
          text += shown;
          length += shown.length;
        }
      }
      if (length > 0) pieces.push({ run, kept, bits: original.styles.bits(kid(run, 'rPr'), nearestParagraph(run)), length });
    }
    // Find the same unit in the original by its text.
    const match = original.info.streams.find((s) => !s.used && s.kind === root.localName && s.stream.text === text);
    if (!match) continue;
    match.used = true;
    let at = 0;
    for (const piece of pieces) {
      const changed: boolean[] = [];
      for (let k = 0; k < piece.length; k++) changed.push(piece.kept && match.stream.bits[at + k] !== piece.bits);
      // Deleted text is shown as it looked in the original. If its surroundings
      // changed (its paragraph's style, say) so that it would now look different,
      // state the original look on the run itself.
      if (!piece.kept) {
        const was = match.stream.bits[at]!;
        const uniform = match.stream.bits.slice(at, at + piece.length).every((b) => b === was);
        if (uniform && was !== piece.bits) {
          for (const [name, bit] of Object.entries(BIT)) {
            if (name === 'strike' || name === 'dstrike' || (was & bit) === (piece.bits & bit)) continue; // strikethrough is set by the deletion itself
            const on = (was & bit) !== 0;
            setRunProperty(runProperties(piece.run), name, name === 'u' ? { val: on ? 'single' : 'none' } : on ? {} : { val: '0' });
          }
        }
      }
      at += piece.length;
      if (changed.some((c) => c)) shadeRun(piece.run, changed);
    }
  }
}

/**
 * Before conversion: text that stays, in a paragraph whose mark is removed. In the modified document that text is
 * part of the paragraph it joins, and is formatted by that paragraph's style (found on a real contract: a word
 * that became the start of a bold heading). The static redline shows it in the paragraph it sits in, so where the
 * two styles would make it look different, the look it has in the modified document is stated on the text itself.
 */
function stateLookOfJoinedText(root: Element, styles: StyleSheet): void {
  const markRemoved = (paragraph: Element): boolean => hasFlag(kid(kid(paragraph, 'pPr'), 'rPr'), ['del', 'moveFrom']);
  const paragraphs = root.getElementsByTagNameNS(W, 'p');
  for (let i = 0; i < paragraphs.length; i++) {
    const paragraph = paragraphs.item(i) as Element;
    if (!markRemoved(paragraph)) continue;
    // The paragraph it joins: the next one whose mark stays. Only paragraphs directly after one another join.
    let joined: Element | undefined = paragraph;
    while (joined && markRemoved(joined)) {
      let next: Node | null = joined.nextSibling;
      while (next && next.nodeType !== 1) next = next.nextSibling;
      joined = next && isWord(next, 'p') ? (next as Element) : undefined;
    }
    if (!joined) continue;
    const runs = paragraph.getElementsByTagNameNS(W, 'r');
    for (let j = 0; j < runs.length; j++) {
      const run = runs.item(j) as Element;
      if (nearestParagraph(run) !== paragraph || insideWrapper(run, ['del', 'moveFrom'])) continue;
      const rPr = kid(run, 'rPr');
      const here = styles.bits(rPr, paragraph);
      const there = styles.bits(rPr, joined);
      if (here === there) continue;
      for (const [name, bit] of Object.entries(BIT)) {
        if ((here & bit) === (there & bit)) continue;
        const on = (there & bit) !== 0;
        setRunProperty(runProperties(run), name, name === 'u' ? { val: on ? 'single' : 'none' } : on ? {} : { val: '0' });
      }
    }
  }
}

/** The original's clause number for each paragraph of the main document that the original had. */
function oldNumbers(root: Element, original: OriginalContext): Map<Element, string> {
  const result = new Map<Element, string>();
  const paragraphs = paragraphsInCountingOrder(
    root,
    (paragraph) => !hasFlag(kid(kid(paragraph, 'pPr'), 'rPr'), ['ins', 'moveTo']),
    (element) => {
      if (!isWord(element)) return true;
      const parent = element.parentNode as Element;
      if ((element.localName === 'ins' || element.localName === 'moveTo') && !(isWord(parent) && FLAG_PARENTS.has(parent.localName!))) return false;
      if (element.localName === 'tr' && hasFlag(kid(element, 'trPr'), ['ins'])) return false;
      if (element.localName === 'tc' && hasFlag(kid(element, 'tcPr'), ['cellIns'])) return false;
      return true;
    },
  );
  // Only if the paragraphs line up one for one with the original's can the numbers be trusted.
  if (paragraphs.length !== original.info.bodyNumbers.length) return result;
  paragraphs.forEach((paragraph, index) => result.set(paragraph, original.info.bodyNumbers[index]!));
  return result;
}

/** Count formatting changes: unbroken stretches of shaded runs within one paragraph. */
function countShadedStretches(root: Element): number {
  let count = 0;
  const paragraphs = root.getElementsByTagNameNS(W, 'p');
  for (let i = 0; i < paragraphs.length; i++) {
    let shadedBefore = false;
    const visit = (parent: Element): void => {
      for (const child of kids(parent)) {
        if (child.namespaceURI === W && child.localName === 'r') {
          const rPr = kid(child, 'rPr');
          if (!kids(child).some((c) => c !== rPr)) continue;
          const shaded = kid(rPr, 'shd')?.getAttributeNS(W, 'fill')?.toUpperCase() === FORMAT_CHANGE_FILL;
          if (shaded && !shadedBefore) count++;
          shadedBefore = shaded;
        } else if (child.namespaceURI === W && ['hyperlink', 'sdt', 'sdtContent', 'smartTag', 'fldSimple', 'customXml'].includes((child as Element).localName ?? '')) visit(child as Element);
      }
    };
    visit(paragraphs.item(i) as Element);
  }
  return count;
}

/** Paragraph properties that come before w:numPr in Word's schema order. */
const PPR_BEFORE_NUMPR = new Set(['pStyle', 'keepNext', 'keepLines', 'pageBreakBefore', 'framePr', 'widowControl']);

class PartConverter {
  readonly tokens: Token[] = [];
  /** Rows and cells that were themselves inserted or deleted. */
  private readonly cellChange = new Map<Element, ChangeKind>();
  /** Where each paragraph's mark sits in `tokens`. */
  private readonly markToken = new Map<Element, number>();
  /** Last paragraphs of cells that were removed whole, where the engine put the removal on the mark before. */
  private readonly removedAtCellEnd = new Map<Element, ChangeKind>();

  constructor(
    private readonly options: StaticRedlineStyleOptions,
    private readonly oldNumberOf: Map<Element, string> = new Map(),
    /** Looks up a paragraph's numbering (its own, or its style's) in the tracked-changes file. */
    private readonly numberingOf: (paragraph: Element) => Element | undefined = (paragraph) => childW(childW(paragraph, 'pPr') ?? paragraph, 'numPr'),
  ) {}

  /** Type the old clause number, struck through like the paragraph it belongs to, at the start of a removed paragraph. */
  private addOldNumber(paragraph: Element, label: string, kind: ChangeKind): void {
    const doc = paragraph.ownerDocument!;
    const prefix = paragraph.prefix ?? 'w';
    const make = (name: string): Element => doc.createElementNS(W, `${prefix}:${name}`);
    const run = make('r');
    const rPr = make('rPr');
    const style = make('rStyle');
    style.setAttributeNS(W, `${prefix}:val`, OLD_NUMBER_STYLE);
    rPr.appendChild(style);
    run.appendChild(rPr);
    applyChangeStyle(rPr, kind, this.options);
    const text = make('t');
    text.setAttribute('xml:space', 'preserve');
    text.appendChild(doc.createTextNode(label));
    run.appendChild(text);
    run.appendChild(make('tab'));
    const first = kids(paragraph).find((c) => !isWord(c, 'pPr'));
    paragraph.insertBefore(run, first ?? null);
  }

  /** Within each unbroken group of insertions and deletions, put the deletions first. */
  private putDeletionsFirst(parent: Element): void {
    const children = elementChildren(parent);
    let i = 0;
    while (i < children.length) {
      let j = i;
      while (j < children.length && (isW(children[j], 'ins') || isW(children[j], 'del')) && children[j]!.previousSibling === (j > i ? children[j - 1] : children[j]!.previousSibling)) j++;
      if (j - i > 1) {
        const group = children.slice(i, j);
        const after = group[group.length - 1]!.nextSibling;
        for (const node of [...group.filter((n) => isW(n, 'del')), ...group.filter((n) => isW(n, 'ins'))]) parent.insertBefore(node, after);
      }
      i = Math.max(j, i + 1);
    }
  }

  /**
   * Is this the paragraph before the last one of a cell that stays, with that last paragraph removed whole (all
   * of it inside the same kind of removal as this paragraph's mark) and its own mark not flagged?
   */
  private endsCellWithRemovedClause(paragraph: Element, markKind: ChangeKind): boolean {
    const cell = paragraph.parentNode as Element;
    if (!isWord(cell, 'tc')) return false;
    const blocks = kids(cell).filter((c) => isWord(c, 'p') || isWord(c, 'tbl'));
    const last = blocks[blocks.length - 1]!;
    if (blocks[blocks.length - 2] !== paragraph || !isWord(last, 'p') || paragraph.nextSibling !== last) return false;
    let row = cell.parentNode as Element | null;
    while (row && isWord(row) && ['sdt', 'sdtContent', 'customXml'].includes(row.localName!)) row = row.parentNode as Element | null;
    if (this.cellChange.has(cell) || (row && this.cellChange.has(row))) return false;
    const lastPPr = childW(last, 'pPr');
    const lastMark = lastPPr && childW(lastPPr, 'rPr');
    if (lastMark && elementChildren(lastMark).some((flag) => isW(flag) && flag.localName! in WRAPPER_KIND)) return false;
    const wrapper = markKind === 'deleted' ? 'del' : 'moveFrom';
    const content = elementChildren(last).filter((c) => !/^(pPr|bookmark|moveFromRange|moveToRange|proofErr|permStart|permEnd)/.test(c.namespaceURI === W ? c.localName! : ''));
    return content.length > 0 && content.every((c) => isW(c, wrapper));
  }

  /** Give `to` the numbering that `from` states for itself (list and level), in place of its own. */
  private takeNumberingFrom(from: Element, to: Element): void {
    // Its own numbering, or its style's; if it has none, the paragraph that takes over has none either.
    const stated = this.numberingOf(from);
    if (!stated) {
      this.removeNumbering(to);
      return;
    }
    const prefix = to.prefix ?? 'w';
    let pPr = childW(to, 'pPr');
    if (!pPr) {
      pPr = to.ownerDocument!.createElementNS(W, `${prefix}:pPr`);
      to.insertBefore(pPr, to.firstChild);
    }
    const existing = childW(pPr, 'numPr');
    if (existing) pPr.removeChild(existing);
    const before = elementChildren(pPr).find((c) => !(isW(c) && PPR_BEFORE_NUMPR.has(c.localName!)));
    pPr.insertBefore(stated.cloneNode(true), before ?? null);
  }

  private removeNumbering(paragraph: Element): void {
    const prefix = paragraph.prefix ?? 'w';
    const doc = paragraph.ownerDocument!;
    let pPr = childW(paragraph, 'pPr');
    if (!pPr) {
      pPr = doc.createElementNS(W, `${prefix}:pPr`);
      paragraph.insertBefore(pPr, paragraph.firstChild);
    }
    const existing = childW(pPr, 'numPr');
    if (existing) pPr.removeChild(existing);
    const numPr = doc.createElementNS(W, `${prefix}:numPr`);
    const numId = doc.createElementNS(W, `${prefix}:numId`);
    numId.setAttributeNS(W, `${prefix}:val`, '0'); // 0 means "no numbering", overriding any the style would give
    numPr.appendChild(numId);
    const before = elementChildren(pPr).find((c) => !(isW(c) && PPR_BEFORE_NUMPR.has(c.localName!)));
    pPr.insertBefore(numPr, before ?? null);
  }

  /** Convert everything under `node`. `kind` is the change the content is inside, if any. */
  convert(node: Element, kind: ChangeKind | undefined): void {
    for (const child of elementChildren(node)) this.visit(child, kind);
  }

  private visit(element: Element, kind: ChangeKind | undefined): void {
    if (!isW(element)) {
      this.convert(element, kind);
      return;
    }
    const name = element.localName!;
    const parent = element.parentNode as Element;

    if (PROPERTY_CHANGE_RECORDS.has(name) || RANGE_MARKERS.has(name) || CELL_FLAGS.has(name)) {
      parent.removeChild(element);
      return;
    }

    if (name in WRAPPER_KIND) {
      if (isW(parent) && FLAG_PARENTS.has(parent.localName!)) {
        parent.removeChild(element); // a flag; paragraph marks are styled in visitParagraph
        return;
      }
      // A wrapper round content: style what is inside, then dissolve the wrapper.
      const inner = WRAPPER_KIND[name]!;
      this.convert(element, inner);
      while (element.firstChild) parent.insertBefore(element.firstChild, element);
      parent.removeChild(element);
      return;
    }

    switch (name) {
      case 'tr':
      case 'tc': {
        // Note whether the row or cell itself was inserted or deleted, before its flags are removed.
        const flags = kid(element, name === 'tr' ? 'trPr' : 'tcPr');
        const deleted = hasFlag(flags, name === 'tr' ? ['del'] : ['cellDel']);
        const inserted = hasFlag(flags, name === 'tr' ? ['ins'] : ['cellIns']);
        if (deleted || inserted) this.cellChange.set(element, deleted ? 'deleted' : 'inserted');
        this.convert(element, kind);
        return;
      }
      case 'p':
        this.visitParagraph(element, kind);
        return;
      case 'r':
        this.visitRun(element, kind);
        return;
      case 'footnote':
      case 'endnote':
      case 'comment':
        this.tokens.push('boundary');
        this.convert(element, kind);
        this.tokens.push('boundary');
        return;
      default:
        this.convert(element, kind);
    }
  }

  private visitRun(run: Element, kind: ChangeKind | undefined): void {
    let hasContent = false;
    for (const child of elementChildren(run)) {
      if (!isW(child)) {
        hasContent = true;
        continue;
      }
      if (child.localName === 'rPr') {
        this.convert(child, kind); // removes rPrChange and any flags
        continue;
      }
      hasContent = true;
      if (child.localName === 'delText') rename(child, 't');
      else if (child.localName === 'delInstrText') rename(child, 'instrText');
    }
    if (kind) applyChangeStyle(runProperties(run), kind, this.options);
    if (hasContent) this.tokens.push(kind ?? 'unchanged');
    // Text boxes and other nested content inside the run.
    for (const child of elementChildren(run)) if (!isW(child, 'rPr')) this.convert(child, kind);
  }

  private visitParagraph(paragraph: Element, kind: ChangeKind | undefined): void {
    // Which change, if any, applies to the paragraph mark itself?
    const pPr = childW(paragraph, 'pPr');
    const markRPr = pPr && childW(pPr, 'rPr');
    let markKind: ChangeKind | undefined = kind;
    if (markRPr) {
      for (const flag of elementChildren(markRPr)) {
        if (isW(flag) && flag.localName! in WRAPPER_KIND) markKind = WRAPPER_KIND[flag.localName!];
      }
    }
    // The engine marks the paragraph mark of a moved paragraph as plainly
    // inserted or deleted. If everything in the paragraph is one move, the mark
    // belongs to that move, so the whole paragraph reads as moved.
    const content = elementChildren(paragraph).filter((c) => isW(c) && (c.localName === 'r' || c.localName! in WRAPPER_KIND || c.localName === 'hyperlink' || c.localName === 'sdt' || c.localName === 'smartTag'));
    if (content.length > 0 && (markKind === 'deleted' || markKind === 'inserted')) {
      const move = markKind === 'deleted' ? 'moveFrom' : 'moveTo';
      if (content.every((c) => isW(c, move))) markKind = WRAPPER_KIND[move];
    }

    // A table cell always keeps its last paragraph mark for as long as the cell
    // exists. So the last mark in a cell is changed only if the cell itself (or
    // its row) was inserted or deleted, whatever the paragraph's own flag says.
    let numberLostAtCellEnd: ChangeKind | undefined;
    let keepsNumber = false;
    const cell = paragraph.parentNode as Element;
    if (!kind && isWord(cell, 'tc') && kids(cell).filter((c) => isWord(c, 'p') || isWord(c, 'tbl')).pop() === paragraph) {
      // The cell's row: its parent, or further up when the cell sits in a content control (a tick-box cell) or custom XML wrapper.
      let row = cell.parentNode as Element | null;
      while (row && isWord(row) && ['sdt', 'sdtContent', 'customXml'].includes(row.localName!)) row = row.parentNode as Element | null;
      const cellKind = this.cellChange.get(cell) ?? (row ? this.cellChange.get(row) : undefined);
      if (!cellKind && markKind) {
        // The cell stays but its last paragraph was added or removed. Since the
        // cell's final mark cannot change, the join or split really happens at
        // the mark of the paragraph before: that is the mark to show as changed.
        const before = kids(cell).filter((c) => isWord(c, 'p') || isWord(c, 'tbl'));
        // Several paragraphs in a row may have been added or removed at the end of the cell (found on a real
        // contract: twelve added after the one that was there). The marks before the last already show the
        // change, so the whole run moves up by one: the mark to show as changed is the one before the run.
        let at = before.length - 2;
        while (at >= 0 && isWord(before[at]!, 'p') && this.markToken.has(before[at]!) && this.tokens[this.markToken.get(before[at]!)!] === markKind) at--;
        const single = at === before.length - 2;
        const removal = markKind === 'deleted' || markKind === 'movedFrom';
        const whollyRemoved = removal && content.length > 0 && content.every((c) => isW(c, markKind === 'deleted' ? 'del' : 'moveFrom'));
        const previous = before[at];
        const index = previous && isWord(previous, 'p') ? this.markToken.get(previous) : undefined;
        // (A run of removed paragraphs ending in one of which part remains is left as it is: which clause number
        // the remaining part should carry cannot be told here, and the Checker judges the result.)
        if (previous && index !== undefined && this.tokens[index] === 'unchanged' && (single || !removal || whollyRemoved || content.length === 0)) {
          applyChangeStyle(markProperties(previous), markKind, this.options);
          this.tokens[index] = markKind;
          // Numbering must follow (found on real contracts: otherwise two paragraphs keep a number where the
          // modified document has one clause, and every later clause is numbered one too high).
          if (removal) {
            if (whollyRemoved || (!single && content.length === 0)) {
              // Nothing of this paragraph remains: it is the clause that was removed, so it is the one that gives up
              // its number (below), and the paragraph before keeps its own.
              numberLostAtCellEnd = markKind;
            } else {
              // Part of this paragraph remains: it is now the tail of the clause before it, which, as with any
              // removed mark, gives up its number to the paragraph it joins. The engine gave the paragraph before
              // the modified document's numbering, so that numbering moves here.
              this.takeNumberingFrom(previous, paragraph);
              const label = this.oldNumberOf.get(previous);
              if (this.options.showOldNumbers && label) this.addOldNumber(previous, label, markKind);
              if (this.options.unnumberRemovedParagraphs) this.removeNumbering(previous);
            }
          }
        }
        markKind = undefined;
      } else if (markKind === 'movedFrom' || markKind === 'movedTo') markKind = cellKind ? markKind : undefined;
      else {
        if (!cellKind) numberLostAtCellEnd = this.removedAtCellEnd.get(paragraph);
        markKind = cellKind;
      }
    } else if (!kind && (markKind === 'deleted' || markKind === 'movedFrom') && this.endsCellWithRemovedClause(paragraph, markKind)) {
      // The same situation as above, written the other way round (found on real contracts): the engine has already
      // put the removal on this mark, the one before the cell's last paragraph, and that last paragraph is removed
      // whole. The clause that went is the last one, so it gives up its number and this paragraph keeps its own.
      this.removedAtCellEnd.set(paragraph.nextSibling as Element, markKind);
      keepsNumber = true;
    }

    if (pPr) this.convert(pPr, kind);
    if (this.options.deletionBeforeInsertion) this.putDeletionsFirst(paragraph);
    for (const child of elementChildren(paragraph)) if (!isW(child, 'pPr')) this.visit(child, kind);

    if (markKind) applyChangeStyle(markProperties(paragraph), markKind, this.options);
    const numberLost = keepsNumber ? undefined : markKind === 'deleted' || markKind === 'movedFrom' ? markKind : numberLostAtCellEnd;
    if (numberLost) {
      const label = this.oldNumberOf.get(paragraph);
      if (this.options.showOldNumbers && label) this.addOldNumber(paragraph, label, numberLost);
      if (this.options.unnumberRemovedParagraphs) this.removeNumbering(paragraph);
    }
    this.markToken.set(paragraph, this.tokens.length);
    this.tokens.push(markKind ?? 'unchanged');
  }
}

/**
 * Count changes in one part, per the definition in PLAN.md: a change is a
 * maximal unbroken stretch, in reading order, of text and paragraph marks of
 * one kind.
 */
function countTokens(tokens: Token[]): Record<ChangeKind, number> {
  const counts: Record<ChangeKind, number> = { inserted: 0, deleted: 0, movedFrom: 0, movedTo: 0 };
  let previous: Token = 'boundary';
  for (const token of tokens) {
    if (token !== previous && token !== 'unchanged' && token !== 'boundary') counts[token]++;
    previous = token;
  }
  return counts;
}

export interface ConvertedPart {
  xml: string;
  counts: Record<ChangeKind, number>;
  formattingChanges: number;
}

/** Convert one Word-markup part (document, header, footer, footnotes, endnotes) from tracked changes to static formatting. */
export function convertPart(xml: string, options: StaticRedlineStyleOptions = DEFAULT_STYLE_OPTIONS, original?: OriginalContext): ConvertedPart {
  const document = new DOMParser().parseFromString(xml, 'application/xml');
  const root = document.documentElement as unknown as Element;
  // These two need the tracked changes still in place, so they come first.
  const numbers = original && original.isMain && options.showOldNumbers ? oldNumbers(root, original) : new Map<Element, string>();
  if (original) stateLookOfJoinedText(root, original.styles);
  if (original && options.markFormattingChanges) markFormattingChanges(root, original);
  const converter = new PartConverter(options, numbers, original ? (paragraph) => original.styles.numPr(paragraph) : undefined);
  converter.convert(root, undefined);
  return { xml: new XMLSerializer().serializeToString(document), counts: countTokens(converter.tokens), formattingChanges: countShadedStretches(root) };
}
