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

const W = NS.w;

export type ChangeKind = 'inserted' | 'deleted' | 'movedFrom' | 'movedTo';

export interface ChangeCounts {
  insertions: number;
  deletions: number;
  /** Counted once per block moved away from its original place. */
  moves: number;
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
const PROPERTY_CHANGE_RECORDS = new Set(['rPrChange', 'pPrChange', 'sectPrChange', 'tblPrChange', 'tblPrExChange', 'tcPrChange', 'trPrChange', 'tblGridChange', 'numberingChange']);
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
function applyChangeStyle(rPr: Element, kind: ChangeKind): void {
  const style = CHANGE_STYLE[kind];
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
}

export const DEFAULT_STYLE_OPTIONS: StaticRedlineStyleOptions = { deletionBeforeInsertion: true, unnumberRemovedParagraphs: true };

/** Paragraph properties that come before w:numPr in Word's schema order. */
const PPR_BEFORE_NUMPR = new Set(['pStyle', 'keepNext', 'keepLines', 'pageBreakBefore', 'framePr', 'widowControl']);

class PartConverter {
  readonly tokens: Token[] = [];

  constructor(private readonly options: StaticRedlineStyleOptions) {}

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
    if (kind) applyChangeStyle(runProperties(run), kind);
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

    if (pPr) this.convert(pPr, kind);
    if (this.options.deletionBeforeInsertion) this.putDeletionsFirst(paragraph);
    for (const child of elementChildren(paragraph)) if (!isW(child, 'pPr')) this.visit(child, kind);

    if (markKind) applyChangeStyle(markProperties(paragraph), markKind);
    if (this.options.unnumberRemovedParagraphs && (markKind === 'deleted' || markKind === 'movedFrom')) this.removeNumbering(paragraph);
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
}

/** Convert one Word-markup part (document, header, footer, footnotes, endnotes) from tracked changes to static formatting. */
export function convertPart(xml: string, options: StaticRedlineStyleOptions = DEFAULT_STYLE_OPTIONS): ConvertedPart {
  const document = new DOMParser().parseFromString(xml, 'application/xml');
  const root = document.documentElement as unknown as Element;
  const converter = new PartConverter(options);
  converter.convert(root, undefined);
  return { xml: new XMLSerializer().serializeToString(document), counts: countTokens(converter.tokens) };
}
