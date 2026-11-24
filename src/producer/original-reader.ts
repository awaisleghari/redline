/**
 * Producer: what the static redline needs to know about the ORIGINAL document
 * that the tracked-changes file no longer says reliably:
 *   - the formatting each character had (to shade text whose formatting changed);
 *   - the automatic number each paragraph had (to show it on removed clauses).
 *
 * This is the Producer's own reading, written separately from the Checker's.
 */
import { DOMParser } from '@xmldom/xmldom';
import { NS } from '../common/namespaces';

const W = NS.w;
const XML_NS = 'http://www.w3.org/XML/1998/namespace';

type El = Element;

export const isWord = (node: Node | null | undefined, name?: string): node is El =>
  !!node && node.nodeType === 1 && (node as El).namespaceURI === W && (name === undefined || (node as El).localName === name);

export function kids(parent: El): El[] {
  const out: El[] = [];
  for (let n = parent.firstChild; n; n = n.nextSibling) if (n.nodeType === 1) out.push(n as El);
  return out;
}

export const kid = (parent: El | undefined, name: string): El | undefined => (parent ? kids(parent).find((c) => isWord(c, name)) : undefined);
const val = (element: El | undefined): string | undefined => element?.getAttributeNS(W, 'val') ?? undefined;
const valOf = (parent: El | undefined, name: string): string | undefined => {
  const v = val(kid(parent, name));
  return v === '' ? undefined : v;
};

export function parse(xml: string): El {
  return new DOMParser().parseFromString(xml, 'application/xml').documentElement as unknown as El;
}

// ---------------------------------------------------------------------------
// Visible formatting
// ---------------------------------------------------------------------------

/** The properties a formatting change is judged on, as bits. */
export const BIT: Record<string, number> = { b: 1, i: 2, u: 4, caps: 8, smallCaps: 16, vanish: 32, strike: 64, dstrike: 128 };

interface Layer {
  set: number;
  on: number;
}

function layerOf(rPr: El | undefined): Layer {
  const layer: Layer = { set: 0, on: 0 };
  if (!rPr) return layer;
  for (const property of kids(rPr)) {
    if (!isWord(property)) continue;
    const bit = BIT[property.localName!];
    if (!bit) continue;
    const v = val(property);
    const on = property.localName === 'u' ? v !== 'none' : v === undefined || v === '' || !['0', 'false', 'off'].includes(v);
    layer.set |= bit;
    layer.on = on ? layer.on | bit : layer.on & ~bit;
  }
  return layer;
}

export class StyleSheet {
  private readonly byId = new Map<string, { basedOn?: string; layer: Layer; numPr?: El }>();
  private defaults: Layer = { set: 0, on: 0 };
  private defaultParagraph: string | undefined;

  constructor(stylesXml: string | undefined) {
    if (!stylesXml) return;
    for (const node of kids(parse(stylesXml))) {
      if (isWord(node, 'docDefaults')) this.defaults = layerOf(kid(kid(node, 'rPrDefault'), 'rPr'));
      if (!isWord(node, 'style')) continue;
      const id = node.getAttributeNS(W, 'styleId') ?? '';
      this.byId.set(id, { basedOn: valOf(node, 'basedOn'), layer: layerOf(kid(node, 'rPr')), numPr: kid(kid(node, 'pPr'), 'numPr') });
      const isDefault = node.getAttributeNS(W, 'default');
      if ((node.getAttributeNS(W, 'type') ?? 'paragraph') === 'paragraph' && isDefault && !['0', 'false', 'off'].includes(isDefault)) this.defaultParagraph = id;
    }
  }

  private lineage(id: string | undefined): { layer: Layer; numPr?: El }[] {
    const out: { layer: Layer; numPr?: El }[] = [];
    const seen = new Set<string>();
    for (let current = id; current !== undefined && !seen.has(current); ) {
      seen.add(current);
      const style = this.byId.get(current);
      if (!style) break;
      out.push(style);
      current = style.basedOn;
    }
    return out;
  }

  /** Bits for a run: its own properties, then its character style, its paragraph's style, the document defaults. */
  bits(rPr: El | undefined, paragraph: El | undefined): number {
    const paragraphStyle = valOf(kid(paragraph, 'pPr'), 'pStyle') ?? this.defaultParagraph;
    const layers = [layerOf(rPr), ...this.lineage(valOf(rPr, 'rStyle')).map((s) => s.layer), ...this.lineage(paragraphStyle).map((s) => s.layer), this.defaults];
    let known = 0;
    let bits = 0;
    for (const layer of layers) {
      bits |= layer.on & layer.set & ~known;
      known |= layer.set;
    }
    return bits;
  }

  /** A paragraph's numbering: its own, or its style's. */
  numPr(paragraph: El): El | undefined {
    const pPr = kid(paragraph, 'pPr');
    const own = kid(pPr, 'numPr');
    if (own) return own;
    return this.lineage(valOf(pPr, 'pStyle') ?? this.defaultParagraph).find((s) => s.numPr)?.numPr;
  }
}

/** A text element's characters as Word shows them (outer whitespace dropped unless preserved). */
export function shownText(element: El): string {
  const raw = element.textContent ?? '';
  return keepsSpace(element) ? raw : raw.replace(/^[ \t\r\n]+|[ \t\r\n]+$/g, '');
}

/**
 * Whether white space in an element is kept: xml:space on the element itself or, failing that, on the
 * nearest element around it that sets it (some programs set it once for the whole document).
 */
export function keepsSpace(element: El): boolean {
  for (let node: El | null = element; node && node.nodeType === 1; node = node.parentNode as El | null) {
    const setting = node.getAttributeNS(XML_NS, 'space');
    if (setting) return setting === 'preserve';
  }
  return false;
}

function nearestParagraph(node: El): El | undefined {
  for (let n: Node | null = node.parentNode; n; n = n.parentNode) if (isWord(n, 'p')) return n;
  return undefined;
}

export interface TextStream {
  text: string;
  /** Formatting bits, one per character of `text`. */
  bits: number[];
}

/** Every character of Word text under `root`, in order, with its formatting. For a document without tracked changes. */
export function plainStream(root: El, styles: StyleSheet): TextStream {
  const stream: TextStream = { text: '', bits: [] };
  const runs = root.getElementsByTagNameNS(W, 'r');
  for (let i = 0; i < runs.length; i++) {
    const run = runs.item(i) as El;
    const bits = styles.bits(kid(run, 'rPr'), nearestParagraph(run));
    for (const child of kids(run)) {
      if (!isWord(child, 't') && !isWord(child, 'delText')) continue;
      const text = shownText(child);
      stream.text += text;
      for (let k = 0; k < text.length; k++) stream.bits.push(bits);
    }
  }
  return stream;
}

// ---------------------------------------------------------------------------
// Automatic numbers
// ---------------------------------------------------------------------------

interface LevelDef {
  start: number;
  format: string;
  text: string;
  restart?: number;
}

function levelDef(lvl: El, inherit?: LevelDef): LevelDef {
  const start = valOf(lvl, 'start');
  const restart = valOf(lvl, 'lvlRestart');
  return {
    start: start === undefined ? (inherit?.start ?? 1) : parseInt(start, 10),
    format: valOf(lvl, 'numFmt') ?? inherit?.format ?? 'decimal',
    text: val(kid(lvl, 'lvlText')) ?? inherit?.text ?? '',
    restart: restart === undefined ? inherit?.restart : parseInt(restart, 10),
  };
}

const ROMAN: [number, string][] = [[1000, 'm'], [900, 'cm'], [500, 'd'], [400, 'cd'], [100, 'c'], [90, 'xc'], [50, 'l'], [40, 'xl'], [10, 'x'], [9, 'ix'], [5, 'v'], [4, 'iv'], [1, 'i']];

function render(n: number, format: string): string | null {
  const alpha = (base: number): string => (n < 1 ? '' : String.fromCharCode(base + ((n - 1) % 26)).repeat(1 + Math.floor((n - 1) / 26)));
  const roman = (): string => {
    if (n < 1 || n > 3999) return '';
    let rest = n;
    let out = '';
    for (const [size, symbol] of ROMAN) for (; rest >= size; rest -= size) out += symbol;
    return out;
  };
  if (format === 'decimal') return `${n}`;
  if (format === 'decimalZero') return n < 10 ? `0${n}` : `${n}`;
  if (format === 'lowerLetter') return alpha(97);
  if (format === 'upperLetter') return alpha(65);
  if (format === 'lowerRoman') return roman();
  if (format === 'upperRoman') return roman().toUpperCase();
  return null;
}

/** Works out automatic paragraph numbers the way Word counts them. Bullets and unsupported formats give ''. */
export class NumberCounter {
  private readonly abstract = new Map<string, Map<number, LevelDef>>();
  private readonly instances = new Map<string, { abstractId: string; levels: Map<number, LevelDef>; startAt: Map<number, number> }>();
  private readonly state = new Map<string, (number | undefined)[]>();
  private readonly used = new Set<string>();

  constructor(numberingXml: string | undefined) {
    if (!numberingXml) return;
    const all = kids(parse(numberingXml));
    for (const node of all) {
      if (!isWord(node, 'abstractNum')) continue;
      const levels = new Map<number, LevelDef>();
      for (const lvl of kids(node)) if (isWord(lvl, 'lvl')) levels.set(parseInt(lvl.getAttributeNS(W, 'ilvl') ?? '0', 10), levelDef(lvl));
      this.abstract.set(node.getAttributeNS(W, 'abstractNumId') ?? '', levels);
    }
    for (const node of all) {
      if (!isWord(node, 'num')) continue;
      const abstractId = valOf(node, 'abstractNumId') ?? '';
      const levels = new Map(this.abstract.get(abstractId) ?? []);
      const startAt = new Map<number, number>();
      for (const override of kids(node)) {
        if (!isWord(override, 'lvlOverride')) continue;
        const ilvl = parseInt(override.getAttributeNS(W, 'ilvl') ?? '0', 10);
        const lvl = kid(override, 'lvl');
        if (lvl) levels.set(ilvl, levelDef(lvl, levels.get(ilvl)));
        const startOverride = valOf(override, 'startOverride');
        if (startOverride !== undefined) startAt.set(ilvl, parseInt(startOverride, 10));
      }
      this.instances.set(node.getAttributeNS(W, 'numId') ?? '', { abstractId, levels, startAt });
    }
  }

  label(numPr: El | undefined): string {
    const numId = valOf(numPr, 'numId');
    if (!numId || numId === '0') return '';
    const instance = this.instances.get(numId);
    if (!instance) return '';
    const ilvl = parseInt(valOf(numPr, 'ilvl') ?? '0', 10);
    const counters = this.state.get(instance.abstractId) ?? [];
    this.state.set(instance.abstractId, counters);
    if (!this.used.has(numId)) {
      this.used.add(numId);
      instance.startAt.forEach((start, level) => {
        counters[level] = start - 1;
        for (let d = level + 1; d < counters.length; d++) counters[d] = undefined;
      });
    }
    const level = instance.levels.get(ilvl);
    if (!level) return '';
    for (let s = 0; s < ilvl; s++) counters[s] ??= instance.levels.get(s)?.start ?? 1;
    counters[ilvl] = (counters[ilvl] ?? level.start - 1) + 1;
    for (let d = ilvl + 1; d < counters.length; d++) {
      const restart = instance.levels.get(d)?.restart;
      if (restart === undefined || (restart !== 0 && ilvl < restart)) counters[d] = undefined;
    }
    if (level.format === 'bullet' || level.format === 'none') return '';
    let ok = true;
    const label = level.text.replace(/%(\d)/g, (_match, digit: string) => {
      const index = parseInt(digit, 10) - 1;
      const piece = render(counters[index] ?? instance.levels.get(index)?.start ?? 1, instance.levels.get(index)?.format ?? 'decimal');
      if (piece === null) ok = false;
      return piece ?? '';
    });
    return ok ? label : '';
  }
}

/**
 * Paragraphs in the order their numbers are counted: a paragraph's own text
 * boxes first, then the paragraph itself.
 */
export function paragraphsInCountingOrder(root: El, include: (paragraph: El) => boolean = () => true, enter: (element: El) => boolean = () => true): El[] {
  const out: El[] = [];
  const walk = (element: El): void => {
    for (const child of kids(element)) {
      if (!enter(child)) continue;
      walk(child);
      if (isWord(child, 'p') && include(child)) out.push(child);
    }
  };
  walk(root);
  return out;
}

export interface OriginalInfo {
  /** Text and formatting of each unit of the original: the body, each header and footer, each footnote and endnote. */
  streams: { kind: string; stream: TextStream; used: boolean }[];
  /** The automatic number of each body paragraph of the original, in counting order ('' when not numbered). */
  bodyNumbers: string[];
}

/** Units of a part: the whole part, or each note of a notes part. */
export function unitsOf(root: El): El[] {
  if (root.localName === 'footnotes' || root.localName === 'endnotes') return kids(root).filter((c) => isWord(c, 'footnote') || isWord(c, 'endnote'));
  return [root];
}

export function readOriginal(parts: { xml: string; isMain: boolean }[], stylesXml: string | undefined, numberingXml: string | undefined): OriginalInfo {
  const styles = new StyleSheet(stylesXml);
  const info: OriginalInfo = { streams: [], bodyNumbers: [] };
  for (const part of parts) {
    const root = parse(part.xml);
    for (const unit of unitsOf(root)) info.streams.push({ kind: root.localName!, stream: plainStream(unit, styles), used: false });
    if (part.isMain) {
      const counter = new NumberCounter(numberingXml);
      info.bodyNumbers = paragraphsInCountingOrder(root).map((paragraph) => counter.label(styles.numPr(paragraph)));
    }
  }
  return info;
}
