/**
 * The Checker's own reading of character formatting: what a reader would see
 * for a piece of text (bold, italic, underlined, capitals, hidden), resolved
 * through the document's styles; and, for a static redline, which kind of
 * change a piece of text is marked as.
 */
import { NS } from '../common/namespaces';
import type { DocxPackage } from './package';
import { attribute, child, is, type XmlElement } from './xml';

const W = NS.w;

/** One bit per property a reader can see. */
export const FORMAT = { bold: 1, italic: 2, underline: 4, caps: 8, smallCaps: 16, hidden: 32, strike: 64, doubleStrike: 128 } as const;
/** All formatting bits together. */
export const FORMAT_MASK = 255;
export type FormatBits = number;

export const FORMAT_NAMES: [number, string][] = [
  [FORMAT.bold, 'bold'],
  [FORMAT.italic, 'italic'],
  [FORMAT.underline, 'underline'],
  [FORMAT.caps, 'capitals'],
  [FORMAT.smallCaps, 'small capitals'],
  [FORMAT.hidden, 'hidden text'],
  [FORMAT.strike, 'struck through'],
  [FORMAT.doubleStrike, 'double struck through'],
];

/** "bold, italic" for a set of bits; "plain" for none. */
export function describeFormat(bits: FormatBits): string {
  const names = FORMAT_NAMES.filter(([bit]) => bits & bit).map(([, name]) => name);
  return names.length > 0 ? names.join(', ') : 'plain';
}

const PROPERTY: Record<string, number> = { b: FORMAT.bold, i: FORMAT.italic, u: FORMAT.underline, caps: FORMAT.caps, smallCaps: FORMAT.smallCaps, vanish: FORMAT.hidden, strike: FORMAT.strike, dstrike: FORMAT.doubleStrike };

/** Is an on/off property element switched on? (No value means on.) */
export function isOn(element: XmlElement): boolean {
  const value = attribute(element, W, 'val');
  if (element.name === 'u') return value !== 'none';
  return value === undefined || !['0', 'false', 'off'].includes(value);
}

/** What one set of run properties says explicitly: which properties it sets, and to what. */
interface Stated {
  mask: number;
  value: number;
}

function stated(rPr: XmlElement | undefined): Stated {
  let mask = 0;
  let value = 0;
  for (const property of rPr?.children ?? []) {
    if (property.ns !== W) continue;
    const bit = PROPERTY[property.name];
    if (!bit) continue;
    mask |= bit;
    if (isOn(property)) value |= bit;
    else value &= ~bit;
  }
  return { mask, value };
}

interface Style {
  type: string;
  basedOn?: string;
  stated: Stated;
  numbering?: XmlElement;
  /** Whether the style sets a tab stop that draws a leader (dots, a line). */
  leader: boolean;
}

/** Whether paragraph properties set a tab stop that draws a leader. */
function setsLeader(pPr: XmlElement | undefined): boolean {
  const tabs = pPr && child(pPr, W, 'tabs');
  return !!tabs && tabs.children.some((tab) => is(tab, W, 'tab') && !['none', undefined].includes(attribute(tab, W, 'leader')));
}

export class Styles {
  private readonly styles = new Map<string, Style>();
  private defaults: Stated = { mask: 0, value: 0 };
  private defaultParagraphStyle: string | undefined;

  constructor(pkg: DocxPackage, mainPart: string) {
    const partName = pkg.relationshipsOfType(mainPart, 'styles')[0]?.partName;
    if (!partName || !pkg.has(partName)) return;
    for (const node of pkg.xml(partName).children) {
      if (is(node, W, 'docDefaults')) {
        const rPrDefault = child(node, W, 'rPrDefault');
        this.defaults = stated(rPrDefault && child(rPrDefault, W, 'rPr'));
      } else if (is(node, W, 'style')) {
        const id = attribute(node, W, 'styleId') ?? '';
        const type = attribute(node, W, 'type') ?? 'paragraph';
        const basedOn = child(node, W, 'basedOn');
        const pPr = child(node, W, 'pPr');
        this.styles.set(id, {
          type,
          basedOn: basedOn && attribute(basedOn, W, 'val'),
          stated: stated(child(node, W, 'rPr')),
          numbering: pPr && child(pPr, W, 'numPr'),
          leader: setsLeader(pPr),
        });
        const isDefault = attribute(node, W, 'default');
        if (type === 'paragraph' && isDefault !== undefined && !['0', 'false', 'off'].includes(isDefault)) this.defaultParagraphStyle = id;
      }
    }
  }

  /** Does the document's own style sheet define this style? */
  defines(id: string): boolean {
    return this.styles.has(id);
  }

  /** A style and everything it is based on, nearest first. */
  private chain(id: string | undefined): Style[] {
    const result: Style[] = [];
    const seen = new Set<string>();
    while (id !== undefined && !seen.has(id)) {
      seen.add(id);
      const style = this.styles.get(id);
      if (!style) break;
      result.push(style);
      id = style.basedOn;
    }
    return result;
  }

  /**
   * The formatting a reader sees for a run: its own properties first, then its
   * character style, then its paragraph's style, then the document defaults.
   */
  formatOf(runProperties: XmlElement | undefined, paragraphStyle: string | undefined): FormatBits {
    const characterStyle = runProperties && child(runProperties, W, 'rStyle');
    const layers: Stated[] = [
      stated(runProperties),
      ...this.chain(characterStyle && attribute(characterStyle, W, 'val')).map((s) => s.stated),
      ...this.chain(paragraphStyle ?? this.defaultParagraphStyle).map((s) => s.stated),
      this.defaults,
    ];
    let decided = 0;
    let bits = 0;
    for (const layer of layers) {
      const fresh = layer.mask & ~decided;
      bits |= layer.value & fresh;
      decided |= layer.mask;
    }
    return bits;
  }

  /** The numbering a paragraph has: its own, or else its style's. Undefined when it has none. */
  numberingOf(paragraphProperties: XmlElement | undefined): XmlElement | undefined {
    const direct = paragraphProperties && child(paragraphProperties, W, 'numPr');
    if (direct) return direct;
    const style = paragraphProperties && child(paragraphProperties, W, 'pStyle');
    for (const s of this.chain((style && attribute(style, W, 'val')) ?? this.defaultParagraphStyle)) if (s.numbering) return s.numbering;
    return undefined;
  }

  /** Whether a tab in this paragraph may draw a leader: one of its own tab stops, or one of its style's, has a leader. */
  hasLeaderTab(paragraphProperties: XmlElement | undefined): boolean {
    if (setsLeader(paragraphProperties)) return true;
    const style = paragraphProperties && child(paragraphProperties, W, 'pStyle');
    return this.chain((style && attribute(style, W, 'val')) ?? this.defaultParagraphStyle).some((s) => s.leader);
  }
}

export type StaticKind = 'inserted' | 'deleted' | 'movedFrom' | 'movedTo';

/**
 * The colour of inserted text. Blue, unless one of the two documents itself contains text formatted
 * blue with a double underline: a reader could not then tell the document's own text from an insertion,
 * so the redline must use the alternative (a dark magenta) and say so in its legend (owner's decision,
 * 2026-11-24).
 */
export const INSERTED_COLOUR = { standard: '0000FF', alternative: 'AA00AA' } as const;

/**
 * The colour of deleted text, on the same rule (owner's rule for insertions, applied to deletions; found on a real
 * contract whose own text held a paragraph mark formatted red and struck through): red, unless one of the two
 * documents itself contains text formatted red with a strikethrough. Then deletions are shown in a dark orange,
 * and the legend says so.
 */
export const DELETED_COLOUR = { standard: 'FF0000', alternative: 'C55A11' } as const;

/** The colours a static redline shows insertions and deletions in. */
export interface ChangeColours {
  inserted: string;
  deleted: string;
}
export const STANDARD_COLOURS: ChangeColours = { inserted: INSERTED_COLOUR.standard, deleted: DELETED_COLOUR.standard };

/** Whether a document's own text or paragraph marks carry the standard look of the given kind of change. */
function usesLook(root: XmlElement, kind: StaticKind): boolean {
  for (const node of [root, ...descendantsOf(root)]) {
    if (node.ns !== W || (node.name !== 'r' && node.name !== 'pPr')) continue;
    if (staticKindOf(child(node, W, 'rPr'), STANDARD_COLOURS) === kind) return true;
  }
  return false;
}
/** Whether a document's own text or paragraph marks carry the standard look of an insertion. */
export const usesInsertionLook = (root: XmlElement): boolean => usesLook(root, 'inserted');
/** Whether a document's own text or paragraph marks carry the standard look of a deletion. */
export const usesDeletionLook = (root: XmlElement): boolean => usesLook(root, 'deleted');
function* descendantsOf(element: XmlElement): Generator<XmlElement> {
  for (const c of element.children) {
    yield c;
    yield* descendantsOf(c);
  }
}

/**
 * Which change, if any, a set of run properties marks in a static redline,
 * per the style spec in PLAN.md. Read from the formatting alone.
 */
export function staticKindOf(rPr: XmlElement | undefined, colours: ChangeColours = STANDARD_COLOURS): StaticKind | undefined {
  if (!rPr) return undefined;
  const colorElement = child(rPr, W, 'color');
  // A theme colour takes precedence over the stated one in Word, so the text would not show in the change colour.
  const themed = !!colorElement && attribute(colorElement, W, 'themeColor') !== undefined;
  const color = themed ? undefined : (colorElement && attribute(colorElement, W, 'val'))?.toUpperCase();
  const on = (name: string): boolean => {
    const element = child(rPr, W, name);
    return !!element && isOn(element);
  };
  const underline = child(rPr, W, 'u');
  const doubleUnderline = !!underline && attribute(underline, W, 'val') === 'double';
  if (color === colours.deleted && on('strike')) return 'deleted';
  if (color === '008000' && on('dstrike')) return 'movedFrom';
  if (color === colours.inserted && doubleUnderline) return 'inserted';
  if (color === '008000' && doubleUnderline) return 'movedTo';
  return undefined;
}
