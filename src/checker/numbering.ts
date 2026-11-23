/**
 * The Checker's own working-out of automatic paragraph numbers ("7.3", "(c)"),
 * from the numbering definitions in a Word file. Written separately from the
 * Producer's.
 *
 * Supported number formats: decimal, decimal with leading zero, lower and upper
 * letters, lower and upper roman. Bullets and unsupported formats give no label.
 */
import { NS } from '../common/namespaces';
import type { DocxPackage } from './package';
import { attribute, child, is, type XmlElement } from './xml';

const W = NS.w;

interface Level {
  start: number;
  format: string;
  text: string;
  /** Restart after a level this shallow or shallower appears (undefined: the normal rule). */
  restartAfter?: number;
}

const value = (parent: XmlElement | undefined, name: string): string | undefined => {
  const element = parent && child(parent, W, name);
  return element ? attribute(element, W, 'val') : undefined;
};

function readLevel(lvl: XmlElement, base?: Level): Level {
  const start = value(lvl, 'start');
  const restart = value(lvl, 'lvlRestart');
  return {
    start: start !== undefined ? Number(start) : (base?.start ?? 1),
    format: value(lvl, 'numFmt') ?? base?.format ?? 'decimal',
    text: value(lvl, 'lvlText') ?? base?.text ?? '',
    restartAfter: restart !== undefined ? Number(restart) : base?.restartAfter,
  };
}

function letters(n: number, upper: boolean): string {
  if (n < 1) return '';
  const letter = String.fromCharCode((upper ? 65 : 97) + ((n - 1) % 26));
  return letter.repeat(Math.floor((n - 1) / 26) + 1);
}

function roman(n: number, upper: boolean): string {
  if (n < 1 || n > 3999) return '';
  const table: [number, string][] = [[1000, 'm'], [900, 'cm'], [500, 'd'], [400, 'cd'], [100, 'c'], [90, 'xc'], [50, 'l'], [40, 'xl'], [10, 'x'], [9, 'ix'], [5, 'v'], [4, 'iv'], [1, 'i']];
  let out = '';
  for (const [size, symbol] of table) {
    while (n >= size) {
      out += symbol;
      n -= size;
    }
  }
  return upper ? out.toUpperCase() : out;
}

/** A number in a given format, or undefined if the format is not supported. */
function formatNumber(n: number, format: string): string | undefined {
  switch (format) {
    case 'decimal':
      return String(n);
    case 'decimalZero':
      return n < 10 ? `0${n}` : String(n);
    case 'lowerLetter':
      return letters(n, false);
    case 'upperLetter':
      return letters(n, true);
    case 'lowerRoman':
      return roman(n, false);
    case 'upperRoman':
      return roman(n, true);
    default:
      return undefined;
  }
}

export class Numbering {
  /** abstractNumId -> its levels. */
  private readonly definitions = new Map<string, Map<number, Level>>();
  /** numId -> the definition it uses, with any per-list overrides applied. */
  private readonly lists = new Map<string, { definition: string; levels: Map<number, Level>; restarts: Map<number, number> }>();
  /** Running counters, per definition (lists that share a definition continue one another). */
  private readonly counters = new Map<string, (number | undefined)[]>();
  private readonly started = new Set<string>();

  constructor(pkg: DocxPackage, mainPart: string) {
    const partName = pkg.relationshipsOfType(mainPart, 'numbering')[0]?.partName;
    if (!partName || !pkg.has(partName)) return;
    const root = pkg.xml(partName);
    for (const node of root.children) {
      if (!is(node, W, 'abstractNum')) continue;
      const levels = new Map<number, Level>();
      for (const lvl of node.children) if (is(lvl, W, 'lvl')) levels.set(Number(attribute(lvl, W, 'ilvl') ?? 0), readLevel(lvl));
      this.definitions.set(attribute(node, W, 'abstractNumId') ?? '', levels);
    }
    for (const node of root.children) {
      if (!is(node, W, 'num')) continue;
      const definition = value(node, 'abstractNumId') ?? '';
      const levels = new Map(this.definitions.get(definition) ?? []);
      const restarts = new Map<number, number>();
      for (const override of node.children) {
        if (!is(override, W, 'lvlOverride')) continue;
        const ilvl = Number(attribute(override, W, 'ilvl') ?? 0);
        const lvl = child(override, W, 'lvl');
        if (lvl) levels.set(ilvl, readLevel(lvl, levels.get(ilvl)));
        const startOverride = value(override, 'startOverride');
        if (startOverride !== undefined) restarts.set(ilvl, Number(startOverride));
      }
      this.lists.set(attribute(node, W, 'numId') ?? '', { definition, levels, restarts });
    }
  }

  /** The bullet a list item is printed with ('' when the paragraph is not a bulleted item). Does not advance any counter. */
  bulletOf(numPr: XmlElement | undefined): string {
    const numId = value(numPr, 'numId');
    if (numId === undefined || numId === '0') return '';
    const level = this.lists.get(numId)?.levels.get(Number(value(numPr, 'ilvl') ?? 0));
    return level?.format === 'bullet' ? level.text : '';
  }

  /**
   * The label of the next paragraph carrying this numbering, advancing the
   * counters. '' when the paragraph is not numbered or its label cannot be
   * worked out (bullets, unsupported formats).
   */
  next(numPr: XmlElement | undefined): string {
    const numId = value(numPr, 'numId');
    if (numId === undefined || numId === '0') return '';
    const list = this.lists.get(numId);
    if (!list) return '';
    const ilvl = Number(value(numPr, 'ilvl') ?? 0);
    let counters = this.counters.get(list.definition);
    if (!counters) this.counters.set(list.definition, (counters = []));

    // A list that says "start again at N" does so the first time it is used.
    if (!this.started.has(numId)) {
      this.started.add(numId);
      for (const [level, start] of list.restarts) {
        counters[level] = start - 1;
        for (let deeper = level + 1; deeper < counters.length; deeper++) counters[deeper] = undefined;
      }
    }

    const level = list.levels.get(ilvl);
    if (!level) return '';
    // Shallower levels that have not been used yet count from their start.
    for (let shallower = 0; shallower < ilvl; shallower++) {
      if (counters[shallower] === undefined) counters[shallower] = list.levels.get(shallower)?.start ?? 1;
    }
    counters[ilvl] = (counters[ilvl] ?? level.start - 1) + 1;
    // Deeper levels start again, unless a level says it restarts only after a particular shallower one.
    for (let deeper = ilvl + 1; deeper < counters.length; deeper++) {
      const restartAfter = list.levels.get(deeper)?.restartAfter;
      if (restartAfter === undefined || (restartAfter !== 0 && ilvl < restartAfter)) counters[deeper] = undefined;
    }

    if (level.format === 'bullet' || level.format === 'none') return '';
    let unsupported = false;
    const label = level.text.replace(/%(\d)/g, (_, digit: string) => {
      const index = Number(digit) - 1;
      const formatted = formatNumber(counters![index] ?? list.levels.get(index)?.start ?? 1, list.levels.get(index)?.format ?? 'decimal');
      if (formatted === undefined) unsupported = true;
      return formatted ?? '';
    });
    return unsupported ? '' : label;
  }
}
