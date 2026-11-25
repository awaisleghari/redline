/**
 * Producer: shows a change of wording in a footnote or endnote separator,
 * which the compare engine does not compare (found on a real contract,
 * held-out run 1).
 *
 * The separator is the short rule Word draws between the text and the notes
 * at the foot of a page. A document can type words into it (the contract
 * found used it as a second footer). The engine copies the original's
 * separators into its output without comparing them, so a change of wording
 * there was missing from the redline, and the Checker withheld it.
 *
 * The repair, on the engine's output: where a paragraph of a separator reads
 * differently in the two uploads, it is rewritten as the modified document's
 * paragraph with the old wording shown removed in front of the new wording
 * shown added. The whole paragraph's wording is shown replaced, not the
 * single words that differ: a separator holds a line or two at most. Only
 * plain paragraphs are handled (text, tabs, line breaks and the separator
 * rule itself), and only where both uploads have the same number of
 * paragraphs in the separator. Anything else is left as the engine wrote it,
 * for the Checker to judge.
 */
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { NS } from '../common/namespaces';

const W = NS.w;
type El = Element;
const SPECIAL = ['separator', 'continuationSeparator', 'continuationNotice'];
const IN_A_RUN = new Set(['rPr', 't', 'tab', 'br', 'cr', 'separator', 'continuationSeparator', 'noBreakHyphen', 'softHyphen', 'lastRenderedPageBreak']);
const BESIDE_RUNS = new Set(['bookmarkStart', 'bookmarkEnd', 'proofErr']);

const kids = (e: El): El[] => {
  const out: El[] = [];
  for (let n = e.firstChild; n; n = n.nextSibling) if (n.nodeType === 1) out.push(n as El);
  return out;
};
const isW = (e: El | null | undefined, name: string): e is El => !!e && e.namespaceURI === W && e.localName === name;

/** The separators of a notes part: kind ("separator" and the like) to the note that holds it. */
function separatorsOf(root: El): Map<string, El> {
  const found = new Map<string, El>();
  for (const note of kids(root)) {
    const type = note.namespaceURI === W ? note.getAttributeNS(W, 'type') : null;
    if (type && SPECIAL.includes(type) && !found.has(type)) found.set(type, note);
  }
  return found;
}

/** What a paragraph says: its text, with tabs, breaks and the separator rule each standing as one character. */
function saidBy(paragraph: El): string {
  let said = '';
  const walk = (element: El): void => {
    for (const c of kids(element)) {
      if (isW(c, 't') || isW(c, 'delText')) said += c.textContent ?? '';
      else if (isW(c, 'tab')) said += '\t';
      else if (isW(c, 'br') || isW(c, 'cr')) said += '\n';
      else if (isW(c, 'separator')) said += '\u0001';
      else if (isW(c, 'continuationSeparator')) said += '\u0002';
      else if (!isW(c, 'pPr') && !isW(c, 'rPr')) walk(c);
    }
  };
  walk(paragraph);
  return said;
}

/** A paragraph of runs holding nothing but text, tabs, breaks and the rule. */
const isPlain = (paragraph: El): boolean =>
  kids(paragraph).every((c) => {
    if (c.namespaceURI !== W) return false;
    const name = c.localName!;
    return name === 'pPr' || BESIDE_RUNS.has(name) || (name === 'r' && kids(c).every((x) => x.namespaceURI === W && IN_A_RUN.has(x.localName!)));
  });

const parse = (xml: string): Document => new DOMParser().parseFromString(xml, 'application/xml') as unknown as Document;
/** The separators as text, for the quick test "is there anything to do". */
const quick = (xml: string): string => (xml.match(/<w:(?:foot|end)note\b[^>]*\bw:type="(?:separator|continuationSeparator|continuationNotice)"[^>]*>[\s\S]*?<\/w:(?:foot|end)note>/g) ?? []).map((note) => note.replace(/<w:tab\b[^>]*>/g, '\t').replace(/<[^>]+>/g, '')).join('\u0000');

/**
 * The engine's notes part (footnotes or endnotes) with changed separator wording shown, or undefined if there is
 * nothing to show (or it cannot be shown safely).
 */
export function repairSeparators(tracked: string, original: string, modified: string, date: string): string | undefined {
  if (quick(original) === quick(modified)) return undefined;
  const document = parse(tracked);
  const root = document.documentElement as unknown as El;
  const here = separatorsOf(root);
  const before = separatorsOf(parse(original).documentElement as unknown as El);
  const after = separatorsOf(parse(modified).documentElement as unknown as El);
  const prefix = root.prefix ? `${root.prefix}:` : '';
  let nextId = 1;
  for (const match of tracked.matchAll(/\bw:id="(\d+)"/g)) nextId = Math.max(nextId, Number(match[1]) + 1);
  const wrapper = (name: 'ins' | 'del'): El => {
    const element = document.createElementNS(W, `${prefix}${name}`) as unknown as El;
    element.setAttributeNS(W, `${prefix}id`, String(nextId++));
    element.setAttributeNS(W, `${prefix}author`, 'Redline');
    element.setAttributeNS(W, `${prefix}date`, date);
    return element;
  };
  let changed = false;
  for (const [type, note] of here) {
    const oldNote = before.get(type);
    const newNote = after.get(type);
    if (!oldNote || !newNote) continue;
    const blocks = [note, oldNote, newNote].map((n) => kids(n));
    // Paragraphs only, the same number in each, and the engine's copy still reading as the original does.
    if (!blocks.every((list) => list.length === blocks[0]!.length && list.every((b) => isW(b, 'p')))) continue;
    const [mine, olds, news] = blocks as [El[], El[], El[]];
    if (!mine.every((p, i) => saidBy(p) === saidBy(olds[i]!))) continue;
    const differing = olds.map((p, i) => saidBy(p) !== saidBy(news[i]!));
    if (!differing.some(Boolean) || !olds.every((p, i) => !differing[i] || (isPlain(p) && isPlain(news[i]!)))) continue;
    differing.forEach((differs, i) => {
      if (!differs) return;
      const paragraph = document.importNode(news[i]! as unknown as Node, true) as unknown as El;
      const added = wrapper('ins');
      const removed = wrapper('del');
      for (const run of kids(paragraph).filter((c) => isW(c, 'r'))) added.appendChild(run as unknown as Node);
      for (const run of kids(olds[i]!).filter((c) => isW(c, 'r'))) {
        const copy = document.importNode(run as unknown as Node, true) as unknown as El;
        for (const text of kids(copy).filter((c) => isW(c, 't'))) {
          const gone = document.createElementNS(W, `${prefix}delText`) as unknown as El;
          for (let a = 0; a < text.attributes.length; a++) gone.setAttribute(text.attributes.item(a)!.name, text.attributes.item(a)!.value);
          while (text.firstChild) gone.appendChild(text.firstChild);
          copy.replaceChild(gone as unknown as Node, text as unknown as Node);
        }
        removed.appendChild(copy as unknown as Node);
      }
      const properties = kids(paragraph).find((c) => isW(c, 'pPr'));
      const first = properties ? properties.nextSibling : paragraph.firstChild;
      if (added.firstChild) paragraph.insertBefore(added as unknown as Node, first);
      if (removed.firstChild) paragraph.insertBefore(removed as unknown as Node, added.parentNode ? (added as unknown as Node) : first);
      note.replaceChild(paragraph as unknown as Node, mine[i]! as unknown as Node);
      changed = true;
    });
  }
  return changed ? new XMLSerializer().serializeToString(document as never) : undefined;
}
