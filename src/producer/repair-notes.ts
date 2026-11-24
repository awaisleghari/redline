/**
 * Producer: repairs a fault in the compare engine's handling of footnotes and
 * endnotes (found on real contracts, Stage 7).
 *
 * When anything changes in a table (a row added or removed, a word changed in
 * another cell), the engine re-issues every note referred to from that table:
 * it marks the note's whole text as inserted, and adds a second note holding
 * the same text marked as deleted, which nothing in the document refers to.
 * Accepting every change still gives the modified document, but rejecting
 * them empties the note, and a reader sees an unchanged footnote as struck out
 * and re-typed.
 *
 * The repair, on the engine's output: such a stray deleted note is paired with
 * the wholly inserted note it belongs to. If their text is the same, the note
 * did not change: it is restored to plain text and the stray copy is removed.
 * If the text differs (the note really was edited as well), the deleted text
 * is put into the note in front of the inserted text, so the note shows both,
 * and the stray copy is removed. The Checker judges the result as it judges
 * everything else.
 */
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { NS } from '../common/namespaces';

const W = NS.w;
type El = Element;
const CHANGE_WRAPPERS = new Set(['ins', 'del', 'moveFrom', 'moveTo']);

const children = (e: El, name?: string): El[] => {
  const out: El[] = [];
  for (let n = e.firstChild; n; n = n.nextSibling) if (n.nodeType === 1 && (!name || ((n as El).namespaceURI === W && (n as El).localName === name))) out.push(n as El);
  return out;
};
const descendants = (e: El, name: string): El[] => {
  const list = e.getElementsByTagNameNS(W, name);
  const out: El[] = [];
  for (let i = 0; i < list.length; i++) out.push(list.item(i) as unknown as El);
  return out;
};
/** The nearest change wrapper around an element, inside `within`. */
function wrapperOf(element: El, within: El): string | undefined {
  for (let n: Node | null = element.parentNode; n && n !== within; n = n.parentNode) {
    const e = n as El;
    if (e.namespaceURI === W && CHANGE_WRAPPERS.has(e.localName!)) return e.localName!;
  }
  return undefined;
}
/** A note's text and whether every run of it lies in one kind of change (ignoring the note's own number mark when it stands alone). */
function describe(note: El): { text: string; all?: 'inserted' | 'deleted' } {
  let text = '';
  const kinds = new Set<string>();
  for (const run of descendants(note, 'r')) {
    const wrapper = wrapperOf(run, note);
    kinds.add(wrapper === 'ins' || wrapper === 'moveTo' ? 'inserted' : wrapper === 'del' || wrapper === 'moveFrom' ? 'deleted' : 'plain');
    for (const t of [...children(run, 't'), ...children(run, 'delText')]) text += t.textContent ?? '';
  }
  const only = kinds.size === 1 ? [...kinds][0] : undefined;
  return { text, all: only === 'inserted' || only === 'deleted' ? only : undefined };
}

/**
 * `parts`: the Word-markup parts of the engine's output, by name. Returns the parts that were changed.
 */
export function repairNotes(parts: Map<string, string>): Map<string, string> {
  const changed = new Map<string, string>();
  for (const kind of ['footnote', 'endnote'] as const) {
    const partName = [...parts.keys()].find((name) => new RegExp(`(^|/)${kind}s\\.xml$`).test(name));
    if (!partName) continue;
    const notesXml = parts.get(partName)!;
    // Quick test first: is there a note that nothing refers to? (Almost always not.)
    const referenced = new Set<string>();
    const reference = new RegExp(`<w:${kind}Reference\\b[^>]*\\bw:id="(-?\\d+)"`, 'g');
    for (const [name, xml] of parts) if (name !== partName) for (const match of xml.matchAll(reference)) referenced.add(match[1]!);
    const ids = [...notesXml.matchAll(new RegExp(`<w:${kind}\\b(?![^>]*w:type="(?:separator|continuationSeparator|continuationNotice)")[^>]*\\bw:id="(-?\\d+)"`, 'g'))].map((m) => m[1]!);
    if (ids.every((id) => referenced.has(id))) continue;

    const doc = new DOMParser().parseFromString(notesXml, 'application/xml');
    const root = doc.documentElement as unknown as El;
    const notes = children(root, kind).filter((note) => !['separator', 'continuationSeparator', 'continuationNotice'].includes(note.getAttributeNS(W, 'type') ?? ''));
    const idOf = (note: El): string => note.getAttributeNS(W, 'id') ?? '';
    const strays = notes.filter((note) => !referenced.has(idOf(note)) && describe(note).all === 'deleted');
    if (strays.length === 0) continue;

    // Notes that are wholly inserted although the mark that refers to them is not itself an insertion.
    const insertedReferences = new Set<string>();
    for (const [name, xml] of parts) {
      if (name === partName || !xml.includes(`<w:${kind}Reference`)) continue;
      const body = new DOMParser().parseFromString(xml, 'application/xml').documentElement as unknown as El;
      for (const mark of descendants(body, `${kind}Reference`)) {
        const wrapper = wrapperOf(mark, body);
        if (wrapper === 'ins' || wrapper === 'moveTo') insertedReferences.add(mark.getAttributeNS(W, 'id') ?? '');
      }
    }
    const reissued = notes.filter((note) => referenced.has(idOf(note)) && !insertedReferences.has(idOf(note)) && describe(note).all === 'inserted');
    if (reissued.length === 0) continue;

    // Pair each stray with its note: the one with the same text if there is one, otherwise in order.
    const pairs: [El, El][] = [];
    const free = [...reissued];
    const unpaired: El[] = [];
    for (const stray of strays) {
      const text = describe(stray).text;
      const at = free.findIndex((note) => describe(note).text === text);
      if (at >= 0) pairs.push([stray, free.splice(at, 1)[0]!]);
      else unpaired.push(stray);
    }
    for (const stray of unpaired) if (free.length > 0) pairs.push([stray, free.shift()!]);

    for (const [stray, note] of pairs) {
      if (describe(stray).text === describe(note).text) {
        // Nothing changed in this note: make it plain again.
        for (const wrapper of [...descendants(note, 'ins'), ...descendants(note, 'moveTo')]) {
          const parent = wrapper.parentNode!;
          if ((parent as El).namespaceURI === W && (parent as El).localName === 'rPr') parent.removeChild(wrapper); // the flag on a paragraph mark
          else {
            while (wrapper.firstChild) parent.insertBefore(wrapper.firstChild, wrapper);
            parent.removeChild(wrapper);
          }
        }
      } else {
        // The note was really edited too: show the old text, deleted, in front of the new.
        const first = children(note, 'p')[0] ?? null;
        for (const paragraph of children(stray)) note.insertBefore(paragraph, first);
      }
      root.removeChild(stray);
    }
    if (pairs.length > 0) changed.set(partName, new XMLSerializer().serializeToString(doc));
  }
  return changed;
}
