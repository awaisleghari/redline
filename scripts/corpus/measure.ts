/**
 * Stage 7 tooling: measures a redline against the right answer.
 *
 *  - How many words really differ between two documents (the smallest possible
 *    set of inserted and deleted words, part by part).
 *  - How many words a tracked-changes file marks as inserted, deleted or moved.
 *
 * A perfect redline marks exactly the words that differ. The ratio of the two
 * numbers says how far a redline is from that.
 */
import { DOMParser } from '@xmldom/xmldom';
import { unzip } from '../../src/common/zip';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const TEXT_PART = /^word\/(document\d*|header\d*|footer\d*|footnotes|endnotes)\.xml$/;
type El = Element;

const parse = (bytes: Uint8Array): Document => new DOMParser().parseFromString(new TextDecoder().decode(bytes), 'application/xml') as unknown as Document;
const wordsIn = (text: string): string[] => text.split(/\s+/).filter(Boolean);

/** The words of each text part of a clean document (one with no tracked changes), in order. Paragraph ends count as a word of their own. */
export async function wordsByPart(docx: Uint8Array): Promise<Map<string, string[]>> {
  const out = new Map<string, string[]>();
  for (const [name, bytes] of await unzip(docx)) {
    if (!TEXT_PART.test(name)) continue;
    const words: string[] = [];
    const paragraphs = parse(bytes).getElementsByTagNameNS(W, 'p');
    for (let i = 0; i < paragraphs.length; i++) {
      const texts = (paragraphs.item(i) as El).getElementsByTagNameNS(W, 't');
      let text = '';
      for (let j = 0; j < texts.length; j++) text += texts.item(j)!.textContent ?? '';
      words.push(...wordsIn(text), '¶');
    }
    out.set(name.replace(/\d+\.xml$/, '.xml').replace(/^word\/(header|footer)\.xml$/, 'word/$1s'), [...(out.get(name) ?? []), ...words]);
  }
  return out;
}

/** The smallest number of words that must be deleted from a and inserted to make b (Myers' algorithm; real words only are counted). */
export function smallestDifference(a: string[], b: string[]): { deleted: number; inserted: number } {
  const ids = new Map<string, number>();
  const id = (w: string): number => ids.get(w) ?? (ids.set(w, ids.size), ids.size - 1);
  const x = Int32Array.from(a, id), y = Int32Array.from(b, id);
  const n = x.length, m = y.length, max = n + m;
  // Furthest-reaching paths, kept per step so the path can be walked back to count real words (not paragraph ends).
  const trace: Int32Array[] = [];
  let v = new Int32Array(2 * max + 2);
  let found = -1;
  for (let d = 0; d <= max && found < 0; d++) {
    const next = new Int32Array(v);
    for (let k = -d; k <= d; k += 2) {
      let i = k === -d || (k !== d && v[max + k - 1]! < v[max + k + 1]!) ? v[max + k + 1]! : v[max + k - 1]! + 1;
      let j = i - k;
      while (i < n && j < m && x[i] === y[j]) { i++; j++; }
      next[max + k] = i;
      if (i >= n && j >= m) { found = d; break; }
    }
    trace.push(v);
    v = next;
  }
  const mark = id('¶');
  let deleted = 0, inserted = 0;
  let i = n, j = m;
  for (let d = found; d > 0; d--) {
    const before = trace[d]!;
    const k = i - j;
    const down = k === -d || (k !== d && before[max + k - 1]! < before[max + k + 1]!);
    const prevK = down ? k + 1 : k - 1;
    const prevI = before[max + prevK]!;
    const prevJ = prevI - prevK;
    if (down) { if (y[prevJ] !== mark) inserted++; } else if (x[prevI] !== mark) deleted++;
    i = prevI; j = prevJ;
  }
  return { deleted, inserted };
}

export async function trueDifference(original: Uint8Array, modified: Uint8Array): Promise<{ deleted: number; inserted: number; words: number }> {
  const a = await wordsByPart(original), b = await wordsByPart(modified);
  let deleted = 0, inserted = 0, words = 0;
  for (const part of new Set([...a.keys(), ...b.keys()])) {
    const d = smallestDifference(a.get(part) ?? [], b.get(part) ?? []);
    deleted += d.deleted;
    inserted += d.inserted;
    words += (b.get(part) ?? []).filter((w) => w !== '¶').length;
  }
  return { deleted, inserted, words };
}

export interface Marked { inserted: number; deleted: number; movedFrom: number; movedTo: number; blocks: number; formattingRecords: number }

/** What a tracked-changes file marks: words inside insertions, deletions and moves, and how many separate marked stretches there are. */
export async function markedIn(tracked: Uint8Array): Promise<Marked> {
  const marked: Marked = { inserted: 0, deleted: 0, movedFrom: 0, movedTo: 0, blocks: 0, formattingRecords: 0 };
  const kindOf: Record<string, keyof Marked> = { ins: 'inserted', del: 'deleted', moveFrom: 'movedFrom', moveTo: 'movedTo' };
  for (const [name, bytes] of await unzip(tracked)) {
    if (!/^word\/[^/]+\.xml$/.test(name) || /styles|numbering|settings|fontTable|theme|webSettings/.test(name)) continue;
    const doc = parse(bytes);
    if (doc.documentElement?.namespaceURI !== W) continue;
    for (const tag of ['t', 'delText']) {
      const texts = doc.getElementsByTagNameNS(W, tag);
      for (let i = 0; i < texts.length; i++) {
        const t = texts.item(i) as El;
        for (let n: Node | null = t.parentNode; n; n = n.parentNode) {
          const e = n as El;
          if (e.namespaceURI === W && e.localName! in kindOf) { marked[kindOf[e.localName!]!] += wordsIn(t.textContent ?? '').length; break; }
        }
      }
    }
    for (const wrapper of Object.keys(kindOf)) {
      const list = doc.getElementsByTagNameNS(W, wrapper);
      for (let i = 0; i < list.length; i++) if ((list.item(i) as El).getElementsByTagNameNS(W, 'r').length > 0) marked.blocks++;
    }
    marked.formattingRecords += doc.getElementsByTagNameNS(W, 'rPrChange').length;
  }
  return marked;
}
