/**
 * EXPERIMENT, not used by the product. A second way of making the
 * tracked-changes file: line the two documents up word by word in reading
 * order (ignoring how tables, sections and pages are cut), keep the modified
 * document exactly as it is, and mark only the words that really differ.
 *
 * Usage: npx tsx scripts/experiments/word-align.ts <original.docx> <modified.docx> <output folder>
 * Known gaps of this rough version are listed in PLAN.md.
 */
import { mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { unzip, zip } from '../../src/common/zip';
import { makeSpacesExplicit } from '../../src/producer/explicit-spaces';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const DATE = '2000-01-01T00:00:00Z';
type El = Element;
interface Token { text: string; mark: boolean; run?: El; start: number; para: El }

const kids = (e: El, name?: string): El[] => {
  const out: El[] = [];
  for (let n = e.firstChild; n; n = n.nextSibling) if (n.nodeType === 1 && (!name || ((n as El).namespaceURI === W && (n as El).localName === name))) out.push(n as El);
  return out;
};
const runText = (run: El): string => kids(run, 't').map((t) => t.textContent ?? '').join('');

/** Every paragraph of the body in reading order, each as its words followed by one paragraph mark. */
function tokens(body: El): Token[] {
  const out: Token[] = [];
  const paragraphs = body.getElementsByTagNameNS(W, 'p');
  for (let i = 0; i < paragraphs.length; i++) {
    const para = paragraphs.item(i) as El;
    const runs = para.getElementsByTagNameNS(W, 'r');
    for (let j = 0; j < runs.length; j++) {
      const run = runs.item(j) as El;
      let owner: Node | null = run.parentNode;
      while (owner && !((owner as El).namespaceURI === W && (owner as El).localName === 'p')) owner = owner.parentNode;
      if (owner !== para) continue; // a run of a paragraph nested inside this one (a text box)
      for (const m of runText(run).matchAll(/\S+/g)) out.push({ text: m[0], mark: false, run, start: m.index!, para });
    }
    out.push({ text: '¶', mark: true, start: 0, para });
  }
  return out;
}

/** Longest common subsequence: which tokens of a and b are the same text, in order. */
function align(a: Token[], b: Token[]): { keepA: boolean[]; keepB: boolean[] } {
  const n = a.length, m = b.length, width = m + 1;
  const table = new Uint16Array((n + 1) * width);
  for (let i = n - 1; i >= 0; i--) for (let j = m - 1; j >= 0; j--) table[i * width + j] = a[i]!.text === b[j]!.text && a[i]!.mark === b[j]!.mark ? table[(i + 1) * width + j + 1]! + 1 : Math.max(table[(i + 1) * width + j]!, table[i * width + j + 1]!);
  const keepA = new Array<boolean>(n).fill(false), keepB = new Array<boolean>(m).fill(false);
  for (let i = 0, j = 0; i < n && j < m; ) {
    if (a[i]!.text === b[j]!.text && a[i]!.mark === b[j]!.mark) { keepA[i] = keepB[j] = true; i++; j++; }
    else if (table[(i + 1) * width + j]! >= table[i * width + j + 1]!) i++;
    else j++;
  }
  return { keepA, keepB };
}

const [originalPath, modifiedPath, folder] = process.argv.slice(2);
if (!originalPath || !modifiedPath || !folder) throw new Error('usage: word-align.ts <original.docx> <modified.docx> <output folder>');
const decoder = new TextDecoder();
const originalFiles = await unzip(await makeSpacesExplicit(new Uint8Array(readFileSync(originalPath))));
const files = await unzip(await makeSpacesExplicit(new Uint8Array(readFileSync(modifiedPath))));
const parse = (f: Map<string, Uint8Array>): Document => new DOMParser().parseFromString(decoder.decode(f.get('word/document.xml')!), 'application/xml') as unknown as Document;
const originalDoc = parse(originalFiles), doc = parse(files);
const bodyOf = (d: Document): El => d.getElementsByTagNameNS(W, 'body').item(0) as El;
const a = tokens(bodyOf(originalDoc)), b = tokens(bodyOf(doc));
const { keepA, keepB } = align(a, b);

// Deleted words of the original, attached to the modified token they come just before.
const deletedBefore = new Map<number, Token[]>();
{
  let i = 0;
  for (let j = 0; j <= b.length; j++) {
    const gone: Token[] = [];
    while (i < a.length && !keepA[i]) gone.push(a[i++]!);
    if (gone.length > 0) deletedBefore.set(Math.min(j, b.length - 1), [...(deletedBefore.get(Math.min(j, b.length - 1)) ?? []), ...gone]);
    while (j < b.length && !keepB[j]) j++;
    i++; // the common token a[i] pairs with b[j]
  }
}

let id = 90000;
const make = (name: string, attributes: Record<string, string> = {}): El => {
  const e = doc.createElementNS(W, `w:${name}`) as unknown as El;
  for (const [k, v] of Object.entries(attributes)) e.setAttribute(k, v);
  return e;
};
const change = (kind: 'ins' | 'del'): El => make(kind, { 'w:id': String(id++), 'w:author': kind === 'ins' ? 'Redline (inserted)' : 'Redline (deleted)', 'w:date': DATE });
const textNode = (name: 't' | 'delText', text: string): El => {
  const t = make(name, { 'xml:space': 'preserve' });
  t.appendChild(doc.createTextNode(text));
  return t;
};
const deletion = (gone: Token[]): El => {
  const del = change('del');
  const run = make('r');
  const rPr = gone.find((g) => g.run) && kids(gone.find((g) => g.run)!.run!, 'rPr')[0];
  if (rPr) {
    const copy = doc.importNode(rPr as unknown as Node, true) as unknown as El;
    for (const style of kids(copy, 'rStyle')) copy.removeChild(style); // the original's named styles may not exist here
    run.appendChild(copy);
  }
  run.appendChild(textNode('delText', gone.filter((g) => !g.mark).map((g) => g.text).join(' ') + ' '));
  del.appendChild(run);
  return del;
};

// Per paragraph of the modified document: mark inserted words, and put deleted words where they were.
const byPara = new Map<El, number[]>();
b.forEach((token, index) => byPara.set(token.para, [...(byPara.get(token.para) ?? []), index]));
let inserted = 0, deleted = 0;
for (const [para, indexes] of byPara) {
  const markIndex = indexes[indexes.length - 1]!;
  const words = indexes.slice(0, -1);
  if (!keepB[markIndex] && words.length > 0) {
    // The paragraph mark itself is new: show the paragraph break as inserted.
    const pPr = kids(para, 'pPr')[0] ?? (para.insertBefore(make('pPr') as unknown as Node, para.firstChild) as unknown as El);
    const rPr = kids(pPr, 'rPr')[0] ?? (pPr.appendChild(make('rPr') as unknown as Node) as unknown as El);
    rPr.insertBefore(change('ins') as unknown as Node, rPr.firstChild);
  }
  const byRun = new Map<El, number[]>();
  for (const index of words) byRun.set(b[index]!.run!, [...(byRun.get(b[index]!.run!) ?? []), index]);
  let previousInserted = false;
  for (const [run, runTokens] of byRun) {
    const untouched = runTokens.every((index) => keepB[index] && !deletedBefore.has(index));
    if (untouched) { previousInserted = false; continue; }
    const text = runText(run);
    const status: boolean[] = new Array(text.length).fill(false); // true = inserted
    for (const index of runTokens) if (!keepB[index]) for (let c = b[index]!.start; c < b[index]!.start + b[index]!.text.length; c++) status[c] = true;
    // White space goes with the word before it.
    let last = previousInserted;
    for (let c = 0; c < text.length; c++) { if (/\s/.test(text[c]!)) status[c] = last; else last = status[c]!; }
    const cuts = new Set<number>([0, text.length]);
    for (let c = 1; c < text.length; c++) if (status[c] !== status[c - 1]) cuts.add(c);
    for (const index of runTokens) if (deletedBefore.has(index)) cuts.add(b[index]!.start);
    const points = [...cuts].sort((x, y) => x - y);
    const rPr = kids(run, 'rPr')[0];
    const others = kids(run).filter((k) => !(k.namespaceURI === W && (k.localName === 't' || k.localName === 'rPr')));
    const parent = run.parentNode!;
    for (let p = 0; p < points.length - 1; p++) {
      const from = points[p]!, to = points[p + 1]!;
      const hook = runTokens.find((index) => b[index]!.start === from && deletedBefore.has(index));
      if (hook !== undefined) { parent.insertBefore(deletion(deletedBefore.get(hook)!) as unknown as Node, run as unknown as Node); deleted += deletedBefore.get(hook)!.filter((g) => !g.mark).length; deletedBefore.delete(hook); }
      const piece = make('r');
      if (rPr) piece.appendChild(rPr.cloneNode(true) as unknown as Node);
      if (p === 0) for (const other of others) piece.appendChild(other.cloneNode(true) as unknown as Node);
      piece.appendChild(textNode('t', text.slice(from, to)) as unknown as Node);
      if (status[from]) { const ins = change('ins'); ins.appendChild(piece as unknown as Node); parent.insertBefore(ins as unknown as Node, run as unknown as Node); }
      else parent.insertBefore(piece as unknown as Node, run as unknown as Node);
    }
    parent.removeChild(run as unknown as Node);
    inserted += runTokens.filter((index) => !keepB[index]).length;
    previousInserted = !keepB[runTokens[runTokens.length - 1]!];
  }
  // Words deleted from the end of this paragraph.
  if (deletedBefore.has(markIndex)) { para.appendChild(deletion(deletedBefore.get(markIndex)!) as unknown as Node); deleted += deletedBefore.get(markIndex)!.filter((g) => !g.mark).length; deletedBefore.delete(markIndex); }
}

files.set('word/document.xml', new TextEncoder().encode(new XMLSerializer().serializeToString(doc as unknown as Parameters<XMLSerializer['serializeToString']>[0])));
mkdirSync(folder, { recursive: true });
writeFileSync(`${folder}/word-aligned (tracked changes).docx`, await zip(files));
console.log(`words: original ${a.filter((t) => !t.mark).length}, modified ${b.filter((t) => !t.mark).length}; marked inserted ${inserted}, deleted ${deleted}`);
