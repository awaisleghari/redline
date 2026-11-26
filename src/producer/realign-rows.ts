/**
 * Producer: pairs the rows of a changed table again, by what they say (found
 * on real contracts and on the owner's own comparison: a fee table shown
 * almost wholly rewritten).
 *
 * The compare engine pairs the rows of a table by position. When a row is
 * added or removed, every row after it is paired with its neighbour, and each
 * is then shown as rewritten although nothing in it changed.
 *
 * This step works out, from the engine's markup, the rows of the table as
 * they were and as they are. It pairs rows that are exactly the same in both
 * (keeping their order), and writes the table again: a row that is the same
 * in both versions is written once, unmarked; a pairing the engine itself
 * made is kept exactly as the engine wrote it; a row only one version has is
 * written as a removed or an added row.
 *
 * Narrow by design. Tables holding moves, merged cells that span rows, cell
 * insertions, nested tables, notes or row-level content controls are left
 * alone, as is any table where this would not pair more rows than the engine
 * did. The rewritten table must give back exactly the same rows for both
 * versions (checked here, and then again by the Checker).
 */
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { NS } from '../common/namespaces';
import { acceptChangesInPart } from './accept-changes';

const W = NS.w;
type El = Element;
const serializer = new XMLSerializer();
const kids = (e: El): El[] => {
  const out: El[] = [];
  for (let n = e.firstChild; n; n = n.nextSibling) if (n.nodeType === 1) out.push(n as El);
  return out;
};
const isW = (e: El | null | undefined, name: string): e is El => !!e && e.namespaceURI === W && e.localName === name;
const child = (e: El | undefined, name: string): El | undefined => (e ? kids(e).find((c) => isW(c, name)) : undefined);
const all = (e: El, name: string): El[] => {
  const list = e.getElementsByTagNameNS(W, name);
  const out: El[] = [];
  for (let i = 0; i < list.length; i++) out.push(list.item(i) as unknown as El);
  return out;
};

/** Things in a table that make pairing its rows again unsafe, or that this step does not rebuild. */
const LEFT_ALONE = ['moveFrom', 'moveTo', 'moveFromRangeStart', 'moveToRangeStart', 'vMerge', 'cellMerge', 'tblPrChange', 'tblGridChange', 'footnoteReference', 'endnoteReference', 'commentReference', 'commentRangeStart', 'sdt', 'sectPr'];
const FLAG_PARENTS = new Set(['rPr', 'trPr', 'tcPr', 'numPr', 'tblPrEx']);
const SWAPPED_TEXT: Record<string, string> = { t: 'delText', delText: 't', instrText: 'delInstrText', delInstrText: 'instrText' };
/** Records of earlier settings, and what in the settings beside them is not part of the record. */
const RECORDS: Record<string, string[]> = { rPrChange: ['ins', 'del'], pPrChange: ['rPr', 'sectPr'], tcPrChange: [], trPrChange: ['ins', 'del'], tblPrExChange: [] };

function renamed(element: El, to: string): El {
  const doc = element.ownerDocument!;
  const copy = doc.createElementNS(W, `${element.prefix ? `${element.prefix}:` : ''}${to}`) as unknown as El;
  for (let a = 0; a < element.attributes.length; a++) copy.setAttribute(element.attributes.item(a)!.name, element.attributes.item(a)!.value);
  while (element.firstChild) copy.appendChild(element.firstChild);
  element.parentNode!.replaceChild(copy as unknown as Node, element as unknown as Node);
  return copy;
}

/** Turn a row's changes round: what was added reads as removed and the other way, and each record of earlier settings changes places with the settings. */
function turnedRound(row: El): El {
  const copy = row.cloneNode(true) as unknown as El;
  for (const name of Object.keys(RECORDS)) {
    for (const record of all(copy, name)) {
      const owner = record.parentNode as El;
      const inner = kids(record)[0];
      if (!inner) continue;
      const stay = new Set([name, ...RECORDS[name]!]);
      const present = kids(owner).filter((c) => !(c.namespaceURI === W && stay.has(c.localName!)));
      const earlier = kids(inner);
      for (const c of present) inner.appendChild(c as unknown as Node);
      const first = owner.firstChild;
      for (const c of earlier) owner.insertBefore(c as unknown as Node, first);
      // (The settings come first in their element, in the order they were recorded; what stays, stays after them.)
    }
  }
  for (const [from, to] of [['ins', 'del'], ['del', 'ins'], ['cellIns', 'cellDel'], ['cellDel', 'cellIns']] as const) {
    for (const element of all(copy, from)) element.setAttribute('data-turn', to);
  }
  for (const name of ['cellIns', 'cellDel']) {
    for (const element of all(copy, name)) {
      const to = element.getAttribute('data-turn');
      if (!to) continue;
      element.removeAttribute('data-turn');
      renamed(element, to);
    }
  }
  for (const name of ['ins', 'del']) {
    for (const element of all(copy, name)) {
      const to = element.getAttribute('data-turn');
      if (!to) continue;
      element.removeAttribute('data-turn');
      const flag = isW_(element.parentNode as El) && FLAG_PARENTS.has((element.parentNode as El).localName!);
      const now = renamed(element, to);
      if (flag) continue;
      for (const [textName, other] of Object.entries(SWAPPED_TEXT)) {
        // Text inside what now reads as removed is written as removed text, and the other way.
        if ((to === 'del') !== (textName === 't' || textName === 'instrText')) continue;
        for (const text of all(now, textName)) renamed(text, other);
      }
    }
  }
  return copy;
}
function isW_(e: El | null | undefined): e is El {
  return !!e && e.namespaceURI === W;
}

const flagOf = (row: El): 'ins' | 'del' | undefined => {
  const trPr = child(row, 'trPr');
  return child(trPr, 'ins') ? 'ins' : child(trPr, 'del') ? 'del' : undefined;
};

/** A row as one version has it: the row's XML with every change accepted. */
function accepted(row: El, declarations: string): string {
  const xml = `<w:document ${declarations}><w:body><w:tbl>${serializer.serializeToString(row as never)}</w:tbl></w:body></w:document>`;
  const done = acceptChangesInPart(xml);
  const start = done.indexOf('<w:tr');
  const end = done.lastIndexOf('</w:tr>');
  return start < 0 || end < 0 ? '' : done.slice(start, end + '</w:tr>'.length);
}
/**
 * What a row is compared by: its XML without the attributes that only identify an editing session or a paragraph
 * (Word's own bookkeeping) and without elements that hold nothing.
 */
const keyOf = (rowXml: string): string =>
  rowXml
    .replace(/ w:rsid\w*="[^"]*"| w14:(?:paraId|textId)="[^"]*"/g, '')
    .replace(/ xmlns:\w+="[^"]*"/g, '')
    .replace(/\s+\/>/g, '/>')
    .replace(/<w:rPr\/>|<w:rPr><\/w:rPr>/g, '')
    .replace(/<w:pPr><\/w:pPr>|<w:pPr\/>|<w:trPr><\/w:trPr>|<w:trPr\/>/g, '');

/** Pairs of places, in order, at which the two lists hold the same thing. */
function pairsOf(a: string[], b: string[]): [number, number][] {
  const m = b.length;
  const table = new Uint32Array((a.length + 1) * (m + 1));
  for (let i = a.length - 1; i >= 0; i--) for (let j = m - 1; j >= 0; j--) table[i * (m + 1) + j] = a[i] === b[j] ? table[(i + 1) * (m + 1) + j + 1]! + 1 : Math.max(table[(i + 1) * (m + 1) + j]!, table[i * (m + 1) + j + 1]!);
  const pairs: [number, number][] = [];
  for (let i = 0, j = 0; i < a.length && j < m; ) {
    if (a[i] === b[j]) pairs.push([i++, j++]);
    else if (table[(i + 1) * (m + 1) + j]! >= table[i * (m + 1) + j + 1]!) i++;
    else j++;
  }
  return pairs;
}

/** A clean row marked as removed or added whole: the flag on the row, every run inside the change, every paragraph mark but each cell's last flagged. */
function wholeRow(rowXml: string, declarations: string, as: 'ins' | 'del', into: Document, stamp: () => string): El {
  const parsed = new DOMParser().parseFromString(`<w:document ${declarations}><w:body><w:tbl>${rowXml}</w:tbl></w:body></w:document>`, 'application/xml');
  const row = into.importNode(all(parsed.documentElement as unknown as El, 'tr')[0]! as unknown as Node, true) as unknown as El;
  const make = (name: string): El => {
    const e = into.createElementNS(W, `w:${name}`) as unknown as El;
    for (const [attribute, value] of stamp().split('|').map((pair) => pair.split('=') as [string, string])) e.setAttributeNS(W, `w:${attribute}`, value);
    return e;
  };
  let trPr = child(row, 'trPr');
  if (!trPr) {
    trPr = into.createElementNS(W, 'w:trPr') as unknown as El;
    const after = kids(row).find((c) => !isW(c, 'tblPrEx')) ?? null;
    row.insertBefore(trPr as unknown as Node, after as unknown as Node | null);
  }
  trPr.appendChild(make(as) as unknown as Node);
  for (const run of all(row, 'r')) {
    const wrapper = make(as);
    run.parentNode!.replaceChild(wrapper as unknown as Node, run as unknown as Node);
    wrapper.appendChild(run as unknown as Node);
    if (as === 'del') for (const name of ['t', 'instrText']) for (const text of all(run, name)) renamed(text, SWAPPED_TEXT[name]!);
  }
  for (const cell of kids(row).filter((c) => isW(c, 'tc'))) {
    const paragraphs = kids(cell).filter((c) => isW(c, 'p'));
    for (const p of paragraphs.slice(0, -1)) {
      let pPr = child(p, 'pPr');
      if (!pPr) {
        pPr = into.createElementNS(W, 'w:pPr') as unknown as El;
        p.insertBefore(pPr as unknown as Node, p.firstChild);
      }
      let mark = child(pPr, 'rPr');
      if (!mark) {
        mark = into.createElementNS(W, 'w:rPr') as unknown as El;
        pPr.appendChild(mark as unknown as Node);
      }
      mark.insertBefore(make(as) as unknown as Node, mark.firstChild);
    }
  }
  return row;
}

/** Mark one clean paragraph as removed or added whole: every run inside the change, and (if asked) its mark flagged. */
function markParagraph(p: El, as: 'ins' | 'del', flagMark: boolean, into: Document, make: (name: string) => El): void {
  for (const run of all(p, 'r')) {
    const wrapper = make(as);
    run.parentNode!.replaceChild(wrapper as unknown as Node, run as unknown as Node);
    wrapper.appendChild(run as unknown as Node);
    if (as === 'del') for (const name of ['t', 'instrText']) for (const text of all(run, name)) renamed(text, SWAPPED_TEXT[name]!);
  }
  if (!flagMark) return;
  let pPr = child(p, 'pPr');
  if (!pPr) {
    pPr = into.createElementNS(W, 'w:pPr') as unknown as El;
    p.insertBefore(pPr as unknown as Node, p.firstChild);
  }
  let mark = child(pPr, 'rPr');
  if (!mark) {
    mark = into.createElementNS(W, 'w:rPr') as unknown as El;
    const after = kids(pPr).find((c) => isW(c, 'sectPr') || isW(c, 'pPrChange')) ?? null;
    pPr.insertBefore(mark as unknown as Node, after as unknown as Node | null);
  }
  mark.insertBefore(make(as) as unknown as Node, mark.firstChild);
}

const wordsOf = (rowXml: string): string[] => rowXml.replace(/<w:tc[ >]/g, ' ').replace(/<[^>]+>/g, '').toLowerCase().match(/[\p{L}\p{N}]+/gu) ?? [];
/** How alike two rows are, by the words they hold (0 to 1). */
function likeness(a: string[], b: string[]): number {
  if (a.length === 0 && b.length === 0) return 1;
  const counts = new Map<string, number>();
  for (const w of a) counts.set(w, (counts.get(w) ?? 0) + 1);
  let both = 0;
  for (const w of b) {
    const left = counts.get(w) ?? 0;
    if (left > 0) {
      both++;
      counts.set(w, left - 1);
    }
  }
  return (2 * both) / (a.length + b.length);
}

const parseRow = (rowXml: string, declarations: string, into: Document): El => {
  const parsed = new DOMParser().parseFromString(`<w:document ${declarations}><w:body><w:tbl>${rowXml}</w:tbl></w:body></w:document>`, 'application/xml');
  return into.importNode(all(parsed.documentElement as unknown as El, 'tr')[0]! as unknown as Node, true) as unknown as El;
};

/**
 * A row for an old row and a new row that belong together but that the engine did not pair: the new row, with each
 * cell that differs holding the old paragraphs removed and the new ones added, in the one form a cell allows (the
 * last old paragraph and the first new one share a paragraph; the cell's last mark is untouched). The refining
 * step then compares those cells word by word. Undefined if the rows are not shaped alike.
 */
function rePaired(oldXml: string, newXml: string, declarations: string, into: Document, stamp: () => string): El | undefined {
  const make = (name: string): El => {
    const e = into.createElementNS(W, `w:${name}`) as unknown as El;
    for (const [attribute, value] of stamp().split('|').map((pair) => pair.split('=') as [string, string])) e.setAttributeNS(W, `w:${attribute}`, value);
    return e;
  };
  const oldRow = parseRow(oldXml, declarations, into);
  const row = parseRow(newXml, declarations, into);
  const shape = (r: El): string => kids(r).filter((c) => isW(c, 'tc')).map((c) => child(child(c, 'tcPr'), 'gridSpan')?.getAttributeNS(W, 'val') ?? '1').join(',');
  const ser = (e: El | undefined): string => (e ? keyOf(serializer.serializeToString(e as never)) : '');
  if (shape(oldRow) !== shape(row) || ser(child(oldRow, 'trPr')) !== ser(child(row, 'trPr'))) return undefined;
  const oldCells = kids(oldRow).filter((c) => isW(c, 'tc'));
  const cells = kids(row).filter((c) => isW(c, 'tc'));
  for (let k = 0; k < cells.length; k++) {
    const cell = cells[k]!;
    const before = oldCells[k]!;
    if (keyOf(serializer.serializeToString(before as never)) === keyOf(serializer.serializeToString(cell as never))) continue;
    if (keyOf(serializer.serializeToString((child(before, 'tcPr') ?? before) as never)) !== keyOf(serializer.serializeToString((child(cell, 'tcPr') ?? cell) as never))) return undefined;
    const olds = kids(before).filter((c) => !isW(c, 'tcPr'));
    const news = kids(cell).filter((c) => !isW(c, 'tcPr'));
    if (olds.length === 0 || news.length === 0 || ![...olds, ...news].every((c) => isW(c, 'p'))) return undefined;
    // The paragraph settings the cell's last paragraph had: recorded on the paragraph whose mark stays, so that
    // rejecting the changes gives them back.
    const settingsOf = (p: El): El[] => kids(child(p, 'pPr') ?? (into.createElementNS(W, 'w:pPr') as unknown as El)).filter((c) => !isW(c, 'rPr') && !isW(c, 'sectPr') && !isW(c, 'pPrChange'));
    const lastOld = olds[olds.length - 1]! as El;
    const lastNew = news[news.length - 1]! as El;
    const written = (list: El[]): string => keyOf(list.map((c) => serializer.serializeToString(c as never)).join(''));
    if (written(settingsOf(lastOld)) !== written(settingsOf(lastNew))) {
      let pPr = child(lastNew, 'pPr');
      if (!pPr) {
        pPr = into.createElementNS(W, 'w:pPr') as unknown as El;
        lastNew.insertBefore(pPr as unknown as Node, lastNew.firstChild);
      }
      const record = make('pPrChange');
      const earlier = into.createElementNS(W, 'w:pPr') as unknown as El;
      for (const c of settingsOf(lastOld)) earlier.appendChild(c.cloneNode(true));
      record.appendChild(earlier as unknown as Node);
      pPr.appendChild(record as unknown as Node);
    }
    // The look of the paragraph mark that stays (it sets the look of a clause number and the height of an empty
    // line): where the old last paragraph's differed, that is recorded on it too (review finding).
    const markLook = (p: El): El[] => kids(child(child(p, 'pPr'), 'rPr') ?? (into.createElementNS(W, 'w:rPr') as unknown as El));
    if (written(markLook(lastOld)) !== written(markLook(lastNew))) {
      let pPr = child(lastNew, 'pPr');
      if (!pPr) {
        pPr = into.createElementNS(W, 'w:pPr') as unknown as El;
        lastNew.insertBefore(pPr as unknown as Node, lastNew.firstChild);
      }
      let mark = child(pPr, 'rPr');
      if (!mark) {
        mark = into.createElementNS(W, 'w:rPr') as unknown as El;
        const after = kids(pPr).find((c) => isW(c, 'sectPr') || isW(c, 'pPrChange')) ?? null;
        pPr.insertBefore(mark as unknown as Node, after as unknown as Node | null);
      }
      const record = make('rPrChange');
      const earlier = into.createElementNS(W, 'w:rPr') as unknown as El;
      for (const c of markLook(lastOld)) earlier.appendChild(c.cloneNode(true));
      record.appendChild(earlier as unknown as Node);
      mark.appendChild(record as unknown as Node);
    }
    (olds as El[]).forEach((p, index) => markParagraph(p, 'del', index < olds.length - 1, into, make));
    (news as El[]).forEach((p, index) => markParagraph(p, 'ins', index < news.length - 1, into, make));
    const firstNew = news[0]! as El;
    for (const p of olds.slice(0, -1)) cell.insertBefore(p as unknown as Node, firstNew as unknown as Node);
    const ownFirst = kids(firstNew).find((c) => !isW(c, 'pPr')) ?? null;
    for (const c of kids(lastOld)) if (!isW(c, 'pPr')) firstNew.insertBefore(c as unknown as Node, ownFirst as unknown as Node | null);
  }
  return row;
}

/** One part of the engine's output with the rows of its changed tables paired again; undefined if nothing was changed. */
export function realignRows(xml: string): string | undefined {
  if (!xml.includes('<w:tbl') || (!xml.includes('<w:ins') && !xml.includes('<w:del'))) return undefined;
  const document = new DOMParser().parseFromString(xml, 'application/xml');
  const root = document.documentElement as unknown as El;
  if (root.prefix !== 'w') return undefined;
  let declarations = '';
  for (let a = 0; a < root.attributes.length; a++) {
    const attribute = root.attributes.item(a)!;
    if (attribute.name.startsWith('xmlns')) declarations += ` ${attribute.name}="${attribute.value}"`;
  }
  let highest = 0;
  for (const match of xml.matchAll(/\bw:id="(\d+)"/g)) highest = Math.max(highest, Number(match[1]));
  const stamped = /<w:(?:ins|del)\b[^>]*\bw:author="([^"|=]*)"[^>]*\bw:date="([^"|=]*)"/.exec(xml);
  const stamp = (): string => `id=${++highest}|author=${stamped?.[1] ?? 'Redline'}|date=${stamped?.[2] ?? new Date().toISOString().replace(/\.\d+Z$/, 'Z')}`;

  let changed = false;
  for (const table of all(root, 'tbl')) {
    if (all(table, 'tbl').length > 0) continue;
    if (all(table, 'ins').length + all(table, 'del').length === 0) continue;
    if (LEFT_ALONE.some((name) => all(table, name).length > 0)) continue;
    if (!kids(table).every((c) => isW_(c) && ['tblPr', 'tblGrid', 'tr', 'bookmarkStart', 'bookmarkEnd'].includes(c.localName!))) continue;
    const rows = kids(table).filter((c) => isW(c, 'tr'));
    if (rows.length < 2 || rows.length > 600) continue;

    // The two versions' rows, and for each of the engine's rows which row of each version it stands for.
    const oldRows: string[] = [];
    const newRows: string[] = [];
    const engine: { row: El; old?: number; new?: number }[] = [];
    let broken = false;
    for (const row of rows) {
      const flag = flagOf(row);
      const entry: { row: El; old?: number; new?: number } = { row };
      if (flag !== 'ins') {
        const before = accepted(turnedRound(row), declarations);
        if (!before) broken = true;
        entry.old = oldRows.push(before) - 1;
      }
      if (flag !== 'del') {
        const after = accepted(row, declarations);
        if (!after) broken = true;
        entry.new = newRows.push(after) - 1;
      }
      engine.push(entry);
    }
    if (broken) continue;
    const oldKeys = oldRows.map(keyOf);
    const newKeys = newRows.map(keyOf);
    const same = pairsOf(oldKeys, newKeys);
    const oldWords = oldRows.map(wordsOf);
    const newWords = newRows.map(wordsOf);
    const ALIKE = 0.5;

    // Write the rows again. Between two rows that are the same in both versions, a pairing the engine made is kept
    // as the engine wrote it if it lies wholly in that gap; the other rows of the gap are removed or added whole.
    const enginePair = new Map(engine.filter((e) => e.old !== undefined && e.new !== undefined).map((e) => [`${e.old} ${e.new}`, e.row]));
    const rebuilt: El[] = [];
    let mine = 0;
    const gap = (oldFrom: number, oldTo: number, newFrom: number, newTo: number): void => {
      // Within the gap, the rows are paired in order so that the rows paired are as alike as can be; only rows at
      // least half alike are paired at all.
      const n = oldTo - oldFrom;
      const m = newTo - newFrom;
      const pairs: [number, number][] = [];
      if (n > 0 && m > 0 && n * m <= 250_000) {
        const score = new Float64Array((n + 1) * (m + 1));
        const like = (x: number, y: number): number => likeness(oldWords[oldFrom + x]!, newWords[newFrom + y]!);
        for (let x = n - 1; x >= 0; x--) {
          for (let y = m - 1; y >= 0; y--) {
            const l = like(x, y);
            score[x * (m + 1) + y] = Math.max(score[(x + 1) * (m + 1) + y]!, score[x * (m + 1) + y + 1]!, l >= ALIKE ? l + score[(x + 1) * (m + 1) + y + 1]! : 0);
          }
        }
        for (let x = 0, y = 0; x < n && y < m; ) {
          const l = like(x, y);
          if (l >= ALIKE && score[x * (m + 1) + y] === l + score[(x + 1) * (m + 1) + y + 1]!) pairs.push([oldFrom + x++, newFrom + y++]);
          else if (score[(x + 1) * (m + 1) + y]! >= score[x * (m + 1) + y + 1]!) x++;
          else y++;
        }
      }
      let i = oldFrom;
      let j = newFrom;
      for (const [a, b] of pairs) {
        while (i < a) rebuilt.push(wholeRow(oldRows[i++]!, declarations, 'del', document as never, stamp));
        while (j < b) rebuilt.push(wholeRow(newRows[j++]!, declarations, 'ins', document as never, stamp));
        const kept = enginePair.get(`${a} ${b}`);
        const row = kept ? (kept.cloneNode(true) as unknown as El) : rePaired(oldRows[a]!, newRows[b]!, declarations, document as never, stamp);
        if (row) {
          rebuilt.push(row);
          mine += likeness(oldWords[a]!, newWords[b]!);
        } else {
          rebuilt.push(wholeRow(oldRows[a]!, declarations, 'del', document as never, stamp));
          rebuilt.push(wholeRow(newRows[b]!, declarations, 'ins', document as never, stamp));
        }
        i = a + 1;
        j = b + 1;
      }
      while (i < oldTo) rebuilt.push(wholeRow(oldRows[i++]!, declarations, 'del', document as never, stamp));
      while (j < newTo) rebuilt.push(wholeRow(newRows[j++]!, declarations, 'ins', document as never, stamp));
    };
    let i = 0;
    let j = 0;
    for (const [a, b] of same) {
      gap(i, a, j, b);
      const kept = enginePair.get(`${a} ${b}`);
      if (kept) rebuilt.push(kept.cloneNode(true) as unknown as El);
      else {
        const parsed = new DOMParser().parseFromString(`<w:document ${declarations}><w:body><w:tbl>${newRows[b]!}</w:tbl></w:body></w:document>`, 'application/xml');
        rebuilt.push((document as unknown as Document).importNode(all(parsed.documentElement as unknown as El, 'tr')[0]! as unknown as Node, true) as unknown as El);
      }
      i = a + 1;
      j = b + 1;
    }
    gap(i, oldRows.length, j, newRows.length);
    // Worth doing only if the rows now paired are clearly more alike than the rows the engine paired.
    mine += same.length;
    const engines = engine.reduce((sum, e) => sum + (e.old !== undefined && e.new !== undefined ? likeness(oldWords[e.old]!, newWords[e.new]!) : 0), 0);
    if (mine < engines + 0.75) continue;

    // The rebuilt rows must stand for exactly the same rows in both versions.
    const backOld: string[] = [];
    const backNew: string[] = [];
    for (const row of rebuilt) {
      const flag = flagOf(row);
      if (flag !== 'ins') backOld.push(keyOf(accepted(turnedRound(row), declarations)));
      if (flag !== 'del') backNew.push(keyOf(accepted(row, declarations)));
    }
    if (backOld.join('\u0000') !== oldKeys.join('\u0000') || backNew.join('\u0000') !== newKeys.join('\u0000')) continue;
    const firstRow = rows[0]!;
    for (const row of rebuilt) table.insertBefore(row as unknown as Node, firstRow as unknown as Node);
    for (const row of rows) table.removeChild(row as unknown as Node);
    changed = true;
  }
  return changed ? serializer.serializeToString(document as never) : undefined;
}
