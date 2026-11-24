/**
 * Stage 7 tooling (not part of the product): makes a "modified" version of a
 * real contract by applying a known list of edits of many kinds, the way a
 * lawyer turning a draft would: words and figures changed, words and sentences
 * added and removed, clauses added, deleted, moved, split and joined,
 * formatting changed, table rows added and removed, headers, footers and
 * footnotes edited.
 *
 * Because every edit is made here, the list of edits is the exact right answer
 * that a perfect redline would show. Edits are chosen by a seeded generator, so
 * the same document always gets the same edits.
 */
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { unzip, zip } from '../../src/common/zip';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
type El = Element;

export interface Edit {
  kind: string;
  /** Where: the part, and a few words of the text concerned (for the report only). */
  part: string;
  detail: string;
}
export interface Mutation {
  docx: Uint8Array;
  edits: Edit[];
  /** Edits by kind. */
  tally: Record<string, number>;
}

/** A small, fast, repeatable random number generator. */
function generator(seed: number): () => number {
  let a = seed >>> 0;
  return () => {
    a = (a + 0x6d2b79f5) >>> 0;
    let t = a;
    t = Math.imul(t ^ (t >>> 15), t | 1);
    t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

const REPLACEMENTS: [RegExp, string[]][] = [
  [/^shall$/, ['will', 'must']], [/^will$/, ['shall']], [/^may$/, ['shall', 'must']], [/^must$/, ['shall']],
  [/^and$/, ['or']], [/^or$/, ['and']], [/^thirty$/i, ['sixty', 'ninety']], [/^reasonable$/, ['commercially reasonable', 'best']],
  [/^written$/, ['prior written']], [/^any$/, ['each', 'all']], [/^all$/, ['any']], [/^not$/, ['never']],
  [/^Supplier$/, ['Contractor']], [/^Customer$/, ['Client']], [/^Company$/, ['Corporation']], [/^Agreement$/, ['Contract']],
  [/^days$/, ['Business Days', 'calendar days']], [/^material$/, ['substantial']], [/^promptly$/, ['immediately', 'without delay']],
];
const SPARE_WORDS = ['notwithstanding', 'the', 'foregoing', 'provided', 'that', 'such', 'party', 'obligation', 'consent', 'liability', 'termination', 'notice', 'period', 'indemnity', 'applicable', 'law', 'subject', 'to', 'clause', 'schedule', 'reasonable', 'efforts', 'without', 'prejudice', 'in', 'writing', 'except', 'as', 'expressly', 'stated'];
const SENTENCES = [
  'Each party shall bear its own costs in connection with this provision.',
  'This obligation shall survive termination or expiry for a period of three years.',
  'Nothing in this clause limits any liability that cannot be limited by law.',
  'The parties shall review this arrangement at least once in every twelve months.',
  'Any such consent shall not be unreasonably withheld, conditioned or delayed.',
  'For the avoidance of doubt, this applies to affiliates and permitted assignees.',
];

const kids = (e: El, name?: string): El[] => {
  const out: El[] = [];
  for (let n = e.firstChild; n; n = n.nextSibling) if (n.nodeType === 1 && (!name || ((n as El).namespaceURI === W && (n as El).localName === name))) out.push(n as El);
  return out;
};
const isW = (n: Node | null, name: string): n is El => !!n && n.nodeType === 1 && (n as El).namespaceURI === W && (n as El).localName === name;

/** A run that holds nothing but properties and one piece of text. */
const plainRun = (run: El): El | undefined => {
  const parts = kids(run).filter((k) => !isW(k, 'rPr'));
  return parts.length === 1 && isW(parts[0]!, 't') ? parts[0] : undefined;
};
const textOfParagraph = (p: El): string => kids(p, 'r').map((r) => plainRun(r)?.textContent ?? '').join('');
/** A paragraph it is safe to restructure: direct runs of plain text only (no fields, pictures, notes, links, section break). */
function simple(p: El): boolean {
  for (const k of kids(p)) {
    if (isW(k, 'pPr')) {
      if (kids(k, 'sectPr').length > 0) return false;
      continue;
    }
    if (isW(k, 'bookmarkStart') || isW(k, 'bookmarkEnd') || isW(k, 'proofErr')) continue;
    if (!isW(k, 'r') || !plainRun(k)) return false;
  }
  return true;
}

export async function mutate(original: Uint8Array, seed: number, wanted = 160): Promise<Mutation> {
  const random = generator(seed);
  const pick = <T,>(items: T[]): T => items[Math.floor(random() * items.length)]!;
  const files = await unzip(original);
  const decoder = new TextDecoder();
  const encoder = new TextEncoder();
  const serializer = new XMLSerializer();
  const edits: Edit[] = [];
  const note = (kind: string, part: string, detail: string): void => void edits.push({ kind, part, detail: detail.slice(0, 60) });

  const setText = (t: El, text: string): void => {
    while (t.firstChild) t.removeChild(t.firstChild);
    t.appendChild(t.ownerDocument!.createTextNode(text));
    t.setAttribute('xml:space', 'preserve');
  };
  /** Runs of a paragraph with enough words to edit. */
  const wordyRuns = (p: El, least: number): { run: El; t: El; words: string[] }[] =>
    kids(p, 'r').flatMap((run) => {
      const t = plainRun(run);
      const words = t ? (t.textContent ?? '').split(' ') : [];
      return t && words.filter((w) => /[A-Za-z]{2,}/.test(w)).length >= least ? [{ run, t, words }] : [];
    });

  // ---- edits inside one piece of text -------------------------------------------------
  const wordEdits: Record<string, (p: El, part: string) => boolean> = {
    'word replaced': (p, part) => {
      const target = pick(wordyRuns(p, 4));
      if (!target) return false;
      const { t, words } = target;
      const candidates = words.map((w, i) => ({ w, i })).filter(({ w }) => /^[A-Za-z]{3,}[.,;:]?$/.test(w));
      if (candidates.length === 0) return false;
      const { w, i } = pick(candidates);
      const bare = w.replace(/[.,;:]$/, '');
      const rule = REPLACEMENTS.find(([pattern]) => pattern.test(bare));
      const replacement = rule ? pick(rule[1]) : pick(SPARE_WORDS.filter((s) => s !== bare.toLowerCase()));
      words[i] = replacement + w.slice(bare.length);
      setText(t, words.join(' '));
      note('word replaced', part, `${bare} -> ${replacement}`);
      return true;
    },
    'figure changed': (p, part) => {
      const runs = kids(p, 'r').map((run) => plainRun(run)).filter((t): t is El => !!t && /\d/.test(t.textContent ?? ''));
      if (runs.length === 0) return false;
      const t = pick(runs);
      const text = t.textContent ?? '';
      const matches = [...text.matchAll(/\d+/g)];
      const m = pick(matches);
      const changed = String(Number(m[0]) + 1 + Math.floor(random() * 8)).padStart(m[0].length, '0');
      setText(t, text.slice(0, m.index!) + changed + text.slice(m.index! + m[0].length));
      note('figure changed', part, `${m[0]} -> ${changed}`);
      return true;
    },
    'words inserted': (p, part) => {
      const target = pick(wordyRuns(p, 4));
      if (!target) return false;
      const { t, words } = target;
      const at = 1 + Math.floor(random() * (words.length - 1));
      const added = Array.from({ length: 1 + Math.floor(random() * 4) }, () => pick(SPARE_WORDS));
      words.splice(at, 0, ...added);
      setText(t, words.join(' '));
      note('words inserted', part, added.join(' '));
      return true;
    },
    'words deleted': (p, part) => {
      const target = pick(wordyRuns(p, 8));
      if (!target) return false;
      const { t, words } = target;
      const count = 1 + Math.floor(random() * Math.min(5, words.length - 4));
      const at = 1 + Math.floor(random() * (words.length - count - 1));
      const removed = words.splice(at, count);
      setText(t, words.join(' '));
      note('words deleted', part, removed.join(' '));
      return true;
    },
    'sentence added': (p, part) => {
      const runs = kids(p, 'r').filter((run) => plainRun(run));
      const last = runs[runs.length - 1];
      if (!last || textOfParagraph(p).trim().length < 40) return false;
      const t = plainRun(last)!;
      const sentence = pick(SENTENCES);
      setText(t, `${(t.textContent ?? '').replace(/\s+$/, '')} ${sentence}`);
      note('sentence added', part, sentence);
      return true;
    },
    'formatting changed': (p, part) => {
      const target = pick(wordyRuns(p, 6));
      if (!target) return false;
      const { run, t, words } = target;
      const count = 1 + Math.floor(random() * 3);
      const at = 1 + Math.floor(random() * (words.length - count - 1));
      const doc = run.ownerDocument!;
      const piece = (text: string, format?: string): El => {
        const copy = run.cloneNode(true) as El;
        setText(plainRun(copy)!, text);
        if (format) {
          let rPr = kids(copy, 'rPr')[0];
          if (!rPr) rPr = copy.insertBefore(doc.createElementNS(W, 'w:rPr'), copy.firstChild) as El;
          for (const existing of kids(rPr, format)) rPr.removeChild(existing);
          const mark = doc.createElementNS(W, `w:${format}`);
          if (format === 'u') mark.setAttribute('w:val', 'single');
          // Schema order inside w:rPr: bold, then italic, before most other things; underline later.
          const order = ['rStyle', 'rFonts', 'b', 'bCs', 'i', 'iCs', 'caps', 'smallCaps', 'strike', 'dstrike', 'outline', 'shadow', 'emboss', 'imprint', 'noProof', 'snapToGrid', 'vanish', 'webHidden', 'color', 'spacing', 'w', 'kern', 'position', 'sz', 'szCs', 'highlight', 'u'];
          const rank = order.indexOf(format);
          const before = kids(rPr).find((k) => k.namespaceURI === W && (order.indexOf(k.localName!) === -1 || order.indexOf(k.localName!) > rank));
          rPr.insertBefore(mark, before ?? null);
        }
        return copy;
      };
      const format = pick(['b', 'i', 'u']);
      const already = kids(run, 'rPr')[0] && kids(kids(run, 'rPr')[0]!, format).length > 0;
      if (already) return false;
      const before = words.slice(0, at).join(' ') + ' ';
      const middle = words.slice(at, at + count).join(' ');
      const after = ' ' + words.slice(at + count).join(' ');
      const parent = run.parentNode!;
      parent.insertBefore(piece(before), run);
      parent.insertBefore(piece(middle, format), run);
      parent.insertBefore(piece(after), run);
      parent.removeChild(run);
      void t;
      note('formatting changed', part, `${format}: ${middle}`);
      return true;
    },
  };

  // ---- the main text -------------------------------------------------------------------
  const mainName = [...files.keys()].find((name) => /^word\/document\d*\.xml$/.test(name));
  if (!mainName) throw new Error('no main document part');
  const doc = new DOMParser().parseFromString(decoder.decode(files.get(mainName)!), 'application/xml') as unknown as Document;
  const body = doc.getElementsByTagNameNS(W, 'body').item(0) as El;
  const all = (name: string): El[] => {
    const list = body.getElementsByTagNameNS(W, name);
    const out: El[] = [];
    for (let i = 0; i < list.length; i++) out.push(list.item(i) as El);
    return out;
  };
  const inTextBox = (e: El): boolean => {
    for (let n: Node | null = e.parentNode; n && n !== body; n = n.parentNode) if (isW(n, 'txbxContent')) return true;
    return false;
  };
  const used = new Set<El>();
  /** Paragraphs still free, simple, attached, with at least so many characters of text. */
  const free = (least: number): El[] => all('p').filter((p) => !used.has(p) && p.parentNode && !inTextBox(p) && simple(p) && textOfParagraph(p).trim().length >= least);
  const soleInContainer = (p: El): boolean => kids(p.parentNode as El, 'p').length < 2;
  const sameContainer = (a: El, b: El | null): b is El => !!b && isW(b, 'p') && b.parentNode === a.parentNode;
  const nextElement = (e: El): El | null => {
    let n: Node | null = e.nextSibling;
    while (n && n.nodeType !== 1) n = n.nextSibling;
    return n as El | null;
  };

  /** A table row it is safe to remove or copy: plain paragraphs only (no note marks, bookmarks, pictures, fields, nested or merged cells). */
  const plainRow = (row: El): boolean => {
    if (row.getElementsByTagNameNS(W, 'tbl').length > 0 || row.getElementsByTagNameNS(W, 'vMerge').length > 0) return false;
    for (const tag of ['footnoteReference', 'endnoteReference', 'bookmarkStart', 'commentRangeStart', 'drawing', 'pict', 'fldChar', 'fldSimple', 'sdt', 'hyperlink']) if (row.getElementsByTagNameNS(W, tag).length > 0) return false;
    const inside = row.getElementsByTagNameNS(W, 'p');
    for (let i = 0; i < inside.length; i++) if (!simple(inside.item(i) as El) || used.has(inside.item(i) as El)) return false;
    return true;
  };

  const structural: Record<string, () => boolean> = {
    'clause added': () => {
      const after = pick(free(40));
      if (!after) return false;
      const fresh = after.cloneNode(true) as El;
      for (const k of kids(fresh)) if (isW(k, 'bookmarkStart') || isW(k, 'bookmarkEnd') || isW(k, 'proofErr')) fresh.removeChild(k); // a bookmark must not exist twice
      const runs = kids(fresh, 'r');
      runs.slice(1).forEach((r) => fresh.removeChild(r));
      const words = textOfParagraph(after).split(' ').filter((w) => /^[A-Za-z]{4,}$/.test(w));
      const text = `${pick(SENTENCES)} ${Array.from({ length: 6 }, () => (words.length ? pick(words) : pick(SPARE_WORDS))).join(' ')}.`;
      setText(plainRun(runs[0]!)!, text);
      after.parentNode!.insertBefore(fresh, after.nextSibling);
      used.add(after).add(fresh);
      note('clause added', mainName, text);
      return true;
    },
    'clause deleted': () => {
      const p = pick(free(40).filter((x) => !soleInContainer(x)));
      if (!p) return false;
      note('clause deleted', mainName, textOfParagraph(p));
      p.parentNode!.removeChild(p);
      used.add(p);
      return true;
    },
    'clause moved': () => {
      const candidates = free(80).filter((x) => x.parentNode === body);
      const p = pick(candidates);
      if (!p || candidates.length < 12) return false;
      const far = candidates.filter((x) => x !== p && Math.abs(candidates.indexOf(x) - candidates.indexOf(p)) > 5);
      const target = pick(far);
      if (!target) return false;
      body.removeChild(p);
      body.insertBefore(p, target);
      used.add(p).add(target);
      note('clause moved', mainName, textOfParagraph(p));
      return true;
    },
    'clause split in two': () => {
      const p = pick(free(160));
      if (!p) return false;
      const target = pick(wordyRuns(p, 12));
      if (!target) return false;
      const { run, t, words } = target;
      const at = Math.floor(words.length / 2);
      const second = p.cloneNode(false) as El;
      const pPr = kids(p, 'pPr')[0];
      if (pPr) second.appendChild(pPr.cloneNode(true));
      const tail = run.cloneNode(true) as El;
      setText(plainRun(tail)!, words.slice(at).join(' '));
      setText(t, words.slice(0, at).join(' '));
      second.appendChild(tail);
      for (let n = run.nextSibling; n; ) {
        const following = n.nextSibling;
        second.appendChild(n);
        n = following;
      }
      p.parentNode!.insertBefore(second, p.nextSibling);
      used.add(p).add(second);
      note('clause split in two', mainName, words.slice(at, at + 5).join(' '));
      return true;
    },
    'two clauses joined': () => {
      const p = pick(free(40).filter((x) => { const n = nextElement(x); return sameContainer(x, n) && !used.has(n) && simple(n) && textOfParagraph(n).trim().length >= 20; }));
      if (!p) return false;
      const next = nextElement(p)!;
      const lastRun = kids(p, 'r').filter((r) => plainRun(r)).pop();
      if (!lastRun) return false;
      const t = plainRun(lastRun)!;
      setText(t, `${(t.textContent ?? '').replace(/\s+$/, '')} `);
      for (const k of kids(next)) if (!isW(k, 'pPr')) p.appendChild(k);
      next.parentNode!.removeChild(next);
      used.add(p).add(next);
      note('two clauses joined', mainName, textOfParagraph(p).slice(-40));
      return true;
    },
    'table row deleted': () => {
      const table = pick(all('tbl').filter((tbl) => kids(tbl, 'tr').length >= 4));
      if (!table) return false;
      const rows = kids(table, 'tr');
      const row = rows[1 + Math.floor(random() * (rows.length - 2))]!;
      if (!plainRow(row)) return false;
      note('table row deleted', mainName, row.textContent ?? '');
      table.removeChild(row);
      return true;
    },
    'table row added': () => {
      const table = pick(all('tbl').filter((tbl) => kids(tbl, 'tr').length >= 2));
      if (!table) return false;
      const rows = kids(table, 'tr');
      const row = rows[Math.floor(random() * rows.length)]!;
      if (!plainRow(row)) return false;
      const fresh = row.cloneNode(true) as El;
      const texts = fresh.getElementsByTagNameNS(W, 't');
      if (texts.length === 0) return false;
      for (let i = 0; i < texts.length; i++) {
        const t = texts.item(i) as El;
        if ((t.textContent ?? '').trim()) setText(t, `${pick(SPARE_WORDS)} ${pick(SPARE_WORDS)} ${Math.floor(random() * 900) + 100}`);
      }
      table.insertBefore(fresh, row.nextSibling);
      const inside = fresh.getElementsByTagNameNS(W, 'p');
      for (let i = 0; i < inside.length; i++) used.add(inside.item(i) as El);
      note('table row added', mainName, fresh.textContent ?? '');
      return true;
    },
    'alignment changed': () => {
      const p = pick(free(40));
      if (!p) return false;
      let pPr = kids(p, 'pPr')[0];
      if (!pPr) pPr = p.insertBefore(doc.createElementNS(W, 'w:pPr'), p.firstChild) as El;
      const existing = kids(pPr, 'jc')[0];
      const value = existing?.getAttributeNS(W, 'val') === 'center' ? 'both' : 'center';
      if (existing) existing.setAttribute('w:val', value);
      else {
        const jc = doc.createElementNS(W, 'w:jc');
        jc.setAttribute('w:val', value);
        const after = ['textAlignment', 'textboxTightWrap', 'outlineLvl', 'divId', 'cnfStyle', 'rPr', 'sectPr', 'pPrChange'];
        pPr.insertBefore(jc, kids(pPr).find((k) => k.namespaceURI === W && after.includes(k.localName!)) ?? null);
      }
      used.add(p);
      note('alignment changed', mainName, value);
      return true;
    },
  };

  // The mix of edits, weighted towards the small changes that make up most of a real turn of a draft.
  const plan: [string, number][] = [
    ['word replaced', 34], ['words inserted', 22], ['words deleted', 22], ['figure changed', 10], ['sentence added', 10], ['formatting changed', 20],
    ['clause added', 10], ['clause deleted', 10], ['clause moved', 6], ['clause split in two', 5], ['two clauses joined', 5],
    ['table row added', 3], ['table row deleted', 3], ['alignment changed', 6],
  ];
  const scale = wanted / plan.reduce((sum, [, n]) => sum + n, 0);
  for (const [kind, count] of plan) {
    const goal = Math.max(1, Math.round(count * scale));
    let made = 0;
    for (let attempt = 0; made < goal && attempt < goal * 6; attempt++) {
      if (kind in structural) {
        if (structural[kind]!()) made++;
      } else {
        const p = pick(free(30));
        if (!p) break;
        if (wordEdits[kind]!(p, mainName)) {
          used.add(p);
          made++;
        }
      }
    }
  }
  // Top up with small text edits if the document could not take enough of some kind.
  for (let attempt = 0; edits.length < wanted && attempt < wanted * 4; attempt++) {
    const p = pick(free(30));
    if (!p) break;
    const kind = pick(['word replaced', 'words inserted', 'words deleted']);
    if (wordEdits[kind]!(p, mainName)) used.add(p);
  }
  // A document with few separate paragraphs: further small edits in paragraphs already touched (as happens in a real turn of a draft).
  const touched = (): El[] => all('p').filter((p) => p.parentNode && !inTextBox(p) && simple(p) && textOfParagraph(p).trim().length >= 60);
  for (let attempt = 0; edits.length < wanted && attempt < wanted * 6; attempt++) {
    const p = pick(touched());
    if (!p) break;
    wordEdits[pick(['word replaced', 'words inserted', 'words deleted', 'figure changed'])]!(p, mainName);
  }
  files.set(mainName, encoder.encode(serializer.serializeToString(doc as unknown as Parameters<XMLSerializer['serializeToString']>[0])));

  // ---- headers, footers and notes: a word or two each ---------------------------------
  for (const [name, bytes] of files) {
    if (!/^word\/(header|footer)\d*\.xml$|^word\/(footnotes|endnotes)\.xml$/.test(name)) continue;
    const part = new DOMParser().parseFromString(decoder.decode(bytes), 'application/xml') as unknown as Document;
    const paragraphs = part.getElementsByTagNameNS(W, 'p');
    const candidates: El[] = [];
    for (let i = 0; i < paragraphs.length; i++) {
      const p = paragraphs.item(i) as El;
      if (simple(p) && textOfParagraph(p).trim().length >= 25) candidates.push(p);
    }
    let changed = false;
    for (let n = 0; n < Math.min(2, candidates.length); n++) {
      const p = candidates.splice(Math.floor(random() * candidates.length), 1)[0]!;
      if (wordEdits[pick(['word replaced', 'words inserted'])]!(p, name)) changed = true;
    }
    if (changed) files.set(name, encoder.encode(serializer.serializeToString(part as unknown as Parameters<XMLSerializer['serializeToString']>[0])));
  }

  const tally: Record<string, number> = {};
  for (const edit of edits) tally[edit.kind] = (tally[edit.kind] ?? 0) + 1;
  return { docx: await zip(files), edits, tally };
}
