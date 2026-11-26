/**
 * Producer: stops text that nobody changed from being shown removed and typed
 * again (found by the owner on a real comparison, and measured on real
 * contracts as redlines marking far more than really changed).
 *
 * The compare engine decides which paragraphs correspond. Where it gets a
 * stretch of paragraphs wrong, or re-issues a line because of something
 * inside it (a link field, for instance), it marks whole lines as removed
 * and added although most of their words are the same.
 *
 * This step looks at each stretch of changed paragraphs in the engine's
 * output and works out, from the markup alone, the two readings it stands
 * for: the paragraphs as they were, and as they are. It compares the two
 * readings word by word (paragraph marks included) and writes the stretch
 * again in one fixed form: words in both readings unmarked, the others
 * removed or added, a paragraph mark removed or added only where the
 * paragraphs really join or split, and the last mark of the stretch left
 * alone (which is also the one form Word allows at the end of a table cell).
 *
 * It is deliberately narrow. A stretch is rewritten only if it holds nothing
 * but plain paragraphs of text, links and simple fields, all with the same
 * paragraph settings; only if the rewritten stretch gives back exactly the
 * same two readings (checked here, before the Checker checks everything
 * again); and only if it marks less than the engine did. Otherwise the
 * engine's own markup is left exactly as it was.
 */
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { NS } from '../common/namespaces';

const W = NS.w;
type El = Element;
type State = 'both' | 'old' | 'new';

const kids = (e: El): El[] => {
  const out: El[] = [];
  for (let n = e.firstChild; n; n = n.nextSibling) if (n.nodeType === 1) out.push(n as El);
  return out;
};
const isW = (e: El | null | undefined, name: string): e is El => !!e && e.namespaceURI === W && e.localName === name;
const child = (e: El | undefined, name: string): El | undefined => (e ? kids(e).find((c) => isW(c, name)) : undefined);
const serializer = new XMLSerializer();
/** An element as text, without the attributes that only say when or by whom something was typed. */
const written = (e: El | undefined): string => (e ? serializer.serializeToString(e as never).replace(/ w:rsid\w*="[^"]*"/g, '').replace(/ xmlns:\w+="[^"]*"/g, '') : '');

/** One unit of a reading: a word, a space, a mark of punctuation, a tab or break, a whole field, or a paragraph mark. */
interface Token {
  /** What it is compared by: its kind, its text, its look and the link it stands in. */
  key: string;
  state: State;
  /** 'mark' for a paragraph mark; otherwise what is needed to write it out again. */
  kind: 'text' | 'item' | 'field' | 'mark' | 'bare';
  text?: string;
  /** The run's formatting, to be copied. */
  rPr?: El;
  /** For an item (tab, break...): the element. For a field: its runs, in order. */
  source?: El[];
  /** The link element the token stands in, if any, and what identifies that link. */
  link?: El;
  linkKey?: string;
  /** True for a word (letters or digits), false for spaces and punctuation. */
  word?: boolean;
}

const IN_RUN_ITEM = new Set(['tab', 'br', 'cr', 'noBreakHyphen', 'softHyphen', 'sym']);
const PASSED_OVER = new Set(['proofErr']);
const WORDS = /[\p{L}\p{N}]+|\s+|[^\p{L}\p{N}\s]/gu;

class Unsupported extends Error {}

/** The paragraph's settings without the flags on its mark. */
function settingsOf(paragraph: El): string {
  const pPr = child(paragraph, 'pPr');
  if (!pPr) return '';
  const copy = pPr.cloneNode(true) as unknown as El;
  const mark = child(copy, 'rPr');
  if (mark) {
    for (const flag of kids(mark)) if (isW(flag, 'ins') || isW(flag, 'del')) mark.removeChild(flag);
    if (kids(mark).length === 0) copy.removeChild(mark);
  }
  return kids(copy).length === 0 ? '' : written(copy);
}

function markState(paragraph: El): State {
  const mark = child(child(paragraph, 'pPr'), 'rPr');
  for (const flag of mark ? kids(mark) : []) {
    if (isW(flag, 'ins')) return 'new';
    if (isW(flag, 'del')) return 'old';
    if (isW(flag, 'moveFrom') || isW(flag, 'moveTo')) throw new Unsupported();
  }
  return 'both';
}

const isChanged = (paragraph: El): boolean => {
  try {
    if (markState(paragraph) !== 'both') return true;
  } catch {
    return true;
  }
  return paragraph.getElementsByTagNameNS(W, 'ins').length + paragraph.getElementsByTagNameNS(W, 'del').length > 0;
};

/** The tokens of one paragraph's content (its mark is added by the caller). Throws on anything this step does not handle. */
function tokensOf(paragraph: El): Token[] {
  const out: Token[] = [];
  let field: { state: State; runs: El[]; text: string } | undefined;
  const run = (r: El, state: State, link: El | undefined): void => {
    const rPr = child(r, 'rPr');
    if (rPr && child(rPr, 'rPrChange')) throw new Unsupported();
    const look = written(rPr);
    const linkKey = link ? written(link.cloneNode(false) as unknown as El) : '';
    for (const c of kids(r)) {
      if (c.namespaceURI === W && c.localName === 'rPr') continue;
      if (c.namespaceURI !== W) throw new Unsupported();
      const name = c.localName!;
      if (name === 'fldChar') {
        const type = c.getAttributeNS(W, 'fldCharType');
        if (kids(c).length > 0 || kids(r).filter((x) => !isW(x, 'rPr')).length !== 1) throw new Unsupported();
        if (type === 'begin') {
          if (field || link) throw new Unsupported(); // a field inside a field, or inside a link: left alone
          field = { state, runs: [r], text: '{' };
        } else {
          if (!field || field.state !== state) throw new Unsupported();
          field.runs.push(r);
          field.text += type === 'separate' ? '|' : '}';
          if (type === 'end') {
            out.push({ key: `field ${field.text} ${field.runs.map((x) => written(child(x, 'rPr'))).join('/')}`, state, kind: 'field', source: field.runs });
            field = undefined;
          } else if (type !== 'separate') throw new Unsupported();
        }
        continue;
      }
      if (field) {
        if (field.state !== state || kids(r).filter((x) => !isW(x, 'rPr')).length !== 1) throw new Unsupported();
        if (!['t', 'delText', 'instrText', 'delInstrText'].includes(name)) throw new Unsupported();
        field.runs.push(r);
        field.text += c.textContent ?? '';
        continue;
      }
      if (name === 't' || name === 'delText') {
        // Text written without "keep white space" loses the spaces at its ends when read; this step writes text
        // with that setting, so such text (other programs write it) is left to the engine (review finding).
        const value = c.textContent ?? '';
        if (c.getAttribute('xml:space') !== 'preserve' && value !== value.trim()) throw new Unsupported();
        for (const piece of (c.textContent ?? '').match(WORDS) ?? []) {
          out.push({ key: `t ${piece} ${look} ${linkKey}`, state, kind: 'text', text: piece, rPr, link, linkKey, word: /[\p{L}\p{N}]/u.test(piece) });
        }
      } else if (IN_RUN_ITEM.has(name)) {
        out.push({ key: `i ${written(c)} ${look} ${linkKey}`, state, kind: 'item', rPr, source: [c], link, linkKey });
      } else throw new Unsupported();
    }
  };
  const inline = (parent: El, state: State, link: El | undefined): void => {
    for (const c of kids(parent)) {
      if (c.namespaceURI !== W) throw new Unsupported();
      const name = c.localName!;
      if (name === 'pPr' && parent === paragraph) continue;
      if (PASSED_OVER.has(name)) continue;
      // A bookmark's start or end holds no text but must stay where it is (a cross-reference may point at it): it
      // is carried as a token that both readings have.
      if (name === 'bookmarkStart' || name === 'bookmarkEnd') {
        if (state !== 'both') throw new Unsupported();
        out.push({ key: `b ${written(c)}`, state: 'both', kind: 'bare', source: [c], link });
        continue;
      }
      if (name === 'r') run(c, state, link);
      else if (name === 'ins' || name === 'del') {
        if (state !== 'both' || !kids(c).every((x) => isW(x, 'r'))) throw new Unsupported();
        inline(c, name === 'ins' ? 'new' : 'old', link);
      } else if (name === 'hyperlink' && !link && parent === paragraph) inline(c, state, c);
      else throw new Unsupported();
    }
  };
  inline(paragraph, 'both', undefined);
  if (field) throw new Unsupported();
  return out;
}

/** One reading of a stretch: the tokens of that version, with a mark token wherever that version ends a paragraph. */
function reading(stream: Token[], side: 'old' | 'new'): string[] {
  // (Bookmark starts and ends hold no text and are not part of either reading; they are put back by place afterwards.)
  return stream.filter((t) => t.kind !== 'bare' && (t.state === 'both' || t.state === side)).map((t) => t.key);
}

/** The longest run of tokens the two readings share, in order: for each pair kept, its places in the two readings. */
function shared(a: string[], b: string[]): [number, number][] {
  const n = a.length;
  const m = b.length;
  const table = new Uint32Array((n + 1) * (m + 1));
  for (let i = n - 1; i >= 0; i--) {
    for (let j = m - 1; j >= 0; j--) {
      table[i * (m + 1) + j] = a[i] === b[j] ? table[(i + 1) * (m + 1) + j + 1]! + 1 : Math.max(table[(i + 1) * (m + 1) + j]!, table[i * (m + 1) + j + 1]!);
    }
  }
  const pairs: [number, number][] = [];
  for (let i = 0, j = 0; i < n && j < m; ) {
    if (a[i] === b[j]) pairs.push([i++, j++]);
    else if (table[(i + 1) * (m + 1) + j]! >= table[i * (m + 1) + j + 1]!) i++;
    else j++;
  }
  return pairs;
}

const MARK = '¶';
const marked = (stream: Token[]): number => stream.filter((t) => t.state !== 'both').length;

/** Rewrite the stretch `paragraphs` (in order, siblings) if that can be done safely and marks less. Returns true if rewritten. */
function refineStretch(blocks: El[], endsContainer: boolean, nextId: () => string, stamp: { author: string; date: string }): boolean {
  // (Between the paragraphs there may be bookmark starts and ends: they are carried along, into the paragraph that follows.)
  const paragraphs = blocks.filter((b) => isW(b, 'p'));
  const settings = settingsOf(paragraphs[0]!);
  if (!paragraphs.every((p) => settingsOf(p) === settings)) return false;
  const stream: Token[] = [];
  try {
    let i = -1;
    blocks.forEach((p) => {
      if (!isW(p, 'p')) {
        if (PASSED_OVER.has((p as El).localName!)) return; // a spelling mark: not kept
        stream.push({ key: `b ${written(p)}`, state: 'both', kind: 'bare', source: [p] });
        return;
      }
      i++;
      stream.push(...tokensOf(p));
      const last = i === paragraphs.length - 1;
      const state = markState(p);
      // The last mark of the stretch stays as it is and is not part of the comparison. (It may be flagged only at
      // the end of a container, where the flag means nothing to Word; the rewritten stretch leaves it unflagged.)
      if (last && state !== 'both' && !endsContainer) throw new Unsupported();
      if (!last) stream.push({ key: MARK, state, kind: 'mark' });
    });
  } catch (error) {
    if (error instanceof Unsupported) return false;
    throw error;
  }
  const before = marked(stream);
  if (before === 0) return false;
  // A flagged last mark at the end of a container: text after the last mark that a version keeps is that version's
  // last paragraph; if a version keeps none, its last paragraph ends at its last kept mark. Only the simple case
  // is taken on: each version has something of its own in the last paragraph, or the mark is not flagged.
  const lastState = markState(paragraphs[paragraphs.length - 1]!);
  if (lastState !== 'both') {
    const lastTokens = tokensOf(paragraphs[paragraphs.length - 1]!);
    const other = lastState === 'new' ? 'old' : 'new';
    if (lastTokens.some((t) => t.state === 'both' || t.state === other)) return false;
    // The version without this paragraph ends at its own last mark: that mark is then the fixed one, so it is dropped
    // from that version's reading by turning it into the shared end. Handled by the comparison below only when the
    // mark before the last paragraph belongs to that version alone or to both; otherwise the stretch is left.
    let k = stream.length - 1;
    while (k >= 0 && stream[k]!.kind !== 'mark') k--;
    if (k < 0 || stream[k]!.state === lastState) return false;
  }
  const oldKeys = reading(stream, 'old');
  const newKeys = reading(stream, 'new');
  // (See above: a version that lacks the last paragraph ends with a mark, which is its fixed end.)
  if (lastState === 'new' && oldKeys[oldKeys.length - 1] === MARK) oldKeys.pop();
  if (lastState === 'old' && newKeys[newKeys.length - 1] === MARK) newKeys.pop();
  if (oldKeys.length * newKeys.length > 6_000_000) return false;
  const oldTokens = stream.filter((t) => t.kind !== 'bare' && t.state !== 'new').slice(0, oldKeys.length);
  const newTokens = stream.filter((t) => t.kind !== 'bare' && t.state !== 'old').slice(0, newKeys.length);

  // The plan: every token of both readings in one order, each kept (in both), removed (old only) or added (new only).
  const plan: Token[] = [];
  let i = 0;
  let j = 0;
  const pairs = shared(oldKeys, newKeys);
  // Words in both readings that stand alone between changes, with no word of three letters or more among them, read
  // as confetti: a stretch like that is shown changed as a whole.
  const keep = pairs.map(() => true);
  for (let s = 0; s < pairs.length; ) {
    let e = s;
    while (e + 1 < pairs.length && pairs[e + 1]![0] === pairs[e]![0] + 1 && pairs[e + 1]![1] === pairs[e]![1] + 1) e++;
    const island = pairs.slice(s, e + 1).map(([a]) => oldTokens[a]!);
    const touchesStart = pairs[s]![0] === 0 && pairs[s]![1] === 0;
    const touchesEnd = pairs[e]![0] === oldKeys.length - 1 && pairs[e]![1] === newKeys.length - 1;
    const weighty = island.some((t) => t.kind === 'mark' || t.kind === 'field' || (t.word && (t.text ?? '').length >= 3));
    if (!weighty && !touchesStart && !touchesEnd) for (let k = s; k <= e; k++) keep[k] = false;
    s = e + 1;
  }
  pairs.forEach(([a, b], index) => {
    if (!keep[index]) return; // its two tokens go out with the removed and added words around them
    while (i < a) plan.push({ ...oldTokens[i++]!, state: 'old' });
    while (j < b) plan.push({ ...newTokens[j++]!, state: 'new' });
    plan.push({ ...newTokens[j]!, state: 'both' });
    i++;
    j++;
  });
  while (i < oldTokens.length) plan.push({ ...oldTokens[i++]!, state: 'old' });
  while (j < newTokens.length) plan.push({ ...newTokens[j++]!, state: 'new' });
  if (marked(plan) >= before) return false;
  // Bookmark starts and ends go back in front of the token of the new reading they stood in front of (or at the end).
  const follows = new Map<Token, Token[]>();
  const atEnd: Token[] = [];
  {
    let waiting: Token[] = [];
    for (const t of stream) {
      if (t.kind === 'bare') waiting.push(t);
      else if (t.state !== 'old' && waiting.length > 0) {
        follows.set(t, waiting);
        waiting = [];
      }
    }
    atEnd.push(...waiting);
  }
  if (follows.size > 0 || atEnd.length > 0) {
    const origin = new Map<Token, Token>(); // plan entries are copies: find each copy's source among the new tokens
    let n = 0;
    for (const entry of plan) if (entry.state !== 'old') origin.set(entry, newTokens[n++]!);
    const withBookmarks: Token[] = [];
    for (const entry of plan) {
      const source = origin.get(entry);
      if (source && follows.has(source)) withBookmarks.push(...follows.get(source)!);
      withBookmarks.push(entry);
    }
    // (A bookmark in front of a token beyond the compared part, or of none: at the end.)
    const placed = new Set(withBookmarks.filter((t) => t.kind === 'bare'));
    for (const list of follows.values()) for (const t of list) if (!placed.has(t)) atEnd.unshift(t);
    withBookmarks.push(...atEnd);
    plan.length = 0;
    plan.push(...withBookmarks);
  }
  // The plan must give back the two readings exactly.
  if (reading(plan, 'old').join('\u0000') !== oldKeys.join('\u0000') || reading(plan, 'new').join('\u0000') !== newKeys.join('\u0000')) return false;

  // Write the stretch again from the plan.
  const first = paragraphs[0]!;
  const doc = first.ownerDocument!;
  const prefix = first.prefix ? `${first.prefix}:` : '';
  const make = (name: string): El => doc.createElementNS(W, `${prefix}${name}`) as unknown as El;
  const change = (name: 'ins' | 'del'): El => {
    const e = make(name);
    e.setAttributeNS(W, `${prefix}id`, nextId());
    e.setAttributeNS(W, `${prefix}author`, stamp.author);
    e.setAttributeNS(W, `${prefix}date`, stamp.date);
    return e;
  };
  const template = child(first, 'pPr');
  const newParagraph = (mark: State): El => {
    const p = make('p');
    const pPr = template ? (template.cloneNode(true) as unknown as El) : make('pPr');
    let markProps = child(pPr, 'rPr');
    for (const flag of markProps ? kids(markProps) : []) if (isW(flag, 'ins') || isW(flag, 'del')) markProps!.removeChild(flag);
    if (mark !== 'both') {
      if (!markProps) {
        markProps = make('rPr');
        // The mark's settings stand after everything but a section break and a record of changed settings.
        const after = kids(pPr).find((c) => isW(c, 'sectPr') || isW(c, 'pPrChange')) ?? null;
        pPr.insertBefore(markProps as unknown as Node, after as unknown as Node | null);
      }
      markProps.insertBefore(change(mark === 'new' ? 'ins' : 'del') as unknown as Node, markProps.firstChild);
    } else if (markProps && kids(markProps).length === 0) pPr.removeChild(markProps);
    if (kids(pPr).length > 0) p.appendChild(pPr as unknown as Node);
    return p;
  };
  const built: El[] = [];
  let content: Token[] = [];
  const close = (mark: State): void => {
    const p = newParagraph(mark);
    // Tokens side by side with the same state, look and link go into one run.
    let holder: El = p;
    let linkOf: string | undefined;
    let wrapper: El | undefined;
    let wrapperState: State = 'both';
    let run: El | undefined;
    let runLook = '';
    let runText: El | undefined;
    for (const t of content) {
      // (One link stays one link: the removed and the added words of the same link share its element. Review finding.)
      if ((t.linkKey ?? '') !== (linkOf ?? '')) {
        linkOf = t.linkKey;
        holder = t.link && t.linkKey ? (t.link.cloneNode(false) as unknown as El) : p;
        if (holder !== p) p.appendChild(holder as unknown as Node);
        wrapper = undefined;
        run = undefined;
      }
      if (t.kind === 'bare') {
        holder.appendChild(t.source![0]!.cloneNode(true));
        wrapper = undefined;
        run = undefined;
        continue;
      }
      if (!wrapper || wrapperState !== t.state) {
        wrapperState = t.state;
        wrapper = t.state === 'both' ? holder : change(t.state === 'new' ? 'ins' : 'del');
        if (wrapper !== holder) holder.appendChild(wrapper as unknown as Node);
        run = undefined;
      }
      if (t.kind === 'field') {
        for (const source of t.source!) {
          const copy = source.cloneNode(true) as unknown as El;
          for (const c of kids(copy)) {
            const want: Record<string, string> = t.state === 'old' ? { t: 'delText', instrText: 'delInstrText' } : { delText: 't', delInstrText: 'instrText' };
            const to = want[c.localName!];
            if (c.namespaceURI === W && to) {
              const renamed = make(to);
              for (let a = 0; a < c.attributes.length; a++) renamed.setAttribute(c.attributes.item(a)!.name, c.attributes.item(a)!.value);
              while (c.firstChild) renamed.appendChild(c.firstChild);
              copy.replaceChild(renamed as unknown as Node, c as unknown as Node);
            }
          }
          wrapper.appendChild(copy as unknown as Node);
        }
        run = undefined;
        continue;
      }
      const look = written(t.rPr);
      if (!run || runLook !== look) {
        run = make('r');
        runLook = look;
        runText = undefined;
        if (t.rPr) run.appendChild(t.rPr.cloneNode(true));
        wrapper.appendChild(run as unknown as Node);
      }
      if (t.kind === 'text') {
        if (!runText) {
          runText = make(t.state === 'old' ? 'delText' : 't');
          runText.setAttribute('xml:space', 'preserve');
          runText.appendChild(doc.createTextNode('') as unknown as Node);
          run.appendChild(runText as unknown as Node);
        }
        (runText.firstChild as unknown as { data: string }).data += t.text!;
      } else {
        run.appendChild(t.source![0]!.cloneNode(true));
        runText = undefined;
      }
    }
    built.push(p);
    content = [];
  };
  for (const t of plan) {
    if (t.kind === 'mark') close(t.state);
    else content.push(t);
  }
  close('both');

  // Checked again on what was actually written, before anything in the document is replaced.
  try {
    const rewritten: Token[] = [];
    built.forEach((p, index) => {
      rewritten.push(...tokensOf(p));
      if (index < built.length - 1) rewritten.push({ key: MARK, state: markState(p), kind: 'mark' });
    });
    if (reading(rewritten, 'old').join('\u0000') !== oldKeys.join('\u0000') || reading(rewritten, 'new').join('\u0000') !== newKeys.join('\u0000')) return false;
  } catch {
    return false;
  }
  const parent = first.parentNode!;
  for (const p of built) parent.insertBefore(p as unknown as Node, first as unknown as Node);
  for (const p of blocks) parent.removeChild(p as unknown as Node);
  return true;
}

/** One part of the engine's output (main text, a header, a footer, the notes) with over-marked stretches rewritten; undefined if nothing was. */
export function refineChanges(xml: string): string | undefined {
  if (!xml.includes('<w:ins') && !xml.includes('<w:del')) return undefined;
  const document = new DOMParser().parseFromString(xml, 'application/xml');
  const root = document.documentElement as unknown as El;
  let highest = 0;
  for (const match of xml.matchAll(/\bw:id="(\d+)"/g)) highest = Math.max(highest, Number(match[1]));
  const nextId = (): string => String(++highest);
  const stamped = /<w:(?:ins|del)\b[^>]*\bw:author="([^"]*)"[^>]*\bw:date="([^"]*)"/.exec(xml);
  const stamp = { author: stamped?.[1] ?? 'Redline', date: stamped?.[2] ?? new Date().toISOString().replace(/\.\d+Z$/, 'Z') };

  // Every element that holds paragraphs directly: the text itself, cells, notes, text boxes, headers and footers.
  const holders = new Set<El>();
  const all = root.getElementsByTagNameNS(W, 'p');
  for (let i = 0; i < all.length; i++) holders.add((all.item(i) as unknown as El).parentNode as El);
  let changed = false;
  for (const holder of holders) {
    const blocks = kids(holder);
    for (let at = 0; at < blocks.length; ) {
      if (!isW(blocks[at], 'p') || !isChanged(blocks[at]!)) {
        at++;
        continue;
      }
      // A stretch: changed paragraphs side by side, and the paragraph a removed or added last mark runs into.
      let end = at;
      const isMarker = (b: El | undefined): boolean => !!b && b.namespaceURI === W && (b.localName === 'bookmarkStart' || b.localName === 'bookmarkEnd' || PASSED_OVER.has(b.localName!));
      /** The next paragraph after `from`, stepping over bookmark starts and ends; -1 if something else comes first. */
      const nextParagraph = (from: number): number => {
        let k = from + 1;
        while (isMarker(blocks[k])) k++;
        return isW(blocks[k], 'p') ? k : -1;
      };
      for (let next = nextParagraph(end); next >= 0 && isChanged(blocks[next]!); next = nextParagraph(end)) end = next;
      let flagged: boolean;
      try {
        flagged = markState(blocks[end]!) !== 'both';
      } catch {
        flagged = true;
      }
      if (flagged && nextParagraph(end) >= 0) end = nextParagraph(end);
      const stretch = blocks.slice(at, end + 1);
      // The end of the container: nothing after the stretch but what holds no text (a cell's or a section's settings).
      const endsContainer = blocks.slice(end + 1).every((b) => !isW(b, 'p') && !isW(b, 'tbl') && !isW(b, 'sdt'));
      let safe = true;
      try {
        if (markState(stretch[stretch.length - 1]!) !== 'both' && !endsContainer) safe = false;
      } catch {
        safe = false;
      }
      if (safe && refineStretch(stretch, endsContainer, nextId, stamp)) changed = true;
      at = end + 1;
    }
  }
  return changed ? serializer.serializeToString(document as never) : undefined;
}
