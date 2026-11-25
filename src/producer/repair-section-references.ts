/**
 * Producer: repairs a fault in the compare engine's handling of headers and
 * footers when a section break is shown removed and added again (found on a
 * real contract, held-out run 1).
 *
 * Each section of a Word document names its headers and footers at its
 * section break (or names none, and carries on with the previous section's).
 * When the paragraph that holds a section break changed so much that the
 * engine shows it removed and put back, its output holds one section break
 * more than either document. The engine then names the headers for the
 * section breaks that follow from the wrong section: on the real contract
 * every later section showed the header of the section after it.
 *
 * The repair, on the engine's output, only where the output holds a different
 * number of section breaks from the uploads and the two uploads arrange their
 * headers and footers identically (the same sections naming the same ones): the section breaks that remain
 * when every change is accepted are, in order, the modified document's; each
 * is given the headers and footers the modified document shows in that
 * section (those it names there, or carries over from an earlier section). A
 * section break that only the original has (shown removed) is given the
 * original's in the same way. Nothing else in a section break is
 * touched. If the breaks cannot be paired off one for one, or a header named
 * by an upload is not in the output under the same name, nothing is changed
 * and the Checker judges the engine's output as it stands.
 */
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { NS } from '../common/namespaces';

const W = NS.w;
const R = NS.r;
type El = Element;

const kids = (e: El): El[] => {
  const out: El[] = [];
  for (let n = e.firstChild; n; n = n.nextSibling) if (n.nodeType === 1) out.push(n as El);
  return out;
};
const isW = (e: El | null | undefined, name: string): e is El => !!e && e.namespaceURI === W && e.localName === name;

interface Reference {
  kind: 'header' | 'footer';
  type: string;
  id: string;
}
interface Break {
  element: El;
  /** Shown removed (only the original has it), shown added (only the modified has it), or neither. */
  change: 'removed' | 'added' | undefined;
  references: Reference[];
}
export interface SectionSource {
  main: string;
  rels: string;
}

/** The section breaks of a main document part, in order: those in paragraphs, then the one that ends the text. */
function breaksOf(root: El): Break[] {
  const found: Break[] = [];
  const all = root.getElementsByTagNameNS(W, 'sectPr');
  for (let i = 0; i < all.length; i++) {
    const element = all.item(i) as unknown as El;
    const parent = element.parentNode as El | null;
    // (A section break recorded inside a tracked "section changed" note is the old one, not a break of its own.)
    if (!isW(parent, 'pPr') && !isW(parent, 'body')) continue;
    if (isW(parent, 'pPr') && !isW(parent.parentNode as El | null, 'p')) continue;
    let change: Break['change'];
    const mark = isW(parent, 'pPr') ? kids(parent).find((c) => isW(c, 'rPr')) : undefined;
    for (const flag of mark ? kids(mark) : []) {
      if (isW(flag, 'del') || isW(flag, 'moveFrom')) change = 'removed';
      else if (isW(flag, 'ins') || isW(flag, 'moveTo')) change = 'added';
    }
    const references: Reference[] = [];
    for (const c of kids(element)) {
      if (!isW(c, 'headerReference') && !isW(c, 'footerReference')) continue;
      references.push({ kind: c.localName === 'headerReference' ? 'header' : 'footer', type: c.getAttributeNS(W, 'type') ?? '', id: c.getAttributeNS(R, 'id') ?? '' });
    }
    found.push({ element, change, references });
  }
  return found;
}

/** The headers and footers a relationships file names: id to kind and target. */
function namedIn(rels: string): Map<string, string> {
  const named = new Map<string, string>();
  for (const relationship of rels.match(/<Relationship\b[^>]*>/g) ?? []) {
    const kind = /\bType\s*=\s*["'][^"']*\/(header|footer)["']/.exec(relationship)?.[1];
    const id = /\bId\s*=\s*["']([^"']+)["']/.exec(relationship)?.[1];
    const target = /\bTarget\s*=\s*["']([^"']+)["']/.exec(relationship)?.[1];
    if (kind && id && target) named.set(id, `${kind} ${target.replace(/^\/word\//, '')}`);
  }
  return named;
}

const parse = (xml: string): El => new DOMParser().parseFromString(xml, 'application/xml').documentElement as unknown as El;
const count = (xml: string): number => xml.split('<w:sectPr').length - 1;
const same = (a: Reference[], b: Reference[]): boolean => {
  const key = (list: Reference[]): string => list.map((r) => `${r.kind} ${r.type} ${r.id}`).sort().join('|');
  return key(a) === key(b);
};

/**
 * The engine's main document part with the headers and footers of its section breaks put right, or undefined
 * if there is nothing to repair (or the repair cannot be made safely).
 */
export function repairSectionReferences(tracked: string, trackedRels: string, original: SectionSource, modified: SectionSource): string | undefined {
  // The usual case, by far: as many section breaks as in each upload. The engine handles that correctly.
  if (count(tracked) === count(modified.main) && count(tracked) === count(original.main)) return undefined;
  const document = new DOMParser().parseFromString(tracked, 'application/xml');
  const root = document.documentElement as unknown as El;
  const breaks = breaksOf(root);
  const inModified = breaksOf(parse(modified.main));
  const inOriginal = breaksOf(parse(original.main));
  const kept = breaks.filter((b) => b.change !== 'removed');
  // What each section of an upload shows: the headers and footers it names, and for the rest those of the nearest
  // section before it that names them. These are written out in full, as the engine itself does, so that a break
  // reads the same whichever other breaks stand before it once the changes are accepted or rejected.
  const shownAt = (list: Break[]): Reference[][] => {
    const current = new Map<string, Reference>();
    return list.map((b) => {
      for (const r of b.references) current.set(`${r.kind} ${r.type}`, r);
      return [...current.values()];
    });
  };
  const shownInModified = shownAt(inModified);
  const shownInOriginal = shownAt(inOriginal);
  const old = breaks.filter((b) => b.change !== 'added');
  if (kept.length !== inModified.length || old.length !== inOriginal.length) return undefined;
  if (breaks.length === inModified.length && breaks.length === inOriginal.length) return undefined;

  const here = namedIn(trackedRels);
  // Only where nothing about the headers and footers changed between the uploads: the same sections naming the
  // same ones under the same names, and the output holding exactly those. Where a section or a header came, went
  // or changed, the engine builds header parts of its own and points sections at them deliberately; what each
  // section should then name cannot be told from the uploads, and its output is left as it is (review finding).
  const sameMap = (a: Map<string, string>, b: Map<string, string>): boolean => a.size === b.size && [...a].every(([id, target]) => b.get(id) === target);
  if (inOriginal.length !== inModified.length || !inOriginal.every((b, i) => same(b.references, inModified[i]!.references))) return undefined;
  if (!sameMap(here, namedIn(modified.rels)) || !sameMap(here, namedIn(original.rels))) return undefined;
  const inUpload = { modified: namedIn(modified.rels), original: namedIn(original.rels) };
  /** Is every header and footer in the list in the output under the same name, pointing at the same part? */
  const present = (list: Reference[], upload: Map<string, string>): boolean => list.every((r) => r.id !== '' && here.has(r.id) && here.get(r.id) === upload.get(r.id));
  if (!inModified.every((b) => present(b.references, inUpload.modified))) return undefined;

  const prefix = root.lookupPrefix(R) ?? 'r';
  const name = (local: string): string => `${root.prefix ? `${root.prefix}:` : ''}${local}`;
  let changed = false;
  const give = (to: Break, references: Reference[]): void => {
    if (same(to.references, references)) return;
    for (const c of kids(to.element)) if (isW(c, 'headerReference') || isW(c, 'footerReference')) to.element.removeChild(c);
    const first = to.element.firstChild;
    for (const reference of references) {
      const element = document.createElementNS(W, name(`${reference.kind}Reference`)) as unknown as El;
      element.setAttributeNS(W, name('type'), reference.type);
      element.setAttributeNS(R, `${prefix}:id`, reference.id);
      to.element.insertBefore(element as unknown as Node, first);
    }
    changed = true;
  };
  kept.forEach((b, i) => give(b, shownInModified[i]!));
  old.forEach((b, i) => {
    // Only the original has this break: it keeps the original's headers, where the output holds them under the same names.
    if (b.change === 'removed' && present(shownInOriginal[i]!, inUpload.original)) give(b, shownInOriginal[i]!);
  });
  return changed ? new XMLSerializer().serializeToString(document) : undefined;
}
