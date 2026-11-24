/**
 * Producer: an upload that already contains tracked changes is compared as if
 * every one of them had been accepted, as Word's own Compare does (the owner's
 * decision of 2026-11-24). This module makes that accepted copy. It is used
 * for everything the Producer reads from an upload: the copy given to the
 * engine, and the original's formatting and numbering for the static redline.
 *
 * The Checker never sees this copy. It reads the uploads as uploaded, with its
 * own separate implementation of "accepted", so a mistake here shows up as an
 * output that does not match and is withheld.
 */
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { NS } from '../common/namespaces';
import { unzip, zip } from '../common/zip';

const W = NS.w;
type El = Element;

/** Wrappers whose content stays when accepted, and those whose content goes. */
const KEPT_WRAPPERS = new Set(['ins', 'moveTo']);
const REMOVED_WRAPPERS = new Set(['del', 'moveFrom']);
/** Records of what properties used to be: the present properties are the accepted ones. */
const CHANGE_RECORDS = new Set(['rPrChange', 'pPrChange', 'sectPrChange', 'tblPrChange', 'tblPrExChange', 'tcPrChange', 'trPrChange', 'tblGridChange', 'numberingChange']);
/** Zero-width markers that exist only alongside tracked changes. */
const MARKERS = /^(moveFromRange|moveToRange|customXml(Ins|Del|MoveFrom|MoveTo)Range)(Start|End)$/;
/** Property containers in which w:ins / w:del are a flag on the owner (a paragraph mark, a row), not a wrapper round content. */
const FLAG_PARENTS = new Set(['rPr', 'trPr', 'tcPr', 'numPr', 'tblPrEx']);

/** The same test the page uses: any tracked-change markup, whatever the namespace prefix. */
const TRACKED = /<(?:[A-Za-z_][\w.-]*:)?(?:ins|del|moveFrom|moveTo|cellIns|cellDel|cellMerge|numberingChange|(?:rPr|pPr|sectPr|tblPr|tblPrEx|tcPr|trPr|tblGrid)Change|customXml(?:Ins|Del|MoveFrom|MoveTo)RangeStart|move(?:From|To)RangeStart)[\s/>]/;

const isW = (node: Node | null | undefined, name?: string): node is El => !!node && node.nodeType === 1 && (node as El).namespaceURI === W && (name === undefined || (node as El).localName === name);
const children = (element: El): El[] => {
  const out: El[] = [];
  for (let n = element.firstChild; n; n = n.nextSibling) if (n.nodeType === 1) out.push(n as El);
  return out;
};
const remove = (element: El): void => void element.parentNode?.removeChild(element);
const unwrap = (element: El): void => {
  const parent = element.parentNode!;
  while (element.firstChild) parent.insertBefore(element.firstChild, element);
  parent.removeChild(element);
};
const flagOn = (owner: El | undefined, names: string[]): El | undefined => (owner ? children(owner).find((c) => isW(c) && names.includes(c.localName!)) : undefined);

export function acceptChangesInPart(xml: string): string {
  if (!TRACKED.test(xml)) return xml;
  const document = new DOMParser().parseFromString(xml, 'application/xml');
  const root = document.documentElement as unknown as El;

  const all = (): El[] => {
    const list = root.getElementsByTagNameNS(W, '*');
    const out: El[] = [];
    for (let i = 0; i < list.length; i++) out.push(list.item(i) as unknown as El);
    return out;
  };

  // 1. Rows and cells: a deleted row or cell goes; an inserted one simply stays.
  for (const row of all().filter((e) => e.localName === 'tr')) {
    const trPr = children(row).find((c) => isW(c, 'trPr'));
    if (flagOn(trPr, ['del'])) remove(row);
  }
  for (const cell of all().filter((e) => e.localName === 'tc')) {
    const tcPr = children(cell).find((c) => isW(c, 'tcPr'));
    if (flagOn(tcPr, ['cellDel'])) remove(cell);
  }

  // 2. Paragraph marks: a deleted (or moved-away) mark joins its paragraph to the next one, which keeps its own properties.
  //    Done before the deleted text is removed, last paragraph first, so that a run of deleted marks collapses correctly.
  const paragraphs = all().filter((e) => e.localName === 'p');
  for (let i = paragraphs.length - 1; i >= 0; i--) {
    const paragraph = paragraphs[i]!;
    const pPr = children(paragraph).find((c) => isW(c, 'pPr'));
    const markProperties = pPr && children(pPr).find((c) => isW(c, 'rPr'));
    if (!flagOn(markProperties, ['del', 'moveFrom'])) continue;
    let next: Node | null = paragraph.nextSibling;
    while (next && !(next.nodeType === 1)) next = next.nextSibling;
    if (!isW(next, 'p')) {
      // The last paragraph of a cell, a note or the text: there is nothing to join it to. If nothing of it
      // remains (all its content is deleted or moved away) and a paragraph stands before it, it goes whole.
      const runs = paragraph.getElementsByTagNameNS(W, 'r');
      let kept = false;
      for (let r = 0; r < runs.length && !kept; r++) {
        let inside: Node | null = runs.item(r)!.parentNode;
        let removed = false;
        for (; inside && inside !== paragraph; inside = inside.parentNode) if (isW(inside) && REMOVED_WRAPPERS.has((inside as El).localName!)) removed = true;
        kept = !removed;
      }
      let previous: Node | null = paragraph.previousSibling;
      while (previous && previous.nodeType !== 1) previous = previous.previousSibling;
      if (!kept && isW(previous, 'p')) remove(paragraph);
      continue;
    }
    const target = next;
    const targetProperties = children(target).find((c) => isW(c, 'pPr'));
    const before = targetProperties ? targetProperties.nextSibling : target.firstChild;
    for (const child of children(paragraph)) {
      if (isW(child, 'pPr')) continue;
      target.insertBefore(child, before);
    }
    remove(paragraph);
  }

  // 3. Content: deleted and moved-away content goes; inserted and moved-in content stays, unwrapped.
  for (const element of all()) {
    if (!element.parentNode) continue;
    const name = element.localName!;
    const parent = element.parentNode as El;
    const isFlag = isW(parent) && FLAG_PARENTS.has(parent.localName!);
    if (REMOVED_WRAPPERS.has(name)) {
      remove(element); // as a wrapper: with its content; as a flag left on a mark that could not go: just the flag
    } else if (KEPT_WRAPPERS.has(name)) {
      if (isFlag) remove(element);
      else unwrap(element);
    } else if (name === 'cellIns' || name === 'cellMerge' || CHANGE_RECORDS.has(name) || MARKERS.test(name)) {
      remove(element);
    }
  }
  return new XMLSerializer().serializeToString(document);
}

/** Whether a Word file holds tracked changes in any of its parts. */
export async function hasTrackedChanges(docx: Uint8Array): Promise<boolean> {
  const decoder = new TextDecoder();
  for (const [name, bytes] of await unzip(docx)) if (/^word\/[^/]+\.xml$/.test(name) && TRACKED.test(decoder.decode(bytes))) return true;
  return false;
}

/** The upload with every tracked change accepted. An upload without tracked changes is returned as the very same bytes. */
export async function acceptTrackedChanges(docx: Uint8Array): Promise<Uint8Array> {
  const files = await unzip(docx);
  const decoder = new TextDecoder();
  const encoder = new TextEncoder();
  let changed = false;
  for (const [name, bytes] of files) {
    if (!/^word\/[^/]+\.xml$/.test(name)) continue;
    const xml = decoder.decode(bytes);
    const accepted = acceptChangesInPart(xml);
    if (accepted !== xml) {
      files.set(name, encoder.encode(accepted));
      changed = true;
    }
  }
  return changed ? zip(files) : docx;
}
