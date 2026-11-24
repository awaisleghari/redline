/**
 * Producer: repairs a fault in the compare engine's paragraph properties
 * (found on real contracts, Stage 7).
 *
 * Where paragraphs are added and removed next to each other, the engine
 * sometimes gives a paragraph that survives the wrong properties: another
 * paragraph's style or numbering, or none where the modified document has
 * some. Accepting every change must give the modified document, so each
 * paragraph that survives acceptance must carry exactly the properties of the
 * corresponding paragraph of the modified document.
 *
 * The repair, on the engine's main text: the paragraphs that survive acceptance
 * are lined up, in order, with the modified document's paragraphs. Only if
 * there are exactly as many of each are any touched; then each one whose
 * properties differ is given the modified document's (its style, numbering,
 * indents, alignment and the rest). The paragraph mark's own tracked-change
 * flag and formatting, a section break and a record of earlier properties
 * stay as the engine wrote them. The Checker judges the result as always.
 */
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { NS } from '../common/namespaces';

const W = NS.w;
type El = Element;
/** Children of w:pPr that are not the paragraph's own layout properties and are never replaced. */
const KEPT = new Set(['rPr', 'sectPr', 'pPrChange']);

const kids = (e: El, name?: string): El[] => {
  const out: El[] = [];
  for (let n = e.firstChild; n; n = n.nextSibling) if (n.nodeType === 1 && (!name || ((n as El).namespaceURI === W && (n as El).localName === name))) out.push(n as El);
  return out;
};
const isW = (n: Node | null, name: string): n is El => !!n && n.nodeType === 1 && (n as El).namespaceURI === W && (n as El).localName === name;
const has = (owner: El | undefined, names: string[]): boolean => !!owner && kids(owner).some((c) => c.namespaceURI === W && names.includes(c.localName!));

/** The paragraphs of a main text that exist once every change is accepted, in order. */
function survivors(body: El): El[] {
  const all = body.getElementsByTagNameNS(W, 'p');
  const out: El[] = [];
  for (let i = 0; i < all.length; i++) {
    const paragraph = all.item(i) as unknown as El;
    const mark = kids(kids(paragraph, 'pPr')[0] ?? paragraph, 'rPr')[0];
    if (has(mark, ['del', 'moveFrom'])) continue; // joins the next paragraph: it is not a paragraph of the accepted text
    let gone = false;
    for (let n: Node | null = paragraph.parentNode; n && n !== body; n = n.parentNode) {
      const e = n as El;
      if (e.namespaceURI !== W) continue;
      if (e.localName === 'del' || e.localName === 'moveFrom') gone = true;
      if (e.localName === 'tr' && has(kids(e, 'trPr')[0], ['del'])) gone = true;
      if (e.localName === 'tc' && has(kids(e, 'tcPr')[0], ['cellDel'])) gone = true;
    }
    if (!gone) out.push(paragraph);
  }
  return out;
}

const serializer = new XMLSerializer();
/** A paragraph's own layout properties, as text, for comparing (attribute spacing and namespace declarations aside). */
const layoutOf = (paragraph: El): string =>
  kids(kids(paragraph, 'pPr')[0] ?? paragraph)
    .filter((c) => paragraph !== c.parentNode && !(c.namespaceURI === W && KEPT.has(c.localName!)))
    .map((c) => serializer.serializeToString(c as unknown as Parameters<XMLSerializer['serializeToString']>[0]).replace(/ xmlns:\w+="[^"]*"/g, '').replace(/<w:numberingChange\b[^>]*\/>/g, '').replace(/\s+\/>/g, '/>').replace(/ w:rsid\w*="[^"]*"/g, ''))
    .join('');

/** Returns the engine's main text with its surviving paragraphs' properties made the modified document's, or undefined if nothing needed changing or the two do not line up. */
export function repairParagraphProperties(trackedDocumentXml: string, modifiedDocumentXml: string): string | undefined {
  const tracked = new DOMParser().parseFromString(trackedDocumentXml, 'application/xml');
  const modified = new DOMParser().parseFromString(modifiedDocumentXml, 'application/xml');
  const body = tracked.getElementsByTagNameNS(W, 'body').item(0) as unknown as El | null;
  const modifiedBody = modified.getElementsByTagNameNS(W, 'body').item(0) as unknown as El | null;
  if (!body || !modifiedBody) return undefined;
  const ours = survivors(body);
  const theirs = survivors(modifiedBody);
  if (ours.length !== theirs.length) return undefined;

  let changed = false;
  ours.forEach((paragraph, index) => {
    const wanted = theirs[index]!;
    if (layoutOf(paragraph) === layoutOf(wanted)) return;
    const source = kids(wanted, 'pPr')[0];
    let pPr = kids(paragraph, 'pPr')[0];
    if (!pPr) {
      if (!source) return;
      pPr = tracked.createElementNS(W, `${paragraph.prefix ?? 'w'}:pPr`) as unknown as El;
      paragraph.insertBefore(pPr, paragraph.firstChild);
    }
    for (const child of kids(pPr)) if (!(child.namespaceURI === W && KEPT.has(child.localName!))) pPr.removeChild(child);
    const before = kids(pPr)[0] ?? null; // the kept children come last in a w:pPr
    if (source) for (const child of kids(source)) if (!(isW(child, 'rPr') || isW(child, 'sectPr') || isW(child, 'pPrChange'))) pPr.insertBefore(tracked.importNode(child as unknown as Parameters<typeof tracked.importNode>[0], true) as unknown as Node, before as unknown as Node | null);
    changed = true;
  });
  return changed ? serializer.serializeToString(tracked as unknown as Parameters<XMLSerializer['serializeToString']>[0]) : undefined;
}
