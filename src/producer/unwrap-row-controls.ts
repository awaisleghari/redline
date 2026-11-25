/**
 * Producer: works round a fault in the compare engine with table rows that sit
 * inside content controls (found on real contracts, Stage 7).
 *
 * Procurement and form templates wrap table rows in "repeating section"
 * content controls, so that a user can add rows. Given such a table, the engine
 * loses the table's ordinary rows from its output and does not mark changes in
 * the wrapped ones. So, in the copies of the documents given to the engine, a
 * content control that wraps whole table rows is taken off and its rows are
 * left in place. The text, the rows and their formatting are untouched.
 *
 * Consequence, declared: the redline does not carry those row-level content
 * controls (the "add a row" control is not in the redline). Content controls
 * inside cells and paragraphs are not touched.
 */
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { NS } from '../common/namespaces';
import { unzip, zip } from '../common/zip';

const W = NS.w;
type El = Element;
const kids = (e: El): El[] => {
  const out: El[] = [];
  for (let n = e.firstChild; n; n = n.nextSibling) if (n.nodeType === 1) out.push(n as El);
  return out;
};
const isW = (e: El | null | undefined, name: string): e is El => !!e && e.namespaceURI === W && e.localName === name;

export function unwrapRowControlsInPart(xml: string): string {
  if (!xml.includes('<w:sdt') || !xml.includes('<w:tbl')) return xml;
  const document = new DOMParser().parseFromString(xml, 'application/xml');
  let changed = false;
  for (let again = true; again; ) {
    again = false;
    const controls = document.getElementsByTagNameNS(W, 'sdt');
    for (let i = 0; i < controls.length; i++) {
      const control = controls.item(i) as unknown as El;
      const parent = control.parentNode as El | null;
      const content = kids(control).find((c) => isW(c, 'sdtContent'));
      // A control standing where a row stands (directly in a table), holding rows or further such controls.
      if (!content || !isW(parent, 'tbl')) continue;
      if (!kids(content).every((c) => isW(c, 'tr') || isW(c, 'sdt'))) continue;
      while (content.firstChild) parent.insertBefore(content.firstChild, control);
      parent.removeChild(control);
      changed = again = true;
      break; // the live list has changed: start again
    }
  }
  return changed ? new XMLSerializer().serializeToString(document) : xml;
}

/** The document with row-level content controls taken off. A document without any is returned as the very same bytes. */
export async function unwrapRowControls(docx: Uint8Array): Promise<Uint8Array> {
  const files = await unzip(docx);
  const decoder = new TextDecoder();
  const encoder = new TextEncoder();
  let changed = false;
  for (const [name, bytes] of files) {
    if (!/^word\/(document|header\d*|footer\d*|footnotes|endnotes)\.xml$/.test(name)) continue;
    const xml = decoder.decode(bytes);
    const unwrapped = unwrapRowControlsInPart(xml);
    if (unwrapped !== xml) {
      files.set(name, encoder.encode(unwrapped));
      changed = true;
    }
  }
  return changed ? zip(files) : docx;
}
