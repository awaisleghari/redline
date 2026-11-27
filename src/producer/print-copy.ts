/**
 * Producer: the copy of the static redline that is sent to the PDF converter.
 *
 * The converter (LibreOffice) is not Word. It lays a Word file out again with its own engine, and in two known
 * ways it then prints something other than what the file holds:
 *
 *   - it works out cross-references again when it opens the file, and prints "Error: Reference source not found"
 *     where it cannot follow one, in place of the clause number stored in the file;
 *   - it cuts off whatever does not fit a table row of fixed height (as Word does on screen).
 *
 * So the converter is given a copy in which neither can happen: every cross-reference is replaced
 * by the text stored for it, and a row of fixed height may grow to hold its text. The copy is never offered to
 * the user, and nothing here is trusted: the Checker compares the PDF with the static redline it has already
 * passed (not with this copy), so a mistake made here withholds the PDF like any other wrong PDF.
 */
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { NS } from '../common/namespaces';
import { unzip, zip } from '../common/zip';

const W = NS.w;
type El = Element;

/** Fields the converter must work out itself, page by page. */
const PAGE_FIELDS = /^\s*(PAGE|NUMPAGES|SECTIONPAGES)\b/i;
/**
 * The fields replaced by their stored text: cross-references. Others are left alone on purpose. A table of contents
 * in particular must stay a field: the converter prints its stored entries as they are, but once they are plain
 * paragraphs it hides their page numbers (which Word files mark "hidden in web view").
 */
const FROZEN_FIELDS = /^\s*(REF|NOTEREF|STYLEREF)\b/i;
const TEXT_PARTS = /^word\/(document|header\d*|footer\d*|footnotes|endnotes)\.xml$/;

const isW = (node: Node | null, name: string): node is El => !!node && node.nodeType === 1 && (node as El).namespaceURI === W && (node as El).localName === name;

interface Field {
  begin: El;
  separate?: El;
  end?: El;
  instruction: string;
  /** Holds another field, or is held by one, that must stay live. */
  keep: boolean;
  parent?: Field;
}

/** Replace fields written with begin / separate / end marks by their stored text. */
function freezeMarkedFields(root: El): void {
  const marks = root.getElementsByTagNameNS(W, '*');
  const fields: Field[] = [];
  const open: Field[] = [];
  /** For every element between a field's begin and its separate mark: the field. */
  const inInstruction: [El, Field][] = [];
  for (let i = 0; i < marks.length; i++) {
    const element = marks.item(i) as unknown as El;
    const current = open[open.length - 1];
    if (element.localName === 'fldChar') {
      const type = element.getAttributeNS(W, 'fldCharType');
      if (type === 'begin') {
        const field: Field = { begin: element, instruction: '', keep: false, parent: current };
        fields.push(field);
        open.push(field);
      } else if (type === 'separate' && current && !current.separate) current.separate = element;
      else if (type === 'end' && current) {
        current.end = element;
        open.pop();
      }
    } else if (current && !current.separate) {
      if (element.localName === 'instrText' || element.localName === 'delInstrText') current.instruction += element.textContent ?? '';
      if (element.localName === 'r') inInstruction.push([element, current]);
    }
  }
  if (open.length > 0) return; // a field left open: the part is left as it is
  for (const field of fields) {
    if (!PAGE_FIELDS.test(field.instruction)) continue;
    for (let f: Field | undefined = field; f; f = f.parent) f.keep = true;
  }
  // (A field held by one that stays live stays too.)
  const kept = (field: Field): boolean => field.keep || (!!field.parent && kept(field.parent));
  const frozen = new Set(fields.filter((field) => field.separate && field.end && FROZEN_FIELDS.test(field.instruction) && !kept(field)));
  if (frozen.size === 0) return;
  const drop = (mark: El): void => {
    const run = mark.parentNode as El;
    run.removeChild(mark);
    // A run left holding nothing but its settings goes too.
    let holds = false;
    for (let c = run.firstChild; c; c = c.nextSibling) if (c.nodeType === 1 && !isW(c, 'rPr')) holds = true;
    if (!holds && isW(run, 'r')) run.parentNode?.removeChild(run);
  };
  for (const [run, field] of inInstruction) {
    if (!frozen.has(field)) continue;
    // The instruction's runs: the instruction text goes; the begin mark is dropped below.
    for (const name of ['instrText', 'delInstrText']) {
      const texts = run.getElementsByTagNameNS(W, name);
      for (let i = texts.length - 1; i >= 0; i--) texts.item(i)!.parentNode!.removeChild(texts.item(i)!);
    }
    let holds = false;
    for (let c = run.firstChild; c; c = c.nextSibling) if (c.nodeType === 1 && !isW(c, 'rPr')) holds = true;
    if (!holds) run.parentNode?.removeChild(run);
  }
  for (const field of frozen) {
    if (field.begin.parentNode) drop(field.begin);
    if (field.separate!.parentNode) drop(field.separate!);
    if (field.end!.parentNode) drop(field.end!);
  }
}

/** Replace fields written as one element (fldSimple) by what they hold. */
function freezeSimpleFields(root: El): void {
  const found = root.getElementsByTagNameNS(W, 'fldSimple');
  const simple: El[] = [];
  for (let i = 0; i < found.length; i++) simple.push(found.item(i) as unknown as El);
  for (const field of simple) {
    if (!FROZEN_FIELDS.test(field.getAttributeNS(W, 'instr') ?? '')) continue;
    // (Not one that holds a live page field inside it.)
    const inner = field.getElementsByTagNameNS(W, 'fldSimple');
    let live = field.getElementsByTagNameNS(W, 'fldChar').length > 0;
    for (let i = 0; i < inner.length; i++) if (PAGE_FIELDS.test((inner.item(i) as unknown as El).getAttributeNS(W, 'instr') ?? '')) live = true;
    if (live) continue;
    const parent = field.parentNode!;
    while (field.firstChild) parent.insertBefore(field.firstChild, field);
    parent.removeChild(field);
  }
}

/** A table row of fixed height may grow to hold its text. */
function letRowsGrow(root: El): void {
  const heights = root.getElementsByTagNameNS(W, 'trHeight');
  for (let i = 0; i < heights.length; i++) {
    const height = heights.item(i) as unknown as El;
    if (height.getAttributeNS(W, 'hRule') === 'exact') height.setAttributeNS(W, `${height.prefix ?? 'w'}:hRule`, 'atLeast');
  }
}

export async function makePrintCopy(staticRedline: Uint8Array): Promise<Uint8Array> {
  const files = await unzip(staticRedline);
  const decoder = new TextDecoder();
  const encoder = new TextEncoder();
  for (const [name, bytes] of files) {
    if (!TEXT_PARTS.test(name)) continue;
    const xml = decoder.decode(bytes);
    if (!/fldChar|fldSimple|trHeight/.test(xml)) continue;
    const document = new DOMParser().parseFromString(xml, 'application/xml');
    const root = document.documentElement as unknown as El;
    freezeSimpleFields(root);
    freezeMarkedFields(root);
    letRowsGrow(root);
    files.set(name, encoder.encode(new XMLSerializer().serializeToString(document)));
  }
  return zip(files);
}
