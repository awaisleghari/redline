/**
 * Test-only reader for the tracked changes in a .docx file.
 *
 * It reads the raw XML itself rather than asking the compare engine what it
 * did, so the tests check the file a user would actually receive. (This is a
 * test helper; the product's independent Checker is built separately in
 * `src/checker/`.)
 */
import { DOMParser } from '@xmldom/xmldom';
import JSZip from 'jszip';
import type { ExpectedChanges, PartKind } from '../fixtures/pairs';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

/** Tracked-change elements that wrap runs of text. */
const WRAPPERS = {
  ins: 'insertions',
  del: 'deletions',
  moveFrom: 'movedFrom',
  moveTo: 'movedTo',
} as const satisfies Record<string, keyof ExpectedChanges>;
type WrapperName = keyof typeof WRAPPERS;

function partKind(path: string): PartKind | undefined {
  if (path === 'word/document.xml') return 'body';
  if (/^word\/header\d*\.xml$/.test(path)) return 'header';
  if (/^word\/footer\d*\.xml$/.test(path)) return 'footer';
  if (path === 'word/footnotes.xml') return 'footnotes';
  if (path === 'word/endnotes.xml') return 'endnotes';
  return undefined;
}

type XmlNode = { nodeType: number; namespaceURI: string | null; localName: string | null; childNodes: ArrayLike<XmlNode>; textContent: string | null };

const isW = (node: XmlNode, name: string): boolean =>
  node.nodeType === 1 && node.namespaceURI === W && node.localName === name;

/** All text inside a node, counting both live text (w:t) and deleted text (w:delText); tabs and breaks included. */
function textOf(node: XmlNode): string {
  if (isW(node, 't') || isW(node, 'delText')) return node.textContent ?? '';
  if (isW(node, 'tab')) return '\t';
  if (isW(node, 'br') || isW(node, 'cr')) return '\n';
  let text = '';
  for (let i = 0; i < node.childNodes.length; i++) text += textOf(node.childNodes[i]!);
  return text;
}

function emptyChanges(): ExpectedChanges {
  return {
    insertions: [],
    deletions: [],
    movedFrom: [],
    movedTo: [],
    moveOrder: [],
    insertedParagraphMarks: 0,
    deletedParagraphMarks: 0,
    insertedRows: 0,
    deletedRows: 0,
  };
}

/** Elements that merely group runs inside a paragraph; changes can sit inside them. */
const CONTAINERS = new Set(['hyperlink', 'sdt', 'sdtContent', 'smartTag', 'fldSimple', 'customXml']);
/** Elements with no width that may sit between two pieces of one change. */
const isInvisible = (name: string): boolean =>
  name === 'pPr' || name === 'sdtPr' || name === 'sdtEndPr' || name === 'proofErr' || /Range(Start|End)$/.test(name) || /^bookmark(Start|End)$/.test(name);

function firstChild(node: XmlNode, name: string): XmlNode | undefined {
  for (let i = 0; i < node.childNodes.length; i++) if (isW(node.childNodes[i]!, name)) return node.childNodes[i]!;
  return undefined;
}

/**
 * Walk one paragraph. Neighbouring wrappers of the same kind (which Word and
 * the engine may split arbitrarily, e.g. at formatting boundaries) are joined
 * into one change; anything else between them ends the change. Wrappers inside
 * hyperlinks and content controls are found too.
 */
function collectParagraph(paragraph: XmlNode, into: ExpectedChanges): void {
  let open: { kind: WrapperName; text: string } | undefined;
  const flush = (): void => {
    if (open && open.text !== '') {
      into[WRAPPERS[open.kind]].push(open.text);
      if (open.kind === 'moveFrom') into.moveOrder.push('from');
      if (open.kind === 'moveTo') into.moveOrder.push('to');
    }
    open = undefined;
  };
  const walk = (parent: XmlNode): void => {
    for (let i = 0; i < parent.childNodes.length; i++) {
      const child = parent.childNodes[i]!;
      if (child.nodeType !== 1) continue;
      const name = child.localName ?? '';
      const isWord = child.namespaceURI === W;
      if (isWord && name in WRAPPERS) {
        const kind = name as WrapperName;
        if (open && open.kind !== kind) flush();
        open ??= { kind, text: '' };
        open.text += textOf(child);
      } else if (isWord && CONTAINERS.has(name)) {
        walk(child);
      } else if (isWord && isInvisible(name)) {
        // Does not interrupt a change.
      } else {
        flush();
      }
    }
  };
  walk(paragraph);
  flush();

  // A tracked paragraph mark: the paragraph itself (not just its text) was added or removed.
  const markProperties = firstChild(paragraph, 'pPr');
  const mark = markProperties && firstChild(markProperties, 'rPr');
  if (mark) {
    if (firstChild(mark, 'ins') || firstChild(mark, 'moveTo')) into.insertedParagraphMarks++;
    if (firstChild(mark, 'del') || firstChild(mark, 'moveFrom')) into.deletedParagraphMarks++;
  }
}

function collect(node: XmlNode, into: ExpectedChanges): void {
  if (isW(node, 'p')) {
    collectParagraph(node, into);
    // Paragraphs can nest (text boxes); keep walking for those.
  }
  if (isW(node, 'tr')) {
    const rowProperties = firstChild(node, 'trPr');
    if (rowProperties && firstChild(rowProperties, 'ins')) into.insertedRows++;
    if (rowProperties && firstChild(rowProperties, 'del')) into.deletedRows++;
  }
  for (let i = 0; i < node.childNodes.length; i++) {
    const child = node.childNodes[i]!;
    if (child.nodeType === 1) collect(child, into);
  }
}

export interface TrackedChanges {
  /** Changes per part kind. Every kind is present, empty when the part has no changes. */
  parts: Record<PartKind, ExpectedChanges>;
  /** Number of formatting-change records (w:rPrChange, w:pPrChange) anywhere in the file. */
  formatChanges: number;
}

export async function readTrackedChanges(docx: Uint8Array): Promise<TrackedChanges> {
  const zip = await JSZip.loadAsync(docx);
  const parts: Record<PartKind, ExpectedChanges> = {
    body: emptyChanges(),
    header: emptyChanges(),
    footer: emptyChanges(),
    footnotes: emptyChanges(),
    endnotes: emptyChanges(),
  };
  let formatChanges = 0;
  const paths = Object.keys(zip.files).sort();
  for (const path of paths) {
    const kind = partKind(path);
    if (!kind) continue;
    const xml = await zip.file(path)!.async('string');
    const dom = new DOMParser().parseFromString(xml, 'application/xml') as unknown as { documentElement: XmlNode & { getElementsByTagNameNS(ns: string, name: string): { length: number } } };
    collect(dom.documentElement, parts[kind]);
    formatChanges +=
      dom.documentElement.getElementsByTagNameNS(W, 'rPrChange').length +
      dom.documentElement.getElementsByTagNameNS(W, 'pPrChange').length;
  }
  return { parts, formatChanges };
}

/** The full expectation for a pair: listed parts as given, every other part empty. */
export function fullExpectation(expected: Partial<Record<PartKind, ExpectedChanges>>): Record<PartKind, ExpectedChanges> {
  return {
    body: expected.body ?? emptyChanges(),
    header: expected.header ?? emptyChanges(),
    footer: expected.footer ?? emptyChanges(),
    footnotes: expected.footnotes ?? emptyChanges(),
    endnotes: expected.endnotes ?? emptyChanges(),
  };
}
