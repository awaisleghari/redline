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
  return { insertions: [], deletions: [], movedFrom: [], movedTo: [] };
}

/**
 * Walk one paragraph. Neighbouring wrappers of the same kind (which Word and
 * the engine may split arbitrarily, e.g. at formatting boundaries) are joined
 * into one change; anything else between them ends the change.
 */
function collectParagraph(paragraph: XmlNode, into: ExpectedChanges): void {
  let open: { kind: WrapperName; text: string } | undefined;
  const flush = (): void => {
    if (open && open.text !== '') into[WRAPPERS[open.kind]].push(open.text);
    open = undefined;
  };
  for (let i = 0; i < paragraph.childNodes.length; i++) {
    const child = paragraph.childNodes[i]!;
    if (child.nodeType !== 1) continue;
    const name = child.localName ?? '';
    if (child.namespaceURI === W && name in WRAPPERS) {
      const kind = name as WrapperName;
      if (open && open.kind !== kind) flush();
      open ??= { kind, text: '' };
      open.text += textOf(child);
    } else if (child.namespaceURI === W && (name === 'pPr' || /Range(Start|End)$/.test(name) || /^bookmark(Start|End)$/.test(name))) {
      // Paragraph properties and zero-width range markers do not interrupt a change.
    } else {
      flush();
    }
  }
  flush();
}

function collect(node: XmlNode, into: ExpectedChanges): void {
  if (isW(node, 'p')) {
    collectParagraph(node, into);
    // Paragraphs can nest (text boxes); keep walking for those.
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
