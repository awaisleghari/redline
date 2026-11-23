/**
 * Producer: build the static redline Word file from the tracked-changes file.
 * Every Word-markup part is converted (body, headers, footers, footnotes,
 * endnotes), the change counts are totalled, and the cover page is added.
 */
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { NS } from '../common/namespaces';
import { unzip, zip } from '../common/zip';
import { coverPageXml, coverStylesXml, type CoverPageContent } from './cover-page';
import { convertPart, DEFAULT_STYLE_OPTIONS, type ChangeCounts, type StaticRedlineStyleOptions } from './static-redline';

const W = NS.w;
const decoder = new TextDecoder();
const encoder = new TextEncoder();

/** Content types of the parts that hold document text and can therefore hold tracked changes. */
const TEXT_PART_TYPES = [
  'wordprocessingml.document.main+xml',
  'wordprocessingml.template.main+xml',
  'ms-word.document.macroEnabled.main+xml',
  'wordprocessingml.header+xml',
  'wordprocessingml.footer+xml',
  'wordprocessingml.footnotes+xml',
  'wordprocessingml.endnotes+xml',
  'wordprocessingml.comments+xml',
];

export interface StaticRedlineOptions {
  originalName: string;
  modifiedName: string;
  /** Date and time of the comparison, already formatted for display. */
  comparedAt: string;
  style?: Partial<StaticRedlineStyleOptions>;
}

export interface StaticRedline {
  docx: Uint8Array;
  counts: ChangeCounts;
}

function attributeValue(tag: string, name: string): string | undefined {
  return tag.match(new RegExp(`\\b${name}="([^"]*)"`))?.[1];
}

/** Part names by declared content type, read from [Content_Types].xml. */
function partsOfTypes(contentTypes: string, names: Iterable<string>, wanted: string[]): { textParts: string[]; main?: string; styles?: string } {
  const overrides = new Map<string, string>();
  for (const tag of contentTypes.match(/<Override\b[^>]*>/g) ?? []) {
    const part = attributeValue(tag, 'PartName');
    const type = attributeValue(tag, 'ContentType');
    if (part && type) overrides.set(part.replace(/^\//, ''), type);
  }
  const textParts: string[] = [];
  let main: string | undefined;
  let styles: string | undefined;
  for (const name of names) {
    const type = overrides.get(name);
    if (!type) continue;
    if (wanted.some((w) => type.endsWith(w))) textParts.push(name);
    if (/\.(document|template)\.(macroEnabled\.)?main\+xml$/.test(type) || type.endsWith('macroEnabledTemplate.main+xml')) main = name;
    if (type.endsWith('wordprocessingml.styles+xml')) styles = name;
  }
  return { textParts: textParts.sort(), main, styles };
}

const isW = (node: Node | null, name: string): node is Element => !!node && node.nodeType === 1 && (node as Element).namespaceURI === W && (node as Element).localName === name;

function firstDescendant(root: Element, name: string): Element | undefined {
  const found = root.getElementsByTagNameNS(W, name);
  return found.length > 0 ? (found.item(0) as Element) : undefined;
}

/**
 * Put the cover page in front of the document, in a section of its own, and
 * make sure the document still starts on a fresh page numbered as before.
 */
function addCoverPage(documentXml: string, content: CoverPageContent): string {
  const parser = new DOMParser();
  const document = parser.parseFromString(documentXml, 'application/xml');
  const root = document.documentElement as unknown as Element;
  const body = firstDescendant(root, 'body');
  if (!body) throw new Error('the document has no body');
  const serializer = new XMLSerializer();

  // The document's first section: page size and margins for the cover; and it must now begin on a new page.
  const firstSection = firstDescendant(body, 'sectPr');
  let pageSetup = '<w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/>';
  if (firstSection) {
    const copied: string[] = [];
    for (let n = firstSection.firstChild; n; n = n.nextSibling) {
      if (isW(n, 'pgSz') || isW(n, 'pgMar')) copied.push(serializer.serializeToString(n as unknown as Parameters<typeof serializer.serializeToString>[0]).replace(/ xmlns:\w+="[^"]*"/g, ''));
    }
    if (copied.length > 0) pageSetup = copied.join('');

    const prefix = firstSection.prefix ?? 'w';
    const children: Element[] = [];
    for (let n = firstSection.firstChild; n; n = n.nextSibling) if (n.nodeType === 1) children.push(n as Element);
    // Schema order inside w:sectPr (the elements relevant here).
    const ORDER = ['headerReference', 'footerReference', 'endnotePr', 'footnotePr', 'type', 'pgSz', 'pgMar', 'paperSrc', 'pgBorders', 'lnNumType', 'pgNumType', 'cols', 'formProt', 'vAlign', 'noEndnote', 'titlePg', 'textDirection', 'bidi', 'rtlGutter', 'docGrid', 'printerSettings', 'sectPrChange'];
    const place = (element: Element): void => {
      const rank = ORDER.indexOf(element.localName!);
      const before = children.find((c) => c.namespaceURI === W && ORDER.indexOf(c.localName!) > rank);
      firstSection.insertBefore(element, before ?? null);
    };
    // Being the first section, its "start on" setting had no effect until now.
    const type = children.find((c) => isW(c, 'type'));
    if (type) type.setAttributeNS(W, `${prefix}:val`, 'nextPage');
    // Page numbers: the document's first page keeps the number it had (1 unless it said otherwise).
    let numbering = children.find((c) => isW(c, 'pgNumType'));
    if (!numbering) {
      numbering = document.createElementNS(W, `${prefix}:pgNumType`) as unknown as Element;
      place(numbering);
    }
    if (!numbering.getAttributeNS(W, 'start')) numbering.setAttributeNS(W, `${prefix}:start`, '1');
  }

  const fragment = parser.parseFromString(`<w:body xmlns:w="${W}">${coverPageXml(content, pageSetup)}</w:body>`, 'application/xml');
  const first = body.firstChild;
  for (let n = (fragment.documentElement as unknown as Element).firstChild; n; n = n.nextSibling) {
    body.insertBefore(document.importNode(n as unknown as Parameters<typeof document.importNode>[0], true) as unknown as Node, first);
  }
  return serializer.serializeToString(document);
}

function addCoverStyles(stylesXml: string): string {
  const defaultStyle = stylesXml.match(/<w:style\b(?=[^>]*w:type="paragraph")(?=[^>]*w:default="(?:1|true|on)")[^>]*>/)?.[0];
  const basedOn = defaultStyle ? attributeValue(defaultStyle, 'w:styleId') : undefined;
  if (!stylesXml.includes('</w:styles>')) throw new Error('the styles part could not be extended');
  return stylesXml.replace('</w:styles>', `${coverStylesXml(basedOn)}</w:styles>`);
}

export async function makeStaticRedline(trackedChanges: Uint8Array, options: StaticRedlineOptions): Promise<StaticRedline> {
  const files = await unzip(trackedChanges);
  const contentTypes = files.get('[Content_Types].xml');
  if (!contentTypes) throw new Error('the tracked-changes file has no content types');
  const { textParts, main, styles } = partsOfTypes(decoder.decode(contentTypes), files.keys(), TEXT_PART_TYPES);
  if (!main) throw new Error('the tracked-changes file has no main document part');
  if (!styles) throw new Error('the tracked-changes file has no styles part');

  const counts: ChangeCounts = { insertions: 0, deletions: 0, moves: 0 };
  const converted = new Map<string, string>();
  for (const part of textParts) {
    const result = convertPart(decoder.decode(files.get(part)!), { ...DEFAULT_STYLE_OPTIONS, ...options.style });
    converted.set(part, result.xml);
    counts.insertions += result.counts.inserted;
    counts.deletions += result.counts.deleted;
    counts.moves += result.counts.movedFrom;
  }

  const content: CoverPageContent = { originalName: options.originalName, modifiedName: options.modifiedName, comparedAt: options.comparedAt, ...counts };
  converted.set(main, addCoverPage(converted.get(main)!, content));
  for (const [part, xml] of converted) files.set(part, encoder.encode(xml));
  files.set(styles, encoder.encode(addCoverStyles(decoder.decode(files.get(styles)!))));

  return { docx: await zip(files), counts };
}
