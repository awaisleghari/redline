/**
 * Some programs (PDF-to-Word converters in particular) say "keep white space"
 * once, on a document's outermost element, instead of on every piece of text.
 * That is valid, and Word honours it. The compare engine does not: it copies
 * the pieces of text into its output without the outer setting, and words
 * there run together ("TERM" + " SHEET" becomes "TERMSHEET").
 *
 * So, before a document is given to the engine, the setting is written onto
 * each piece of text it applies to. Nothing else is changed, and a document
 * that does not use the outer setting is passed on byte for byte.
 */
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { unzip, zip } from '../common/zip';
import { keepsSpace } from './original-reader';

/** Elements whose characters are the document's text or field codes. */
const TEXT_ELEMENTS = new Set(['t', 'delText', 'instrText', 'delInstrText']);
/** xml:space on any element that is not itself one of those. */
const OUTER_SETTING = /<(?![\w.-]+:(?:t|delText|instrText|delInstrText)[\s>\/])[^<>]*\sxml:space\s*=/;

export function makeSpacesExplicitInPart(xml: string): string {
  if (!OUTER_SETTING.test(xml)) return xml;
  const document = new DOMParser().parseFromString(xml, 'application/xml');
  const all = document.getElementsByTagName('*');
  let changed = false;
  for (let i = 0; i < all.length; i++) {
    const element = all.item(i)!;
    if (!TEXT_ELEMENTS.has(element.localName ?? '') || element.hasAttribute('xml:space')) continue;
    if (keepsSpace(element as unknown as Parameters<typeof keepsSpace>[0])) {
      element.setAttribute('xml:space', 'preserve');
      changed = true;
    }
  }
  return changed ? new XMLSerializer().serializeToString(document) : xml;
}

export async function makeSpacesExplicit(docx: Uint8Array): Promise<Uint8Array> {
  const files = await unzip(docx);
  const decoder = new TextDecoder();
  const encoder = new TextEncoder();
  let changed = false;
  for (const [name, bytes] of files) {
    if (!name.endsWith('.xml') || name === '[Content_Types].xml') continue;
    const xml = decoder.decode(bytes);
    const explicit = makeSpacesExplicitInPart(xml);
    if (explicit !== xml) {
      files.set(name, encoder.encode(explicit));
      changed = true;
    }
  }
  return changed ? zip(files) : docx;
}
