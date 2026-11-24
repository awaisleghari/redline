/**
 * Some programs (PDF-to-Word converters in particular) say "keep white space"
 * once, on a document's outermost element, instead of on every piece of text.
 * That is valid, and Word honours it. The compare engine does not: it copies
 * the pieces of text into its output without the outer setting, and words
 * there run together ("TERM" + " SHEET" becomes "TERMSHEET").
 *
 * So, before a document is given to the engine, the setting is written onto
 * each piece of text it applies to and taken off everything else. The text
 * reads exactly as before, and nothing the engine produces can depend on an
 * outer setting. Only the opening tags concerned are edited: every other
 * character of the part is passed on as it was. A document that does not use
 * an outer setting is passed on byte for byte.
 */
import { NS } from '../common/namespaces';
import { unzip, zip } from '../common/zip';

const MATH = 'http://schemas.openxmlformats.org/officeDocument/2006/math';
/** Elements whose characters are the document's text or field codes, and which may carry the setting themselves. */
const isTextElement = (ns: string | undefined, local: string): boolean =>
  (ns === NS.w && ['t', 'delText', 'instrText', 'delInstrText'].includes(local)) || (ns === MATH && local === 't');

/** xml:space on any element that is not named like one of those: the quick test for "this part needs looking at". */
const OUTER_SETTING = /<(?![\w.-]+:(?:t|delText|instrText|delInstrText)[\s>\/])[^<>!?][^<>]*\sxml:space\s*=/;
/** Comments, CDATA, processing instructions and declarations are stepped over; tags are captured. */
const TOKEN = /<!--[\s\S]*?-->|<!\[CDATA\[[\s\S]*?\]\]>|<\?[\s\S]*?\?>|<![^>]*>|<(\/?)([^\s\/>]+)((?:"[^"]*"|'[^']*'|[^>"'])*?)(\/?)>/g;
const ATTRIBUTE = /([^\s=]+)\s*=\s*(?:"([^"]*)"|'([^']*)')/g;
const SPACE_ATTRIBUTE = /\s+xml:space\s*=\s*(?:"[^"]*"|'[^']*')/;

export function makeSpacesExplicitInPart(xml: string): string {
  if (!OUTER_SETTING.test(xml)) return xml;
  const stack: { preserve: boolean; prefixes: Map<string, string> }[] = [];
  let out = '';
  let last = 0;
  let wordMarkup: boolean | undefined;
  TOKEN.lastIndex = 0;
  for (let match = TOKEN.exec(xml); match; match = TOKEN.exec(xml)) {
    const [tag, closing, name, attributeText, selfClosing] = match;
    if (name === undefined) continue; // a comment, CDATA section, instruction or declaration
    if (closing) {
      stack.pop();
      continue;
    }
    const parent = stack[stack.length - 1];
    const prefixes = new Map(parent?.prefixes);
    let own: string | undefined;
    for (const attribute of (attributeText ?? '').matchAll(ATTRIBUTE)) {
      const value = attribute[2] ?? attribute[3] ?? '';
      if (attribute[1] === 'xmlns') prefixes.set('', value);
      else if (attribute[1]!.startsWith('xmlns:')) prefixes.set(attribute[1]!.slice(6), value);
      else if (attribute[1] === 'xml:space') own = value;
    }
    const colon = name.indexOf(':');
    const ns = prefixes.get(colon === -1 ? '' : name.slice(0, colon));
    const local = colon === -1 ? name : name.slice(colon + 1);
    // Only parts written in Word's own markup are touched (a document, header, footer or notes part).
    wordMarkup ??= ns === NS.w;
    if (!wordMarkup) return xml;
    const preserve = own === undefined ? (parent?.preserve ?? false) : own === 'preserve';
    if (!selfClosing) stack.push({ preserve, prefixes });

    let edited = tag;
    if (isTextElement(ns, local)) {
      if (own === undefined && preserve) edited = `<${name}${attributeText ?? ''} xml:space="preserve"${selfClosing ?? ''}>`;
    } else if (own !== undefined) {
      edited = tag.replace(SPACE_ATTRIBUTE, '');
    }
    if (edited !== tag) {
      out += xml.slice(last, match.index) + edited;
      last = match.index + tag.length;
    }
  }
  return last === 0 ? xml : out + xml.slice(last);
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
