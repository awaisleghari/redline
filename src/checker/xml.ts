/**
 * The Checker's own XML reader: a strict, namespace-aware parser (saxes)
 * feeding a minimal read-only tree. Web Workers have no DOMParser, and the
 * Checker must not share document-reading code with the Producer.
 */
import { SaxesParser } from 'saxes';

export interface XmlElement {
  /** Namespace name ('' when none). */
  ns: string;
  /** Local name, without prefix. */
  name: string;
  /** Attributes keyed "namespace|local" (namespace '' when none). */
  attributes: Map<string, string>;
  children: XmlElement[];
  /** Text directly inside this element (all text nodes joined). */
  text: string;
  /**
   * Whether white space here is to be kept: the nearest xml:space setting on this element or on any
   * element around it, as the XML standard defines it. Some programs set it once, on the document's
   * outermost element, instead of on every piece of text.
   */
  preserveSpace: boolean;
}

export class XmlError extends Error {}

export function decodeXml(bytes: Uint8Array): string {
  if (bytes.length >= 2 && bytes[0] === 0xff && bytes[1] === 0xfe) return new TextDecoder('utf-16le').decode(bytes);
  if (bytes.length >= 2 && bytes[0] === 0xfe && bytes[1] === 0xff) return new TextDecoder('utf-16be').decode(bytes);
  return new TextDecoder('utf-8', { fatal: true }).decode(bytes);
}

export function parseXml(source: string | Uint8Array): XmlElement {
  let text: string;
  try {
    text = typeof source === 'string' ? source : decodeXml(source);
  } catch {
    throw new XmlError('the part is not valid UTF-8 text');
  }
  if (text.charCodeAt(0) === 0xfeff) text = text.slice(1);

  const parser = new SaxesParser({ xmlns: true });
  const stack: XmlElement[] = [];
  let root: XmlElement | undefined;
  let failure: Error | undefined;

  parser.on('error', (error) => {
    failure ??= error;
  });
  parser.on('opentag', (tag) => {
    const element: XmlElement = { ns: tag.uri, name: tag.local, attributes: new Map(), children: [], text: '', preserveSpace: stack[stack.length - 1]?.preserveSpace ?? false };
    for (const attribute of Object.values(tag.attributes)) {
      if (attribute.prefix === 'xmlns' || (attribute.prefix === '' && attribute.local === 'xmlns')) continue;
      element.attributes.set(`${attribute.uri}|${attribute.local}`, attribute.value);
      if (attribute.uri === 'http://www.w3.org/XML/1998/namespace' && attribute.local === 'space') element.preserveSpace = attribute.value === 'preserve';
    }
    const parent = stack[stack.length - 1];
    if (parent) parent.children.push(element);
    else root = element;
    stack.push(element);
  });
  parser.on('closetag', () => {
    stack.pop();
  });
  const addText = (value: string): void => {
    const current = stack[stack.length - 1];
    if (current) current.text += value;
  };
  parser.on('text', addText);
  parser.on('cdata', addText);

  try {
    parser.write(text).close();
  } catch (error) {
    failure ??= error instanceof Error ? error : new Error(String(error));
  }
  if (failure) throw new XmlError(failure.message.split('\n')[0] ?? 'not well-formed XML');
  if (!root) throw new XmlError('the part has no XML content');
  return root;
}

export const is = (element: XmlElement, ns: string, name: string): boolean => element.ns === ns && element.name === name;

export function child(element: XmlElement, ns: string, name: string): XmlElement | undefined {
  return element.children.find((c) => c.ns === ns && c.name === name);
}

export function attribute(element: XmlElement, ns: string, name: string): string | undefined {
  return element.attributes.get(`${ns}|${name}`);
}

/** Every element in the subtree, parents before children, in document order. */
export function* descendants(element: XmlElement): Generator<XmlElement> {
  for (const c of element.children) {
    yield c;
    yield* descendants(c);
  }
}
