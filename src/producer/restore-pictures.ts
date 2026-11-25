/**
 * Producer: repairs a fault in the compare engine's handling of pictures
 * stored as text (SVG; found on real contracts, held-out run 1).
 *
 * The engine copies ordinary pictures byte for byte, but it reads an SVG
 * picture as XML and writes it out again: the same drawing, with a different
 * heading line, indentation and line ends. The picture a reader sees is the
 * same, but the file is no longer the one that was uploaded, and nothing can
 * then show that the picture was not altered.
 *
 * The repair, on the engine's output: an SVG picture that says exactly what an
 * uploaded SVG picture says (the same elements, attributes and text, in the
 * same order) is given back the uploaded picture's own bytes. A picture that
 * matches no upload is left as the engine wrote it, for the Checker to judge.
 */
import { DOMParser } from '@xmldom/xmldom';

const decoder = new TextDecoder();
const isSvg = (name: string): boolean => /\.svg$/i.test(name);

/** What an SVG picture says, whatever its layout: undefined if it cannot be read. */
function saidBy(bytes: Uint8Array): string | undefined {
  let broken = false;
  let root: Element | null | undefined;
  try {
    root = new DOMParser({ onError: () => (broken = true) }).parseFromString(decoder.decode(bytes).replace(/^\uFEFF/, ''), 'application/xml').documentElement as unknown as Element | null;
  } catch {
    return undefined; // not readable as XML (the parser throws on a fatal fault): left exactly as it is
  }
  if (broken || !root) return undefined;
  const say = (node: Node): unknown => {
    if (node.nodeType === 3 || node.nodeType === 4) return (node.nodeValue ?? '').replace(/\r\n?/g, '\n');
    if (node.nodeType !== 1) return null; // comments and instructions draw nothing
    const element = node as unknown as Element;
    const attributes: string[][] = [];
    for (let i = 0; i < element.attributes.length; i++) {
      const a = element.attributes.item(i)!;
      attributes.push([a.namespaceURI ?? '', a.localName ?? a.name, a.value]);
    }
    attributes.sort((x, y) => (x.join('\u0000') < y.join('\u0000') ? -1 : 1));
    const inside: unknown[] = [];
    for (let child = element.firstChild; child; child = child.nextSibling) {
      const said = say(child as unknown as Node);
      // Line breaks and indentation between elements are layout, not content.
      if (said === null || (typeof said === 'string' && said.trim() === '')) continue;
      inside.push(said);
    }
    return [element.namespaceURI ?? '', element.localName, attributes, inside];
  };
  return JSON.stringify(say(root as unknown as Node));
}

/** Give each rewritten SVG picture in `output` the bytes of the uploaded picture it is a rewriting of. */
export function restorePictures(output: Map<string, Uint8Array>, uploads: Map<string, Uint8Array>[]): void {
  const rewritten = [...output.keys()].filter(isSvg);
  if (rewritten.length === 0) return;
  const uploaded = new Map<string, Uint8Array>();
  for (const upload of uploads) {
    for (const [name, bytes] of upload) {
      if (!isSvg(name)) continue;
      const said = saidBy(bytes);
      if (said !== undefined && !uploaded.has(said)) uploaded.set(said, bytes);
    }
  }
  for (const name of rewritten) {
    const said = saidBy(output.get(name)!);
    const own = said === undefined ? undefined : uploaded.get(said);
    if (own) output.set(name, own);
  }
}
