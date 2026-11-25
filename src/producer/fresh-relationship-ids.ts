/**
 * Producer: works round a fault in the compare engine with documents whose
 * internal reference names look like the engine's own (found on real
 * contracts, held-out run 1).
 *
 * Inside a Word file, a link or a picture is tied to its address or its image
 * file by a short name (a "relationship id": rId7, R18). Word writes names
 * beginning "rId". The engine writes names beginning "R" for what it adds. A
 * Word file that the engine itself produced (a redline made by this tool, sent
 * out and then compared again) therefore already holds R-names, and the
 * engine's new names collide with them: a link ended up tied to a picture
 * instead of its web address.
 *
 * So, in the copies of the documents given to the engine, every name of that
 * form is replaced by one the engine never writes. Nothing a reader sees
 * changes: the names are internal, and each still ties the same thing to the
 * same target. Only the names themselves are edited; every other character is
 * passed on as it was, and a document with no such name is passed on byte for
 * byte.
 */
import { NS } from '../common/namespaces';
import { unzip, zip } from '../common/zip';

/** The form of name the engine writes: "R" and a number (decimal or hexadecimal). */
const ENGINE_STYLE = /^R[0-9A-Fa-f]+$/;
const RELATIONSHIP_NAMESPACES = [NS.r, 'http://purl.oclc.org/ooxml/officeDocument/relationships'];
const OFFICE_NAMESPACE = 'urn:schemas-microsoft-com:office:office';
const DECLARATION = /\sxmlns:([\w.-]+)\s*=\s*(?:"([^"]*)"|'([^']*)')/g;
const ID_IN_RELS = /(<Relationship\b[^>]*?\sId\s*=\s*)(?:"([^"]*)"|'([^']*)')/g;

/** Comments, CDATA, processing instructions, declarations and closing tags are stepped over; an opening tag's attributes are captured. */
const TAG = /<!--[\s\S]*?-->|<!\[CDATA\[[\s\S]*?\]\]>|<\?[\s\S]*?\?>|<![^>]*>|<\/[^>]*>|<[^\s\/>]+((?:"[^"]*"|'[^']*'|[^>"'])*)>/g;

const escapeForPattern = (text: string): string => text.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');

/** The new names for one part's relationships: engine-style names only, each mapped to a name not in use. */
function renamings(rels: string): Map<string, string> {
  const ids: string[] = [];
  for (const match of rels.matchAll(ID_IN_RELS)) ids.push(match[2] ?? match[3] ?? '');
  const taken = new Set(ids);
  const renamed = new Map<string, string>();
  let next = 1;
  for (const id of ids) {
    if (!ENGINE_STYLE.test(id) || renamed.has(id)) continue;
    while (taken.has(`rIdU${next}`)) next++;
    renamed.set(id, `rIdU${next}`);
    taken.add(`rIdU${next}`);
  }
  return renamed;
}

/** A part's relationships file with the names replaced. */
export function renameInRels(rels: string, renamed: Map<string, string>): string {
  return rels.replace(ID_IN_RELS, (whole, before: string, double: string | undefined, single: string | undefined) => {
    const to = renamed.get(double ?? single ?? '');
    return to ? `${before}"${to}"` : whole;
  });
}

/**
 * A part with every use of those names replaced. A name is used by an attribute in the relationships namespace
 * (r:id, r:embed, r:link and the rest) or, in older drawings, by o:relid.
 */
export function renameInPart(xml: string, renamed: Map<string, string>): string {
  const names: string[] = [];
  for (const match of xml.matchAll(DECLARATION)) {
    const namespace = match[2] ?? match[3] ?? '';
    if (RELATIONSHIP_NAMESPACES.includes(namespace)) names.push(`${escapeForPattern(match[1]!)}:[\\w.-]+`);
    else if (namespace === OFFICE_NAMESPACE) names.push(`${escapeForPattern(match[1]!)}:relid`);
  }
  if (names.length === 0) return xml;
  const use = new RegExp(`(\\s(?:${[...new Set(names)].join('|')})\\s*=\\s*)(?:"([^"]*)"|'([^']*)')`, 'g');
  // Only inside opening tags: the same characters in the document's text are text, and stay as they are.
  return xml.replace(TAG, (tag, attributes: string | undefined) =>
    attributes === undefined
      ? tag
      : tag.replace(use, (whole, before: string, double: string | undefined, single: string | undefined) => {
          const to = renamed.get(double ?? single ?? '');
          return to ? `${before}"${to}"` : whole;
        }),
  );
}

/** The document with engine-style relationship names replaced. A document without any is returned as the very same bytes. */
export async function freshRelationshipIds(docx: Uint8Array): Promise<Uint8Array> {
  const files = await unzip(docx);
  const decoder = new TextDecoder();
  const encoder = new TextEncoder();
  let changed = false;
  for (const [name, bytes] of [...files]) {
    const place = /^(.*\/)?_rels\/([^/]+)\.rels$/.exec(name);
    if (!place) continue;
    const partName = `${place[1] ?? ''}${place[2]}`;
    const part = files.get(partName);
    // Only the parts that are XML can use the names (a picture or an embedded file has no relationships of its own to name).
    if (!part || !/\.(xml|vml)$/i.test(partName)) continue;
    const rels = decoder.decode(bytes);
    const renamed = renamings(rels);
    if (renamed.size === 0) continue;
    files.set(name, encoder.encode(renameInRels(rels, renamed)));
    files.set(partName, encoder.encode(renameInPart(decoder.decode(part), renamed)));
    changed = true;
  }
  return changed ? zip(files) : docx;
}
