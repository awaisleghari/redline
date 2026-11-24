/**
 * Check 7, structure: the output is a sound Word package.
 *   - it unzips and every XML part is well-formed;
 *   - the required parts exist and every part has a declared content type;
 *   - every relationship points to a part that exists, and every relationship
 *     id used in the text exists;
 *   - comment anchors are balanced and point to real comments;
 *   - every footnote and endnote reference resolves;
 *   - move markers are paired.
 * (Validation against the Office Open XML schema is done in the test suite
 * with Microsoft's Open XML SDK; see PLAN.md.)
 */
import { NS } from '../common/namespaces';
import { DocxPackage, PackageError, relsPathFor } from './package';
import { attribute, descendants, is, type XmlElement } from './xml';

const W = NS.w;

const WML = 'application/vnd.openxmlformats-officedocument.wordprocessingml';
/**
 * The content type Word requires for a part reached through each kind of
 * relationship. A part declared as anything else (for example as generic XML)
 * makes Word report the file as corrupt.
 */
const REQUIRED_TYPE: Record<string, string[]> = {
  officeDocument: [`${WML}.document.main+xml`, `${WML}.template.main+xml`, 'application/vnd.ms-word.document.macroEnabled.main+xml', 'application/vnd.ms-word.template.macroEnabledTemplate.main+xml'],
  header: [`${WML}.header+xml`],
  footer: [`${WML}.footer+xml`],
  footnotes: [`${WML}.footnotes+xml`],
  endnotes: [`${WML}.endnotes+xml`],
  comments: [`${WML}.comments+xml`],
  styles: [`${WML}.styles+xml`],
  numbering: [`${WML}.numbering+xml`],
  settings: [`${WML}.settings+xml`],
};

function contentTypeProblems(pkg: DocxPackage): string[] {
  const problems: string[] = [];
  if (!pkg.has('[Content_Types].xml')) return ['The file has no [Content_Types].xml part.'];
  const types = pkg.xml('[Content_Types].xml');
  const defaults = new Map<string, string>();
  const overrides = new Map<string, string>();
  for (const entry of types.children) {
    if (entry.ns !== NS.contentTypes) continue;
    const type = attribute(entry, '', 'ContentType') ?? '';
    if (entry.name === 'Default') defaults.set((attribute(entry, '', 'Extension') ?? '').toLowerCase(), type);
    if (entry.name === 'Override') overrides.set((attribute(entry, '', 'PartName') ?? '').replace(/^\//, ''), type);
  }
  const typeOf = (partName: string): string | undefined => {
    const extension = partName.includes('.') ? partName.slice(partName.lastIndexOf('.') + 1).toLowerCase() : '';
    return overrides.get(partName) ?? defaults.get(extension);
  };
  for (const partName of pkg.partNames) {
    if (partName === '[Content_Types].xml') continue;
    if (typeOf(partName) === undefined) problems.push(`The part "${partName}" has no declared content type.`);
  }
  for (const partName of overrides.keys()) {
    if (!pkg.has(partName)) problems.push(`A content type is declared for "${partName}", which does not exist.`);
  }
  // Parts with a defined role must be declared as what they are.
  const sources = ['', ...pkg.partNames.filter((p) => !p.endsWith('.rels'))];
  for (const source of sources) {
    for (const rel of pkg.relationships(source)) {
      if (rel.external || !rel.partName || !pkg.has(rel.partName)) continue;
      const suffix = rel.type.slice(rel.type.lastIndexOf('/') + 1);
      const required = REQUIRED_TYPE[suffix];
      const actual = typeOf(rel.partName);
      if (required && actual !== undefined && !required.includes(actual)) {
        problems.push(`The part "${rel.partName}" is used as the document's ${suffix} but is declared as a different kind of file.`);
      }
    }
  }
  return problems;
}

function relationshipProblems(pkg: DocxPackage): string[] {
  const problems: string[] = [];
  const sources = ['', ...pkg.partNames.filter((p) => !p.endsWith('.rels'))];
  for (const source of sources) {
    const seen = new Set<string>();
    for (const rel of pkg.relationships(source)) {
      const owner = source === '' ? 'the package' : `"${source}"`;
      if (rel.id === '') problems.push(`A relationship of ${owner} has no id.`);
      if (seen.has(rel.id)) problems.push(`Two relationships of ${owner} share the id "${rel.id}".`);
      seen.add(rel.id);
      if (!rel.external && (!rel.partName || !pkg.has(rel.partName))) {
        problems.push(`A relationship of ${owner} points to "${rel.target}", which does not exist.`);
      }
    }
  }
  // A .rels file must belong to a part that exists.
  for (const partName of pkg.partNames) {
    if (!partName.endsWith('.rels') || partName === '_rels/.rels') continue;
    const owner = partName.replace(/_rels\/([^/]+)\.rels$/, '$1');
    if (relsPathFor(owner) === partName && !pkg.has(owner)) problems.push(`"${partName}" holds relationships for "${owner}", which does not exist.`);
  }
  return problems;
}

/** Problems inside one Word-markup part (document, header, footer, notes, comments). */
function markupProblems(pkg: DocxPackage, partName: string, root: XmlElement, commentIds: Set<string> | undefined, noteIds: Record<'footnote' | 'endnote', Set<string> | undefined>): string[] {
  const problems: string[] = [];
  const relIds = new Set(pkg.relationships(partName).map((r) => r.id));
  const commentStarts = new Map<string, number>();
  const commentEnds = new Map<string, number>();
  const moveNames: Record<'From' | 'To', Map<string, number>> = { From: new Map(), To: new Map() };
  const rangeStarts: Record<string, Set<string>> = { moveFrom: new Set(), moveTo: new Set() };
  const rangeEnds: Record<string, string[]> = { moveFrom: [], moveTo: [] };
  const bump = (map: Map<string, number>, key: string): void => void map.set(key, (map.get(key) ?? 0) + 1);

  // Deleted text must be marked as deleted text, and only there: Word's treatment
  // of the mismatched forms is not defined, so what a reader would see is unknown.
  const misplaced = (element: XmlElement, inDeletion: boolean): void => {
    for (const c of element.children) {
      if (c.ns !== W) {
        misplaced(c, inDeletion);
        continue;
      }
      if (c.name === 'delText' && !inDeletion) problems.push(`"${partName}" has deleted-text markup outside any deletion.`);
      if (c.name === 't' && inDeletion) problems.push(`"${partName}" has ordinary text directly inside a deletion (it should be marked as deleted text).`);
      // (A moved-from block keeps ordinary text markup; only w:del uses deleted-text markup.)
      const next = c.name === 'del' ? true : ['txbxContent', 'ins', 'moveTo', 'moveFrom'].includes(c.name) ? false : inDeletion;
      misplaced(c, next);
    }
  };
  misplaced(root, false);

  // Whether spaces are kept must be stated on each piece of text itself. An output that left it to a
  // setting on some outer element would read differently to any program that does not apply such a
  // setting to everything inside, so it is refused. (Uploads may be written that way; outputs may not.)
  const MATH = 'http://schemas.openxmlformats.org/officeDocument/2006/math';
  const carriesText = (e: XmlElement): boolean => (e.ns === W && ['t', 'delText', 'instrText', 'delInstrText'].includes(e.name)) || (e.ns === MATH && e.name === 't');
  for (const node of [root, ...descendants(root)]) {
    if (node.attributes.has('http://www.w3.org/XML/1998/namespace|space') && !carriesText(node)) {
      problems.push(`"${partName}" sets how white space is treated on something other than a piece of text.`);
      break;
    }
  }

  for (const node of descendants(root)) {
    // Every relationship id used in the markup must exist.
    for (const [key, value] of node.attributes) {
      if (key.startsWith(`${NS.r}|`) && value !== '' && !relIds.has(value)) {
        problems.push(`"${partName}" uses relationship "${value}", which is not defined.`);
      }
    }
    if (node.ns !== W) continue;
    const id = attribute(node, W, 'id') ?? '';
    switch (node.name) {
      case 'commentRangeStart':
        bump(commentStarts, id);
        break;
      case 'commentRangeEnd':
        bump(commentEnds, id);
        break;
      case 'commentReference':
        if (!commentIds?.has(id)) problems.push(`A comment marker in "${partName}" points to comment ${id}, which does not exist.`);
        break;
      case 'footnoteReference':
        if (!noteIds.footnote?.has(id)) problems.push(`A footnote reference in "${partName}" points to footnote ${id}, which does not exist.`);
        break;
      case 'endnoteReference':
        if (!noteIds.endnote?.has(id)) problems.push(`An endnote reference in "${partName}" points to endnote ${id}, which does not exist.`);
        break;
      case 'moveFromRangeStart':
        bump(moveNames.From, attribute(node, W, 'name') ?? '');
        rangeStarts.moveFrom!.add(id);
        break;
      case 'moveToRangeStart':
        bump(moveNames.To, attribute(node, W, 'name') ?? '');
        rangeStarts.moveTo!.add(id);
        break;
      case 'moveFromRangeEnd':
        rangeEnds.moveFrom!.push(id);
        break;
      case 'moveToRangeEnd':
        rangeEnds.moveTo!.push(id);
        break;
      default:
        break;
    }
  }

  for (const id of new Set([...commentStarts.keys(), ...commentEnds.keys()])) {
    const starts = commentStarts.get(id) ?? 0;
    const ends = commentEnds.get(id) ?? 0;
    if (starts !== 1 || ends !== 1) {
      problems.push(`Comment ${id} in "${partName}" is not anchored properly (${starts} start marker${starts === 1 ? '' : 's'}, ${ends} end marker${ends === 1 ? '' : 's'}).`);
    }
    if (!commentIds?.has(id)) problems.push(`A comment anchor in "${partName}" points to comment ${id}, which does not exist.`);
  }
  for (const kind of ['moveFrom', 'moveTo'] as const) {
    const starts = rangeStarts[kind]!;
    const ends = rangeEnds[kind]!;
    if (starts.size !== ends.length || ends.some((id) => !starts.has(id))) {
      problems.push(`A moved block in "${partName}" has unbalanced range markers.`);
    }
  }
  // Every move origin needs a destination of the same name, and vice versa. The
  // engine gives each paragraph of a multi-paragraph move its own pair of
  // markers under one shared name, so a name may be used several times; what
  // must hold is that origins and destinations match up one for one.
  for (const name of new Set([...moveNames.From.keys(), ...moveNames.To.keys()])) {
    const origins = moveNames.From.get(name) ?? 0;
    const destinations = moveNames.To.get(name) ?? 0;
    if (name === '' || origins !== destinations) {
      problems.push(`A moved block in "${partName}" is not paired: its origin and its destination do not match up.`);
    }
  }
  return problems;
}

function idsIn(pkg: DocxPackage, partName: string | undefined, element: string): Set<string> | undefined {
  if (!partName || !pkg.has(partName)) return undefined;
  const ids = new Set<string>();
  for (const node of pkg.xml(partName).children) if (is(node, W, element)) ids.add(attribute(node, W, 'id') ?? '');
  return ids;
}

/** All structural problems of a package, most fundamental first. Empty when sound. */
export function structureProblems(pkg: DocxPackage): string[] {
  const problems: string[] = [];

  // Every XML part must be well-formed before anything else can be trusted.
  for (const partName of pkg.partNames) {
    if (!/\.(xml|rels)$/.test(partName)) continue;
    try {
      pkg.xml(partName);
    } catch (error) {
      if (error instanceof PackageError) problems.push(error.message);
      else throw error;
    }
  }
  if (problems.length > 0) return problems;

  if (!pkg.has('_rels/.rels')) return ['The file has no package relationships (_rels/.rels).'];
  let main: string;
  try {
    main = pkg.mainDocumentPart();
  } catch (error) {
    return [error instanceof PackageError ? error.message : String(error)];
  }
  if (!is(pkg.xml(main), W, 'document')) problems.push('The main document part is not a Word document.');

  problems.push(...contentTypeProblems(pkg), ...relationshipProblems(pkg));

  const partOf = (suffix: string): string | undefined => pkg.relationshipsOfType(main, suffix)[0]?.partName;
  const commentIds = idsIn(pkg, partOf('comments'), 'comment');
  const noteIds = { footnote: idsIn(pkg, partOf('footnotes'), 'footnote'), endnote: idsIn(pkg, partOf('endnotes'), 'endnote') };

  for (const partName of pkg.partNames) {
    if (!partName.endsWith('.xml') || partName === '[Content_Types].xml') continue;
    const root = pkg.xml(partName);
    if (root.ns !== W) continue;
    if (!['document', 'hdr', 'ftr', 'footnotes', 'endnotes', 'comments'].includes(root.name)) continue;
    problems.push(...markupProblems(pkg, partName, root, commentIds, noteIds));
  }
  return problems;
}
