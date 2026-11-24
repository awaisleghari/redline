/**
 * The Checker's own reading of a .docx as a package: its parts, their declared
 * content types and their relationships. Nothing here trusts the Producer.
 */
import { NS } from '../common/namespaces';
import { unzip, type ZipEntries } from '../common/zip';
import { attribute, parseXml, XmlError, type XmlElement } from './xml';

export interface Relationship {
  id: string;
  /** Full relationship type URI. */
  type: string;
  /** Raw Target attribute. */
  target: string;
  external: boolean;
  /** Resolved part name for internal targets (no leading slash). */
  partName?: string;
}

export class PackageError extends Error {}

/** "word/document.xml" -> "word/_rels/document.xml.rels"; "" (package root) -> "_rels/.rels". */
export function relsPathFor(partName: string): string {
  const slash = partName.lastIndexOf('/');
  const folder = slash >= 0 ? partName.slice(0, slash + 1) : '';
  const file = slash >= 0 ? partName.slice(slash + 1) : partName;
  return `${folder}_rels/${file}.rels`;
}

/** Resolve a relationship target against the part that owns the relationship. */
export function resolveTarget(sourcePart: string, target: string): string {
  let path = target;
  try {
    path = decodeURIComponent(target);
  } catch {
    // Keep the raw target; it will simply fail to resolve.
  }
  path = path.split('#')[0] ?? '';
  const base = path.startsWith('/') ? [] : sourcePart.split('/').slice(0, -1);
  for (const segment of path.split('/')) {
    if (segment === '' || segment === '.') continue;
    if (segment === '..') base.pop();
    else base.push(segment);
  }
  return base.join('/');
}

export class DocxPackage {
  private readonly parsed = new Map<string, XmlElement>();
  private readonly rels = new Map<string, Relationship[]>();

  private constructor(readonly entries: ZipEntries) {}

  static async open(bytes: Uint8Array): Promise<DocxPackage> {
    let entries: ZipEntries;
    try {
      entries = await unzip(bytes);
    } catch {
      throw new PackageError('The file is not a readable .docx (it could not be unzipped).');
    }
    return new DocxPackage(entries);
  }

  get partNames(): string[] {
    return [...this.entries.keys()].sort();
  }

  /**
   * Whether the package holds this part. Part names are compared without regard to capitals, as the file format
   * requires (a document-management system may write "/customXML/item2.xml" for the part "customXml/item2.xml",
   * and Word opens it).
   */
  has(partName: string): boolean {
    if (this.entries.has(partName)) return true;
    const wanted = partName.toLowerCase();
    for (const name of this.entries.keys()) if (name.toLowerCase() === wanted) return true;
    return false;
  }

  /** Parsed XML of a part. Throws PackageError if the part is missing or not well-formed. */
  xml(partName: string): XmlElement {
    const cached = this.parsed.get(partName);
    if (cached) return cached;
    const bytes = this.entries.get(partName);
    if (!bytes) throw new PackageError(`The part "${partName}" is missing.`);
    try {
      const root = parseXml(bytes);
      this.parsed.set(partName, root);
      return root;
    } catch (error) {
      if (error instanceof XmlError) throw new PackageError(`The part "${partName}" is not well-formed XML: ${error.message}`);
      throw error;
    }
  }

  /** Relationships owned by a part ("" for the package itself). Empty when it has none. */
  relationships(sourcePart: string): Relationship[] {
    const cached = this.rels.get(sourcePart);
    if (cached) return cached;
    const path = relsPathFor(sourcePart);
    const list: Relationship[] = [];
    if (this.entries.has(path)) {
      for (const element of this.xml(path).children) {
        if (element.ns !== NS.rels || element.name !== 'Relationship') continue;
        const target = attribute(element, '', 'Target') ?? '';
        const external = attribute(element, '', 'TargetMode') === 'External';
        list.push({
          id: attribute(element, '', 'Id') ?? '',
          type: attribute(element, '', 'Type') ?? '',
          target,
          external,
          partName: external ? undefined : resolveTarget(sourcePart, target),
        });
      }
    }
    this.rels.set(sourcePart, list);
    return list;
  }

  /** Relationships of a part whose type ends with "/<suffix>". */
  relationshipsOfType(sourcePart: string, suffix: string): Relationship[] {
    return this.relationships(sourcePart).filter((r) => r.type.endsWith(`/${suffix}`));
  }

  /** The main document part, found through the package relationships (never assumed). */
  mainDocumentPart(): string {
    const rel = this.relationshipsOfType('', 'officeDocument')[0];
    if (!rel?.partName) throw new PackageError('The file has no main document part.');
    if (!this.entries.has(rel.partName)) throw new PackageError(`The main document part "${rel.partName}" is missing.`);
    return rel.partName;
  }
}
