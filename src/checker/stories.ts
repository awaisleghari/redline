/**
 * The Checker's own reading of a Word file's text.
 *
 * A "story" is one independent flow of text: the body, a header, a footer, a
 * footnote, an endnote, or the text boxes anchored in one of those. Each story
 * is a list of paragraphs, each paragraph a plain string.
 *
 * A file is read in one of two views:
 *   - 'accepted': as if every tracked change were accepted;
 *   - 'rejected': as if every tracked change were rejected.
 * This is the Checker's own accept-all / reject-all, written separately from
 * anything in the Producer.
 *
 * Only the way text is split into runs is normalised away. Every character
 * (spaces, punctuation, digits, tabs, line breaks) is kept.
 */
import { NS } from '../common/namespaces';
import { DocxPackage } from './package';
import { attribute, child, is, type XmlElement } from './xml';

export type View = 'accepted' | 'rejected';

export type StoryKind = 'body' | 'header' | 'footer' | 'footnote' | 'endnote' | 'text box';

export interface Story {
  /** Stable identity used to match a story across two files. */
  key: string;
  kind: StoryKind;
  /** Plain-English name for reports, e.g. "footnote 2" or "header (section 1)". */
  label: string;
  paragraphs: string[];
}

export interface StoryReading {
  stories: Story[];
  /** Every part of the package this reading took text from. */
  partsRead: Set<string>;
  /** Footnote / endnote ids that were reached through a reference, per notes part. */
  noteIdsRead: Map<string, Set<string>>;
  /** Things that prevented a complete reading (dangling references and the like). */
  problems: string[];
  /**
   * Where each hyperlink points, in reading order, with the text it covers.
   * Links whose text is entirely removed in this view are left out. (A changed
   * destination is tracked as the old link's text deleted and the new link's
   * text inserted, so each view has exactly its own document's links.)
   */
  links: string[];
}

/** Markers standing in for things that are not characters but must not vanish unnoticed. */
export const MARK = {
  footnote: '⟦footnote⟧',
  endnote: '⟦endnote⟧',
  object: '⟦object⟧',
  pageBreak: '⟦page break⟧',
  columnBreak: '⟦column break⟧',
  fieldBegin: '⟦field⟧',
  fieldResult: '⟦field result⟧',
  fieldEnd: '⟦end of field⟧',
  /** Table structure, as entries of their own between the paragraphs. */
  table: '⟦table⟧',
  row: '⟦table row⟧',
  cell: '⟦table cell⟧',
  tableEnd: '⟦end of table⟧',
  /** Start of one text box, as an entry of its own in a text-box story. */
  textBox: '⟦text box⟧',
} as const;

const STRUCTURE_MARKS: ReadonlySet<string> = new Set([MARK.table, MARK.row, MARK.cell, MARK.tableEnd, MARK.textBox]);
/** Is this story entry a table-structure marker rather than a paragraph? */
export const isStructureMark = (entry: string): boolean => STRUCTURE_MARKS.has(entry);

const XML_NS = 'http://www.w3.org/XML/1998/namespace';
/**
 * Namespaces whose elements may sit between a paragraph and its text: Word's
 * own, and those of drawings, shapes and text boxes. Text reached only through
 * an element of any other namespace may or may not be shown by Word (it drops
 * unknown elements it is told it may ignore), so it cannot be verified.
 */
const KNOWN_WRAPPER_NAMESPACES: ReadonlySet<string> = new Set([
  NS.w,
  NS.mc,
  'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
  'http://schemas.openxmlformats.org/drawingml/2006/main',
  'http://schemas.openxmlformats.org/drawingml/2006/picture',
  'http://schemas.microsoft.com/office/word/2010/wordprocessingShape',
  'http://schemas.microsoft.com/office/word/2010/wordprocessingGroup',
  'http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas',
  'http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing',
  'http://schemas.openxmlformats.org/officeDocument/2006/math',
  'urn:schemas-microsoft-com:vml',
  'urn:schemas-microsoft-com:office:office',
  'urn:schemas-microsoft-com:office:word',
]);
const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const M_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/math';
const V_NS = 'urn:schemas-microsoft-com:vml';

/** Does an element hold any document markup at all (text, a note reference, a tab, a field, an equation, a shape...)? */
function holdsWordMarkup(element: XmlElement): boolean {
  return element.children.some((c) => c.ns === W || c.ns === M_NS || c.ns === A_NS || c.ns === V_NS || holdsWordMarkup(c));
}

/** A short, stable fingerprint of a file's bytes (FNV-1a, two passes), to tell whether a picture or embedded object changed. */
function fingerprint(bytes: Uint8Array): string {
  let a = 0x811c9dc5;
  let b = 0x01000193;
  for (let i = 0; i < bytes.length; i++) {
    a = Math.imul(a ^ bytes[i]!, 0x01000193) >>> 0;
    b = Math.imul(b ^ bytes[bytes.length - 1 - i]!, 0x811c9dc5) >>> 0;
  }
  return `${bytes.length}:${a.toString(16)}:${b.toString(16)}`;
}

function allText(element: XmlElement, into: string[]): void {
  if (element.text.trim() !== '') into.push(element.text);
  for (const c of element.children) allText(c, into);
}

/** The text of a w:t or w:delText as Word shows it: outer whitespace is dropped unless the element says to preserve it. */
function textOf(node: XmlElement): string {
  if (attribute(node, XML_NS, 'space') === 'preserve') return node.text;
  return node.text.replace(/^[ \t\r\n]+|[ \t\r\n]+$/g, '');
}

const W = NS.w;
const w = (element: XmlElement, name: string): boolean => is(element, W, name);

/** Tracked-change wrappers and which view drops their content. */
function droppedIn(element: XmlElement, view: View): boolean {
  if (element.ns !== W) return false;
  if (element.name === 'ins' || element.name === 'moveTo') return view === 'rejected';
  if (element.name === 'del' || element.name === 'moveFrom') return view === 'accepted';
  return false;
}

/** Does a property element (w:rPr of a paragraph mark, w:trPr, w:tcPr) carry a change that removes its owner in this view? */
function removedIn(properties: XmlElement | undefined, view: View, inserted: string[], deleted: string[]): boolean {
  if (!properties) return false;
  const names = view === 'rejected' ? inserted : deleted;
  return properties.children.some((c) => c.ns === W && names.includes(c.name));
}

interface NoteRef {
  kind: 'footnote' | 'endnote';
  id: string;
}

interface Context {
  view: View;
  /** Paragraphs of the story being read. */
  paragraphs: string[];
  /** Text of paragraphs whose marks were removed, waiting to join the next paragraph. */
  pending: string;
  /** Paragraphs of text boxes met while reading this story. */
  textBoxes: string[];
  notes: NoteRef[];
  /** Section properties met, in order (for headers and footers). */
  sections: XmlElement[];
  /** Things found that make the text impossible to verify. Shared by every nested context. */
  problems: string[];
  /** The package and the part being read, for following links to pictures, charts and the like. */
  source: Source;
}

interface Source {
  pkg: DocxPackage;
  part: string;
  partsRead: Set<string>;
  links: string[];
}

function newContext(view: View, problems: string[], source: Source): Context {
  return { view, paragraphs: [], pending: '', textBoxes: [], notes: [], sections: [], problems, source };
}

/**
 * What an element links to through a relationship (r:id, r:embed, ...): a
 * chart or diagram (its text is read), a picture or embedded file (its bytes
 * are fingerprinted), or an external address. Read at the place the document
 * anchors it, so a swapped chart or picture shows up as a difference there.
 */
function linkedContent(node: XmlElement, context: Context): string {
  let text = '';
  for (const [key, id] of node.attributes) {
    if (!key.startsWith(`${NS.r}|`) || id === '') continue;
    const { pkg, part, partsRead } = context.source;
    const rel = pkg.relationships(part).find((r) => r.id === id);
    if (!rel) continue; // a dangling id is reported by the structure check
    if (rel.external) {
      text += `⟦link to ${rel.target}⟧`;
      continue;
    }
    if (!rel.partName || !pkg.has(rel.partName)) continue;
    if (/\.xml$/i.test(rel.partName)) {
      const root = pkg.xml(rel.partName);
      if (root.ns === W) continue; // headers, footers and notes are read as stories of their own
      partsRead.add(rel.partName);
      const texts: string[] = [];
      allText(root, texts);
      text += `⟦linked content: ${texts.join(' | ')}⟧`;
    } else {
      partsRead.add(rel.partName);
      text += `⟦file ${fingerprint(pkg.entries.get(rel.partName)!)}⟧`;
    }
  }
  return text;
}

/** Fold a nested reading (a cell, a text box) back into the one that contains it. */
function absorb(outer: Context, inner: Context): void {
  outer.notes.push(...inner.notes);
  outer.sections.push(...inner.sections);
}

/**
 * mc:AlternateContent offers the same content in two forms (modern and
 * legacy), and which one is shown depends on the program opening the file.
 * Read the first, and insist that the other says exactly the same, so the text
 * is the same whichever is shown.
 */
function alternateContentBranch(element: XmlElement, context: Context, read: (branch: XmlElement, into: Context) => string): XmlElement | undefined {
  const branches = element.children.filter((c) => c.ns === NS.mc && (c.name === 'Choice' || c.name === 'Fallback'));
  if (branches.length > 1) {
    // Every branch (there may be several Choices) must read the same: text, text boxes and note references.
    const readings = branches.map((branch) => {
      const probe = newContext(context.view, [], { ...context.source, partsRead: new Set(), links: [] });
      const inline = read(branch, probe);
      // Pictures are stored differently in the modern and legacy forms, so linked files are left out of this comparison.
      return JSON.stringify([inline, probe.paragraphs, probe.pending, probe.textBoxes, probe.notes]).replace(/⟦file [^⟧]*⟧/g, '');
    });
    if (new Set(readings).size > 1) {
      context.problems.push('A drawing or text box is stored in alternative forms whose text differs, so what a reader would see cannot be verified.');
    }
  }
  return branches[0];
}

function unknownWrapperProblem(node: XmlElement, context: Context): void {
  if (!KNOWN_WRAPPER_NAMESPACES.has(node.ns) && holdsWordMarkup(node)) {
    context.problems.push(`Text sits inside an element this tool does not know (${node.name}), so whether Word shows it cannot be verified.`);
  }
}

/** Read the inline content of a paragraph (or of anything inside one) into text. */
function inlineText(element: XmlElement, context: Context): string {
  let text = '';
  for (const node of element.children) {
    if (droppedIn(node, context.view)) continue;
    if (node.ns === NS.mc && node.name === 'AlternateContent') {
      const branch = alternateContentBranch(node, context, inlineText);
      if (branch) text += inlineText(branch, context);
      continue;
    }
    // Text that is not Word text but is still shown: drawing text, equations, watermarks.
    if ((node.ns === A_NS || node.ns === M_NS) && node.name === 't') {
      text += node.text;
      continue;
    }
    if (node.ns === V_NS && node.name === 'textpath') {
      text += `⟦shape text: ${attribute(node, '', 'string') ?? ''}⟧`;
      continue;
    }
    unknownWrapperProblem(node, context);
    if (node.ns !== W) text += linkedContent(node, context);
    if (node.ns === W) {
      switch (node.name) {
        case 'hyperlink': {
          // Where the link goes is part of what the document says.
          const anchor = attribute(node, W, 'anchor');
          const covered = inlineText(node, context);
          // A link whose text is entirely removed in this view is not there to be followed.
          if (covered !== '') context.source.links.push(`"${covered}" links to${anchor ? ` bookmark ${anchor}` : ''}${linkedContent(node, context) || (anchor ? '' : ' nothing')}`);
          text += covered;
          continue;
        }
        case 'object':
          text += MARK.object + linkedContent(node, context) + inlineText(node, context);
          continue;
        case 't':
        case 'delText':
          text += textOf(node);
          continue;
        case 'tab':
          text += '\t';
          continue;
        case 'ptab':
          text += '⟦alignment tab⟧';
          continue;
        case 'br': {
          const type = attribute(node, W, 'type');
          text += type === 'page' ? MARK.pageBreak : type === 'column' ? MARK.columnBreak : '\n';
          continue;
        }
        case 'cr':
          text += '\n';
          continue;
        case 'fldChar': {
          const type = attribute(node, W, 'fldCharType');
          text += type === 'begin' ? MARK.fieldBegin : type === 'separate' ? MARK.fieldResult : type === 'end' ? MARK.fieldEnd : '';
          continue;
        }
        case 'instrText':
        case 'delInstrText':
          text += node.text; // the field's code: compared, so a cross-reference cannot be silently repointed
          continue;
        case 'fldSimple':
          text += `${MARK.fieldBegin}${attribute(node, W, 'instr') ?? ''}${MARK.fieldResult}${inlineText(node, context)}${MARK.fieldEnd}`;
          continue;
        case 'altChunk':
        case 'subDoc':
          context.problems.push('The document pulls in content from another file (an embedded or linked sub-document), which cannot be checked.');
          continue;
        case 'noBreakHyphen':
          text += '‑';
          continue;
        case 'softHyphen':
          text += '­';
          continue;
        case 'sym':
          text += `⟦symbol ${attribute(node, W, 'font') ?? ''} ${attribute(node, W, 'char') ?? ''}⟧`;
          continue;
        case 'footnoteReference':
          context.notes.push({ kind: 'footnote', id: attribute(node, W, 'id') ?? '' });
          text += MARK.footnote;
          continue;
        case 'endnoteReference':
          context.notes.push({ kind: 'endnote', id: attribute(node, W, 'id') ?? '' });
          text += MARK.endnote;
          continue;
        case 'drawing':
        case 'pict':
          // The marker, then any text shown inside the object (watermark or shape text).
          // Text boxes inside it are read into the text-box story as they are met.
          text += MARK.object + inlineText(node, context);
          continue;
        case 'txbxContent': {
          // A text box is its own flow of paragraphs, read into the text-box story.
          const box = newContext(context.view, context.problems, context.source);
          readBlocks(node, box);
          context.textBoxes.push(MARK.textBox, ...box.paragraphs, ...box.textBoxes);
          absorb(context, box);
          continue;
        }
        case 'pPr':
        case 'rPr':
          continue; // properties are not document text
        default:
          break;
      }
    }
    text += inlineText(node, context);
  }
  return text;
}

function readParagraph(paragraph: XmlElement, context: Context): void {
  const properties = child(paragraph, W, 'pPr');
  const text = inlineText(paragraph, context);
  const markRemoved = removedIn(properties && child(properties, W, 'rPr'), context.view, ['ins', 'moveTo'], ['del', 'moveFrom']);
  context.pending += text;
  if (markRemoved) return; // this paragraph joins the next one in this view
  context.paragraphs.push(context.pending);
  context.pending = '';
  const section = properties && child(properties, W, 'sectPr');
  if (section) context.sections.push(section);
}

/** Read block-level content (paragraphs, tables, containers) of a body, cell, note, header or text box. */
function readBlocks(container: XmlElement, context: Context): void {
  for (const node of container.children) {
    if (droppedIn(node, context.view)) continue;
    if (node.ns === NS.mc && node.name === 'AlternateContent') {
      const branch = alternateContentBranch(node, context, (b, into) => {
        readBlocks(b, into);
        return '';
      });
      if (branch) readBlocks(branch, context);
      continue;
    }
    if (node.ns !== W) {
      if (holdsWordMarkup(node)) context.problems.push(`Text sits inside an element this tool does not know (${node.name}), so whether Word shows it cannot be verified.`);
      continue;
    }
    switch (node.name) {
      case 'altChunk':
      case 'subDoc':
        context.problems.push('The document pulls in content from another file (an embedded or linked sub-document), which cannot be checked.');
        break;
      case 'p':
        readParagraph(node, context);
        break;
      case 'tbl':
        flushPending(context);
        context.paragraphs.push(MARK.table);
        for (const row of node.children) {
          if (droppedIn(row, context.view)) continue;
          if (w(row, 'tr')) readRow(row, context);
          else if (w(row, 'sdt') || w(row, 'customXml') || w(row, 'ins') || w(row, 'del') || w(row, 'moveFrom') || w(row, 'moveTo')) readTableContainer(row, context);
        }
        context.paragraphs.push(MARK.tableEnd);
        break;
      case 'sdt': {
        const content = child(node, W, 'sdtContent');
        if (content) readBlocks(content, context);
        break;
      }
      case 'customXml':
      case 'smartTag':
      case 'ins':
      case 'del':
      case 'moveFrom':
      case 'moveTo':
        readBlocks(node, context);
        break;
      case 'sectPr':
        context.sections.push(node);
        break;
      default:
        break;
    }
  }
  flushPending(context);
}

/** Text whose paragraph mark was removed with nothing after it to join: keep it, never lose it. */
function flushPending(context: Context): void {
  if (context.pending !== '') {
    context.paragraphs.push(context.pending);
    context.pending = '';
  }
}

/** Rows wrapped in a content control or similar. */
function readTableContainer(wrapper: XmlElement, context: Context): void {
  const inner = w(wrapper, 'sdt') ? child(wrapper, W, 'sdtContent') : wrapper;
  if (!inner) return;
  for (const row of inner.children) {
    if (droppedIn(row, context.view)) continue;
    if (w(row, 'tr')) readRow(row, context);
    else if (row.ns === W) readTableContainer(row, context);
  }
}

function readRow(row: XmlElement, context: Context): void {
  if (removedIn(child(row, W, 'trPr'), context.view, ['ins'], ['del'])) return;
  context.paragraphs.push(MARK.row);
  for (const cell of row.children) {
    if (w(cell, 'tc')) readCell(cell, context);
    else if (w(cell, 'sdt')) {
      const content = child(cell, W, 'sdtContent');
      for (const inner of content?.children ?? []) if (w(inner, 'tc')) readCell(inner, context);
    }
  }
}

function readCell(cell: XmlElement, context: Context): void {
  if (removedIn(child(cell, W, 'tcPr'), context.view, ['cellIns'], ['cellDel'])) return;
  // A cell is its own scope for joining paragraphs: text never leaks out of a cell.
  const inner = newContext(context.view, context.problems, context.source);
  readBlocks(cell, inner);
  // A cell that survives in this view always holds at least one paragraph, as it
  // does in Word. So a row or cell that should have vanished, but was left behind
  // empty, still shows up as a difference.
  if (inner.paragraphs.length === 0) inner.paragraphs.push('');
  context.paragraphs.push(MARK.cell, ...inner.paragraphs);
  context.textBoxes.push(...inner.textBoxes);
  absorb(context, inner);
}

const isBlank = (paragraphs: string[]): boolean => paragraphs.every((p) => p === '' || isStructureMark(p));

/**
 * Read every story of a Word file in the given view.
 *
 * Stories with no text at all are left out, so "a header that exists but is
 * empty" and "no header" compare as equal.
 */
export function readStories(pkg: DocxPackage, view: View): StoryReading {
  const stories: Story[] = [];
  const partsRead = new Set<string>();
  const noteIdsRead = new Map<string, Set<string>>();
  const problems: string[] = [];
  const links: string[] = [];
  const add = (story: Story): void => {
    if (!isBlank(story.paragraphs)) stories.push(story);
  };

  const main = pkg.mainDocumentPart();
  const body = child(pkg.xml(main), W, 'body');
  if (!body) {
    problems.push('The main document has no body.');
    return { stories, partsRead, noteIdsRead, problems, links };
  }
  partsRead.add(main);

  const bodyContext = newContext(view, problems, { pkg, part: main, partsRead, links });
  readBlocks(body, bodyContext);
  add({ key: 'body', kind: 'body', label: 'main text', paragraphs: bodyContext.paragraphs });
  add({ key: 'body/text boxes', kind: 'text box', label: 'text boxes in the main text', paragraphs: bodyContext.textBoxes });

  const notes: NoteRef[] = [...bodyContext.notes];

  // Headers and footers, through each section's references.
  const mainRels = pkg.relationships(main);
  bodyContext.sections.forEach((section, sectionIndex) => {
    for (const reference of section.children) {
      const kind = w(reference, 'headerReference') ? 'header' : w(reference, 'footerReference') ? 'footer' : undefined;
      if (!kind) continue;
      const type = attribute(reference, W, 'type') ?? 'default';
      const relId = attribute(reference, NS.r, 'id') ?? '';
      const label = `${kind} (section ${sectionIndex + 1}${type === 'default' ? '' : `, ${type} page`})`;
      const rel = mainRels.find((r) => r.id === relId);
      if (!rel?.partName || !pkg.has(rel.partName)) {
        problems.push(`The ${label} points to a part that does not exist.`);
        continue;
      }
      partsRead.add(rel.partName);
      const context = newContext(view, problems, { pkg, part: rel.partName, partsRead, links });
      readBlocks(pkg.xml(rel.partName), context);
      const key = `${kind}/${sectionIndex + 1}/${type}`;
      add({ key, kind, label, paragraphs: context.paragraphs });
      add({ key: `${key}/text boxes`, kind: 'text box', label: `text boxes in the ${label}`, paragraphs: context.textBoxes });
      notes.push(...context.notes);
    }
  });

  // Footnotes and endnotes, in the order the text refers to them. Matching by
  // order of reference (not by id) is deliberate: ids may be renumbered.
  for (const kind of ['footnote', 'endnote'] as const) {
    const partName = pkg.relationshipsOfType(main, `${kind}s`)[0]?.partName;
    const refs = notes.filter((n) => n.kind === kind);
    if (!partName || !pkg.has(partName)) {
      if (refs.length > 0) problems.push(`The text refers to ${kind}s but the file has no ${kind}s part.`);
      continue;
    }
    partsRead.add(partName);
    const byId = new Map<string, XmlElement>();
    const idsRead = noteIdsRead.get(partName) ?? new Set<string>();
    noteIdsRead.set(partName, idsRead);
    for (const note of pkg.xml(partName).children) {
      if (!w(note, kind)) continue;
      const id = attribute(note, W, 'id') ?? '';
      byId.set(id, note);
      // Separators and continuation notices are printed on the page too, so
      // they are read as stories of their own (normally they hold no text).
      const type = attribute(note, W, 'type');
      if (type && type !== 'normal') {
        idsRead.add(id);
        const context = newContext(view, problems, { pkg, part: partName, partsRead, links });
        readBlocks(note, context);
        add({ key: `${kind} ${type}`, kind, label: `${kind} separator`, paragraphs: [...context.paragraphs, ...context.textBoxes] });
      }
    }
    refs.forEach((ref, index) => {
      const note = byId.get(ref.id);
      const label = `${kind} ${index + 1}`;
      if (!note) {
        problems.push(`A ${kind} reference (${label}) points to a ${kind} that does not exist.`);
        return;
      }
      idsRead.add(ref.id);
      const context = newContext(view, problems, { pkg, part: partName, partsRead, links });
      readBlocks(note, context);
      // Always keep the note's story, even if blank, so a missing note cannot hide.
      stories.push({ key: `${kind}/${index + 1}`, kind, label, paragraphs: context.paragraphs });
      add({ key: `${kind}/${index + 1}/text boxes`, kind: 'text box', label: `text boxes in ${label}`, paragraphs: context.textBoxes });
      if (context.notes.length > 0) problems.push(`A ${kind} contains a note reference, which this tool does not check.`);
    });
  }

  return { stories, partsRead, noteIdsRead, problems, links };
}
