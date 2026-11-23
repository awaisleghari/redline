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
import { staticKindOf, Styles, type StaticKind } from './formatting';
import { Numbering } from './numbering';
import { DocxPackage } from './package';
import { attribute, child, is, type XmlElement } from './xml';

export type View = 'accepted' | 'rejected';

/**
 * How changes are expressed in the file being read:
 *   - 'tracked': as Word tracked changes (w:ins, w:del, ...);
 *   - 'static': as fixed formatting (the static redline's colours and strikes).
 */
export type Markup = 'tracked' | 'static';

/**
 * What is known about one character of a paragraph, packed into a character
 * code: the low six bits are its visible formatting (see FORMAT); then:
 */
export const CHAR = {
  /** A real character of text (as opposed to a marker standing for a tab-like thing or an object). */
  isText: 64,
  /** Carries the static redline's "formatting changed" shading. */
  shaded: 128,
  /** Part of a change that this view keeps (inserted text when accepting, deleted text when rejecting). */
  changed: 256,
} as const;

/** One paragraph (or structure marker) of a story, with what the checks beyond plain text need. */
export interface Para {
  text: string;
  /** One code per character of `text` (see CHAR). Empty for structure markers. */
  meta: string;
  /** The automatic number this paragraph carries ('' if none). */
  number: string;
  /** In a static redline: the old clause number typed in front of a removed paragraph ('' if none). */
  oldNumber: string;
  /** True when this paragraph itself is removed by the changes (so it exists only in the rejected view). */
  removed: boolean;
}

const marker = (text: string): Para => ({ text, meta: '', number: '', oldNumber: '', removed: false });

/** Style ids the static redline uses for its own additions. The Checker recognises them by name. */
export const STATIC_STYLE = { coverPrefix: 'RedlineCover', oldNumber: 'RedlineOldNumber', shadeFill: 'FFF2CC' } as const;

export type StoryKind = 'body' | 'header' | 'footer' | 'footnote' | 'endnote' | 'text box';

export interface Story {
  /** Stable identity used to match a story across two files. */
  key: string;
  kind: StoryKind;
  /** Plain-English name for reports, e.g. "footnote 2" or "header (section 1)". */
  label: string;
  paragraphs: string[];
  /** The same entries as `paragraphs`, with per-character detail. */
  details: Para[];
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
  /** In a static redline: the text of each line of the cover page, in order. */
  cover: string[];
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
  paragraphs: Para[];
  /** Text of paragraphs whose marks were removed, waiting to join the next paragraph. */
  pending: string;
  /** Per-character codes for everything emitted since the last paragraph was closed (parallel to `pending` plus the text being read). */
  meta: string;
  /** Code applied to characters being emitted now (formatting of the current run, and whether it is part of a change). */
  code: number;
  /** Style of the paragraph being read. */
  paragraphStyle: string | undefined;
  /** Old clause number met in the paragraph being read (static redline). */
  oldNumber: string;
  /** Whether the last paragraph read had its mark removed in this view. */
  lastMarkRemoved: boolean;
  /** Paragraphs of text boxes met while reading this story. */
  textBoxes: Para[];
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
  markup: Markup;
  styles: Styles;
  numbering: Numbering;
}

function newContext(view: View, problems: string[], source: Source): Context {
  return { view, paragraphs: [], pending: '', meta: '', code: 0, paragraphStyle: undefined, oldNumber: '', lastMarkRemoved: false, textBoxes: [], notes: [], sections: [], problems, source };
}

/** Record text being emitted, with the current character code, and hand it back. */
function emit(context: Context, text: string, isText = false): string {
  if (text !== '') context.meta += String.fromCharCode(context.code | (isText ? CHAR.isText : 0)).repeat(text.length);
  return text;
}

/** Is a change of this kind dropped in this view? */
function staticDropped(kind: StaticKind | undefined, view: View): boolean {
  if (!kind) return false;
  return view === 'accepted' ? kind === 'deleted' || kind === 'movedFrom' : kind === 'inserted' || kind === 'movedTo';
}

/** Read a run: its formatting applies to everything it emits. */
function readRun(run: XmlElement, context: Context): string {
  const rPr = child(run, W, 'rPr');
  const { markup, styles } = context.source;
  let code = styles.formatOf(rPr, context.paragraphStyle) | (context.code & CHAR.changed);
  if (markup === 'static') {
    const kind = staticKindOf(rPr);
    const styleElement = rPr && child(rPr, W, 'rStyle');
    if (styleElement && attribute(styleElement, W, 'val') === STATIC_STYLE.oldNumber) {
      // The old number of a removed clause: shown struck through, and not part of either document's text.
      if (kind !== 'deleted' && kind !== 'movedFrom') context.problems.push('An old clause number is not marked as deleted text.');
      const saved = context.meta;
      const label = inlineText(run, context);
      context.meta = saved;
      if (context.view === 'rejected') context.oldNumber += label;
      return '';
    }
    if (staticDropped(kind, context.view)) return '';
    if (kind) code |= CHAR.changed;
    const shading = rPr && child(rPr, W, 'shd');
    if (shading && attribute(shading, W, 'fill')?.toUpperCase() === STATIC_STYLE.shadeFill) code |= CHAR.shaded;
  }
  const outer = context.code;
  context.code = code;
  const text = inlineText(run, context);
  context.code = outer;
  return text;
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
      const texts = (list: Para[]): string[] => list.map((entry) => entry.text);
      return JSON.stringify([inline, texts(probe.paragraphs), probe.pending, texts(probe.textBoxes), probe.notes]).replace(/⟦file [^⟧]*⟧/g, '');
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
      text += emit(context, node.text, true);
      continue;
    }
    if (node.ns === V_NS && node.name === 'textpath') {
      text += emit(context, `⟦shape text: ${attribute(node, '', 'string') ?? ''}⟧`);
      continue;
    }
    unknownWrapperProblem(node, context);
    if (node.ns !== W) text += emit(context, linkedContent(node, context));
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
        case 'r':
          text += readRun(node, context);
          continue;
        case 'ins':
        case 'del':
        case 'moveFrom':
        case 'moveTo': {
          // A tracked change this view keeps: what is inside is part of a change.
          const outer = context.code;
          context.code |= CHAR.changed;
          text += inlineText(node, context);
          context.code = outer;
          continue;
        }
        case 'object':
          text += emit(context, MARK.object + linkedContent(node, context));
          text += inlineText(node, context);
          continue;
        case 't':
        case 'delText':
          text += emit(context, textOf(node), true);
          continue;
        case 'tab':
          text += emit(context, '\t');
          continue;
        case 'ptab':
          text += emit(context, '⟦alignment tab⟧');
          continue;
        case 'br': {
          const type = attribute(node, W, 'type');
          text += emit(context, type === 'page' ? MARK.pageBreak : type === 'column' ? MARK.columnBreak : '\n');
          continue;
        }
        case 'cr':
          text += emit(context, '\n');
          continue;
        case 'fldChar': {
          const type = attribute(node, W, 'fldCharType');
          text += emit(context, type === 'begin' ? MARK.fieldBegin : type === 'separate' ? MARK.fieldResult : type === 'end' ? MARK.fieldEnd : '');
          continue;
        }
        case 'instrText':
        case 'delInstrText':
          text += emit(context, node.text); // the field's code: compared, so a cross-reference cannot be silently repointed
          continue;
        case 'fldSimple':
          text += emit(context, `${MARK.fieldBegin}${attribute(node, W, 'instr') ?? ''}${MARK.fieldResult}`);
          text += inlineText(node, context);
          text += emit(context, MARK.fieldEnd);
          continue;
        case 'altChunk':
        case 'subDoc':
          context.problems.push('The document pulls in content from another file (an embedded or linked sub-document), which cannot be checked.');
          continue;
        case 'noBreakHyphen':
          text += emit(context, '‑');
          continue;
        case 'softHyphen':
          text += emit(context, '­');
          continue;
        case 'sym':
          text += emit(context, `⟦symbol ${attribute(node, W, 'font') ?? ''} ${attribute(node, W, 'char') ?? ''}⟧`);
          continue;
        case 'footnoteReference':
          context.notes.push({ kind: 'footnote', id: attribute(node, W, 'id') ?? '' });
          text += emit(context, MARK.footnote);
          continue;
        case 'endnoteReference':
          context.notes.push({ kind: 'endnote', id: attribute(node, W, 'id') ?? '' });
          text += emit(context, MARK.endnote);
          continue;
        case 'drawing':
        case 'pict':
          // The marker, then any text shown inside the object (watermark or shape text).
          // Text boxes inside it are read into the text-box story as they are met.
          text += emit(context, MARK.object);
          text += inlineText(node, context);
          continue;
        case 'txbxContent': {
          // A text box is its own flow of paragraphs, read into the text-box story.
          const box = newContext(context.view, context.problems, context.source);
          box.code = context.code & CHAR.changed;
          readBlocks(node, box);
          context.textBoxes.push(marker(MARK.textBox), ...box.paragraphs, ...box.textBoxes);
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

/** Close the paragraph being assembled and add it to the story. */
function closeParagraph(context: Context, number: string, removed: boolean): void {
  if (context.meta.length !== context.pending.length) throw new Error('internal error: character details are out of step with the text');
  context.paragraphs.push({ text: context.pending, meta: context.meta, number, oldNumber: context.oldNumber, removed });
  context.pending = '';
  context.meta = '';
  context.oldNumber = '';
}

function readParagraph(paragraph: XmlElement, context: Context): void {
  const properties = child(paragraph, W, 'pPr');
  const { markup, styles, numbering } = context.source;
  const styleElement = properties && child(properties, W, 'pStyle');
  context.paragraphStyle = styleElement && attribute(styleElement, W, 'val');
  if (markup === 'static' && (context.paragraphStyle ?? '').startsWith(STATIC_STYLE.coverPrefix)) {
    context.problems.push('A paragraph inside the document carries the cover-page style, which would exempt it from the counts.');
  }
  const text = inlineText(paragraph, context);
  context.paragraphStyle = undefined;

  const markProperties = properties && child(properties, W, 'rPr');
  const markKind = markup === 'static' ? staticKindOf(markProperties) : undefined;
  const markRemoved =
    markup === 'static' ? staticDropped(markKind, context.view) : removedIn(markProperties, context.view, ['ins', 'moveTo'], ['del', 'moveFrom']);
  // Is the paragraph itself removed by the changes (whatever this view shows)?
  const removed = markup === 'static' ? markKind === 'deleted' || markKind === 'movedFrom' : removedIn(markProperties, 'accepted', [], ['del', 'moveFrom']);

  // Automatic numbering. In a static redline every paragraph in the file is
  // live and takes a number if it is numbered, so every paragraph is counted.
  // With tracked changes, a paragraph this view removes takes no number.
  const number = markup === 'static' || !markRemoved ? numbering.next(styles.numberingOf(properties)) : '';

  context.pending += text;
  context.lastMarkRemoved = markRemoved;
  if (markRemoved) return; // this paragraph joins the next one in this view
  closeParagraph(context, number, removed);
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
        const tableStart = context.paragraphs.length;
        context.paragraphs.push(marker(MARK.table));
        for (const row of node.children) {
          if (droppedIn(row, context.view)) continue;
          if (w(row, 'tr')) readRow(row, context);
          else if (w(row, 'sdt') || w(row, 'customXml') || w(row, 'ins') || w(row, 'del') || w(row, 'moveFrom') || w(row, 'moveTo')) readTableContainer(row, context);
        }
        // A table none of whose rows exists in this view (wholly inserted or wholly deleted) is not there at all.
        if (!context.paragraphs.some((entry, index) => index > tableStart && entry.text === MARK.row && entry.meta === '')) context.paragraphs.length = tableStart;
        else context.paragraphs.push(marker(MARK.tableEnd));
        break;
      case 'sdt': {
        const content = child(node, W, 'sdtContent');
        if (content) readBlocks(content, context);
        break;
      }
      case 'customXml':
      case 'smartTag':
        readBlocks(node, context);
        break;
      case 'ins':
      case 'del':
      case 'moveFrom':
      case 'moveTo': {
        const outer = context.code;
        context.code |= CHAR.changed;
        readBlocks(node, context);
        context.code = outer;
        break;
      }
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
  if (context.pending !== '' || context.oldNumber !== '') closeParagraph(context, '', true);
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
  if (context.source.markup === 'tracked' && removedIn(child(row, W, 'trPr'), context.view, ['ins'], ['del'])) return;
  const cells: XmlElement[] = [];
  for (const cell of row.children) {
    if (w(cell, 'tc')) cells.push(cell);
    else if (w(cell, 'sdt')) {
      const content = child(cell, W, 'sdtContent');
      for (const inner of content?.children ?? []) if (w(inner, 'tc')) cells.push(inner);
    }
  }
  // Read the cells first: in a static redline a row is gone exactly when every one of its cells is gone.
  const kept = cells.map((cell) => readCell(cell, context)).filter((cell): cell is Context => cell !== undefined);
  if (kept.length === 0 && context.source.markup === 'static' && cells.length > 0) return;
  context.paragraphs.push(marker(MARK.row));
  for (const inner of kept) {
    context.paragraphs.push(marker(MARK.cell), ...inner.paragraphs);
    context.textBoxes.push(...inner.textBoxes);
    absorb(context, inner);
  }
}

/** Read one cell. Returns undefined when the cell does not exist in this view. */
function readCell(cell: XmlElement, context: Context): Context | undefined {
  if (context.source.markup === 'tracked' && removedIn(child(cell, W, 'tcPr'), context.view, ['cellIns'], ['cellDel'])) return undefined;
  // A cell is its own scope for joining paragraphs: text never leaks out of a cell.
  const inner = newContext(context.view, context.problems, context.source);
  inner.code = context.code & CHAR.changed;
  readBlocks(cell, inner);
  // In a static redline there are no flags on rows and cells: a cell is gone
  // in this view exactly when its last paragraph's mark is (Word never lets a
  // surviving cell lose its last paragraph mark) and nothing of it remains.
  if (context.source.markup === 'static' && inner.lastMarkRemoved && inner.paragraphs.length === 0) return undefined;
  // A cell that survives in this view always holds at least one paragraph, as it
  // does in Word. So a row or cell that should have vanished, but was left behind
  // empty, still shows up as a difference.
  if (inner.paragraphs.length === 0) inner.paragraphs.push(marker(''));
  return inner;
}

const isBlank = (paragraphs: Para[]): boolean => paragraphs.every((p) => p.text === '' || isStructureMark(p.text));
const storyOf = (key: string, kind: StoryKind, label: string, details: Para[]): Story => ({ key, kind, label, details, paragraphs: details.map((p) => p.text) });

/**
 * Read every story of a Word file in the given view.
 *
 * Stories with no text at all are left out, so "a header that exists but is
 * empty" and "no header" compare as equal.
 */
export function readStories(pkg: DocxPackage, view: View, markup: Markup = 'tracked'): StoryReading {
  const stories: Story[] = [];
  const partsRead = new Set<string>();
  const noteIdsRead = new Map<string, Set<string>>();
  const problems: string[] = [];
  const links: string[] = [];
  const cover: string[] = [];
  const add = (key: string, kind: StoryKind, label: string, details: Para[]): void => {
    if (!isBlank(details)) stories.push(storyOf(key, kind, label, details));
  };

  const main = pkg.mainDocumentPart();
  let body = child(pkg.xml(main), W, 'body');
  if (!body) {
    problems.push('The main document has no body.');
    return { stories, partsRead, noteIdsRead, problems, links, cover };
  }
  partsRead.add(main);
  const shared = { pkg, partsRead, links, markup, styles: new Styles(pkg, main), numbering: new Numbering(pkg, main) };

  if (markup === 'static') {
    // The cover page: the leading paragraphs in the cover styles. It is the
    // redline's own addition, so it is set aside here and verified separately.
    const isCover = (node: XmlElement): boolean => {
      const pPr = w(node, 'p') ? child(node, W, 'pPr') : undefined;
      const style = pPr && child(pPr, W, 'pStyle');
      return !!style && (attribute(style, W, 'val') ?? '').startsWith(STATIC_STYLE.coverPrefix);
    };
    let count = 0;
    while (count < body.children.length && isCover(body.children[count]!)) count++;
    const rawText = (element: XmlElement): string => (w(element, 't') ? element.text : element.children.map(rawText).join(''));
    for (const paragraph of body.children.slice(0, count)) cover.push(rawText(paragraph));
    if (body.children.slice(count).some((node) => isCover(node))) problems.push('A cover-page paragraph appears inside the document.');
    body = { ...body, children: body.children.slice(count) };
  }

  const bodyContext = newContext(view, problems, { ...shared, part: main });
  readBlocks(body, bodyContext);
  // (The cover page's own section ended inside the paragraphs set aside above, so it is not in the list of sections.)
  add('body', 'body', 'main text', bodyContext.paragraphs);
  add('body/text boxes', 'text box', 'text boxes in the main text', bodyContext.textBoxes);

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
      const context = newContext(view, problems, { ...shared, part: rel.partName });
      readBlocks(pkg.xml(rel.partName), context);
      const key = `${kind}/${sectionIndex + 1}/${type}`;
      add(key, kind, label, context.paragraphs);
      add(`${key}/text boxes`, 'text box', `text boxes in the ${label}`, context.textBoxes);
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
        const context = newContext(view, problems, { ...shared, part: partName });
        readBlocks(note, context);
        add(`${kind} ${type}`, kind, `${kind} separator`, [...context.paragraphs, ...context.textBoxes]);
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
      const context = newContext(view, problems, { ...shared, part: partName });
      readBlocks(note, context);
      // Always keep the note's story, even if blank, so a missing note cannot hide.
      stories.push(storyOf(`${kind}/${index + 1}`, kind, label, context.paragraphs));
      add(`${kind}/${index + 1}/text boxes`, 'text box', `text boxes in ${label}`, context.textBoxes);
      if (context.notes.length > 0) problems.push(`A ${kind} contains a note reference, which this tool does not check.`);
    });
  }

  return { stories, partsRead, noteIdsRead, problems, links, cover };
}
