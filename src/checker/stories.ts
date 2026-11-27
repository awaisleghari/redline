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
import { FORMAT, STANDARD_COLOURS, staticKindOf, Styles, type ChangeColours, type StaticKind } from './formatting';
import { Numbering } from './numbering';
import { DocxPackage } from './package';
import { attribute, child, is, type XmlElement } from './xml';

/**
 * 'accepted' and 'rejected' are the two documents a redline stands for. 'all'
 * reads a static redline exactly as it is printed: every visible character,
 * inserted and deleted alike, cover page included (used to check the PDF).
 */
export type View = 'accepted' | 'rejected' | 'all';

/**
 * How changes are expressed in the file being read:
 *   - 'tracked': as Word tracked changes (w:ins, w:del, ...);
 *   - 'static': as fixed formatting (the static redline's colours and strikes).
 */
export type Markup = 'tracked' | 'static';

/**
 * What is known about one character of a paragraph, packed into a character
 * code: the low eight bits are its visible formatting (see FORMAT); then:
 */
export const CHAR = {
  /** A real character of text (as opposed to a marker standing for a tab-like thing or an object). */
  isText: 256,
  /** Carries the static redline's "formatting changed" shading. */
  shaded: 512,
  /** Part of a change that this view keeps (inserted text when accepting, deleted text when rejecting). */
  changed: 1024,
  /** Which kind of change (static redline): multiply the index in STATIC_KINDS, plus one, by this. */
  kindUnit: 2048,
} as const;

/** The change kinds in the order their index is stored in a character code. */
export const STATIC_KINDS: readonly StaticKind[] = ['inserted', 'deleted', 'movedFrom', 'movedTo'];
/** The kind of change a character code records, if any. */
export const kindOfCode = (code: number): StaticKind | undefined => STATIC_KINDS[Math.floor(code / CHAR.kindUnit) - 1];

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
  /** True when the paragraph is a list item (numbered or bulleted), whether or not its label could be worked out. */
  listed: boolean;
  /** The bullet character of a bulleted item ('' otherwise). */
  bullet: string;
  /** True when a tab in the paragraph may draw a leader (dots, a line): one of its tab stops has one. */
  leader?: boolean;
  /** In a static redline: the kind of change the paragraph's own mark shows (its number or bullet is printed the same way). */
  markKind?: StaticKind;
  /** The links in the paragraph: which characters of `text` each covers, and where it points. Absent when there are none. */
  links?: LinkSpan[];
}

/** One link within a paragraph: characters `start` (included) to `end` (not included) point to `target`. */
export interface LinkSpan {
  start: number;
  end: number;
  target: string;
}

/** Where character `at` of a paragraph links to ('' when it is not part of a link). */
export function linkAt(paragraph: Para, at: number): string {
  for (const span of paragraph.links ?? []) if (at >= span.start && at < span.end) return span.target;
  return '';
}

const marker = (text: string): Para => ({ text, meta: '', number: '', oldNumber: '', removed: false, listed: false, bullet: '' });

/** Style ids the static redline uses for its own additions. The Checker recognises them by name. */
export const STATIC_STYLE = { coverPrefix: 'RedlineCover', oldNumber: 'RedlineOldNumber', layoutNote: 'RedlineLayoutNote', shadeFill: 'FFF2CC' } as const;

/**
 * In a static redline: a note of the redline's own, saying that the wording after it is unchanged and only its
 * layout changed (see layout-notes.ts). `at` is the place in the story's entries where the noted wording begins.
 * Either the next `count` paragraphs were a table in the original, or the table that follows was `count` paragraphs.
 */
export interface LayoutNote {
  at: number;
  was: 'table' | 'paragraphs';
  count: number;
}

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
  /** A header or footer the section does not name itself but carries over from the section before. */
  carriedOver?: boolean;
  /** In a static redline: the layout notes met in this story, in order. */
  layoutNotes?: LayoutNote[];
}

export interface StoryReading {
  stories: Story[];
  /** The colour in which a static redline shows insertions (only meaningful for a static redline's reading). */
  colours: ChangeColours;
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
  /** Handwriting (pen strokes) the document holds, in reading order: a fingerprint of each. */
  inks: string[];
  /**
   * Clause numbers in the stored forms of drawings and text boxes that are NOT the form read (Word stores a text
   * box twice; a program may show either). One entry per such form, in reading order.
   */
  otherForms: string[];
  /** Templates that legacy shapes in the document ask for and the document does not define: such a shape may not be drawn as intended. */
  missingShapeTypes: string[];
  /** The text of each table row marked to repeat at the top of every page the table runs onto. */
  repeatRows: string[];
  /** How the pages are set up: sections, their headers and footers, page numbering, note numbering. */
  layout: Layout;
  /** In a static redline: the text of each line of the cover page, in order. */
  cover: string[];
  /** For each cover line, how its text is formatted: 'plain', a change kind, 'shaded', with ':hidden' if it would not be shown. */
  coverLooks: string[];
}

/** Markers standing in for things that are not characters but must not vanish unnoticed. */
export type PageKind = 'default' | 'first' | 'even';

/** One section of the main text, as far as the printed page depends on it. */
export interface SectionLayout {
  /** How many entries of the main text's `details` belong to this section and those before it. */
  end: number;
  /** The section begins on a new page (anything but a "continuous" break). */
  newPage: boolean;
  /** Its first page has a header and footer of its own. */
  titlePage: boolean;
  /** The page number its first page is given, when it says. */
  start: number | undefined;
  /** The page-number format (Word's name for it: 'decimal', 'lowerRoman', ...). */
  format: string;
  /** The story key of each header and footer the section itself names. */
  header: Partial<Record<PageKind, string>>;
  footer: Partial<Record<PageKind, string>>;
}

export interface Layout {
  sections: SectionLayout[];
  /** Even pages have headers and footers of their own. */
  evenAndOdd: boolean;
  /** How footnotes and endnotes are numbered (Word's names). */
  noteFormat: { footnote: string; endnote: string };
}

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
const INK_NS = 'http://www.w3.org/2003/InkML';

/** Does an element hold any document markup at all (text, a note reference, a tab, a field, an equation, a shape...)? */
/**
 * Whether an element could hold something a reader sees as text: Word's own markup, a piece of text in a
 * drawing or an equation, or a legacy text box. Drawing settings that hold no text (colours, line and fill
 * styles, such as the "hidden fill" and "hidden line" extensions Word writes on pictures) do not count.
 */
function holdsWordMarkup(element: XmlElement): boolean {
  return element.children.some(
    (c) => c.ns === W || ((c.ns === A_NS || c.ns === M_NS) && (c.name === 't' || c.name === 'fld')) || (c.ns === V_NS && (c.name === 'textbox' || c.name === 'textpath')) || holdsWordMarkup(c),
  );
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

/**
 * The text of a w:t or w:delText as Word shows it: outer whitespace is dropped unless the element, or
 * the nearest element around it that says anything on the matter, says to preserve it.
 */
function textOf(node: XmlElement): string {
  if (node.preserveSpace) return node.text;
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
  /** Links met since the last paragraph was closed (positions are those of `meta`). */
  spans: LinkSpan[];
  /** Layout notes met (static redline). They may stand only directly in the main text, a header or a footer: `topLevel`. */
  layoutNotes: LayoutNote[];
  topLevel: boolean;
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
  /** For each of `sections`: how many paragraphs had been read when it ended. */
  sectionEnds: number[];
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
  inks: string[];
  otherForms: string[];
  /** Shape templates (legacy drawings) defined anywhere in the document, and those its shapes ask for. */
  shapeTypes: { defined: Set<string>; used: Set<string> };
  markup: Markup;
  styles: Styles;
  numbering: Numbering;
  repeatRows: string[];
  /** The colour in which a static redline shows insertions. */
  colours: ChangeColours;
}

function newContext(view: View, problems: string[], source: Source): Context {
  return { view, paragraphs: [], pending: '', meta: '', code: 0, spans: [], layoutNotes: [], topLevel: false, paragraphStyle: undefined, oldNumber: '', lastMarkRemoved: false, textBoxes: [], notes: [], sections: [], sectionEnds: [], problems, source };
}

/** Record text being emitted, with the current character code, and hand it back. */
function emit(context: Context, text: string, isText = false): string {
  if (text !== '') context.meta += String.fromCharCode(context.code | (isText ? CHAR.isText : 0)).repeat(text.length);
  return text;
}

/** Is a change of this kind dropped in this view? */
function staticDropped(kind: StaticKind | undefined, view: View): boolean {
  if (!kind || view === 'all') return false;
  return view === 'accepted' ? kind === 'deleted' || kind === 'movedFrom' : kind === 'inserted' || kind === 'movedTo';
}

/** Read a run: its formatting applies to everything it emits. */
function readRun(run: XmlElement, context: Context): string {
  const rPr = child(run, W, 'rPr');
  const { markup, styles } = context.source;
  let code = styles.formatOf(rPr, context.paragraphStyle) | (context.code & CHAR.changed);
  if (markup === 'static') {
    const kind = staticKindOf(rPr, context.source.colours);
    const styleElement = rPr && child(rPr, W, 'rStyle');
    if (styleElement && attribute(styleElement, W, 'val') === STATIC_STYLE.oldNumber && context.view !== 'all') {
      // The old number of a removed clause: shown struck through, and not part of either document's text.
      if (kind !== 'deleted' && kind !== 'movedFrom') context.problems.push('An old clause number is not marked as deleted text.');
      if (styles.formatOf(rPr, context.paragraphStyle) & FORMAT.hidden) context.problems.push('An old clause number is formatted as hidden text.');
      const undefinedStyle = 'An old clause number uses a style that the document\'s style sheet does not define.';
      if (!styles.defines(STATIC_STYLE.oldNumber) && !context.problems.includes(undefinedStyle)) context.problems.push(undefinedStyle);
      const saved = context.meta;
      const label = inlineText(run, context);
      context.meta = saved;
      if (context.view === 'rejected') context.oldNumber += label;
      return '';
    }
    if (staticDropped(kind, context.view)) {
      // Everything inside a dropped run (a text box, say) must be marked the same way, or a reader would see it as ordinary text.
      const inner = (element: XmlElement): void => {
        for (const c of element.children) {
          if (w(c, 'r') && c.children.some((x) => !w(x, 'rPr')) && staticKindOf(child(c, W, 'rPr'), context.source.colours) !== kind) {
            context.problems.push('Text inside a deleted or inserted object is not itself marked as deleted or inserted.');
          }
          inner(c);
        }
      };
      inner(run);
      return '';
    }
    if (kind) code |= CHAR.changed | ((STATIC_KINDS.indexOf(kind) + 1) * CHAR.kindUnit);
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
      if (root.ns === INK_NS) {
        // Handwriting: pen strokes, stored as coordinates. There is no text in it. The strokes are noted, to be
        // compared between the files; the picture of them that the file carries alongside is read as a picture
        // (see alternateContentBranch).
        partsRead.add(rel.partName);
        context.source.inks.push(fingerprint(pkg.entries.get(rel.partName)!));
        continue;
      }
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
  for (const section of inner.sections) {
    outer.sections.push(section);
    outer.sectionEnds.push(outer.paragraphs.length);
  }
}

/** A field with nothing nested in it: its instruction, the words it shows (captured), and its end. */
/** Marks an entry of `otherForms` that is a form's full reading (the others record which paragraphs are numbered). */
export const FORM_READING = 'reading:';
const SIMPLE_FIELD = new RegExp(`${MARK.fieldBegin}[^⟦]*${MARK.fieldResult}([^⟦]*)${MARK.fieldEnd}`, 'g');

/**
 * mc:AlternateContent offers the same content in two forms (modern and
 * legacy), and which one is shown depends on the program opening the file.
 * Read the first, and insist that the other says exactly the same, so the text
 * is the same whichever is shown.
 */
function alternateContentBranch(element: XmlElement, context: Context, read: (branch: XmlElement, into: Context) => string): XmlElement | undefined {
  const branches = element.children.filter((c) => c.ns === NS.mc && (c.name === 'Choice' || c.name === 'Fallback'));
  const formNumbers: string[] = [];
  const formReadings: string[] = [];
  /** True when the forms show the same words and differ only in field instructions. */
  let instructionsDiffer = false;
  if (branches.length > 1) {
    // Every branch (there may be several Choices) must read the same: text, text boxes and note references.
    const readings = branches.map((branch) => {
      const probe = newContext(context.view, [], { ...context.source, partsRead: new Set(), links: [], inks: [], repeatRows: [] });
      // (These readings only compare the forms; they must not use up clause numbers.)
      const inline = context.source.numbering.withoutCounting(() => read(branch, probe));
      // Pictures are stored differently in the modern and legacy forms, so linked files are left out of this comparison.
      const texts = (list: Para[]): string[] => list.map((entry) => entry.text);
      // Clause numbers, form by form, each read from the same starting count. A published document's two forms
      // need not agree on them (found on a real contract: the legacy form on a different list), so they are not
      // compared with each other; the forms that are not read are recorded, and must be the same in the output
      // as in the modified document. Otherwise a program showing the other form would print numbers nothing
      // has checked (review finding).
      // What is recorded is which paragraphs are in a list, not the label each prints: the label of a form Word
      // never shows depends on list definitions that a published file may leave dangling and the engine rebuilds
      // (found on a real contract: no label could be worked out in the upload, "1." in the output).
      const numbers = (list: Para[]): string[] => list.map((entry) => (entry.listed ? 'listed' : ''));
      // A text box holding nothing but empty paragraphs shows no text: a drawing saved as shapes with such a box
      // and, for older programs, as a flat picture of itself reads the same either way.
      const shown = (list: Para[]): string[] => {
        const boxes: string[][] = [];
        for (const text of texts(list)) {
          if (text === '⟦text box⟧' || boxes.length === 0) boxes.push([]);
          if (text !== '⟦text box⟧') boxes[boxes.length - 1]!.push(text);
        }
        return boxes.filter((box) => box.some((text) => text.trim() !== '')).flatMap((box) => ['⟦text box⟧', ...box]);
      };
      const boxNumbers = numbers(probe.textBoxes.filter((entry) => entry.text.trim() !== '' && entry.text !== '⟦text box⟧'));
      // (Only what is numbered is of interest: a form with no numbered paragraph in this view records nothing, so
      // a box shown removed, or a drawing with no list in it, leaves no trace.)
      const numbered = [...numbers(probe.paragraphs), ...boxNumbers].filter((_entry, _index, all) => all.some((e) => e !== ''));
      formNumbers.push(numbered.length > 0 ? JSON.stringify(numbered) : '');
      const reading = JSON.stringify([inline, texts(probe.paragraphs), probe.pending, shown(probe.textBoxes), probe.notes]).replace(/⟦file [^⟧]*⟧/g, '');
      formReadings.push(reading);
      // What a reader sees is compared: the words a field shows, not the instruction behind it. (Found on a real
      // contract: a web address typed as plain text in one form and as a link showing the same words in the
      // other.) The instructions are not dropped: where the forms differ in them, the forms that are not read
      // are recorded in full, and must be the same in the output as in the modified document.
      // (Only a whole, simple field is reduced to the words it shows; anything else stays as it is and is compared in full.)
      return reading.replace(SIMPLE_FIELD, '$1');
    });
    instructionsDiffer = new Set(readings).size === 1 && new Set(formReadings).size > 1;
    if (new Set(readings).size > 1) {
      context.problems.push('A drawing or text box is stored in alternative forms whose text differs, so what a reader would see cannot be verified.');
    }
  }
  // Handwriting (found on a real contract: a stray pen mark) is stored as pen strokes with, as the other form, a
  // picture of them. The strokes cannot be compared with anything, and a program that rewrites the file may keep
  // only the picture. So the picture is the form that is read: it is what is compared between the files.
  const fallback = branches.find((branch) => branch.name === 'Fallback');
  // The strokes are still noted: a file that keeps them must keep them unchanged (the Checker compares them).
  if (fallback && branches[0] !== fallback) {
    const strokes = inkIn(branches[0]!, context);
    if (strokes.length > 0) {
      context.source.inks.push(...strokes);
      context.source.otherForms.push(...formNumbers.filter((form, i) => form !== '' && branches[i] !== fallback));
      if (instructionsDiffer) context.source.otherForms.push(...formReadings.filter((_, i) => branches[i] !== fallback).map((r) => FORM_READING + r));
      return fallback;
    }
  }
  context.source.otherForms.push(...formNumbers.slice(1).filter((form) => form !== ''));
  if (instructionsDiffer) context.source.otherForms.push(...formReadings.slice(1).map((r) => FORM_READING + r));
  return branches[0];
}

/** The pen strokes anything in this element links to: a fingerprint of each such part. */
function inkIn(element: XmlElement, context: Context): string[] {
  const { pkg, part, partsRead } = context.source;
  const found: string[] = [];
  for (const [key, id] of element.attributes) {
    if (!key.startsWith(`${NS.r}|`) || id === '') continue;
    const rel = pkg.relationships(part).find((r) => r.id === id);
    if (rel && !rel.external && rel.partName && /\.xml$/i.test(rel.partName) && pkg.has(rel.partName) && pkg.xml(rel.partName).ns === INK_NS) {
      partsRead.add(rel.partName);
      found.push(fingerprint(pkg.entries.get(rel.partName)!));
    }
  }
  for (const c of element.children) found.push(...inkIn(c, context));
  return found;
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
    // A shape type is a template for shapes (a definition, like a style): it draws nothing itself, and a file may
    // carry the same definition more than once. Only an actual shape's text is shown.
    if (node.ns === V_NS && node.name === 'shapetype') {
      const id = attribute(node, '', 'id');
      if (id) context.source.shapeTypes.defined.add(id);
      continue;
    }
    // A legacy shape drawn from a template must still have its template (see missingShapeTypes).
    if (node.ns === V_NS) {
      const type = attribute(node, '', 'type');
      if (type?.startsWith('#')) context.source.shapeTypes.used.add(type.slice(1));
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
          const start = context.meta.length;
          const covered = inlineText(node, context);
          // A link whose text is entirely removed in this view is not there to be followed.
          if (covered !== '') {
            const target = `${anchor ? ` bookmark ${anchor}` : ''}${linkedContent(node, context) || (anchor ? '' : ' nothing')}`;
            context.source.links.push(`"${covered}" links to${target}`);
            if (context.meta.length > start) context.spans.push({ start, end: context.meta.length, target });
          }
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
          // A legacy picture element holding nothing at all (found on real contracts: what the engine leaves behind
          // when it moves a template out), or nothing but shape types (templates for shapes, see below), draws
          // nothing: it is a place where definitions are kept. Found on real contracts: a program that rewrites the
          // file may gather a definition into such an element of its own, and a reader sees no difference.
          if (node.name === 'pict' && node.children.every((child) => child.ns === V_NS && child.name === 'shapetype')) {
            // What it defines is still noted: shapes elsewhere may be drawn from it (see missingShapeTypes).
            for (const definition of node.children) {
              const id = attribute(definition, '', 'id');
              if (id) context.source.shapeTypes.defined.add(id);
            }
            continue;
          }
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
function closeParagraph(context: Context, number: string, removed: boolean, listed = false, bullet = '', markKind?: StaticKind): void {
  if (context.meta.length !== context.pending.length) throw new Error('internal error: character details are out of step with the text');
  context.paragraphs.push({ text: context.pending, meta: context.meta, number, oldNumber: context.oldNumber, removed, listed, bullet, markKind });
  if (context.spans.length > 0) context.paragraphs[context.paragraphs.length - 1]!.links = context.spans;
  context.spans = [];
  context.pending = '';
  context.meta = '';
  context.oldNumber = '';
}

/**
 * The paragraph whose properties this paragraph's text has in this view. Paragraph properties belong to the
 * paragraph mark. When tracked changes remove a mark, the text before it joins the next paragraph and takes that
 * paragraph's properties (so text that joins a heading is formatted as the heading). `siblings` is the content of
 * the paragraph's container, and `index` the paragraph's place in it.
 */
const BETWEEN_PARAGRAPHS = /^(bookmarkStart|bookmarkEnd|proofErr|permStart|permEnd|commentRangeStart|commentRangeEnd|move(From|To)Range(Start|End))$/;

function survivingParagraph(paragraph: XmlElement, siblings: XmlElement[], index: number, context: Context): XmlElement {
  if (context.source.markup !== 'tracked' || context.view === 'all') return paragraph;
  const markRemoved = (p: XmlElement): boolean => {
    const properties = child(p, W, 'pPr');
    return removedIn(properties && child(properties, W, 'rPr'), context.view, ['ins', 'moveTo'], ['del', 'moveFrom']);
  };
  let current = paragraph;
  for (let i = index + 1; markRemoved(current); i++) {
    // Markers that hold no content (bookmarks and the like) do not keep paragraphs apart.
    while (siblings[i] && siblings[i]!.ns === W && BETWEEN_PARAGRAPHS.test(siblings[i]!.name)) i++;
    const next = siblings[i];
    // Only a paragraph directly after can be joined; anything else (a table, a wrapper, the end) ends the run of paragraphs.
    if (!next || !w(next, 'p')) return paragraph;
    current = next;
  }
  return current;
}

/** What a layout note may say, and what it then means. */
const LAYOUT_NOTE_WORDING: [RegExp, LayoutNote['was']][] = [
  [/^Layout changed: the next (paragraph) was previously in a table\.$/, 'table'],
  [/^Layout changed: the next ([1-9]\d*) paragraphs were previously in a table\.$/, 'table'],
  [/^Layout changed: the table below was previously a (single) paragraph\.$/, 'paragraphs'],
  [/^Layout changed: the table below was previously ([1-9]\d*) separate paragraphs\.$/, 'paragraphs'],
];

/**
 * A layout note in a static redline. It is the redline's own addition (like the cover page), so it is not part
 * of either document's text; it is recorded, and what it claims is verified against the original (layout-notes.ts).
 * It must look like what it is: shaded, visible, not dressed as a change, and worded exactly.
 */
function readLayoutNote(paragraph: XmlElement, context: Context): void {
  const { styles } = context.source;
  const problem = (text: string): void => void context.problems.push(text);
  if (!context.topLevel) problem('A layout note stands where none may (inside a table, a note or a text box).');
  if (context.pending !== '') problem('A layout note interrupts a paragraph.');
  if (!styles.defines(STATIC_STYLE.layoutNote)) problem("A layout note uses a style that the document's style sheet does not define.");
  const properties = child(paragraph, W, 'pPr');
  const numId = styles.numberingOf(properties) && child(styles.numberingOf(properties)!, W, 'numId');
  if (numId && attribute(numId, W, 'val') !== '0') problem('A layout note carries a clause number.');
  if (staticKindOf(properties && child(properties, W, 'rPr'), context.source.colours)) problem('A layout note is marked as a change.');
  let text = '';
  for (const run of paragraph.children) {
    if (run === properties) continue;
    const rPr = child(run, W, 'rPr');
    if (!w(run, 'r') || run.children.some((c) => c !== rPr && !w(c, 't'))) {
      problem('A layout note holds something other than its own words.');
      continue;
    }
    const shading = rPr && child(rPr, W, 'shd');
    if (!shading || attribute(shading, W, 'fill')?.toUpperCase() !== STATIC_STYLE.shadeFill) problem('A layout note is not shaded.');
    if (staticKindOf(rPr, context.source.colours)) problem('A layout note is marked as a change.');
    if (styles.formatOf(rPr, STATIC_STYLE.layoutNote) & (FORMAT.hidden | FORMAT.strike | FORMAT.doubleStrike)) problem('A layout note is hidden or struck through.');
    text += run.children.filter((c) => w(c, 't')).map(textOf).join('');
  }
  for (const [wording, was] of LAYOUT_NOTE_WORDING) {
    const match = wording.exec(text);
    if (!match) continue;
    context.layoutNotes.push({ at: context.paragraphs.length, was, count: /^\d+$/.test(match[1]!) ? Number(match[1]) : 1 });
    return;
  }
  problem('A layout note does not say what it should.');
}

function readParagraph(paragraph: XmlElement, context: Context, siblings: XmlElement[] = [], index = 0): void {
  const properties = child(paragraph, W, 'pPr');
  const { markup, styles, numbering } = context.source;
  const ownStyle = properties && child(properties, W, 'pStyle');
  if (markup === 'static' && context.view !== 'all' && ownStyle && attribute(ownStyle, W, 'val') === STATIC_STYLE.layoutNote) {
    readLayoutNote(paragraph, context);
    return;
  }
  const survivingProperties = child(survivingParagraph(paragraph, siblings, index, context), W, 'pPr');
  const survivingStyle = survivingProperties && child(survivingProperties, W, 'pStyle');
  context.paragraphStyle = survivingStyle && attribute(survivingStyle, W, 'val');
  if (markup === 'static' && context.view !== 'all' && (context.paragraphStyle ?? '').startsWith(STATIC_STYLE.coverPrefix)) {
    context.problems.push('A paragraph inside the document carries the cover-page style, which would exempt it from the counts.');
  }
  const text = inlineText(paragraph, context);
  context.paragraphStyle = undefined;

  const markProperties = properties && child(properties, W, 'rPr');
  const markKind = markup !== 'static' ? undefined : staticKindOf(markProperties, context.source.colours);
  const markRemoved =
    markup === 'static' ? staticDropped(markKind, context.view) : removedIn(markProperties, context.view, ['ins', 'moveTo'], ['del', 'moveFrom']);
  // Is the paragraph itself removed by the changes (whatever this view shows)?
  const removed = markup === 'static' ? markKind === 'deleted' || markKind === 'movedFrom' : removedIn(markProperties, 'accepted', [], ['del', 'moveFrom']);

  // Automatic numbering. In a static redline every paragraph in the file is
  // live and takes a number if it is numbered, so every paragraph is counted.
  // With tracked changes, a paragraph this view removes takes no number.
  const numPr = styles.numberingOf(properties);
  const number = markup === 'static' || !markRemoved ? numbering.next(numPr) : '';
  const numIdElement = numPr && child(numPr, W, 'numId');
  const listed = !!numIdElement && !['0', undefined].includes(attribute(numIdElement, W, 'val'));

  context.pending += text;
  context.lastMarkRemoved = markRemoved;
  if (markRemoved) return; // this paragraph joins the next one in this view
  closeParagraph(context, number, removed, listed, numbering.bulletOf(numPr), markKind);
  if (styles.hasLeaderTab(properties)) context.paragraphs[context.paragraphs.length - 1]!.leader = true;
  const section = properties && child(properties, W, 'sectPr');
  if (section) {
    context.sections.push(section);
    context.sectionEnds.push(context.paragraphs.length);
  }
}

/** Read block-level content (paragraphs, tables, containers) of a body, cell, note, header or text box. */
function readBlocks(container: XmlElement, context: Context): void {
  for (const [index, node] of container.children.entries()) {
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
        readParagraph(node, context, container.children, index);
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
        context.sectionEnds.push(context.paragraphs.length);
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
  const rowProperties = child(row, W, 'trPr');
  if (rowProperties && child(rowProperties, W, 'tblHeader')) {
    // A heading row: Word prints it again at the top of each page the table continues onto.
    context.source.repeatRows.push(kept.flatMap((cell) => cell.paragraphs.map((p) => [...p.text].filter((_, i) => p.meta.charCodeAt(i) & CHAR.isText).join(''))).join(''));
  }
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
export function readStories(pkg: DocxPackage, view: View, markup: Markup = 'tracked', colours: ChangeColours = STANDARD_COLOURS): StoryReading {
  const stories: Story[] = [];
  const partsRead = new Set<string>();
  const noteIdsRead = new Map<string, Set<string>>();
  const problems: string[] = [];
  const links: string[] = [];
  const inks: string[] = [];
  const otherForms: string[] = [];
  const shapeTypes = { defined: new Set<string>(), used: new Set<string>() };
  const cover: string[] = [];
  const coverLooks: string[] = [];
  const repeatRows: string[] = [];
  const layout: Layout = { sections: [], evenAndOdd: false, noteFormat: { footnote: 'decimal', endnote: 'lowerRoman' } };
  const add = (key: string, kind: StoryKind, label: string, details: Para[], layoutNotes: LayoutNote[] = []): void => {
    if (!isBlank(details)) stories.push(storyOf(key, kind, label, details));
    // (A note in a story with no text at all has nothing to stand over.)
    if (layoutNotes.length > 0 && isBlank(details)) problems.push('A layout note stands in a part of the document that holds no text.');
    if (layoutNotes.length > 0 && !isBlank(details)) stories[stories.length - 1]!.layoutNotes = layoutNotes;
  };

  const main = pkg.mainDocumentPart();
  let body = child(pkg.xml(main), W, 'body');
  if (!body) {
    problems.push('The main document has no body.');
    return { stories, partsRead, noteIdsRead, problems, links, inks, otherForms, missingShapeTypes: [...shapeTypes.used].filter((id) => !shapeTypes.defined.has(id)).sort(), cover, coverLooks, repeatRows, layout, colours };
  }
  partsRead.add(main);
  const shared = { pkg, partsRead, links, inks, otherForms, shapeTypes, markup, styles: new Styles(pkg, main), numbering: new Numbering(pkg, main), repeatRows, colours };

  if (markup === 'static' && view !== 'all') {
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
    for (const paragraph of body.children.slice(0, count)) {
      cover.push(rawText(paragraph));
      // How the line looks: the legend's samples must really be in each style, and nothing may be hidden.
      const pPr = child(paragraph, W, 'pPr');
      const styleElement = pPr && child(pPr, W, 'pStyle');
      // The cover's styles are the redline's own: they must be in the style sheet the document really uses,
      // or Word shows the cover page with none of its layout.
      const coverStyle = (styleElement && attribute(styleElement, W, 'val')) ?? '';
      const undefinedStyle = 'The cover page uses a style that the document\'s style sheet does not define.';
      if (!shared.styles.defines(coverStyle) && !problems.includes(undefinedStyle)) problems.push(undefinedStyle);
      const looks = new Set<string>();
      const look = (element: XmlElement): void => {
        for (const c of element.children) {
          if (w(c, 'r') && rawText(c) !== '') {
            const rPr = child(c, W, 'rPr');
            const shading = rPr && child(rPr, W, 'shd');
            const shaded = !!shading && attribute(shading, W, 'fill')?.toUpperCase() === STATIC_STYLE.shadeFill;
            const hidden = (shared.styles.formatOf(rPr, styleElement && attribute(styleElement, W, 'val')) & FORMAT.hidden) !== 0;
            looks.add(`${staticKindOf(rPr, colours) ?? (shaded ? 'shaded' : 'plain')}${hidden ? ':hidden' : ''}`);
          } else look(c);
        }
      };
      look(paragraph);
      coverLooks.push([...looks].sort().join(','));
    }
    if (body.children.slice(count).some((node) => isCover(node))) problems.push('A cover-page paragraph appears inside the document.');
    body = { ...body, children: body.children.slice(count) };
  }

  const bodyContext = newContext(view, problems, { ...shared, part: main });
  bodyContext.topLevel = true;
  readBlocks(body, bodyContext);
  // (The cover page's own section ended inside the paragraphs set aside above, so it is not in the list of sections.)
  add('body', 'body', 'main text', bodyContext.paragraphs, bodyContext.layoutNotes);
  add('body/text boxes', 'text box', 'text boxes in the main text', bodyContext.textBoxes);

  const notes: NoteRef[] = [...bodyContext.notes];

  // Headers and footers, through each section's references.
  const mainRels = pkg.relationships(main);
  /** The header or footer part in force for each kind and page type, as the sections go by. */
  const inForce = new Map<string, string>();
  bodyContext.sections.forEach((section, sectionIndex) => {
    const numbers = child(section, W, 'pgNumType');
    const start = Number(numbers ? (attribute(numbers, W, 'start') ?? NaN) : NaN);
    const breakType = child(section, W, 'type');
    const titlePage = child(section, W, 'titlePg');
    const sectionLayout: SectionLayout = {
      end: bodyContext.sectionEnds[sectionIndex] ?? bodyContext.paragraphs.length,
      newPage: (breakType && attribute(breakType, W, 'val')) !== 'continuous',
      titlePage: !!titlePage && !['0', 'false', 'off'].includes(attribute(titlePage, W, 'val') ?? '1'),
      start: Number.isInteger(start) ? start : undefined,
      format: (numbers && attribute(numbers, W, 'fmt')) ?? 'decimal',
      header: {},
      footer: {},
    };
    layout.sections.push(sectionLayout);
    // What this section names itself.
    const named = new Set<string>();
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
      named.add(`${kind}/${type}`);
      inForce.set(`${kind}/${type}`, rel.partName);
      if (type === 'default' || type === 'first' || type === 'even') sectionLayout[kind][type] = `${kind}/${sectionIndex + 1}/${type}`;
    }
    // What the section shows: a header or footer it does not name is the one in force from the section
    // before ("same as previous"), so each section is read with the header and footer it really carries.
    // A file that names it again in every section and a file that leaves it to carry over then read alike.
    for (const [slot, partName] of inForce) {
      const [kind, type] = slot.split('/') as ['header' | 'footer', string];
      const label = `${kind} (section ${sectionIndex + 1}${type === 'default' ? '' : `, ${type} page`})`;
      const firstReading = !partsRead.has(partName);
      partsRead.add(partName);
      const context = newContext(view, problems, { ...shared, part: partName });
      context.topLevel = true;
      readBlocks(pkg.xml(partName), context);
      const key = `${kind}/${sectionIndex + 1}/${type}`;
      const before = stories.length;
      add(key, kind, label, context.paragraphs, context.layoutNotes);
      add(`${key}/text boxes`, 'text box', `text boxes in the ${label}`, context.textBoxes);
      if (!named.has(slot)) for (const story of stories.slice(before)) story.carriedOver = true;
      if (firstReading) notes.push(...context.notes);
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
      // A note always holds at least one paragraph, as a cell does: Word never lets a note lose its last paragraph
      // mark. (Found on a real contract: a draft whose emptied footnote had its one paragraph mark flagged as
      // removed. Read as no paragraph at all, it could never match the same note in the other files.)
      // (Only a note that does hold a paragraph: one with none at all is damaged, and must still show as different.)
      if (context.paragraphs.length === 0 && note.children.some((c) => w(c, 'p'))) context.paragraphs.push(marker(''));
      // Always keep the note's story, even if blank, so a missing note cannot hide.
      stories.push(storyOf(`${kind}/${index + 1}`, kind, label, context.paragraphs));
      add(`${kind}/${index + 1}/text boxes`, 'text box', `text boxes in ${label}`, context.textBoxes);
      if (context.notes.length > 0) problems.push(`A ${kind} contains a note reference, which this tool does not check.`);
    });
  }

  // Settings that change what a page prints.
  const settingsPart = pkg.relationshipsOfType(main, 'settings')[0]?.partName;
  if (settingsPart && pkg.has(settingsPart)) {
    const settings = pkg.xml(settingsPart);
    const flag = child(settings, W, 'evenAndOddHeaders');
    layout.evenAndOdd = !!flag && !['0', 'false', 'off'].includes(attribute(flag, W, 'val') ?? '1');
    for (const kind of ['footnote', 'endnote'] as const) {
      const properties = child(settings, W, `${kind}Pr`);
      const format = properties && child(properties, W, 'numFmt');
      if (format) layout.noteFormat[kind] = attribute(format, W, 'val') ?? layout.noteFormat[kind];
    }
  }
  // A section's own note numbering comes before the document's. Where the sections disagree, or the numbers
  // start again on each page or section, the Checker does not work the numbers out ('mixed' matches nothing).
  for (const kind of ['footnote', 'endnote'] as const) {
    const stated = new Set<string>();
    let restarts = false;
    for (const section of bodyContext.sections) {
      const properties = child(section, W, `${kind}Pr`);
      const format = properties && child(properties, W, 'numFmt');
      if (format) stated.add(attribute(format, W, 'val') ?? '');
      const restart = properties && child(properties, W, 'numRestart');
      if (restart && attribute(restart, W, 'val') !== 'continuous') restarts = true;
    }
    if (stated.size === 1) layout.noteFormat[kind] = [...stated][0]!;
    if (stated.size > 1 || restarts) layout.noteFormat[kind] = 'mixed';
  }

  return { stories, partsRead, noteIdsRead, problems, links, inks, otherForms, missingShapeTypes: [...shapeTypes.used].filter((id) => !shapeTypes.defined.has(id)).sort(), cover, coverLooks, repeatRows, layout, colours };
}
