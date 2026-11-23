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
}

/** Markers standing in for things that are not characters but must not vanish unnoticed. */
export const MARK = {
  footnote: '⟦footnote⟧',
  endnote: '⟦endnote⟧',
  object: '⟦object⟧',
} as const;

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
}

function newContext(view: View): Context {
  return { view, paragraphs: [], pending: '', textBoxes: [], notes: [], sections: [] };
}

/** mc:AlternateContent offers the same content twice (modern and legacy); read exactly one. */
function alternateContentBranch(element: XmlElement): XmlElement | undefined {
  return child(element, NS.mc, 'Choice') ?? child(element, NS.mc, 'Fallback');
}

/** Read the inline content of a paragraph (or of anything inside one) into text. */
function inlineText(element: XmlElement, context: Context): string {
  let text = '';
  for (const node of element.children) {
    if (droppedIn(node, context.view)) continue;
    if (node.ns === NS.mc && node.name === 'AlternateContent') {
      const branch = alternateContentBranch(node);
      if (branch) text += inlineText(branch, context);
      continue;
    }
    if (node.ns === W) {
      switch (node.name) {
        case 't':
        case 'delText':
          text += node.text;
          continue;
        case 'tab':
          text += '\t';
          continue;
        case 'br':
        case 'cr':
          text += '\n';
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
        case 'object':
          text += MARK.object;
          inlineText(node, context); // only to reach any text boxes inside
          continue;
        case 'txbxContent': {
          // A text box is its own flow of paragraphs, read into the text-box story.
          const box = newContext(context.view);
          readBlocks(node, box);
          context.textBoxes.push(...box.paragraphs, ...box.textBoxes);
          context.notes.push(...box.notes);
          continue;
        }
        case 'pPr':
        case 'rPr':
        case 'instrText':
        case 'delInstrText':
          continue; // properties and field codes are not document text
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
      const branch = alternateContentBranch(node);
      if (branch) readBlocks(branch, context);
      continue;
    }
    if (node.ns !== W) continue;
    switch (node.name) {
      case 'p':
        readParagraph(node, context);
        break;
      case 'tbl':
        for (const row of node.children) {
          if (droppedIn(row, context.view)) continue;
          if (w(row, 'tr')) readRow(row, context);
          else if (w(row, 'sdt') || w(row, 'customXml') || w(row, 'ins') || w(row, 'del') || w(row, 'moveFrom') || w(row, 'moveTo')) readTableContainer(row, context);
        }
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
  // Text whose paragraph mark was removed with nothing after it to join: keep it, never lose it.
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
  const inner = newContext(context.view);
  readBlocks(cell, inner);
  // A cell that survives in this view always holds at least one paragraph, as it
  // does in Word. So a row or cell that should have vanished, but was left behind
  // empty, still shows up as a difference.
  if (inner.paragraphs.length === 0) inner.paragraphs.push('');
  context.paragraphs.push(...inner.paragraphs);
  context.textBoxes.push(...inner.textBoxes);
  context.notes.push(...inner.notes);
  context.sections.push(...inner.sections);
}

const isBlank = (paragraphs: string[]): boolean => paragraphs.every((p) => p === '');

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
  const add = (story: Story): void => {
    if (!isBlank(story.paragraphs)) stories.push(story);
  };

  const main = pkg.mainDocumentPart();
  const body = child(pkg.xml(main), W, 'body');
  if (!body) {
    problems.push('The main document has no body.');
    return { stories, partsRead, noteIdsRead, problems };
  }
  partsRead.add(main);

  const bodyContext = newContext(view);
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
      const context = newContext(view);
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
    if (refs.length === 0) continue;
    if (!partName || !pkg.has(partName)) {
      problems.push(`The text refers to ${kind}s but the file has no ${kind}s part.`);
      continue;
    }
    partsRead.add(partName);
    const byId = new Map<string, XmlElement>();
    for (const note of pkg.xml(partName).children) {
      if (w(note, kind)) byId.set(attribute(note, W, 'id') ?? '', note);
    }
    const idsRead = noteIdsRead.get(partName) ?? new Set<string>();
    noteIdsRead.set(partName, idsRead);
    refs.forEach((ref, index) => {
      const note = byId.get(ref.id);
      const label = `${kind} ${index + 1}`;
      if (!note) {
        problems.push(`A ${kind} reference (${label}) points to a ${kind} that does not exist.`);
        return;
      }
      idsRead.add(ref.id);
      const context = newContext(view);
      readBlocks(note, context);
      // Always keep the note's story, even if blank, so a missing note cannot hide.
      stories.push({ key: `${kind}/${index + 1}`, kind, label, paragraphs: context.paragraphs });
      add({ key: `${kind}/${index + 1}/text boxes`, kind: 'text box', label: `text boxes in ${label}`, paragraphs: context.textBoxes });
      if (context.notes.length > 0) problems.push(`A ${kind} contains a note reference, which this tool does not check.`);
    });
  }

  return { stories, partsRead, noteIdsRead, problems };
}
