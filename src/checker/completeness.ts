/**
 * Check 6, completeness: no text anywhere in the two uploads or the output was
 * left unread. The text checks only mean something if they covered everything.
 */
import { NS } from '../common/namespaces';
import type { DocxPackage } from './package';
import type { CheckFailure } from './report';
import { isStructureMark, type Story, type StoryReading } from './stories';
import { attribute, descendants, is, type XmlElement } from './xml';

const W = NS.w;

function hasText(element: XmlElement): boolean {
  for (const node of descendants(element)) {
    if ((is(node, W, 't') || is(node, W, 'delText')) && node.text !== '') return true;
  }
  return false;
}

function hasAnyText(element: XmlElement): boolean {
  return element.text.trim() !== '' || element.children.some(hasAnyText);
}

/**
 * Parts that legitimately hold w:t text which is not document text:
 * reviewer comments (annotations, verified structurally instead) and the
 * glossary (a template's stock of building blocks, never shown in the document).
 */
function isExempt(partName: string, root: XmlElement): boolean {
  return is(root, W, 'comments') || is(root, W, 'glossaryDocument') || partName.includes('/glossary/');
}

/**
 * Every problem that means some text in this file was not read.
 * `readings` are all the readings taken of this file (one view for an upload,
 * both views for an output).
 */
export function unreadText(pkg: DocxPackage, readings: StoryReading[], fileName: string): string[] {
  const problems: string[] = [];
  for (const reading of readings) for (const problem of reading.problems) problems.push(`In the ${fileName}: ${problem}`);

  const partsRead = new Set<string>();
  for (const reading of readings) for (const part of reading.partsRead) partsRead.add(part);

  for (const partName of pkg.partNames) {
    // Embedded documents (HTML, RTF, text, another Word file) could carry text Word would import.
    if (/\.(html?|mht|mhtml|rtf|txt|docx?|docm|dotx?|odt)$/i.test(partName)) {
      problems.push(`The ${fileName} embeds another document ("${partName}"), whose content cannot be checked.`);
      continue;
    }
    if (!partName.endsWith('.xml') || partName === '[Content_Types].xml') continue;
    const root = pkg.xml(partName);
    if (isExempt(partName, root)) continue;
    if (root.ns !== W) {
      // Charts, diagrams and the like beside the document are read as "graphics" text.
      // Anything else outside Word's markup that holds text was not read.
      const inDocumentFolder = /^word\//.test(partName) || partsRead.has(partName);
      if (!partsRead.has(partName) && inDocumentFolder && !partName.includes('/theme/') && hasAnyText(root)) {
        problems.push(`The ${fileName} contains text in "${partName}" that was not checked.`);
      }
      continue;
    }

    if (!partsRead.has(partName)) {
      if (hasText(root)) problems.push(`The ${fileName} contains text in "${partName}" that was not checked (nothing in the document refers to it).`);
      continue;
    }

    // Notes that hold text but that no reference in the text leads to.
    const noteKind = is(root, W, 'footnotes') ? 'footnote' : is(root, W, 'endnotes') ? 'endnote' : undefined;
    if (!noteKind) continue;
    const idsRead = new Set<string>();
    for (const reading of readings) for (const id of reading.noteIdsRead.get(partName) ?? []) idsRead.add(id);
    for (const note of root.children) {
      if (!is(note, W, noteKind)) continue;
      const type = attribute(note, W, 'type');
      if (type && type !== 'normal') continue; // separators and continuation notices
      const id = attribute(note, W, 'id') ?? '';
      if (!idsRead.has(id) && hasText(note)) {
        problems.push(`The ${fileName} contains a ${noteKind} with text that nothing in the document refers to, so it was not checked.`);
      }
    }
  }
  return problems;
}

/** One line saying what was covered, for the report. */
export function coverageNote(stories: Story[]): string {
  const count = (kind: Story['kind']): number => stories.filter((s) => s.kind === kind).length;
  const paragraphs = stories.reduce((sum, s) => sum + s.paragraphs.filter((p) => !isStructureMark(p)).length, 0);
  const plural = (n: number, word: string): string => `${n} ${word}${n === 1 ? '' : 's'}`;
  const parts = [
    'main text',
    plural(count('header'), 'header'),
    plural(count('footer'), 'footer'),
    plural(count('footnote'), 'footnote'),
    plural(count('endnote'), 'endnote'),
    plural(count('text box'), 'text-box group'),
  ];
  return `Checked ${plural(paragraphs, 'paragraph')} in: ${parts.join(', ')} (tables included).`;
}

export function completenessFailure(problems: string[]): CheckFailure | undefined {
  if (problems.length === 0) return undefined;
  return { message: problems[0]! + (problems.length > 1 ? ` (and ${problems.length - 1} more of the same kind)` : '') };
}
