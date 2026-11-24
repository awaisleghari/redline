/**
 * Producer: the cover page of the static redline. Plain and professional, in
 * the document's own body font (its styles are based on the document's
 * default paragraph style).
 *
 * The cover page sits in a section of its own, before the document, so the
 * document's headers, footers and page numbering are untouched by it.
 */
import { CHANGE_STYLE, FORMAT_CHANGE_FILL, OLD_NUMBER_STYLE, type ChangeKind } from './static-redline';


export interface CoverPageContent {
  originalName: string;
  modifiedName: string;
  /** Date and time of the comparison, already formatted for display. */
  comparedAt: string;
  insertions: number;
  deletions: number;
  moves: number;
  formattingChanges: number;
  /** Which uploads contained tracked changes of their own (compared as if accepted); undefined when neither did. */
  trackedChangesIn?: 'original' | 'modified' | 'both';
}

/** Style ids of the cover page. The Checker recognises the cover page by these. */
export const COVER_STYLE = { title: 'RedlineCoverTitle', heading: 'RedlineCoverHeading', text: 'RedlineCoverText' } as const;

export const LEGEND: Record<ChangeKind, string> = {
  inserted: 'Inserted text',
  deleted: 'Deleted text',
  movedFrom: 'Text moved from here',
  movedTo: 'Text moved to here',
};

const escapeXml = (text: string): string => text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

function paragraph(style: string, text: string, runProperties = ''): string {
  return `<w:p><w:pPr><w:pStyle w:val="${style}"/></w:pPr><w:r>${runProperties}<w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r></w:p>`;
}

function changeRunProperties(kind: ChangeKind): string {
  const style = CHANGE_STYLE[kind];
  const strike = style.strike === 'single' ? '<w:strike/>' : style.strike === 'double' ? '<w:dstrike/>' : '';
  const underline = style.underline ? `<w:u w:val="${style.underline}"/>` : '';
  return `<w:rPr>${strike}<w:color w:val="${style.color}"/>${underline}</w:rPr>`;
}

/** Said on the cover page when an upload came with tracked changes of its own: the reader must know what was compared. */
export const ACCEPTED_NOTE = {
  original: 'Note: the original contained tracked changes. It was compared as if they had all been accepted.',
  modified: 'Note: the modified document contained tracked changes. It was compared as if they had all been accepted.',
  both: 'Note: both documents contained tracked changes. They were compared as if those had all been accepted.',
} as const;

/** The lines of the cover page, in order. Exported so tests can state the expected wording once. */
export function coverLines(content: CoverPageContent): { style: string; text: string; kind?: ChangeKind; shaded?: boolean }[] {
  return [
    { style: COVER_STYLE.title, text: 'Comparison of documents' },
    { style: COVER_STYLE.text, text: `Original: ${content.originalName}` },
    { style: COVER_STYLE.text, text: `Modified: ${content.modifiedName}` },
    { style: COVER_STYLE.text, text: `Compared: ${content.comparedAt}` },
    ...(content.trackedChangesIn ? [{ style: COVER_STYLE.text, text: ACCEPTED_NOTE[content.trackedChangesIn] }] : []),
    { style: COVER_STYLE.heading, text: 'Legend' },
    ...(Object.keys(LEGEND) as ChangeKind[]).map((kind) => ({ style: COVER_STYLE.text, text: LEGEND[kind], kind })),
    { style: COVER_STYLE.text, text: 'Text whose formatting changed', shaded: true },
    { style: COVER_STYLE.heading, text: 'Changes' },
    { style: COVER_STYLE.text, text: `Insertions: ${content.insertions}` },
    { style: COVER_STYLE.text, text: `Deletions: ${content.deletions}` },
    { style: COVER_STYLE.text, text: `Moves: ${content.moves}` },
    { style: COVER_STYLE.text, text: `Formatting changes: ${content.formattingChanges}` },
  ];
}

/**
 * The cover page as Word XML: its paragraphs, then an empty paragraph that ends
 * the cover's own section. `pageSetup` is the page size and margins XML copied
 * from the document's first section, so the cover matches the document's paper.
 */
export function coverPageXml(content: CoverPageContent, pageSetup: string): string {
  const body = coverLines(content)
    .map((line) => paragraph(line.style, line.text, line.kind ? changeRunProperties(line.kind) : line.shaded ? `<w:rPr><w:shd w:val="clear" w:color="auto" w:fill="${FORMAT_CHANGE_FILL}"/></w:rPr>` : ''))
    .join('');
  const sectionEnd = `<w:p><w:pPr><w:pStyle w:val="${COVER_STYLE.text}"/><w:sectPr>${pageSetup}</w:sectPr></w:pPr></w:p>`;
  return body + sectionEnd;
}

/** The cover page's paragraph styles, to be added to the document's styles part. */
export function coverStylesXml(basedOn: string | undefined): string {
  const base = basedOn ? `<w:basedOn w:val="${escapeXml(basedOn)}"/>` : '';
  const style = (id: string, name: string, pPr: string, rPr: string): string =>
    `<w:style w:type="paragraph" w:customStyle="1" w:styleId="${id}"><w:name w:val="${name}"/>${base}<w:semiHidden/><w:pPr>${pPr}</w:pPr><w:rPr>${rPr}</w:rPr></w:style>`;
  // Each style states its own layout fully, so the cover looks the same whatever
  // the document's body style does. Properties are listed in schema order.
  const pPr = (before: number, after: number): string =>
    `<w:keepNext w:val="0"/><w:pageBreakBefore w:val="0"/><w:numPr><w:numId w:val="0"/></w:numPr><w:spacing w:before="${before}" w:after="${after}" w:line="240" w:lineRule="auto"/><w:ind w:left="0" w:right="0" w:firstLine="0"/><w:jc w:val="left"/>`;
  const rPr = (bold: boolean, size?: number): string =>
    `<w:b w:val="${bold ? 1 : 0}"/><w:i w:val="0"/><w:caps w:val="0"/><w:smallCaps w:val="0"/><w:strike w:val="0"/><w:dstrike w:val="0"/><w:vanish w:val="0"/><w:color w:val="000000"/>${size ? `<w:sz w:val="${size}"/><w:szCs w:val="${size}"/>` : ''}<w:u w:val="none"/>`;
  // The old clause number shown on a removed paragraph: a character style that changes nothing by itself.
  const oldNumber = `<w:style w:type="character" w:customStyle="1" w:styleId="${OLD_NUMBER_STYLE}"><w:name w:val="Redline Old Number"/><w:semiHidden/></w:style>`;
  return (
    oldNumber +
    style(COVER_STYLE.title, 'Redline Cover Title', pPr(0, 360), rPr(true, 36)) +
    style(COVER_STYLE.heading, 'Redline Cover Heading', pPr(360, 120), rPr(true)) +
    style(COVER_STYLE.text, 'Redline Cover Text', pPr(0, 80), rPr(false))
  );
}
