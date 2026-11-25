/**
 * Document shapes found in real use that the generated pairs in pairs.ts do not
 * have. Each is made by reshaping a generated document (never from a real one):
 * the reshaped file must read exactly as before in Word.
 */
import { unzip, zip } from '../../src/common/zip';

const decoder = new TextDecoder();
const encoder = new TextEncoder();

async function editParts(docx: Uint8Array, edit: (name: string, xml: string) => string): Promise<Uint8Array> {
  const files = await unzip(docx);
  for (const [name, bytes] of files) {
    if (!name.endsWith('.xml')) continue;
    const xml = decoder.decode(bytes);
    const edited = edit(name, xml);
    if (edited !== xml) files.set(name, encoder.encode(edited));
  }
  return zip(files);
}

/**
 * As PDF-to-Word converters write text: "keep white space" stated once, on the
 * outermost element, and each word and each space in a run of its own with no
 * setting of its own.
 */
export function withDocumentWideSpaces(docx: Uint8Array): Promise<Uint8Array> {
  return editParts(docx, (name, xml) => {
    if (!/^word\/(document|header\d*|footer\d*|footnotes|endnotes)\.xml$/.test(name)) return xml;
    let out = xml.replace(/<w:r>((?:<w:rPr>(?:(?!<\/w:rPr>).)*<\/w:rPr>)?)<w:t(?: xml:space="preserve")?>([^<]*)<\/w:t><\/w:r>/g, (_run, rPr: string, text: string) =>
      text
        .split(/(?= )|(?<= )/)
        .map((piece) => `<w:r>${rPr}<w:t>${piece}</w:t></w:r>`)
        .join(''),
    );
    out = out.replace(/ xml:space="preserve"/g, '');
    return out.replace(/<(w:(?:document|hdr|ftr|footnotes|endnotes))\b/, '<$1 xml:space="preserve"');
  });
}

/** Give the first paragraph of the text a named style, defined as stated (the same name can then be defined differently in two documents). */
export function withClauseStyle(docx: Uint8Array, definition: string): Promise<Uint8Array> {
  return editParts(docx, (name, xml) => {
    if (name === 'word/styles.xml') return xml.replace('</w:styles>', `<w:style w:type="paragraph" w:styleId="Clause"><w:name w:val="Clause"/>${definition}</w:style></w:styles>`);
    if (name === 'word/document.xml') return xml.replace(/<w:p>(<w:pPr>)?/, (_m, pPr?: string) => (pPr ? '<w:p><w:pPr><w:pStyle w:val="Clause"/>' : '<w:p><w:pPr><w:pStyle w:val="Clause"/></w:pPr>'));
    return xml;
  });
}

/**
 * Cut the text into three sections. Only the first names the header and
 * footer; the later two carry them over ("same as previous"), as Word and
 * converters commonly leave them. With `nameInEvery`, every section names them
 * again instead: the two forms print identically.
 */
export function withSections(docx: Uint8Array, nameInEvery = false): Promise<Uint8Array> {
  return editParts(docx, (name, xml) => {
    if (name !== 'word/document.xml') return xml;
    const last = /<w:sectPr\b[^>]*>((?:(?!<\/w:sectPr>).)*)<\/w:sectPr>(?=<\/w:body>)/.exec(xml);
    if (!last) throw new Error('withSections: no final section found');
    const named = last[1]!;
    const bare = nameInEvery ? named : named.replace(/<w:(header|footer)Reference\b[^>]*\/>/g, '');
    if (!nameInEvery && bare === named) throw new Error('withSections: the document names no header or footer');
    const breakWith = (inner: string): string => `<w:p><w:pPr><w:sectPr>${inner}</w:sectPr></w:pPr></w:p>`;
    let count = 0;
    let out = xml.slice(0, last.index).replace(/<\/w:p>(?=<w:p[ >])/g, (end) => {
      count++;
      return count === 1 ? end + breakWith(named) : count === 2 ? end + breakWith(bare) : end;
    });
    if (count < 2) throw new Error('withSections: the document needs at least three paragraphs');
    out += `<w:sectPr>${bare}</w:sectPr>` + xml.slice(last.index + last[0].length);
    return out;
  });
}

/**
 * Put the paragraph that carries the first footnote mark into a table cell, with
 * a second row beneath it holding `secondRow`. (A table that has a footnote hanging
 * off one of its cells: common in definitions and particulars tables.)
 */
export function withNoteInTable(docx: Uint8Array, secondRow: string): Promise<Uint8Array> {
  return editParts(docx, (name, xml) => {
    if (name !== 'word/document.xml') return xml;
    const mark = xml.indexOf('<w:footnoteReference');
    if (mark < 0) throw new Error('withNoteInTable: the document has no footnote');
    const start = Math.max(xml.lastIndexOf('<w:p>', mark), xml.lastIndexOf('<w:p ', mark));
    const end = xml.indexOf('</w:p>', mark) + '</w:p>'.length;
    const cell = (content: string): string => `<w:tc><w:tcPr><w:tcW w:w="9000" w:type="dxa"/></w:tcPr>${content}</w:tc>`;
    const second = `<w:p><w:r><w:t xml:space="preserve">${secondRow}</w:t></w:r></w:p>`;
    const table = `<w:tbl><w:tblPr><w:tblW w:w="9000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="9000"/></w:tblGrid><w:tr>${cell(xml.slice(start, end))}</w:tr><w:tr>${cell(second)}</w:tr></w:tbl>`;
    return xml.slice(0, start) + table + xml.slice(end);
  });
}

/**
 * Give some of a document's own text the look the redline normally uses for insertions: blue with a
 * double underline. `where` chooses a stretch of text (the last run of the main text) or a paragraph mark only.
 */
export function withBlueDoubleUnderline(docx: Uint8Array, where: 'text' | 'paragraph mark' = 'text'): Promise<Uint8Array> {
  return withLook(docx, '<w:color w:val="0000FF"/><w:u w:val="double"/>', where);
}

/** The same with the look of a deletion: the last stretch of text, or the first paragraph mark, formatted red and struck through. */
export function withRedStrikethrough(docx: Uint8Array, where: 'text' | 'paragraph mark' = 'text'): Promise<Uint8Array> {
  return withLook(docx, '<w:strike/><w:color w:val="FF0000"/>', where);
}

function withLook(docx: Uint8Array, look: string, where: 'text' | 'paragraph mark'): Promise<Uint8Array> {
  return editParts(docx, (name, xml) => {
    if (name !== 'word/document.xml') return xml;
    if (where === 'paragraph mark') {
      const at = xml.indexOf('<w:p>');
      if (at < 0) throw new Error('withLook: no plain paragraph found');
      return `${xml.slice(0, at)}<w:p><w:pPr><w:rPr>${look}</w:rPr></w:pPr>${xml.slice(at + '<w:p>'.length)}`;
    }
    const at = xml.lastIndexOf('<w:r><w:t');
    if (at < 0) throw new Error('withLook: no plain run found');
    return `${xml.slice(0, at)}<w:r><w:rPr>${look}</w:rPr>${xml.slice(at + '<w:r>'.length)}`;
  });
}

/**
 * Put the first three numbered clauses into one table cell (clauses inside a table, as in schedules and
 * definitions tables). `lastTwo` says what then happens to the second and third: nothing; they are run
 * together into one clause (an editor deleting the paragraph break between them); or the third is deleted.
 */
export function withNumberedClausesInCell(docx: Uint8Array, lastTwo: 'kept' | 'joined' | 'last deleted' | 'last deleted, second shortened' = 'kept'): Promise<Uint8Array> {
  return editParts(docx, (name, xml) => {
    if (name !== 'word/document.xml') return xml;
    const paragraphs = [...xml.matchAll(/<w:p>(?:(?!<\/w:p>).)*<w:numPr>(?:(?!<\/w:p>).)*<\/w:p>/gs)];
    if (paragraphs.length < 4) throw new Error('withNumberedClausesInCell: the document needs at least four numbered clauses');
    const [a, b, c] = paragraphs;
    if (a!.index! + a![0].length !== b!.index || b!.index! + b![0].length !== c!.index) throw new Error('withNumberedClausesInCell: the first three numbered clauses are not consecutive');
    let third = c![0];
    let second = b![0];
    if (lastTwo === 'joined') {
      const runs = third.slice(third.indexOf('</w:pPr>') + '</w:pPr>'.length, third.lastIndexOf('</w:p>'));
      second = second.slice(0, second.lastIndexOf('</w:p>')) + '<w:r><w:t xml:space="preserve"> </w:t></w:r>' + runs + '</w:p>';
      third = '';
    } else if (lastTwo === 'last deleted') third = '';
    else if (lastTwo === 'last deleted, second shortened') {
      third = '';
      // Words go from the middle of the second clause; its ending stays as it was.
      const reworded = second.replace(/(<w:t(?: [^>]*)?>)(\S+ \S+ )\S+ \S+ /, '$1$2');
      if (reworded === second) throw new Error('withNumberedClausesInCell: the second clause is too short to shorten');
      second = reworded;
    }
    const table = `<w:tbl><w:tblPr><w:tblW w:w="9000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="9000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="9000" w:type="dxa"/></w:tcPr>${a![0]}${second}${third}</w:tc></w:tr></w:tbl>`;
    return xml.slice(0, a!.index) + table + xml.slice(c!.index! + c![0].length);
  });
}

/** Wrap the last cell of every table row in a content control, as Word stores a tick-box cell in a form or schedule. */
export function withCellsInContentControls(docx: Uint8Array): Promise<Uint8Array> {
  return editParts(docx, (name, xml) => {
    if (name !== 'word/document.xml') return xml;
    let id = 700;
    const out = xml.replace(/(<w:tc>(?:(?!<\/w:tc>).)*<\/w:tc>)(<\/w:tr>)/gs, (_all, cell: string, end: string) => `<w:sdt><w:sdtPr><w:id w:val="${id++}"/></w:sdtPr><w:sdtContent>${cell}</w:sdtContent></w:sdt>${end}`);
    if (out === xml) throw new Error('withCellsInContentControls: the document has no table');
    return out;
  });
}

/**
 * Wrap every table row but the first in a content control, as a "repeating section" in a procurement or form
 * template does (the first row stays an ordinary heading row).
 */
export function withRowsInContentControls(docx: Uint8Array): Promise<Uint8Array> {
  return editParts(docx, (name, xml) => {
    if (name !== 'word/document.xml') return xml;
    let id = 900;
    const out = xml.replace(/<w:tbl>.*?<\/w:tbl>/gs, (table) => {
      let row = 0;
      return table.replace(/<w:tr[ >].*?<\/w:tr>/gs, (tr) => (row++ === 0 ? tr : `<w:sdt><w:sdtPr><w:id w:val="${id++}"/></w:sdtPr><w:sdtContent>${tr}</w:sdtContent></w:sdt>`));
    });
    if (out === xml) throw new Error('withRowsInContentControls: the document has no table with more than one row');
    return out;
  });
}

/**
 * Add a glossary (the store of building blocks a Word template carries, never shown in the document) with a
 * style sheet of its own, as law-firm templates have. The file then holds two style sheets; only the one the
 * main document links to is the document's.
 */
export async function withGlossary(docx: Uint8Array): Promise<Uint8Array> {
  const files = await unzip(docx);
  const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
  const WML = 'application/vnd.openxmlformats-officedocument.wordprocessingml';
  const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  files.set('word/glossary/document.xml', encoder.encode(`<w:glossaryDocument ${W}><w:docParts><w:docPart><w:docPartPr><w:name w:val="Block"/></w:docPartPr><w:docPartBody><w:p><w:r><w:t>A building block</w:t></w:r></w:p></w:docPartBody></w:docPart></w:docParts></w:glossaryDocument>`));
  files.set('word/glossary/styles.xml', encoder.encode(`<w:styles ${W}><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>`));
  files.set('word/glossary/_rels/document.xml.rels', encoder.encode(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/styles" Target="styles.xml"/></Relationships>`));
  files.set('[Content_Types].xml', encoder.encode(decoder.decode(files.get('[Content_Types].xml')!).replace('</Types>', `<Override PartName="/word/glossary/document.xml" ContentType="${WML}.document.glossary+xml"/><Override PartName="/word/glossary/styles.xml" ContentType="${WML}.styles+xml"/></Types>`)));
  files.set('word/_rels/document.xml.rels', encoder.encode(decoder.decode(files.get('word/_rels/document.xml.rels')!).replace('</Relationships>', `<Relationship Id="rIdGlossary" Type="${REL}/glossaryDocument" Target="glossary/document.xml"/></Relationships>`)));
  return zip(files);
}

/** A 1 x 1 pixel picture (PNG). */
const DOT_PNG = Uint8Array.from(atob('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg=='), (c) => c.charCodeAt(0));

/**
 * Put a pen mark at the start of the first paragraph, as Word stores handwriting made with a stylus: the pen
 * strokes themselves (for programs that understand ink) and, as the other form, a picture of them.
 * `picture` is the picture's bytes, so a test can make two documents whose pen marks look different.
 */
export async function withInk(docx: Uint8Array, picture: Uint8Array = DOT_PNG): Promise<Uint8Array> {
  const files = await unzip(docx);
  const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const NAMESPACES =
    'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ' +
    `xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" xmlns:r="${REL}" xmlns:aink="http://schemas.microsoft.com/office/drawing/2016/ink"`;
  const drawing = (graphic: string): string =>
    `<w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="18000" cy="108000"/><wp:docPr id="901" name="Ink 1"/><wp:cNvGraphicFramePr/><a:graphic>${graphic}</a:graphic></wp:inline></w:drawing>`;
  const strokes = '<a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingInk"><w14:contentPart bwMode="auto" r:id="rIdInk"><w14:nvContentPartPr><w14:cNvContentPartPr/></w14:nvContentPartPr><w14:xfrm><a:off x="0" y="0"/><a:ext cx="18000" cy="108000"/></w14:xfrm></w14:contentPart></a:graphicData>';
  const image =
    '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="901" name="Ink 1"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rIdInkPicture"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>' +
    '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="18000" cy="108000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData>';
  const run = `<w:r><mc:AlternateContent ${NAMESPACES}><mc:Choice Requires="aink">${drawing(strokes)}</mc:Choice><mc:Fallback>${drawing(image)}</mc:Fallback></mc:AlternateContent></w:r>`;
  const xml = decoder.decode(files.get('word/document.xml')!);
  const at = xml.indexOf('<w:r>');
  if (at < 0) throw new Error('withInk: the document has no text run to put the pen mark before');
  files.set('word/document.xml', encoder.encode(xml.slice(0, at) + run + xml.slice(at)));
  files.set('word/ink/ink1.xml', encoder.encode('<inkml:ink xmlns:inkml="http://www.w3.org/2003/InkML"><inkml:definitions><inkml:brush xml:id="br0"/></inkml:definitions><inkml:trace contextRef="#ctx0" brushRef="#br0">0 1,\'0\'0</inkml:trace></inkml:ink>'));
  files.set('word/media/ink1.png', picture);
  files.set('[Content_Types].xml', encoder.encode(decoder.decode(files.get('[Content_Types].xml')!).replace('</Types>', '<Override PartName="/word/ink/ink1.xml" ContentType="application/inkml+xml"/><Override PartName="/word/media/ink1.png" ContentType="image/png"/></Types>')));
  files.set(
    'word/_rels/document.xml.rels',
    encoder.encode(decoder.decode(files.get('word/_rels/document.xml.rels')!).replace('</Relationships>', `<Relationship Id="rIdInk" Type="${REL}/customXml" Target="ink/ink1.xml"/><Relationship Id="rIdInkPicture" Type="${REL}/image" Target="media/ink1.png"/></Relationships>`)),
  );
  return zip(files);
}

/**
 * Put the first paragraph into a table cell, followed in the same cell by `added` further paragraphs (each a
 * line of text, or '' for an empty paragraph): a cell to which lines are added, or from which they are removed, at the end.
 */
export function withParagraphsAtCellEnd(docx: Uint8Array, added: string[]): Promise<Uint8Array> {
  return editParts(docx, (name, xml) => {
    if (name !== 'word/document.xml') return xml;
    const first = /<w:p>.*?<\/w:p>/s.exec(xml);
    if (!first) throw new Error('withParagraphsAtCellEnd: the document has no paragraph');
    const more = added.map((text) => (text === '' ? '<w:p/>' : `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`)).join('');
    const table = `<w:tbl><w:tblPr><w:tblW w:w="9000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="9000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="9000" w:type="dxa"/></w:tcPr>${first[0]}${more}</w:tc></w:tr></w:tbl>`;
    return xml.slice(0, first.index) + table + xml.slice(first.index + first[0].length);
  });
}

/** An SVG picture as Word writes one: no heading line, Windows line ends, a style block. */
export const SVG_PICTURE = encoder.encode(
  '<svg viewBox="0 0 96 96" xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink" id="Icons_Information" overflow="hidden"><style>\r\n.Fill {\r\n fill:#558ED5; \r\n}\r\n</style>\r\n<path d="M48 10C27 10 10 27 10 48 10 69 27 86 48 86 69 86 86 69 86 48 86 27 69 10 48 10Z" fill="#558ED5"/></svg>',
);

/**
 * Put an icon at the start of the first paragraph, as Word stores one: a picture stored as text (SVG), with an
 * ordinary picture alongside for programs that cannot draw SVG.
 */
export async function withSvgPicture(docx: Uint8Array, svg: Uint8Array = SVG_PICTURE): Promise<Uint8Array> {
  const files = await unzip(docx);
  const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const run =
    `<w:r><w:drawing xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" xmlns:r="${REL}">` +
    '<wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="180000" cy="180000"/><wp:docPr id="911" name="Graphic 1"/><wp:cNvGraphicFramePr/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
    '<pic:pic><pic:nvPicPr><pic:cNvPr id="911" name="Graphic 1"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rIdIconPicture"><a:extLst><a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}">' +
    '<asvg:svgBlip xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main" r:embed="rIdIconSvg"/></a:ext></a:extLst></a:blip><a:stretch><a:fillRect/></a:stretch></pic:blipFill>' +
    '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="180000" cy="180000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>';
  const xml = decoder.decode(files.get('word/document.xml')!);
  const at = xml.indexOf('<w:r>');
  if (at < 0) throw new Error('withSvgPicture: the document has no text run to put the icon before');
  files.set('word/document.xml', encoder.encode(xml.slice(0, at) + run + xml.slice(at)));
  files.set('word/media/icon1.png', DOT_PNG);
  files.set('word/media/icon1.svg', svg);
  files.set('[Content_Types].xml', encoder.encode(decoder.decode(files.get('[Content_Types].xml')!).replace('</Types>', '<Override PartName="/word/media/icon1.png" ContentType="image/png"/><Override PartName="/word/media/icon1.svg" ContentType="image/svg+xml"/></Types>')));
  files.set(
    'word/_rels/document.xml.rels',
    encoder.encode(decoder.decode(files.get('word/_rels/document.xml.rels')!).replace('</Relationships>', `<Relationship Id="rIdIconPicture" Type="${REL}/image" Target="media/icon1.png"/><Relationship Id="rIdIconSvg" Type="${REL}/image" Target="media/icon1.svg"/></Relationships>`)),
  );
  return zip(files);
}
