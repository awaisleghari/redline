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
  const look = '<w:color w:val="0000FF"/><w:u w:val="double"/>';
  return editParts(docx, (name, xml) => {
    if (name !== 'word/document.xml') return xml;
    if (where === 'paragraph mark') {
      const at = xml.indexOf('<w:p>');
      if (at < 0) throw new Error('withBlueDoubleUnderline: no plain paragraph found');
      return `${xml.slice(0, at)}<w:p><w:pPr><w:rPr>${look}</w:rPr></w:pPr>${xml.slice(at + '<w:p>'.length)}`;
    }
    const at = xml.lastIndexOf('<w:r><w:t');
    if (at < 0) throw new Error('withBlueDoubleUnderline: no plain run found');
    return `${xml.slice(0, at)}<w:r><w:rPr>${look}</w:rPr>${xml.slice(at + '<w:r>'.length)}`;
  });
}

/**
 * Put the first three numbered clauses into one table cell (clauses inside a table, as in schedules and
 * definitions tables). `lastTwo` says what then happens to the second and third: nothing; they are run
 * together into one clause (an editor deleting the paragraph break between them); or the third is deleted.
 */
export function withNumberedClausesInCell(docx: Uint8Array, lastTwo: 'kept' | 'joined' | 'last deleted' = 'kept'): Promise<Uint8Array> {
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
    const table = `<w:tbl><w:tblPr><w:tblW w:w="9000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="9000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="9000" w:type="dxa"/></w:tcPr>${a![0]}${second}${third}</w:tc></w:tr></w:tbl>`;
    return xml.slice(0, a!.index) + table + xml.slice(c!.index! + c![0].length);
  });
}
