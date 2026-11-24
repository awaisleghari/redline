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
