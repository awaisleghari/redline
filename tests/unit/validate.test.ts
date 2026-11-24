/** The page's refusal of uploads it cannot compare, and its recognition of every kind of existing tracked change (compared as accepted, with a notice). */
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { hasTrackedChanges, MAX_FILE_BYTES, containsTrackedChanges, refusalFor } from '../../src/app/validate';

const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

async function docx(body: string, extra: Record<string, string> = {}): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<Types/>');
  zip.file('word/document.xml', `<w:document ${NS}><w:body>${body}</w:body></w:document>`);
  for (const [name, xml] of Object.entries(extra)) zip.file(name, xml);
  return zip.generateAsync({ type: 'uint8array' });
}
const file = (name: string, bytes: Uint8Array, size = bytes.length) => ({ name, size, bytes: async () => bytes });
const refusal = async (body: string, extra?: Record<string, string>): Promise<string | undefined> => refusalFor(file('a.docx', await docx(body, extra)), 'original');

describe('uploads: what is refused, and what is recognised', () => {
  it('accepts a clean Word document, whatever the case of its extension', async () => {
    const clean = await docx('<w:p><w:r><w:t>Insert and delete are just words here: ins del.</w:t></w:r></w:p>');
    expect(await refusalFor(file('a.docx', clean), 'original')).toBeUndefined();
    expect(await refusalFor(file('A.DOCX', clean), 'modified')).toBeUndefined();
  });

  it.each([
    ['inserted text', '<w:p><w:ins w:id="1"><w:r><w:t>x</w:t></w:r></w:ins></w:p>'],
    ['deleted text', '<w:p><w:del w:id="1"><w:r><w:delText>x</w:delText></w:r></w:del></w:p>'],
    ['moved text', '<w:p><w:moveFrom w:id="1"><w:r><w:t>x</w:t></w:r></w:moveFrom></w:p>'],
    ['a tracked paragraph mark', '<w:p><w:pPr><w:rPr><w:ins w:id="1"/></w:rPr></w:pPr></w:p>'],
    ['a character formatting change', '<w:p><w:r><w:rPr><w:b/><w:rPrChange w:id="1"><w:rPr/></w:rPrChange></w:rPr><w:t>x</w:t></w:r></w:p>'],
    ['a paragraph formatting change', '<w:p><w:pPr><w:pPrChange w:id="1"><w:pPr/></w:pPrChange></w:pPr></w:p>'],
    ['a section change', '<w:sectPr><w:sectPrChange w:id="1"><w:sectPr/></w:sectPrChange></w:sectPr>'],
    ['a table property change', '<w:tbl><w:tblPr><w:tblPrChange w:id="1"><w:tblPr/></w:tblPrChange></w:tblPr></w:tbl>'],
    ['a tracked table cell', '<w:tbl><w:tr><w:tc><w:tcPr><w:cellIns w:id="1"/></w:tcPr><w:p/></w:tc></w:tr></w:tbl>'],
    ['a deleted table cell', '<w:tbl><w:tr><w:tc><w:tcPr><w:cellDel w:id="1"/></w:tcPr><w:p/></w:tc></w:tr></w:tbl>'],
    ['a numbering change', '<w:p><w:pPr><w:numPr><w:numberingChange w:id="1" w:original="1."/></w:numPr></w:pPr></w:p>'],
  ])('accepts a document with %s, and recognises that it holds tracked changes', async (_name, body) => {
    // Owner's decision of 2026-11-24: such a document is no longer refused; it is compared as if its changes had been accepted, with a notice.
    expect(await refusal(body)).toBeUndefined();
    expect(await containsTrackedChanges(await docx(body))).toBe(true);
  });

  it('recognises tracked changes under a different namespace prefix, and in headers or footnotes; and none in a clean document', async () => {
    const other = '<x:document xmlns:x="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><x:body><x:p><x:ins x:id="1"><x:r><x:t>y</x:t></x:r></x:ins></x:p></x:body></x:document>';
    expect(hasTrackedChanges(other)).toBe(true);
    const inHeader = { 'word/header1.xml': `<w:hdr ${NS}><w:p><w:del w:id="1"><w:r><w:delText>x</w:delText></w:r></w:del></w:p></w:hdr>` };
    expect(await containsTrackedChanges(await docx('<w:p/>', inHeader))).toBe(true);
    const inNotes = { 'word/footnotes.xml': `<w:footnotes ${NS}><w:footnote w:id="1"><w:p><w:ins w:id="1"><w:r><w:t>x</w:t></w:r></w:ins></w:p></w:footnote></w:footnotes>` };
    expect(await containsTrackedChanges(await docx('<w:p/>', inNotes))).toBe(true);
    expect(await containsTrackedChanges(await docx('<w:p><w:r><w:t>Insert and delete are just words here: ins del.</w:t></w:r></w:p>'))).toBe(false);
  });

  it('refuses old, foreign, oversized, empty and unreadable files in plain words, without reading an oversized file', async () => {
    const never = { bytes: async (): Promise<Uint8Array> => { throw new Error('must not be read'); } };
    expect(await refusalFor({ name: 'old.doc', size: 10, ...never }, 'original')).toMatch(/old-format Word file \(\.doc\)/);
    expect(await refusalFor({ name: 'scan.pdf', size: 10, ...never }, 'modified')).toMatch(/^The modified document is not a Word document/);
    expect(await refusalFor({ name: 'macro.docm', size: 10, ...never }, 'original')).toMatch(/not a Word document/);
    expect(await refusalFor({ name: 'big.docx', size: MAX_FILE_BYTES + 1, ...never }, 'original')).toMatch(/larger than 25 MB/);
    expect(await refusalFor({ name: 'empty.docx', size: 0, ...never }, 'original')).toMatch(/is empty/);
    expect(await refusalFor(file('fake.docx', new TextEncoder().encode('not a zip')), 'original')).toMatch(/could not be read as a Word document/);
  });
});
