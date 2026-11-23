/**
 * The tracked-changes reader is what the Stage 1 tests rely on, so it gets its
 * own test against hand-written XML with a known answer.
 */
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { readTrackedChanges } from '../helpers/tracked-changes';

const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

async function docxWithBody(body: string): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file('word/document.xml', `<?xml version="1.0"?><w:document ${NS}><w:body>${body}</w:body></w:document>`);
  return zip.generateAsync({ type: 'uint8array' });
}

describe('readTrackedChanges', () => {
  it('reads insertions, deletions and moves with their exact text', async () => {
    const docx = await docxWithBody(
      '<w:p><w:r><w:t xml:space="preserve">Pay </w:t></w:r>' +
        '<w:del w:id="1"><w:r><w:delText>$300,000</w:delText></w:r></w:del>' +
        '<w:ins w:id="2"><w:r><w:t>$30,000</w:t></w:r></w:ins>' +
        '<w:r><w:t xml:space="preserve"> now.</w:t></w:r></w:p>' +
        '<w:p><w:moveFromRangeStart w:id="3" w:name="m"/><w:moveFrom w:id="4"><w:r><w:t>Moved clause.</w:t></w:r></w:moveFrom><w:moveFromRangeEnd w:id="3"/></w:p>' +
        '<w:p><w:moveTo w:id="5"><w:r><w:t>Moved clause.</w:t></w:r></w:moveTo></w:p>',
    );
    const found = await readTrackedChanges(docx);
    expect(found.parts.body).toEqual({
      insertions: ['$30,000'],
      deletions: ['$300,000'],
      movedFrom: ['Moved clause.'],
      movedTo: ['Moved clause.'],
      moveOrder: ['from', 'to'],
      insertedParagraphMarks: 0,
      deletedParagraphMarks: 0,
      insertedRows: 0,
      deletedRows: 0,
    });
  });

  it('joins neighbouring pieces of one change but keeps separated changes apart', async () => {
    const docx = await docxWithBody(
      '<w:p><w:ins w:id="1"><w:r><w:t xml:space="preserve">one </w:t></w:r></w:ins><w:ins w:id="2"><w:r><w:t>two</w:t></w:r></w:ins>' +
        '<w:r><w:t xml:space="preserve"> kept </w:t></w:r>' +
        '<w:ins w:id="3"><w:r><w:t>three</w:t></w:r></w:ins></w:p>' +
        '<w:p><w:ins w:id="4"><w:r><w:t>four</w:t></w:r></w:ins></w:p>',
    );
    const found = await readTrackedChanges(docx);
    expect(found.parts.body.insertions).toEqual(['one two', 'three', 'four']);
  });

  it('does not count unchanged text, and counts formatting changes', async () => {
    const docx = await docxWithBody(
      '<w:p><w:r><w:rPr><w:b/><w:rPrChange w:id="1"><w:rPr/></w:rPrChange></w:rPr><w:t>Bold now</w:t></w:r></w:p>',
    );
    const found = await readTrackedChanges(docx);
    expect(found.parts.body.insertions).toEqual([]);
    expect(found.parts.body.deletions).toEqual([]);
    expect(found.formatChanges).toBe(1);
  });

  it('records which way round a move is', async () => {
    const docx = await docxWithBody(
      '<w:p><w:moveTo w:id="1"><w:r><w:t>Clause.</w:t></w:r></w:moveTo></w:p>' +
        '<w:p><w:r><w:t>Kept.</w:t></w:r></w:p>' +
        '<w:p><w:moveFrom w:id="2"><w:r><w:t>Clause.</w:t></w:r></w:moveFrom></w:p>',
    );
    const found = await readTrackedChanges(docx);
    expect(found.parts.body.moveOrder).toEqual(['to', 'from']);
  });

  it('counts tracked paragraph marks and tracked table rows', async () => {
    const docx = await docxWithBody(
      '<w:p><w:pPr><w:rPr><w:ins w:id="1"/></w:rPr></w:pPr><w:ins w:id="2"><w:r><w:t>New paragraph.</w:t></w:r></w:ins></w:p>' +
        '<w:p><w:pPr><w:rPr><w:del w:id="3"/></w:rPr></w:pPr><w:del w:id="4"><w:r><w:delText>Old paragraph.</w:delText></w:r></w:del></w:p>' +
        '<w:p><w:ins w:id="5"><w:r><w:t>text only</w:t></w:r></w:ins></w:p>' +
        '<w:tbl><w:tr><w:trPr><w:ins w:id="6"/></w:trPr><w:tc><w:p><w:ins w:id="7"><w:r><w:t>cell</w:t></w:r></w:ins></w:p></w:tc></w:tr>' +
        '<w:tr><w:trPr><w:del w:id="8"/></w:trPr><w:tc><w:p><w:del w:id="9"><w:r><w:delText>gone</w:delText></w:r></w:del></w:p></w:tc></w:tr></w:tbl>',
    );
    const found = await readTrackedChanges(docx);
    expect(found.parts.body.insertedParagraphMarks).toBe(1);
    expect(found.parts.body.deletedParagraphMarks).toBe(1);
    expect(found.parts.body.insertedRows).toBe(1);
    expect(found.parts.body.deletedRows).toBe(1);
    expect(found.parts.body.insertions).toEqual(['New paragraph.', 'text only', 'cell']);
    expect(found.parts.body.deletions).toEqual(['Old paragraph.', 'gone']);
  });

  it('finds changes inside hyperlinks and content controls', async () => {
    const docx = await docxWithBody(
      '<w:p><w:hyperlink><w:ins w:id="1"><w:r><w:t>linked</w:t></w:r></w:ins></w:hyperlink>' +
        '<w:sdt><w:sdtPr/><w:sdtContent><w:del w:id="2"><w:r><w:delText>controlled</w:delText></w:r></w:del></w:sdtContent></w:sdt></w:p>',
    );
    const found = await readTrackedChanges(docx);
    expect(found.parts.body.insertions).toEqual(['linked']);
    expect(found.parts.body.deletions).toEqual(['controlled']);
  });
});
