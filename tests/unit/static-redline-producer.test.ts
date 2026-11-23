/**
 * The Producer's conversion of tracked changes into static redline formatting,
 * tested on hand-written Word XML. (Whether the result is *right* is the
 * Checker's job, tested separately; these tests pin down the conversion itself.)
 */
import { describe, expect, it } from 'vitest';
import { coverLines } from '../../src/producer/cover-page';
import { convertPart, DEFAULT_STYLE_OPTIONS } from '../../src/producer/static-redline';

const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const doc = (body: string): string => `<w:document ${NS}><w:body>${body}</w:body></w:document>`;
const body = (xml: string): string => xml.replace(/^[\s\S]*<w:body>/, '').replace(/<\/w:body>[\s\S]*$/, '');
const convert = (inner: string, options = DEFAULT_STYLE_OPTIONS): { xml: string; counts: ReturnType<typeof convertPart>['counts'] } => {
  const result = convertPart(doc(inner), options);
  return { xml: body(result.xml), counts: result.counts };
};
const asWritten = { deletionBeforeInsertion: false, unnumberRemovedParagraphs: false };

const INS = '<w:rPr><w:color w:val="0000FF"/><w:u w:val="double"/></w:rPr>';
const DEL = '<w:rPr><w:strike/><w:color w:val="FF0000"/></w:rPr>';
const FROM = '<w:rPr><w:dstrike/><w:color w:val="008000"/></w:rPr>';
const TO = '<w:rPr><w:color w:val="008000"/><w:u w:val="double"/></w:rPr>';

describe('static redline conversion', () => {
  it('styles an insertion blue with a double underline and removes the tracked change', () => {
    const { xml, counts } = convert('<w:p><w:r><w:t>a </w:t></w:r><w:ins w:id="1"><w:r><w:t>new</w:t></w:r></w:ins></w:p>');
    expect(xml).toBe(`<w:p><w:r><w:t>a </w:t></w:r><w:r>${INS}<w:t>new</w:t></w:r></w:p>`);
    expect(counts).toEqual({ inserted: 1, deleted: 0, movedFrom: 0, movedTo: 0 });
  });

  it('styles a deletion red with a strikethrough and turns deleted text into ordinary text, keeping its spaces', () => {
    const { xml, counts } = convert('<w:p><w:del w:id="1"><w:r><w:delText xml:space="preserve"> old </w:delText></w:r></w:del></w:p>');
    expect(xml).toBe(`<w:p><w:r>${DEL}<w:t xml:space="preserve"> old </w:t></w:r></w:p>`);
    expect(counts.deleted).toBe(1);
  });

  it('styles moves green: double strikethrough at the origin, double underline at the destination, and removes the range markers', () => {
    const { xml, counts } = convert(
      '<w:p><w:pPr><w:rPr><w:del w:id="1"/></w:rPr></w:pPr><w:moveFromRangeStart w:id="2" w:name="m"/><w:moveFrom w:id="3"><w:r><w:t>Clause.</w:t></w:r></w:moveFrom><w:moveFromRangeEnd w:id="2"/></w:p>' +
        '<w:p><w:r><w:t>Kept.</w:t></w:r></w:p>' +
        '<w:p><w:pPr><w:rPr><w:ins w:id="4"/></w:rPr></w:pPr><w:moveToRangeStart w:id="5" w:name="m"/><w:moveTo w:id="6"><w:r><w:t>Clause.</w:t></w:r></w:moveTo><w:moveToRangeEnd w:id="5"/></w:p>',
      asWritten,
    );
    expect(xml).toBe(
      `<w:p><w:pPr>${FROM}</w:pPr><w:r>${FROM}<w:t>Clause.</w:t></w:r></w:p>` + '<w:p><w:r><w:t>Kept.</w:t></w:r></w:p>' + `<w:p><w:pPr>${TO}</w:pPr><w:r>${TO}<w:t>Clause.</w:t></w:r></w:p>`,
    );
    // The paragraph mark of a moved paragraph belongs to the move: one move, no stray insertion or deletion.
    expect(counts).toEqual({ inserted: 0, deleted: 0, movedFrom: 1, movedTo: 1 });
  });

  it('keeps existing formatting and overrides only what the change style defines', () => {
    const { xml } = convert('<w:p><w:del w:id="1"><w:r><w:rPr><w:b/><w:dstrike/><w:color w:val="333333"/><w:sz w:val="28"/><w:u w:val="single"/></w:rPr><w:delText>x</w:delText></w:r></w:del></w:p>');
    expect(xml).toBe('<w:p><w:r><w:rPr><w:b/><w:strike/><w:color w:val="FF0000"/><w:sz w:val="28"/><w:u w:val="single"/></w:rPr><w:t>x</w:t></w:r></w:p>');
  });

  it('styles the paragraph mark of an inserted or deleted paragraph, so its number takes the colour too', () => {
    const { xml } = convert('<w:p><w:pPr><w:jc w:val="left"/><w:rPr><w:ins w:id="1"/></w:rPr></w:pPr><w:ins w:id="2"><w:r><w:t>New.</w:t></w:r></w:ins></w:p>');
    expect(xml).toBe(`<w:p><w:pPr><w:jc w:val="left"/>${INS}</w:pPr><w:r>${INS}<w:t>New.</w:t></w:r></w:p>`);
  });

  it('removes formatting-change and numbering-change records, keeping the new state', () => {
    const { xml } = convert(
      '<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="2"/><w:numberingChange w:id="1" w:original="(c)"/></w:numPr><w:jc w:val="center"/><w:pPrChange w:id="2"><w:pPr/></w:pPrChange></w:pPr>' +
        '<w:r><w:rPr><w:b/><w:rPrChange w:id="3"><w:rPr/></w:rPrChange></w:rPr><w:t>Bold now</w:t></w:r></w:p>',
    );
    expect(xml).toBe('<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="2"/></w:numPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t>Bold now</w:t></w:r></w:p>');
  });

  it('removes row and cell change flags, leaving the rows and cells in place', () => {
    const { xml } = convert(
      '<w:tbl><w:tr><w:trPr><w:ins w:id="1"/></w:trPr><w:tc><w:tcPr><w:cellIns w:id="2"/></w:tcPr><w:p><w:pPr><w:rPr><w:ins w:id="3"/></w:rPr></w:pPr><w:ins w:id="4"><w:r><w:t>x</w:t></w:r></w:ins></w:p></w:tc></w:tr></w:tbl>',
    );
    expect(xml).toBe(`<w:tbl><w:tr><w:trPr/><w:tc><w:tcPr/><w:p><w:pPr>${INS}</w:pPr><w:r>${INS}<w:t>x</w:t></w:r></w:p></w:tc></w:tr></w:tbl>`);
  });

  it('leaves no tracked-change markup of any kind', () => {
    const { xml } = convert(
      '<w:p><w:pPr><w:rPr><w:del w:id="1"/></w:rPr><w:sectPr><w:sectPrChange w:id="9"><w:sectPr/></w:sectPrChange></w:sectPr></w:pPr><w:ins w:id="2"><w:r><w:t>a</w:t></w:r></w:ins><w:del w:id="3"><w:r><w:delText>b</w:delText><w:delInstrText>REF</w:delInstrText></w:r></w:del>' +
        '<w:moveFrom w:id="4"><w:r><w:t>c</w:t></w:r></w:moveFrom><w:moveTo w:id="5"><w:r><w:t>d</w:t></w:r></w:moveTo></w:p>',
    );
    expect(xml).not.toMatch(/<w:(ins|del|moveFrom|moveTo|delText|delInstrText|\w+Change|move\w+Range\w+)\b/);
  });

  it('comment anchors are kept exactly where they were', () => {
    const { xml } = convert('<w:p><w:commentRangeStart w:id="0"/><w:ins w:id="1"><w:r><w:t>a</w:t></w:r></w:ins><w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r></w:p>');
    expect(xml).toBe(`<w:p><w:commentRangeStart w:id="0"/><w:r>${INS}<w:t>a</w:t></w:r><w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r></w:p>`);
  });

  it('styles everything inside a deleted text box, including its own paragraphs', () => {
    const { xml, counts } = convert('<w:p><w:del w:id="1"><w:r><w:pict><w:txbxContent><w:p><w:r><w:t>in box</w:t></w:r></w:p></w:txbxContent></w:pict></w:r></w:del></w:p>', asWritten);
    expect(xml).toBe(`<w:p><w:r>${DEL}<w:pict><w:txbxContent><w:p><w:pPr>${DEL}</w:pPr><w:r>${DEL}<w:t>in box</w:t></w:r></w:p></w:txbxContent></w:pict></w:r></w:p>`);
    expect(counts.deleted).toBe(1);
  });
});

describe('layout options', () => {
  const replaced = '<w:p><w:r><w:t>with </w:t></w:r><w:ins w:id="1"><w:r><w:t>new</w:t></w:r></w:ins><w:del w:id="2"><w:r><w:delText>old</w:delText></w:r></w:del><w:r><w:t> care</w:t></w:r></w:p>';

  it('by default shows the deleted text before the inserted text', () => {
    expect(convert(replaced).xml).toBe(`<w:p><w:r><w:t>with </w:t></w:r><w:r>${DEL}<w:t>old</w:t></w:r><w:r>${INS}<w:t>new</w:t></w:r><w:r><w:t> care</w:t></w:r></w:p>`);
  });

  it('can keep the order the engine wrote', () => {
    expect(convert(replaced, asWritten).xml).toBe(`<w:p><w:r><w:t>with </w:t></w:r><w:r>${INS}<w:t>new</w:t></w:r><w:r>${DEL}<w:t>old</w:t></w:r><w:r><w:t> care</w:t></w:r></w:p>`);
  });

  it('does not reorder changes that are separated by unchanged text', () => {
    const apart = '<w:p><w:ins w:id="1"><w:r><w:t>new</w:t></w:r></w:ins><w:r><w:t> kept </w:t></w:r><w:del w:id="2"><w:r><w:delText>old</w:delText></w:r></w:del></w:p>';
    expect(convert(apart).xml).toBe(`<w:p><w:r>${INS}<w:t>new</w:t></w:r><w:r><w:t> kept </w:t></w:r><w:r>${DEL}<w:t>old</w:t></w:r></w:p>`);
  });

  it('by default takes automatic numbering off a deleted paragraph, and only off a deleted paragraph', () => {
    const numbered = (mark: string, content: string): string => `<w:p><w:pPr><w:pStyle w:val="List"/><w:numPr><w:ilvl w:val="1"/><w:numId w:val="2"/></w:numPr>${mark}</w:pPr>${content}</w:p>`;
    const input = numbered('<w:rPr><w:del w:id="1"/></w:rPr>', '<w:del w:id="2"><w:r><w:delText>gone</w:delText></w:r></w:del>') + numbered('', '<w:r><w:t>stays</w:t></w:r>') + numbered('<w:rPr><w:ins w:id="3"/></w:rPr>', '<w:ins w:id="4"><w:r><w:t>new</w:t></w:r></w:ins>');
    const { xml } = convert(input);
    const paragraphs = xml.match(/<w:p>.*?<\/w:p>/g)!;
    expect(paragraphs[0]).toContain('<w:pStyle w:val="List"/><w:numPr><w:numId w:val="0"/></w:numPr>');
    expect(paragraphs[1]).toContain('<w:numPr><w:ilvl w:val="1"/><w:numId w:val="2"/></w:numPr>');
    expect(paragraphs[2]).toContain('<w:numPr><w:ilvl w:val="1"/><w:numId w:val="2"/></w:numPr>');
  });
});

describe('counting changes', () => {
  const p = (inner: string, mark = ''): string => `<w:p>${mark ? `<w:pPr><w:rPr><w:${mark} w:id="9"/></w:rPr></w:pPr>` : ''}${inner}</w:p>`;
  const ins = (t: string): string => `<w:ins w:id="1"><w:r><w:t>${t}</w:t></w:r></w:ins>`;
  const del = (t: string): string => `<w:del w:id="2"><w:r><w:delText>${t}</w:delText></w:r></w:del>`;
  const plain = (t: string): string => `<w:r><w:t>${t}</w:t></w:r>`;

  it('a replaced word is one deletion and one insertion', () => {
    expect(convert(p(plain('a ') + ins('new') + del('old') + plain(' b'))).counts).toEqual({ inserted: 1, deleted: 1, movedFrom: 0, movedTo: 0 });
  });

  it('pieces of one change split across runs count once', () => {
    expect(convert(p(ins('one ') + ins('two'))).counts.inserted).toBe(1);
  });

  it('changes separated by unchanged text count separately', () => {
    expect(convert(p(ins('one') + plain(' kept ') + ins('two'))).counts.inserted).toBe(2);
  });

  it('several whole paragraphs inserted together are one insertion', () => {
    expect(convert(p(plain('before')) + p(ins('first'), 'ins') + p(ins('second'), 'ins') + p(plain('after'))).counts.inserted).toBe(1);
  });

  it('an inserted table row is one insertion however many cells it has', () => {
    const cell = (t: string): string => `<w:tc>${p(ins(t), 'ins')}</w:tc>`;
    expect(convert(`<w:tbl><w:tr><w:tc>${p(plain('kept'))}</w:tc></w:tr><w:tr><w:trPr><w:ins w:id="5"/></w:trPr>${cell('a')}${cell('b')}</w:tr></w:tbl>`).counts.inserted).toBe(1);
  });

  it('changes in different footnotes count separately', () => {
    const xml = `<w:footnotes ${NS}><w:footnote w:id="1">${p(ins('a'))}</w:footnote><w:footnote w:id="2">${p(ins('b'))}</w:footnote></w:footnotes>`;
    expect(convertPart(xml).counts.inserted).toBe(2);
  });
});

describe('cover page', () => {
  it('states both file names, the time, a legend of each style, and the counts', () => {
    const lines = coverLines({ originalName: 'v1.docx', modifiedName: 'v2.docx', comparedAt: '23 November 2026 at 14:05', insertions: 3, deletions: 2, moves: 1 }).map((l) => l.text);
    expect(lines).toEqual([
      'Comparison of documents',
      'Original: v1.docx',
      'Modified: v2.docx',
      'Compared: 23 November 2026 at 14:05',
      'Legend',
      'Inserted text',
      'Deleted text',
      'Text moved from here',
      'Text moved to here',
      'Changes',
      'Insertions: 3',
      'Deletions: 2',
      'Moves: 1',
    ]);
  });
});
