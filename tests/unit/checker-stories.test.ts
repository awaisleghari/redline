/**
 * The Checker's own accept-all / reject-all reading, tested directly on
 * hand-written Word XML with known answers.
 */
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { DocxPackage } from '../../src/checker/package';
import { MARK, readStories, type View } from '../../src/checker/stories';

const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

interface Extra {
  footnotes?: string;
  endnotes?: string;
  header?: string;
}

async function pkg(body: string, extra: Extra = {}): Promise<DocxPackage> {
  const zip = new JSZip();
  zip.file('_rels/.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/officeDocument" Target="word/document.xml"/></Relationships>`);
  const rels: string[] = [];
  if (extra.footnotes !== undefined) {
    rels.push(`<Relationship Id="rFn" Type="${REL}/footnotes" Target="footnotes.xml"/>`);
    zip.file('word/footnotes.xml', `<w:footnotes ${NS}>${extra.footnotes}</w:footnotes>`);
  }
  if (extra.endnotes !== undefined) {
    rels.push(`<Relationship Id="rEn" Type="${REL}/endnotes" Target="endnotes.xml"/>`);
    zip.file('word/endnotes.xml', `<w:endnotes ${NS}>${extra.endnotes}</w:endnotes>`);
  }
  if (extra.header !== undefined) {
    rels.push(`<Relationship Id="rHd" Type="${REL}/header" Target="header1.xml"/>`);
    zip.file('word/header1.xml', `<w:hdr ${NS}>${extra.header}</w:hdr>`);
  }
  zip.file('word/_rels/document.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels.join('')}</Relationships>`);
  zip.file('word/document.xml', `<w:document ${NS}><w:body>${body}</w:body></w:document>`);
  return DocxPackage.open(await zip.generateAsync({ type: 'uint8array' }));
}

async function read(body: string, view: View, extra: Extra = {}): Promise<Record<string, string[]>> {
  const reading = readStories(await pkg(body, extra), view);
  expect(reading.problems).toEqual([]);
  return Object.fromEntries(reading.stories.map((s) => [s.key, s.paragraphs]));
}

const r = (text: string): string => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
const p = (inner: string): string => `<w:p>${inner}</w:p>`;
const ins = (inner: string): string => `<w:ins w:id="1" w:author="a">${inner}</w:ins>`;
const del = (text: string): string => `<w:del w:id="2" w:author="a"><w:r><w:delText xml:space="preserve">${text}</w:delText></w:r></w:del>`;
const mark = (kind: string): string => `<w:pPr><w:rPr><w:${kind} w:id="3" w:author="a"/></w:rPr></w:pPr>`;

describe('reading text', () => {
  it('joins runs and keeps every character, including spaces, tabs and line breaks', async () => {
    const body = p(`${r('Pay ')}${r('$300,000')}<w:r><w:tab/><w:t xml:space="preserve"> now</w:t><w:br/><w:t>then</w:t><w:noBreakHyphen/><w:t>later</w:t></w:r>`);
    expect((await read(body, 'accepted')).body).toEqual(['Pay $300,000\t now\nthen‑later']);
  });

  it('keeps empty paragraphs as paragraphs', async () => {
    expect((await read(p(r('One')) + p('') + p(r('Two')), 'accepted')).body).toEqual(['One', '', 'Two']);
  });

  it('ignores field codes but keeps field results', async () => {
    const body = p(`${r('See clause ')}<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> REF _Ref1 </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>${r('7.2')}<w:r><w:fldChar w:fldCharType="end"/></w:r>`);
    expect((await read(body, 'accepted')).body).toEqual(['See clause 7.2']);
  });

  it('reads text inside hyperlinks and content controls', async () => {
    const body = p(`<w:hyperlink r:id="x">${r('link')}</w:hyperlink><w:sdt><w:sdtContent>${r(' control')}</w:sdtContent></w:sdt>`) + `<w:sdt><w:sdtContent>${p(r('block control'))}</w:sdtContent></w:sdt>`;
    expect((await read(body, 'accepted')).body).toEqual(['link control', 'block control']);
  });
});

describe('accept all / reject all', () => {
  it('a replaced word: accepting keeps the new word, rejecting keeps the old', async () => {
    const body = p(`${r('with ')}${ins(r('exceptional'))}${del('reasonable')}${r(' care')}`);
    expect((await read(body, 'accepted')).body).toEqual(['with exceptional care']);
    expect((await read(body, 'rejected')).body).toEqual(['with reasonable care']);
  });

  it('an inserted paragraph disappears entirely when rejected', async () => {
    const body = p(r('One')) + p(mark('ins') + ins(r('New'))) + p(r('Two'));
    expect((await read(body, 'accepted')).body).toEqual(['One', 'New', 'Two']);
    expect((await read(body, 'rejected')).body).toEqual(['One', 'Two']);
  });

  it('a deleted paragraph disappears entirely when accepted', async () => {
    const body = p(r('One')) + p(mark('del') + del('Old')) + p(r('Two'));
    expect((await read(body, 'accepted')).body).toEqual(['One', 'Two']);
    expect((await read(body, 'rejected')).body).toEqual(['One', 'Old', 'Two']);
  });

  it('a deleted paragraph mark alone joins two paragraphs when accepted', async () => {
    const body = p(mark('del') + r('First half, ')) + p(r('second half.'));
    expect((await read(body, 'accepted')).body).toEqual(['First half, second half.']);
    expect((await read(body, 'rejected')).body).toEqual(['First half, ', 'second half.']);
  });

  it('an inserted paragraph mark alone splits a paragraph when accepted', async () => {
    const body = p(mark('ins') + r('First half, ')) + p(r('second half.'));
    expect((await read(body, 'accepted')).body).toEqual(['First half, ', 'second half.']);
    expect((await read(body, 'rejected')).body).toEqual(['First half, second half.']);
  });

  it('a move: the clause is at its destination when accepted and at its origin when rejected', async () => {
    const from = p(`${mark('del')}<w:moveFromRangeStart w:id="5" w:name="m"/><w:moveFrom w:id="6">${r('Moved clause.')}</w:moveFrom><w:moveFromRangeEnd w:id="5"/>`);
    const to = p(`${mark('ins')}<w:moveToRangeStart w:id="7" w:name="m"/><w:moveTo w:id="8">${r('Moved clause.')}</w:moveTo><w:moveToRangeEnd w:id="7"/>`);
    const body = from + p(r('Kept.')) + to;
    expect((await read(body, 'accepted')).body).toEqual(['Kept.', 'Moved clause.']);
    expect((await read(body, 'rejected')).body).toEqual(['Moved clause.', 'Kept.']);
  });

  it('text left at the very end by a removed paragraph mark is never lost', async () => {
    const body = p(r('One')) + p(mark('del') + r('Tail'));
    expect((await read(body, 'accepted')).body).toEqual(['One', 'Tail']);
  });

  it('ordinary text inside a deletion (as in a deleted text box) counts as deleted', async () => {
    const body = p(`<w:del w:id="1">${r('plain inside deletion')}</w:del>${r('kept')}`);
    expect((await read(body, 'accepted')).body).toEqual(['kept']);
    expect((await read(body, 'rejected')).body).toEqual(['plain inside deletionkept']);
  });
});

describe('tables', () => {
  const cell = (inner: string, props = ''): string => `<w:tc>${props}${inner}</w:tc>`;
  const row = (cells: string, props = ''): string => `<w:tr>${props}${cells}</w:tr>`;

  it('reads cells in order, each paragraph separately', async () => {
    const body = `<w:tbl>${row(cell(p(r('A1'))) + cell(p(r('B1')) + p(r('B1 second'))))}${row(cell(p(r('A2'))) + cell(p('')))}</w:tbl>`;
    expect((await read(body, 'accepted')).body).toEqual(['A1', 'B1', 'B1 second', 'A2', '']);
  });

  it('an inserted row vanishes when rejected; a deleted row vanishes when accepted', async () => {
    const inserted = row(cell(p(mark('ins') + ins(r('new')))), '<w:trPr><w:ins w:id="1"/></w:trPr>');
    const deleted = row(cell(p(mark('del') + del('old'))), '<w:trPr><w:del w:id="2"/></w:trPr>');
    const body = `<w:tbl>${row(cell(p(r('kept'))))}${inserted}${deleted}</w:tbl>`;
    expect((await read(body, 'accepted')).body).toEqual(['kept', 'new']);
    expect((await read(body, 'rejected')).body).toEqual(['kept', 'old']);
  });

  it('a row whose text is tracked but which is not itself tracked leaves an empty cell behind', async () => {
    const body = `<w:tbl>${row(cell(p(r('kept'))))}${row(cell(p(mark('ins') + ins(r('new')))))}</w:tbl>`;
    expect((await read(body, 'rejected')).body).toEqual(['kept', '']);
  });

  it('an inserted or deleted cell follows the same rule', async () => {
    const body = `<w:tbl>${row(cell(p(r('kept'))) + cell(p(r('added')), '<w:tcPr><w:cellIns w:id="1"/></w:tcPr>') + cell(p(r('removed')), '<w:tcPr><w:cellDel w:id="2"/></w:tcPr>'))}</w:tbl>`;
    expect((await read(body, 'accepted')).body).toEqual(['kept', 'added']);
    expect((await read(body, 'rejected')).body).toEqual(['kept', 'removed']);
  });

  it('text never leaks from one cell into the next', async () => {
    const body = `<w:tbl>${row(cell(p(mark('del') + r('left'))) + cell(p(r('right'))))}</w:tbl>`;
    expect((await read(body, 'accepted')).body).toEqual(['left', 'right']);
  });
});

describe('notes, headers and text boxes', () => {
  const footnotes = '<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:id="7"><w:p><w:r><w:footnoteRef/></w:r><w:r><w:t>Second note.</w:t></w:r></w:p></w:footnote><w:footnote w:id="3"><w:p><w:r><w:t>First note.</w:t></w:r></w:p></w:footnote>';

  it('reads footnotes in the order the text refers to them, whatever their ids', async () => {
    const body = p(`${r('A')}<w:r><w:footnoteReference w:id="3"/></w:r>${r('B')}<w:r><w:footnoteReference w:id="7"/></w:r>`);
    const stories = await read(body, 'accepted', { footnotes });
    expect(stories.body).toEqual([`A${MARK.footnote}B${MARK.footnote}`]);
    expect(stories['footnote/1']).toEqual(['First note.']);
    expect(stories['footnote/2']).toEqual(['Second note.']);
  });

  it('a footnote whose reference is deleted exists only when rejected', async () => {
    const body = p(`${r('A')}<w:del w:id="1"><w:r><w:footnoteReference w:id="3"/></w:r></w:del>`);
    expect(Object.keys(await read(body, 'accepted', { footnotes }))).toEqual(['body']);
    expect((await read(body, 'rejected', { footnotes }))['footnote/1']).toEqual(['First note.']);
  });

  it('reports a reference to a footnote that does not exist', async () => {
    const reading = readStories(await pkg(p(`${r('A')}<w:r><w:footnoteReference w:id="99"/></w:r>`), { footnotes }), 'accepted');
    expect(reading.problems.join(' ')).toContain('does not exist');
  });

  it('reads endnotes the same way', async () => {
    const body = p(`${r('A')}<w:r><w:endnoteReference w:id="2"/></w:r>`);
    const stories = await read(body, 'accepted', { endnotes: '<w:endnote w:id="2"><w:p><w:r><w:t>End.</w:t></w:r></w:p></w:endnote>' });
    expect(stories.body).toEqual([`A${MARK.endnote}`]);
    expect(stories['endnote/1']).toEqual(['End.']);
  });

  it('reads headers through the section that refers to them, applying tracked changes', async () => {
    const body = p(r('Body')) + '<w:sectPr><w:headerReference w:type="default" r:id="rHd"/></w:sectPr>';
    const header = p(`${r('Draft ')}${ins(r('2'))}${del('1')}`);
    expect((await read(body, 'accepted', { header }))['header/1/default']).toEqual(['Draft 2']);
    expect((await read(body, 'rejected', { header }))['header/1/default']).toEqual(['Draft 1']);
  });

  it('reads a text box as its own flow, and exactly once when Word stores it twice', async () => {
    const box = '<w:txbxContent><w:p><w:r><w:t>Boxed</w:t></w:r></w:p><w:p><w:r><w:t>text</w:t></w:r></w:p></w:txbxContent>';
    const body = p(`${r('Host')}<w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing><wps:txbx>${box}</wps:txbx></w:drawing></mc:Choice><mc:Fallback><w:pict><v:textbox>${box}</v:textbox></w:pict></mc:Fallback></mc:AlternateContent></w:r>`);
    const stories = await read(body, 'accepted');
    expect(stories.body).toEqual([`Host${MARK.object}`]);
    expect(stories['body/text boxes']).toEqual(['Boxed', 'text']);
  });

  it('a picture or other object leaves a marker, so its removal is noticed', async () => {
    expect((await read(p(`${r('Logo: ')}<w:r><w:drawing/></w:r>`), 'accepted')).body).toEqual([`Logo: ${MARK.object}`]);
  });
});
