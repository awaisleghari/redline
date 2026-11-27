/**
 * The copy of the static redline sent to the PDF converter (src/producer/print-copy.ts): cross-references hold
 * their stored text, rows of fixed height may grow, and everything else is as it was.
 * Every document here is generated; none comes from a real contract.
 */
import { describe, expect, it } from 'vitest';
import { unzip, zip } from '../../src/common/zip';
import { makePrintCopy } from '../../src/producer/print-copy';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const encoder = new TextEncoder();
const decoder = new TextDecoder();
const mark = (type: string): string => `<w:r><w:fldChar w:fldCharType="${type}"/></w:r>`;
const field = (instruction: string, result: string): string => `${mark('begin')}<w:r><w:instrText xml:space="preserve"> ${instruction} </w:instrText></w:r>${mark('separate')}${result}${mark('end')}`;
const text = (t: string): string => `<w:r><w:t xml:space="preserve">${t}</w:t></w:r>`;

async function copyOf(body: string, part = 'word/document.xml'): Promise<string> {
  const xml = `<w:document xmlns:w="${W}"><w:body>${body}</w:body></w:document>`;
  const copy = await makePrintCopy(await zip(new Map([[part, encoder.encode(xml)]])));
  return decoder.decode((await unzip(copy)).get(part)!);
}
const shown = (xml: string): string => xml.replace(/<w:instrText[^>]*>[^<]*<\/w:instrText>/g, '').replace(/<[^>]+>/g, '');

describe('the copy sent to the PDF converter', () => {
  it('replaces a cross-reference by its stored text', async () => {
    const xml = await copyOf(`<w:p>${text('See clause ')}${field('REF _Ref123 \\r \\h', text('4.2'))}${text(' below.')}</w:p>`);
    expect(xml).not.toContain('fldChar');
    expect(xml).not.toContain('instrText');
    expect(shown(xml)).toBe('See clause 4.2 below.');
  });

  it('does the same for a repeated heading, a note reference and a cross-reference written as one element', async () => {
    const xml = await copyOf(`<w:p>${field('STYLEREF "Heading 1"', text('Preface'))}${field('NOTEREF _Ref9 \\h', text('3'))}<w:fldSimple w:instr=" REF _Ref5 \\h ">${text('7.1')}</w:fldSimple></w:p>`);
    expect(xml).not.toMatch(/fldChar|instrText|fldSimple/);
    expect(shown(xml)).toBe('Preface37.1');
  });

  it('leaves page numbers live, in a footer and inside another field', async () => {
    const footer = `<w:p>${text('Page ')}${field('PAGE', text('1'))}${text(' of ')}<w:fldSimple w:instr=" NUMPAGES ">${text('9')}</w:fldSimple></w:p>`;
    expect(await copyOf(footer, 'word/footer1.xml')).toContain(footer);
    const holding = `<w:p>${mark('begin')}<w:r><w:instrText> REF x </w:instrText></w:r>${mark('separate')}${field('PAGE', text('1'))}${mark('end')}</w:p>`;
    expect(await copyOf(holding)).toContain(holding);
  });

  it('leaves a table of contents a field, with its page references', async () => {
    const entry = `<w:p>${mark('begin')}<w:r><w:instrText> TOC \\h </w:instrText></w:r>${mark('separate')}${text('1. Definitions')}${field('PAGEREF _Toc1 \\h', text('4'))}</w:p><w:p>${text('2. Term')}${field('PAGEREF _Toc2 \\h', text('5'))}${mark('end')}</w:p>`;
    expect(await copyOf(entry)).toContain(entry);
  });

  it('freezes a cross-reference that stands inside another field\'s result', async () => {
    const xml = await copyOf(`<w:p>${mark('begin')}<w:r><w:instrText> HYPERLINK "https://example.com" </w:instrText></w:r>${mark('separate')}${text('see ')}${field('REF a', text('2.1'))}${mark('end')}</w:p>`);
    expect(xml.match(/<w:fldChar /g)?.length).toBe(3);
    expect(xml).toContain('HYPERLINK');
    expect(xml).not.toContain(' REF a ');
    expect(shown(xml)).toBe('see 2.1');
  });

  it('leaves a field without an end as it was', async () => {
    const open = `<w:p>${mark('begin')}<w:r><w:instrText> REF a </w:instrText></w:r>${mark('separate')}${text('2.1')}</w:p>`;
    expect(await copyOf(open)).toContain(open);
  });

  it('lets a row of fixed height grow, and leaves other rows alone', async () => {
    const xml = await copyOf('<w:tbl><w:tr><w:trPr><w:trHeight w:val="284" w:hRule="exact"/></w:trPr><w:tc><w:p/></w:tc></w:tr><w:tr><w:trPr><w:trHeight w:val="500"/></w:trPr><w:tc><w:p/></w:tc></w:tr></w:tbl>');
    expect(xml).toContain('<w:trHeight w:val="284" w:hRule="atLeast"/>');
    expect(xml).toContain('<w:trHeight w:val="500"/>');
  });

  it('does not touch a part that needs nothing', async () => {
    const plain = `<w:document xmlns:w="${W}"><w:body><w:p>${text('Nothing to do.')}</w:p></w:body></w:document>`;
    const copy = await makePrintCopy(await zip(new Map([['word/document.xml', encoder.encode(plain)], ['word/media/image1.png', new Uint8Array([1, 2, 3])]])));
    const files = await unzip(copy);
    expect(decoder.decode(files.get('word/document.xml')!)).toBe(plain);
    expect([...files.get('word/media/image1.png')!]).toEqual([1, 2, 3]);
  });
});
