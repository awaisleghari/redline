/**
 * The one form of paragraph marks (src/producer/paragraph-marks.ts): lines that end up ending no paragraph in
 * either reading. Found on a real contract (2026-11-27): an added clause holding some of the original's words,
 * followed by the empty line the original's paragraph used to end on. The static redline printed that empty line
 * as a struck-out paragraph of its own, and the Checker rightly withheld the file.
 *
 * Every document here is generated; none comes from a real contract.
 */
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { describe, expect, it } from 'vitest';
import { settleParagraphMarks } from '../../src/producer/paragraph-marks';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const body = (inner: string): string => `<w:document xmlns:w="${W}"><w:body>${inner}</w:body></w:document>`;
const run = (text: string): string => `<w:r><w:t>${text}</w:t></w:r>`;
const added = (text: string): string => `<w:ins w:id="1" w:author="a">${run(text)}</w:ins>`;
const removed = (text: string): string => `<w:del w:id="2" w:author="a"><w:r><w:delText>${text}</w:delText></w:r></w:del>`;
const line = (flag: 'ins' | 'del' | '', content: string, settings = ''): string =>
  `<w:p><w:pPr>${settings}${flag ? `<w:rPr><w:${flag} w:id="3" w:author="a"/></w:rPr>` : ''}</w:pPr>${content}</w:p>`;

function settle(inner: string): { text: string; marks: string[]; xml: string } {
  const document = new DOMParser().parseFromString(body(inner), 'application/xml');
  settleParagraphMarks(document.documentElement as unknown as Element);
  const paragraphs = Array.from(document.getElementsByTagNameNS(W, 'p'));
  const marks = paragraphs.map((p) => {
    const settings = Array.from(p.getElementsByTagNameNS(W, 'pPr')[0]?.getElementsByTagNameNS(W, 'rPr') ?? [])[0];
    const flag = settings && Array.from(settings.childNodes).find((c) => ['ins', 'del'].includes((c as Element).localName ?? ''));
    return `${(flag as Element | undefined)?.localName ?? 'both'}:${p.textContent}`;
  });
  return { text: document.documentElement!.textContent ?? '', marks, xml: new XMLSerializer().serializeToString(document) };
}

describe('lines that end no paragraph in either reading', () => {
  it('drops the empty line the original paragraph used to end on, after an added line holding its words', () => {
    // Original: "Old words kept" (one paragraph, ending on the empty removed line). Modified: "New kept" and what follows.
    const { marks } = settle(line('', run('Before')) + line('ins', removed('Old words ') + added('New ') + run('kept')) + line('del', '') + line('', run('After')));
    expect(marks).toEqual(['both:Before', 'both:Old words New kept', 'both:After']);
  });

  it('keeps what such a line holds, at the front of the next line, in order', () => {
    // The modified document's one paragraph is the first two lines (its words are all on the first); the original has
    // "One" and then "two three". The middle line ends neither's paragraph once the marks are settled.
    const { marks, text } = settle(line('del', added('New. ') + removed('One')) + line('ins', removed('two ')) + line('del', removed('three')) + line('', run('After')));
    expect(text).toBe('New. Onetwo threeAfter');
    expect(marks).toEqual(['both:New. One', 'del:two three', 'both:After']);
  });

  it('leaves the run as it was written when the spare line carries a section break', () => {
    const inner = line('ins', removed('Old ') + added('New')) + line('del', '', '<w:sectPr/>') + line('', run('After'));
    expect(settle(inner).marks).toEqual(['ins:Old New', 'del:', 'both:After']);
  });
});
