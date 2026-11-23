/**
 * The Checker's reading of a PDF: every visible character, in the order the
 * file draws it, with the colour it is drawn in. Uses PDF.js (Mozilla's PDF
 * reader), which the Producer does not use.
 */
import * as pdfjs from 'pdfjs-dist/legacy/build/pdf.mjs';
import * as pdfWorker from 'pdfjs-dist/legacy/build/pdf.worker.mjs';

// Run PDF.js's parser in this thread: the Checker already lives in its own
// Web Worker (and in Node for the tests), so no second worker is started.
(globalThis as unknown as { pdfjsWorker: unknown }).pdfjsWorker = pdfWorker;

export interface PdfChar {
  /** One character (never white space). */
  c: string;
  /** Fill colour as upper-case RRGGBB, or '?' when it is not a plain colour. */
  color: string;
  /** Page number, from 1. */
  page: number;
}

export interface PdfReading {
  pages: number;
  chars: PdfChar[];
}

export class PdfError extends Error {}

const hex = (n: number): string => Math.max(0, Math.min(255, Math.round(n))).toString(16).padStart(2, '0').toUpperCase();

function colourOf(args: unknown[]): string {
  if (typeof args[0] === 'string') return args[0].replace('#', '').toUpperCase();
  if (args.length >= 3 && args.every((v) => typeof v === 'number')) return hex(args[0] as number) + hex(args[1] as number) + hex(args[2] as number);
  return '?';
}

/** Characters that take no ink: spaces of every kind and invisible joiners. */
export const isBlank = (c: string): boolean => /[\s ­​-‍⁠﻿]/.test(c);

export async function readPdf(bytes: Uint8Array): Promise<PdfReading> {
  if (!(bytes.length > 4 && bytes[0] === 0x25 && bytes[1] === 0x50 && bytes[2] === 0x44 && bytes[3] === 0x46)) throw new PdfError('The file is not a PDF.');
  const task = pdfjs.getDocument({ data: bytes.slice(), disableFontFace: true, useSystemFonts: false, verbosity: 0 });
  let document: pdfjs.PDFDocumentProxy;
  try {
    document = await task.promise;
  } catch {
    throw new PdfError('The PDF could not be opened.');
  }
  const OPS = pdfjs.OPS;
  const chars: PdfChar[] = [];
  try {
    for (let pageNumber = 1; pageNumber <= document.numPages; pageNumber++) {
      const page = await document.getPage(pageNumber);
      const list = await page.getOperatorList();
      let colour = '000000';
      let invisible = false;
      const stack: { colour: string; invisible: boolean }[] = [];
      for (let i = 0; i < list.fnArray.length; i++) {
        const fn = list.fnArray[i];
        const args = (list.argsArray[i] ?? []) as unknown[];
        if (fn === OPS.save) stack.push({ colour, invisible });
        else if (fn === OPS.restore) ({ colour, invisible } = stack.pop() ?? { colour, invisible });
        else if (fn === OPS.setFillRGBColor) colour = colourOf(args);
        else if (fn === OPS.setFillGray || fn === OPS.setFillCMYKColor || fn === OPS.setFillColor || fn === OPS.setFillColorN) colour = colourOf(args);
        else if (fn === OPS.setTextRenderingMode) invisible = args[0] === 3;
        else if (fn === OPS.showText || fn === OPS.showSpacedText || fn === OPS.nextLineShowText || fn === OPS.nextLineSetSpacingShowText) {
          if (invisible) continue;
          const glyphs = (Array.isArray(args[0]) ? args[0] : []) as unknown[];
          for (const glyph of glyphs) {
            if (typeof glyph !== 'object' || glyph === null) continue;
            const unicode = (glyph as { unicode?: unknown }).unicode;
            if (typeof unicode !== 'string') continue;
            // Ligatures and compatibility forms are spelled out, so "ﬁ" compares equal to "fi".
            for (const c of unicode.normalize('NFKC')) if (!isBlank(c)) chars.push({ c, color: colour, page: pageNumber });
          }
        }
      }
      page.cleanup();
    }
  } catch (error) {
    if (error instanceof PdfError) throw error;
    throw new PdfError('The PDF could not be read.');
  } finally {
    await task.destroy();
  }
  return { pages: document.numPages, chars };
}
