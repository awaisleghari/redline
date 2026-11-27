/**
 * The Checker's reading of a PDF: every visible character, in the order the
 * file draws it, with the colour it is drawn in and the place on the page
 * where it is drawn; and every thin straight line (the strokes of an
 * underline or a strikethrough). Uses PDF.js (Mozilla's PDF reader), which
 * the Producer does not use.
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
  /** Where the character starts on its baseline, in points from the left and from the bottom of the page. */
  x: number;
  y: number;
  /** How wide it is drawn, and the size of its type, in points. */
  width: number;
  size: number;
  /** Set when the file marks the character as page furniture (a header, a footer, a repeated table heading), not document content. */
  artifact: boolean;
}

/** A thin horizontal line or bar drawn on a page: an underline, a strikethrough, a border, a rule. */
export interface PdfRule {
  page: number;
  /** From `x1` to `x2`, at height `y` (its middle), `thickness` thick: all in points, measured as for characters. */
  x1: number;
  x2: number;
  y: number;
  thickness: number;
  color: string;
}

export interface PdfReading {
  pages: number;
  chars: PdfChar[];
  rules: PdfRule[];
}

type Matrix = [number, number, number, number, number, number];
const IDENTITY: Matrix = [1, 0, 0, 1, 0, 0];
/** The matrix that applies `first` and then `second`. */
const combine = (first: Matrix, second: Matrix): Matrix => [
  first[0] * second[0] + first[1] * second[2],
  first[0] * second[1] + first[1] * second[3],
  first[2] * second[0] + first[3] * second[2],
  first[2] * second[1] + first[3] * second[3],
  first[4] * second[0] + first[5] * second[2] + second[4],
  first[4] * second[1] + first[5] * second[3] + second[5],
];
const numbers = (values: unknown[]): number[] | undefined => (values.every((v) => typeof v === 'number') ? (values as number[]) : undefined);

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
  const rules: PdfRule[] = [];
  try {
    for (let pageNumber = 1; pageNumber <= document.numPages; pageNumber++) {
      const page = await document.getPage(pageNumber);
      const list = await page.getOperatorList();
      // The drawing state, as a PDF keeps it: colours, the page transform, and where the next character goes.
      let state = { colour: '000000', stroke: '000000', invisible: false, ctm: IDENTITY, lineWidth: 1, fontSize: 0, charSpacing: 0, wordSpacing: 0, scale: 1, rise: 0, leading: 0, widthUnit: 0.001 };
      const stack: (typeof state)[] = [];
      let textMatrix: Matrix = IDENTITY;
      let [x, y, lineX, lineY] = [0, 0, 0, 0];
      /** How deep the drawing is inside content marked as page furniture (0: document content). */
      let artifact = 0;
      const marks: boolean[] = [];
      const moveText = (dx: number, dy: number): void => {
        x = lineX += dx;
        y = lineY += dy;
      };
      for (let i = 0; i < list.fnArray.length; i++) {
        const fn = list.fnArray[i];
        const args = (list.argsArray[i] ?? []) as unknown[];
        if (fn === OPS.save) stack.push({ ...state });
        else if (fn === OPS.restore) state = stack.pop() ?? state;
        else if (fn === OPS.transform) state.ctm = combine((numbers(args.slice(0, 6)) ?? IDENTITY) as Matrix, state.ctm);
        else if (fn === OPS.paintFormXObjectBegin) {
          stack.push({ ...state });
          const matrix = Array.isArray(args[0]) || ArrayBuffer.isView(args[0]) ? numbers(Array.from(args[0] as ArrayLike<unknown>)) : undefined;
          if (matrix && matrix.length === 6) state.ctm = combine(matrix as Matrix, state.ctm);
        } else if (fn === OPS.paintFormXObjectEnd) state = stack.pop() ?? state;
        else if (fn === OPS.beginMarkedContent || fn === OPS.beginMarkedContentProps) {
          const tag = typeof args[0] === 'string' ? args[0] : (args[0] as { name?: unknown } | null)?.name;
          marks.push(tag === 'Artifact');
          if (tag === 'Artifact') artifact++;
        } else if (fn === OPS.endMarkedContent) {
          if (marks.pop()) artifact--;
        } else if (fn === OPS.setFillRGBColor) state.colour = colourOf(args);
        else if (fn === OPS.setFillGray || fn === OPS.setFillCMYKColor || fn === OPS.setFillColor || fn === OPS.setFillColorN) state.colour = colourOf(args);
        else if (fn === OPS.setStrokeRGBColor || fn === OPS.setStrokeGray || fn === OPS.setStrokeCMYKColor || fn === OPS.setStrokeColor || fn === OPS.setStrokeColorN) state.stroke = colourOf(args);
        else if (fn === OPS.setLineWidth) state.lineWidth = typeof args[0] === 'number' ? args[0] : state.lineWidth;
        else if (fn === OPS.setTextRenderingMode) state.invisible = args[0] === 3;
        else if (fn === OPS.beginText) {
          textMatrix = IDENTITY;
          [x, y, lineX, lineY] = [0, 0, 0, 0];
        } else if (fn === OPS.setFont) {
          state.fontSize = typeof args[1] === 'number' ? args[1] : state.fontSize;
          // (A font's own units: a thousandth of the type size, except in fonts drawn by the file itself.)
          const font = typeof args[0] === 'string' && page.commonObjs.has(args[0]) ? (page.commonObjs.get(args[0]) as { fontMatrix?: number[] } | null) : null;
          state.widthUnit = font?.fontMatrix?.[0] ?? 0.001;
        } else if (fn === OPS.setCharSpacing) state.charSpacing = Number(args[0]) || 0;
        else if (fn === OPS.setWordSpacing) state.wordSpacing = Number(args[0]) || 0;
        else if (fn === OPS.setHScale) state.scale = (Number(args[0]) || 100) / 100;
        else if (fn === OPS.setTextRise) state.rise = Number(args[0]) || 0;
        else if (fn === OPS.setLeading) state.leading = -(Number(args[0]) || 0);
        else if (fn === OPS.moveText) moveText(Number(args[0]) || 0, Number(args[1]) || 0);
        else if (fn === OPS.setLeadingMoveText) {
          state.leading = Number(args[1]) || 0;
          moveText(Number(args[0]) || 0, Number(args[1]) || 0);
        } else if (fn === OPS.setTextMatrix) {
          const matrix = Array.isArray(args[0]) || ArrayBuffer.isView(args[0]) ? numbers(Array.from(args[0] as ArrayLike<unknown>)) : numbers(args.slice(0, 6));
          textMatrix = matrix && matrix.length === 6 ? (matrix as Matrix) : IDENTITY;
          [x, y, lineX, lineY] = [0, 0, 0, 0];
        } else if (fn === OPS.nextLine) moveText(0, state.leading);
        else if (fn === OPS.showText || fn === OPS.showSpacedText || fn === OPS.nextLineShowText || fn === OPS.nextLineSetSpacingShowText) {
          const glyphs = (Array.isArray(args[0]) ? args[0] : []) as unknown[];
          const toPage = combine(textMatrix, state.ctm);
          // The size the type is really drawn at: the stated size, as the two transforms scale it.
          const size = Math.abs(state.fontSize) * Math.hypot(toPage[2], toPage[3]);
          const stretch = Math.hypot(toPage[0], toPage[1]) * state.scale;
          for (const glyph of glyphs) {
            if (typeof glyph === 'number') {
              x -= (glyph * state.fontSize * state.scale) / 1000;
              continue;
            }
            if (typeof glyph !== 'object' || glyph === null) continue;
            const { unicode, width, isSpace } = glyph as { unicode?: unknown; width?: unknown; isSpace?: unknown };
            const advance = (typeof width === 'number' ? width : 0) * state.fontSize * state.widthUnit + state.charSpacing + (isSpace ? state.wordSpacing : 0);
            if (typeof unicode === 'string' && !state.invisible) {
              const [px, py] = [toPage[0] * x + toPage[2] * (y + state.rise) + toPage[4], toPage[1] * x + toPage[3] * (y + state.rise) + toPage[5]];
              // Ligatures and compatibility forms are spelled out, so "ﬁ" compares equal to "fi".
              const spelled = [...unicode.normalize('NFKC')].filter((c) => !isBlank(c));
              const each = (((typeof width === 'number' ? width : 0) * state.fontSize * state.widthUnit) / Math.max(1, spelled.length)) * stretch;
              spelled.forEach((c, n) => chars.push({ c, color: state.colour, page: pageNumber, x: px + n * each, y: py, width: each, size, artifact: artifact > 0 }));
            }
            x += advance * state.scale;
          }
        } else if (fn === OPS.constructPath) {
          // A filled bar or a stroked line that is thin and level: the stroke of an underline or a strikethrough (or a border).
          const paint = args[0];
          const box = ArrayBuffer.isView(args[2]) || Array.isArray(args[2]) ? numbers(Array.from(args[2] as ArrayLike<unknown>)) : undefined;
          const filled = paint === OPS.fill || paint === OPS.eoFill || paint === OPS.fillStroke || paint === OPS.eoFillStroke || paint === OPS.closeFillStroke || paint === OPS.closeEOFillStroke;
          const stroked = paint === OPS.stroke || paint === OPS.closeStroke;
          if (!box || box.length !== 4 || !(filled || stroked)) continue;
          const corner = (cx: number, cy: number): [number, number] => [state.ctm[0] * cx + state.ctm[2] * cy + state.ctm[4], state.ctm[1] * cx + state.ctm[3] * cy + state.ctm[5]];
          const [ax, ay] = corner(box[0]!, box[1]!);
          const [bx, by] = corner(box[2]!, box[3]!);
          const thickness = Math.abs(by - ay) + (stroked ? state.lineWidth * Math.hypot(state.ctm[2], state.ctm[3]) : 0);
          if (Math.abs(bx - ax) >= 1 && thickness <= 3) rules.push({ page: pageNumber, x1: Math.min(ax, bx), x2: Math.max(ax, bx), y: (ay + by) / 2, thickness, color: filled ? state.colour : state.stroke });
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
  return { pages: document.numPages, chars, rules };
}
