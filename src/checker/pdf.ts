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
  /** Set when the file's own structure says the character is part of a list label or a note's number (not of the text). */
  label?: boolean;
  /** Set when the file's own structure says the character is inside a footnote or endnote. */
  note?: boolean;
  /**
   * Which table cell the file's own structure puts the character in: a number that is the same for every character
   * of one cell on one page, and 0 for a character in no cell. Absent when the file does not describe its structure.
   */
  cell?: number;
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

/** A node of a PDF's own description of its structure, as PDF.js hands it over. */
interface StructNode {
  type?: string;
  role?: string;
  id?: string;
  children?: StructNode[];
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
  let cells = 0;
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
      /** The piece of marked content being drawn (its number on this page), and the piece each character of the page belongs to. */
      const pieces: (number | undefined)[] = [];
      const pieceOf: (number | undefined)[] = [];
      const pageStart = chars.length;
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
          pieces.push(fn === OPS.beginMarkedContentProps && typeof args[1] === 'number' ? args[1] : pieces[pieces.length - 1]);
        } else if (fn === OPS.endMarkedContent) {
          if (marks.pop()) artifact--;
          pieces.pop();
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
              spelled.forEach((c, n) => {
                chars.push({ c, color: state.colour, page: pageNumber, x: px + n * each, y: py, width: each, size, artifact: artifact > 0 });
                pieceOf.push(pieces[pieces.length - 1]);
              });
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
      // The order a reader takes the page in. A PDF draws its text in whatever order suits the program that made
      // it (a table carried over from the page before may be drawn last although it stands at the top). A file
      // that describes its own structure says in which order its pieces are read; where it does, the page's
      // characters are put in that order: the page furniture above the text, the text, the furniture below it.
      // This only decides the order in which the characters are compared. Every character is still compared.
      let tree: StructNode | null = null;
      try {
        tree = (await page.getStructTree()) as StructNode | null;
      } catch {
        tree = null;
      }
      const order: { piece: number; label: boolean; note: boolean; cell: number }[] = [];
      const walk = (node: StructNode, label: boolean, note: boolean, cell: number): void => {
        if (node.type === 'content') {
          const piece = /_mc(\d+)$/.exec(node.id ?? '');
          if (piece) order.push({ piece: Number(piece[1]), label, note, cell });
          return;
        }
        const inside = node.role === 'TD' || node.role === 'TH' ? ++cells : cell;
        for (const inner of node.children ?? []) walk(inner, label || node.role === 'Lbl', note || node.role === 'Note', inside);
      };
      if (tree) walk(tree, false, false, 0);
      if (order.length > 0) {
        const mine = chars.splice(pageStart);
        const byPiece = new Map<number, PdfChar[]>();
        const furniture: PdfChar[] = [];
        const loose: PdfChar[] = [];
        /** Furniture drawn in the middle of the text (the dots of a tab leader): kept behind the character it was drawn after. */
        const drawnAfter = new Map<PdfChar, PdfChar[]>();
        let previous: PdfChar | undefined;
        mine.forEach((char, n) => {
          const piece = pieceOf[n];
          if (char.artifact) {
            furniture.push(char);
            if (previous) drawnAfter.set(previous, [...(drawnAfter.get(previous) ?? []), char]);
          } else if (piece === undefined) loose.push(char);
          else byPiece.set(piece, [...(byPiece.get(piece) ?? []), char]);
          if (!char.artifact) previous = char;
        });
        const text: PdfChar[] = [];
        for (const { piece, label, note, cell } of order) {
          for (const char of byPiece.get(piece) ?? []) {
            if (label) char.label = true;
            if (note) char.note = true;
            char.cell = cell;
            text.push(char);
          }
          byPiece.delete(piece);
        }
        // (Content the structure does not mention is kept, after the rest: it will have to be accounted for like everything else.)
        for (const rest of byPiece.values()) text.push(...rest);
        text.push(...loose);
        const heights = text.map((char) => char.y);
        const [top, bottom] = heights.length > 0 ? [Math.max(...heights), Math.min(...heights)] : [page.view[3]! / 2, page.view[3]! / 2];
        const within = (char: PdfChar): boolean => char.y <= top && char.y >= bottom;
        const placed = new Set<PdfChar>();
        const body: PdfChar[] = [];
        for (const char of text) {
          body.push(char);
          for (const dots of drawnAfter.get(char) ?? []) {
            if (!within(dots)) continue;
            body.push(dots);
            placed.add(dots);
          }
        }
        // (Furniture within the text that was drawn before any of the page's text, a table's repeated heading row for one, comes first.)
        chars.push(...furniture.filter((char) => char.y > top), ...furniture.filter((char) => within(char) && !placed.has(char)), ...body, ...furniture.filter((char) => char.y < bottom));
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
