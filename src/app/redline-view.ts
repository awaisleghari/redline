/**
 * The redline as the page shows it.
 *
 * The compare engine turns the static redline (a Word file the Checker has
 * passed) into a web page. This module rebuilds that page inside ours: it keeps
 * the document's words and structure, drops everything else (scripts, links,
 * outside resources, the document's own fonts and colours), and marks each
 * change with the design's classes so it can be styled and stepped through.
 *
 * App code: neither Producer nor Checker. It never alters a word, and it shows
 * only a file it was handed.
 */

export type MarkKind = 'ins' | 'del' | 'mvf' | 'mvt' | 'fmt';

/**
 * How the Producer formats each kind of change (src/producer/static-redline.ts).
 * Copied here, not imported, so the page does not have to load the Producer;
 * a unit test holds the two to each other.
 */
export const REDLINE_LOOK = {
  inserted: ['0000FF', 'AA00AA'],
  deleted: ['FF0000', 'C55A11'],
  moved: '008000',
  formatFill: 'FFF2CC',
} as const;

const ALLOWED_TAGS = new Set(['p', 'span', 'div', 'table', 'thead', 'tbody', 'tfoot', 'tr', 'td', 'th', 'colgroup', 'col', 'caption', 'ul', 'ol', 'li', 'br', 'hr', 'sup', 'sub', 'a', 'b', 'strong', 'i', 'em', 'u', 's', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'img', 'header', 'footer', 'section', 'ins', 'del', 'blockquote', 'pre']);
/** Removed with everything inside them. Any other unknown element is unwrapped: its words are kept. */
const DROPPED_TAGS = new Set(['script', 'style', 'link', 'meta', 'title', 'head', 'iframe', 'frame', 'object', 'embed', 'form', 'input', 'button', 'select', 'textarea', 'audio', 'video', 'source', 'template', 'noscript', 'base']);
const KEPT_ATTRIBUTES = new Set(['dir', 'colspan', 'rowspan', 'span', 'value', 'start', 'type', 'alt', 'width', 'height']);
/** Layout the document needs. Its own fonts, sizes, colours and vertical spacing give way to the design's. */
const ALLOWED_PROPERTIES = new Set([
  'font-weight', 'font-style', 'font-variant-caps', 'text-transform', 'text-align', 'text-indent', 'margin-left', 'margin-right',
  'padding-left', 'padding-right', 'padding-top', 'padding-bottom', 'display', 'width', 'min-width', 'max-width', 'vertical-align', 'white-space',
  'float', 'clear', 'border-collapse', 'border-spacing', 'list-style-type', 'direction', 'unicode-bidi', 'table-layout', 'overflow-wrap', 'word-break',
  'border-top-width', 'border-right-width', 'border-bottom-width', 'border-left-width', 'border-top-style', 'border-right-style', 'border-bottom-style', 'border-left-style',
]);

interface Look {
  color?: string;
  line?: string;
  style?: string;
  fill?: string;
  bold?: boolean;
}

function hex(color: string): string | undefined {
  const rgb = /^rgba?\(\s*(\d+)[,\s]+(\d+)[,\s]+(\d+)/.exec(color);
  if (rgb) return [rgb[1], rgb[2], rgb[3]].map((n) => Number(n).toString(16).padStart(2, '0')).join('').toUpperCase();
  const short = /^#([0-9a-f]{3})$/i.exec(color);
  if (short) return short[1]!.split('').map((c) => c + c).join('').toUpperCase();
  return /^#([0-9a-f]{6})$/i.exec(color)?.[1]?.toUpperCase();
}

function lookOf(style: CSSStyleDeclaration): Look {
  const weight = style.getPropertyValue('font-weight');
  return {
    color: hex(style.getPropertyValue('color')),
    line: style.getPropertyValue('text-decoration-line') || undefined,
    style: style.getPropertyValue('text-decoration-style') || undefined,
    fill: hex(style.getPropertyValue('background-color')),
    bold: weight === 'bold' || Number(weight) >= 600 ? true : weight ? false : undefined,
  };
}

/** The declarations of a rule (or of an inline style) that the page keeps, as CSS text. */
function keptDeclarations(style: CSSStyleDeclaration, isChange: boolean): string {
  const out: string[] = [];
  for (let i = 0; i < style.length; i++) {
    const name = style.item(i);
    const value = style.getPropertyValue(name).trim();
    if (!value || /url\(|expression\(|@import|[<>{}]/i.test(value)) continue;
    if (name === 'font-size') {
      // Only sizes relative to the surrounding text (superscripts and the like).
      if (/^[\d.]+(em|%)$/.test(value)) out.push(`font-size: ${value}`);
    } else if (name === 'font-weight') {
      const bold = value === 'bold' || Number(value) >= 600;
      out.push(bold ? 'font-weight: 500' : 'font-weight: 400');
      if (bold && !isChange) out.push('color: var(--rl-strong)');
    } else if (name === 'text-decoration-line') {
      // A change's underline or strike is the design's; the document's own plain underline is kept.
      if (!isChange && /^(underline|line-through|underline line-through)$/.test(value)) out.push(`text-decoration-line: ${value}`);
    } else if (name === 'position') {
      if (value === 'relative' || value === 'absolute') out.push(`position: ${value}`);
    } else if (ALLOWED_PROPERTIES.has(name)) out.push(`${name}: ${value}`);
  }
  return out.join('; ');
}

/**
 * Which change a piece of text is, from how it is formatted. The text may carry a line of its own as well
 * (a deleted underlined heading is underlined and struck), so the strike is looked for first.
 */
function kindOfLook(look: Look, inserted: string, deleted: string): MarkKind | undefined {
  const double = look.style === 'double';
  const lines = (look.line ?? '').split(/\s+/);
  const struck = lines.includes('line-through');
  const underlined = lines.includes('underline');
  if (struck && double && look.color === REDLINE_LOOK.moved) return 'mvf';
  if (struck && !double && look.color === deleted) return 'del';
  if (underlined && double && look.color === inserted) return 'ins';
  if (underlined && double && look.color === REDLINE_LOOK.moved) return 'mvt';
  if (look.fill === REDLINE_LOOK.formatFill) return 'fmt';
  return undefined;
}

function parse(html: string): Document {
  // The engine writes XHTML. Read it strictly; if that fails, read it as ordinary HTML.
  const strict = new DOMParser().parseFromString(html, 'application/xhtml+xml');
  if (!strict.querySelector('parsererror') && strict.body) return strict;
  return new DOMParser().parseFromString(html, 'text/html');
}

/** Only the engine's own class names are kept, so nothing in a document can borrow one of this page's. */
const classesOf = (element: Element): string[] => (element.getAttribute('class') ?? '').split(/\s+/).filter((token) => /^(docx-[A-Za-z0-9_-]+|footnote-ref|fn-backref|footnotes|endnotes|endnote-ref|document-header|document-footer)$/.test(token));
/** The engine's attributes that its own layout rules rely on. */
const KEPT_DATA = new Set(['data-list-marker', 'data-docx-tab', 'data-docx-tab-width', 'data-note-run', 'data-footnote-separator']);
const isCover = (classes: string[]): boolean => classes.some((token) => token.startsWith('docx-RedlineCover'));
const isHeadingClass = (classes: string[]): boolean => classes.some((token) => /^docx-(heading|title|subtitle)/i.test(token));

export interface RedlineModel {
  /** The sanitised document: one detached container, cloned for each layout. */
  master: HTMLElement;
  /** The document's own layout rules, scoped to the viewer. */
  css: string;
  /** How many changes can be stepped through. */
  stops: number;
}

/** Rebuild the engine's page as the viewer's document. */
export function buildRedline(html: string): RedlineModel {
  const source = parse(html);

  // 1. The engine's style rules: what each class looks like, and the layout to keep.
  const looks = new Map<string, Look>();
  const rules: { selector: string; style: CSSStyleDeclaration; className?: string }[] = [];
  const sheet = new CSSStyleSheet();
  sheet.replaceSync(Array.from(source.querySelectorAll('style')).map((style) => style.textContent ?? '').join('\n'));
  for (const rule of Array.from(sheet.cssRules)) {
    if (!(rule instanceof CSSStyleRule)) continue;
    const single = /^[a-z0-9]*\.([A-Za-z0-9_-]+)$/.exec(rule.selectorText.trim());
    if (single) looks.set(single[1]!, lookOf(rule.style));
    rules.push({ selector: rule.selectorText, style: rule.style, className: single?.[1] });
  }

  // 2. Which colours this redline uses: the Producer switches to its second colour when the
  //    document's own text already looks like an insertion or a deletion. The cover page's legend shows which.
  const coverColors = new Set<string>();
  for (const span of Array.from(source.querySelectorAll('p span'))) {
    const paragraph = span.closest('p');
    if (!paragraph || !isCover(classesOf(paragraph))) continue;
    for (const token of classesOf(span)) {
      const color = looks.get(token)?.color;
      if (color) coverColors.add(color);
    }
  }
  const inserted = coverColors.has(REDLINE_LOOK.inserted[1]) && !coverColors.has(REDLINE_LOOK.inserted[0]) ? REDLINE_LOOK.inserted[1] : REDLINE_LOOK.inserted[0];
  const deleted = coverColors.has(REDLINE_LOOK.deleted[1]) && !coverColors.has(REDLINE_LOOK.deleted[0]) ? REDLINE_LOOK.deleted[1] : REDLINE_LOOK.deleted[0];
  const kindOfClass = new Map<string, MarkKind>();
  for (const [name, look] of looks) {
    const kind = kindOfLook(look, inserted, deleted);
    if (kind) kindOfClass.set(name, kind);
  }

  const css = rules
    .map(({ selector, style, className }) => {
      const body = keptDeclarations(style, !!className && kindOfClass.has(className) && kindOfClass.get(className) !== 'fmt');
      if (!body) return '';
      // Every rule is confined to what is inside a sheet: no rule may reach the sheet's siblings or the page around it.
      const scoped = selector
        .split(',')
        .map((part) => part.trim().replace(/^(html|body)(?![\w-])\s*/, ''))
        .filter((one) => one && !/^[>~+]/.test(one) && !/[~+]/.test(one) && !/\b(html|body)\b|:root|:host|\.rl-|\.is-|\.sr-only|#/.test(one))
        .map((one) => `:where(.rl-sheet) ${one}`);
      if (!scoped.length) return '';
      return `${scoped.join(', ')} { ${body} }`;
    })
    .filter(Boolean)
    .join('\n');

  // 3. The document itself, element by element.
  const master = document.createElement('div');
  const scratch = document.createElement('span');
  const copy = (node: Node, into: Node, inCover: boolean, inherited: string | undefined): void => {
    if (node.nodeType === Node.TEXT_NODE) {
      if (node.nodeValue) into.appendChild(document.createTextNode(node.nodeValue));
      return;
    }
    if (node.nodeType !== Node.ELEMENT_NODE) return;
    const from = node as Element;
    const tag = from.localName.toLowerCase();
    if (DROPPED_TAGS.has(tag)) return;
    const classes = classesOf(from);
    // The "back to the text" arrow of a footnote leads nowhere here.
    if (tag === 'a' && classes.includes('fn-backref')) return;
    if (!ALLOWED_TAGS.has(tag)) {
      for (const child of Array.from(from.childNodes)) copy(child, into, inCover, inherited);
      return;
    }
    const part = tag === 'header' ? 'rl-part rl-part-header' : tag === 'footer' ? 'rl-part rl-part-footer' : tag === 'section' ? 'rl-notes' : '';
    const to = document.createElement(part ? 'div' : tag === 'ins' || tag === 'del' ? 'span' : tag);
    const mine: string[] = part ? part.split(' ') : [];
    const cover = inCover || isCover(classes);
    if (cover && tag === 'p') {
      mine.push('rl-cover');
      if (classes.some((token) => token.startsWith('docx-RedlineCoverTitle'))) mine.push('rl-heading');
      if (classes.some((token) => token.startsWith('docx-RedlineCoverHeading'))) mine.push('rl-cover-heading');
    } else if (/^h[1-6]$/.test(tag) || (tag === 'p' && isHeadingClass(classes))) mine.push('rl-heading');

    // Inline style: the same filter as the style rules. A change may also be marked inline.
    let look: Look = {};
    const inline = from.getAttribute('style');
    if (inline) {
      scratch.setAttribute('style', inline);
      look = lookOf(scratch.style);
      const kept = keptDeclarations(scratch.style, false);
      if (kept) to.setAttribute('style', kept);
      scratch.removeAttribute('style');
    }
    // The element's whole look: its classes' rules, then its own style; its colour may come from the span around it
    // (the engine wraps a moved-away underlined phrase as a green underlined span around a double-struck one).
    for (const token of classes) {
      const rule = looks.get(token);
      if (rule) look = { color: look.color ?? rule.color, line: look.line ?? rule.line, style: look.style ?? rule.style, fill: look.fill ?? rule.fill };
    }
    const color = look.color ?? inherited;
    let kind: MarkKind | undefined;
    for (const token of classes) kind ??= kindOfClass.get(token);
    kind ??= look.line || look.fill ? kindOfLook({ ...look, color }, inserted, deleted) : undefined;
    if (kind) {
      mine.push(`rl-${kind}`);
      if (cover) to.setAttribute('data-legend', '');
    }

    to.className = [...classes, ...mine].join(' ');
    if (!to.className) to.removeAttribute('class');
    for (const attribute of Array.from(from.attributes)) {
      const name = attribute.name.toLowerCase();
      if (KEPT_ATTRIBUTES.has(name) || KEPT_DATA.has(name)) to.setAttribute(name, attribute.value);
    }
    if (tag === 'img') {
      const src = from.getAttribute('src') ?? '';
      // Pictures stored inside the document only; nothing is fetched from anywhere.
      if (!/^data:image\/(png|jpe?g|gif|webp|bmp);base64,/i.test(src)) return;
      to.setAttribute('src', src);
    }
    into.appendChild(to);
    for (const child of Array.from(from.childNodes)) copy(child, to, cover, color);
  };
  for (const child of Array.from(source.body.childNodes)) copy(child, master, false, undefined);

  // The cover page is set apart from the document by a rule under its last line.
  const coverLines = master.querySelectorAll('p.rl-cover');
  const lastCover = coverLines[coverLines.length - 1]?.parentElement;
  if (lastCover && lastCover !== master) lastCover.classList.add('rl-cover-end');

  return { master, css, stops: numberChanges(master) };
}

const BLOCKS = 'p, li, td, th, h1, h2, h3, h4, h5, h6, pre, blockquote';
const MARKS = '.rl-ins, .rl-del, .rl-mvf, .rl-mvt, .rl-fmt';
const kindOf = (element: Element): MarkKind => (['ins', 'del', 'mvf', 'mvt', 'fmt'] as const).find((kind) => element.classList.contains(`rl-${kind}`))!;

/**
 * Number the changes, in reading order, the way the Producer counts them: a
 * change is an unbroken stretch of text of one kind. A stretch runs on into the
 * next paragraph only when the whole of the earlier paragraph belongs to it. A
 * move is one change with two ends. Each marked element gets `data-chg`.
 */
function numberChanges(master: HTMLElement): number {
  interface Group { kind: MarkKind; elements: Element[]; text: string }
  const groups: Group[] = [];
  let current: Group | undefined;
  let lastBlock: Element | null = null;
  let lastPart: Element | null = null;
  let blockAllOneKind = false;

  const walker = document.createTreeWalker(master, NodeFilter.SHOW_TEXT);
  for (let node = walker.nextNode(); node; node = walker.nextNode()) {
    const parent = node.parentElement!;
    if (!node.nodeValue || parent.closest('[data-list-marker], [data-legend]')) continue;
    const mark = parent.closest(MARKS);
    const block = parent.closest(BLOCKS);
    const part = parent.closest('.rl-part, .rl-notes');
    if (parent.closest('.rl-cover')) continue;
    const kind = mark ? kindOf(mark) : undefined;
    const newBlock = block !== lastBlock;
    if (newBlock) {
      // Crossing a paragraph mark: the stretch survives only if the paragraph just left was all of this kind.
      if (!(current && blockAllOneKind && kind === current.kind && kind !== 'fmt' && part === lastPart)) current = undefined;
      blockAllOneKind = true;
      lastBlock = block;
      lastPart = part;
    }
    if (!kind) {
      current = undefined;
      blockAllOneKind = false;
      continue;
    }
    if (!current || current.kind !== kind) {
      if (current) blockAllOneKind = false;
      current = { kind, elements: [], text: '' };
      groups.push(current);
    }
    if (current.elements[current.elements.length - 1] !== mark) current.elements.push(mark!);
    // The old clause number the Producer types in front of a removed or moved-away clause is part of the mark, not of the clause's words.
    if (!parent.closest('[class*="docx-RedlineOldNumber"]')) current.text += node.nodeValue;
  }

  // A move is one change: pair each "moved from" with the "moved to" that holds the same words.
  const squash = (text: string): string => text.replace(/\s+/g, ' ').trim();
  const destinations = groups.filter((group) => group.kind === 'mvt');
  const partner = new Map<Group, Group>();
  const taken = new Set<Group>();
  for (const origin of groups) {
    if (origin.kind !== 'mvf') continue;
    const match = destinations.find((candidate) => !taken.has(candidate) && squash(candidate.text) === squash(origin.text));
    if (match) {
      taken.add(match);
      partner.set(origin, match);
      partner.set(match, origin);
    }
  }
  let count = 0;
  const numbered = new Map<Group, number>();
  for (const group of groups) {
    const other = partner.get(group);
    const number = (other && numbered.get(other)) || ++count;
    numbered.set(group, number);
    for (const element of group.elements) element.setAttribute('data-chg', String(number));
  }
  return count;
}

export interface PdfPages {
  /** Width divided by height of the PDF's first page. */
  ratio: number;
  /** The text of each page, in order. */
  pages: string[];
}

/** Characters as compared between the page and the PDF: no white space, ligatures and the like unfolded. */
function squeeze(char: string): string {
  return /\s| |­|​/.test(char) ? '' : char.normalize('NFKC').replace(/\s/g, '');
}

/**
 * Show the document in `container`: as one continuous sheet, or, given the text
 * of each page of its PDF, split into the same pages as the PDF.
 */
export function layOut(container: HTMLElement, model: RedlineModel, pdf?: PdfPages): void {
  const sheet = document.createElement('article');
  sheet.className = 'rl-sheet';
  for (const child of Array.from(model.master.childNodes)) sheet.appendChild(child.cloneNode(true));
  container.replaceChildren(sheet);
  container.classList.toggle('is-paged', !!pdf);
  if (!pdf) return;
  container.style.setProperty('--rl-page-ratio', String(pdf.ratio));

  // The document's running text (headers, footers and notes sit outside it: a PDF repeats or moves them).
  const nodes: Text[] = [];
  const nodeOf: number[] = [];
  const offsetOf: number[] = [];
  let stream = '';
  const walker = document.createTreeWalker(sheet, NodeFilter.SHOW_TEXT);
  for (let node = walker.nextNode() as Text | null; node; node = walker.nextNode() as Text | null) {
    if (node.parentElement!.closest('.rl-part, .rl-notes')) continue;
    const text = node.data;
    nodes.push(node);
    for (let i = 0; i < text.length; i++) {
      const kept = squeeze(text[i]!);
      for (let k = 0; k < kept.length; k++) {
        nodeOf.push(nodes.length - 1);
        offsetOf.push(i);
      }
      stream += kept;
    }
  }

  // Where each page after the first begins: the first run of its text found in the document, near where it is due.
  const PROBE = 24;
  const CONFIRM = 72;
  const starts: number[] = [];
  let previous = 0;
  let previousLength = 0;
  pdf.pages.forEach((raw, index) => {
    const page = Array.from(raw).map(squeeze).join('');
    if (index > 0) {
      const due = Math.min(stream.length, previous + previousLength);
      const window = stream.slice(previous, Math.min(stream.length, due + 400));
      let found = -1;
      for (let o = 0; o + PROBE <= page.length && o < 4000 && found < 0; o++) {
        const probe = page.slice(o, o + PROBE);
        // Of the places the probe occurs and the page's next words follow, the one nearest to where the page is due.
        const need = Math.min(CONFIRM, page.length - o);
        let best = -1;
        for (let at = window.indexOf(probe); at >= 0; at = window.indexOf(probe, at + 1)) {
          if (window.slice(at, at + need) !== page.slice(o, o + need) && at + need <= window.length) continue;
          if (best < 0 || Math.abs(previous + at - due) < Math.abs(previous + best - due)) best = at;
        }
        if (best >= 0) found = previous + best;
      }
      previous = found >= 0 ? found : due;
      starts.push(previous);
    }
    previousLength = page.length;
  });

  // Cut the sheet at each start, last first, so earlier positions stay valid.
  const sheets: HTMLElement[] = [sheet];
  for (let i = starts.length - 1; i >= 0; i--) {
    const at = starts[i]!;
    const next = document.createElement('article');
    next.className = 'rl-sheet';
    if (at < stream.length && sheet.lastChild) {
      const range = document.createRange();
      const node = nodes[nodeOf[at]!]!;
      range.setStart(node, offsetOf[at]!);
      const block = node.parentElement!.closest(BLOCKS);
      if (block && sheet.contains(block)) {
        // Nothing but the clause number before the cut: the whole paragraph belongs to the new page.
        const before = document.createRange();
        before.setStartBefore(block);
        before.setEnd(node, offsetOf[at]!);
        const lead = before.toString().trim();
        if (!lead || lead === (block.querySelector('[data-list-marker]')?.textContent ?? '').trim()) range.setStartBefore(block);
      }
      // The table rows and tables the cut passes through, outermost first.
      const crossed = (name: string): Element[] => {
        const found: Element[] = [];
        const from = range.startContainer;
        for (let up = from.nodeType === Node.ELEMENT_NODE ? (from as Element) : from.parentElement; up && up !== sheet; up = up.parentElement) if (up.localName === name) found.unshift(up);
        return found;
      };
      // Nothing of the outermost row before the cut: the page starts with that row, whole.
      const outer = crossed('tr')[0];
      if (outer) {
        const before = document.createRange();
        before.setStartBefore(outer);
        before.setEnd(range.startContainer, range.startOffset);
        if (!before.toString().trim()) range.setStartBefore(outer);
      }
      const rows = crossed('tr');
      const tables = crossed('table');
      const cutIn = range.startContainer;
      range.setEndAfter(sheet.lastChild);
      next.appendChild(range.extractContents());

      // A row that runs over the page is now in two halves, as in Word. Give both halves all their columns:
      // the second half gets empty cells for the columns before the cut, the first for those after it.
      const halves = (name: string): Element[] => {
        const found: Element[] = [];
        for (let down = next.firstElementChild; down; down = down.firstElementChild) if (down.localName === name) found.push(down);
        return found;
      };
      const rowHalves = halves('tr');
      rows.forEach((first, index) => {
        const second = rowHalves[index];
        if (!second) return;
        const later = Array.from(second.children).slice(1);
        for (const cell of Array.from(first.children).slice(0, -1).reverse()) second.prepend(cell.cloneNode(false));
        for (const cell of later) first.append(cell.cloneNode(false));
      });
      const tableHalves = halves('table');
      tables.forEach((first, index) => {
        const widths = Array.from(first.children).find((child) => child.localName === 'colgroup');
        const second = tableHalves[index];
        if (widths && second && !Array.from(second.children).some((child) => child.localName === 'colgroup')) second.prepend(widths.cloneNode(true));
      });

      // The cut can leave an empty shell of the paragraph it fell at the start of (a table cell is kept: it holds its column).
      for (let shell: Node | null = cutIn; shell && shell !== sheet; ) {
        const parent: Node | null = shell.parentNode;
        const element = shell.nodeType === Node.ELEMENT_NODE ? (shell as Element) : null;
        if (element && /^(td|th|tr)$/.test(element.localName)) break;
        const empty = shell.nodeType === Node.TEXT_NODE ? !(shell as Text).data : !!element && !element.textContent && !element.querySelector('img, br, hr');
        if (!empty) break;
        parent?.removeChild(shell);
        shell = parent;
      }
    }
    sheet.after(next);
    sheets.splice(1, 0, next);
  }
  sheets.forEach((page, index) => {
    const label = document.createElement('span');
    label.className = 'rl-pagelabel';
    label.textContent = `page ${index + 1} of ${sheets.length}`;
    page.appendChild(label);
  });
}
