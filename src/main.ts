/**
 * The workspace: two documents in, a checked redline out.
 *
 * Nothing here decides whether an output is right. The flow (src/app/flow.ts)
 * hands back a file only when the independent Checker has passed it, and this
 * page offers for download, and shows in the viewer, only files it was handed.
 */
import { makePrintCopy } from './producer/print-copy';
import { ProducerError, runComparison, runPdf, type FlowResult } from './app/flow';
import { displayName, formatComparedAt } from './app/format';
import { bandGrid, type GridHandle } from './app/grid';
import { readPdfPages } from './app/pdf-pages';
import { buildRedline, layOut, type PdfPages, type RedlineModel } from './app/redline-view';
import { containsTrackedChanges, refusalFor, TRACKED_NOTICE } from './app/validate';
import { OUTPUT_NAME, reportAsText, type CheckReport } from './checker';
import { CheckerClient } from './checker/checker-client';
import { CompareClient } from './producer/compare-client';

function element<T extends HTMLElement>(id: string): T {
  const found = document.getElementById(id);
  if (!found) throw new Error(`missing element #${id}`);
  return found as T;
}
/** Show or hide anything, including the drawn icons (which are not HTML elements). */
const setHidden = (id: string, hidden: boolean): void => void document.getElementById(id)!.toggleAttribute('hidden', hidden);

type Role = 'original' | 'modified';
type Format = 'pdf' | 'static' | 'tracked';
type Stage = 'idle' | 'loading' | 'done';
const ROLES = ['original', 'modified'] as const;
const FORMATS = ['pdf', 'static', 'tracked'] as const;
const DOCX = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
const PDF_ENDPOINT = '/api/pdf';
/** The creating screen stays up at least this long, so it does not just flash. */
const MIN_CREATING_MS = 600;

const inputs = { original: element<HTMLInputElement>('original'), modified: element<HTMLInputElement>('modified') };
const chosen: { original?: File; modified?: File } = {};
const createButton = element<HTMLButtonElement>('create');
const createAgain = element<HTMLButtonElement>('create-again');
const status = element<HTMLParagraphElement>('status');
const flowError = element<HTMLParagraphElement>('flow-error');
const viewer = element<HTMLDivElement>('viewer');
const doc = element<HTMLDivElement>('doc');
const viewerMessage = element<HTMLParagraphElement>('viewer-message');
const mainPanel = document.querySelector<HTMLElement>('.rl-main')!;
const downloadButton = element<HTMLButtonElement>('download');
const moreButton = element<HTMLButtonElement>('download-more');
const menu = element<HTMLDivElement>('download-menu');
const pdfStatus = element<HTMLParagraphElement>('pdf-status');
const checksToggle = element<HTMLButtonElement>('checks-toggle');
const counter = element<HTMLSpanElement>('change-counter');

const producer = new CompareClient();
const checker = new CheckerClient();
/** True while a comparison or a PDF is being made: the files cannot be changed and no second run can start. */
let running = false;
let stage: Stage = 'idle';

/** What the current comparison released. A file is here only if the Checker passed it. */
interface Released {
  tracked?: Uint8Array;
  staticRedline?: Uint8Array;
  pdf?: Uint8Array;
  /** The PDF was asked for and was withheld, or could not be made. */
  pdfFailed: boolean;
}
let released: Released | undefined;
let model: RedlineModel | undefined;
let pdfPages: PdfPages | undefined;

// ---------- small things remembered between visits ----------

function recall<T extends string>(key: string, allowed: readonly T[], fallback: T): T {
  try {
    const stored = localStorage.getItem(key);
    return allowed.includes(stored as T) ? (stored as T) : fallback;
  } catch {
    return fallback;
  }
}
function remember(key: string, value: string): void {
  try {
    localStorage.setItem(key, value);
  } catch {
    // Private browsing or storage switched off: the choice simply lasts for this visit.
  }
}
let format = recall('redline.format', FORMATS, 'pdf');
let view = recall('redline.view', ['continuous', 'pages'] as const, 'continuous');
let paper = recall('redline.page', ['dark', 'light'] as const, 'dark');

// ---------- status, stages and the bitmap band ----------

function say(state: string, text: string): void {
  status.dataset.state = state;
  status.textContent = text;
}

let band: GridHandle | undefined;
function restartBand(): void {
  band?.stop();
  band = bandGrid(element<HTMLCanvasElement>('band'), stage === 'loading');
}
window.addEventListener('pagehide', () => {
  band?.stop();
  band = undefined;
});
window.addEventListener('pageshow', () => {
  if (!band) restartBand();
});

const baseName = (file: File): string => displayName(file.name).replace(/\.docx?$/i, '');
const sizeOf = (bytes: number): string => (bytes >= 1024 * 1024 ? `${(bytes / 1024 / 1024).toFixed(1)} MB` : `${Math.max(1, Math.round(bytes / 1024))} KB`);

function render(): void {
  const both = !!(chosen.original && chosen.modified);
  const done = stage === 'done';
  const loading = stage === 'loading';
  for (const role of ROLES) {
    const file = chosen[role];
    setHidden(`pick-${role}`, !!file);
    setHidden(`file-${role}`, !file);
    element(`name-${role}`).textContent = file ? baseName(file) : '';
    element(`name-${role}`).title = file ? displayName(file.name) : '';
    element(`size-${role}`).textContent = file ? sizeOf(file.size) : '';
    element<HTMLButtonElement>(`pick-${role}`).disabled = running;
    element<HTMLButtonElement>(`remove-${role}`).disabled = running;
    inputs[role].disabled = running;
  }
  createButton.hidden = done;
  createAgain.hidden = !done;
  createButton.disabled = !both || running || validating > 0;
  createAgain.disabled = !both || running || validating > 0;
  setHidden('create-spinner', !loading);
  setHidden('create-hint', loading);
  element('create-label').textContent = loading ? 'creating redline' : 'create redline';
  element<HTMLButtonElement>('new-redline').disabled = running;

  const shown = done && !!model;
  element('viewer-title').textContent = stage === 'idle' || !chosen.modified ? 'new redline' : baseName(chosen.modified);
  setHidden('viewer-subtitle', !done);
  setHidden('viewer-tools', !shown);
  viewer.hidden = !shown;
  setHidden('idle', stage !== 'idle');
  setHidden('loading', !loading);
  setHidden('withheld', !(done && !model));
  setHidden('side-hint', done);
  for (const id of ['section-export', 'section-checks']) setHidden(id, !done);
  if (!done) setHidden('section-changes', true);
  mainPanel.classList.toggle('is-loading', loading);
}

function setStage(next: Stage): void {
  const bandChanges = (stage === 'loading') !== (next === 'loading');
  stage = next;
  render();
  if (bandChanges) restartBand();
}

// ---------- the two documents ----------

/** Results on screen belong to the files they were made from: a new choice clears them. */
function clearResults(): void {
  released = undefined;
  model = undefined;
  pdfPages = undefined;
  reports = [];
  reportContext = undefined;
  doc.replaceChildren();
  const link = element<HTMLAnchorElement>('download-report');
  if (link.href) URL.revokeObjectURL(link.href);
  link.removeAttribute('href');
  link.hidden = true;
  for (const kind of ['static', 'tracked', 'pdf']) element(`report-${kind}`).replaceChildren();
  setHidden('group-pdf', true);
  setHidden('tracked-notice', true);
  pdfStatus.textContent = '';
  pdfStatus.dataset.state = '';
  setHidden('withheld-note', true);
  flowError.hidden = true;
  flowError.textContent = '';
  closeMenu(false);
  setChecksOpen(false);
}

const choiceId = { original: 0, modified: 0 };
/** Files being read to see whether they can be compared. Nothing is created until each has its answer. */
let validating = 0;
async function choose(role: Role, file: File | undefined): Promise<void> {
  if (running) return;
  const id = ++choiceId[role];
  const problem = element(`error-${role}`);
  problem.hidden = true;
  problem.textContent = '';
  // A file that cannot be compared is refused here, in plain English, and not kept.
  let refusal: string | undefined;
  validating++;
  render();
  try {
    refusal = file ? await refusalFor({ name: file.name, size: file.size, bytes: async () => new Uint8Array(await file.arrayBuffer()) }, role) : undefined;
  } catch {
    refusal = `${role === 'original' ? 'The original' : 'The modified document'} could not be read.`;
  } finally {
    validating--;
    render();
  }
  if (id !== choiceId[role] || running) return;
  chosen[role] = refusal ? undefined : file;
  clearResults();
  setStage('idle');
  if (refusal) {
    problem.textContent = refusal;
    problem.hidden = false;
  }
  // The status line speaks for both slots: a refusal stands until its own slot is put right.
  const standing = ROLES.map((one) => element(`error-${one}`)).find((message) => !message.hidden);
  say(standing ? 'refused' : '', standing?.textContent ?? '');
}

for (const role of ROLES) {
  const input = inputs[role];
  const zone = element(`slot-${role}`);
  element(`pick-${role}`).addEventListener('click', () => input.click());
  element(`remove-${role}`).addEventListener('click', () => {
    void choose(role, undefined).then(() => element(`pick-${role}`).focus());
  });
  input.addEventListener('change', () => {
    const file = input.files?.[0];
    input.value = '';
    if (file) void choose(role, file);
  });
  zone.addEventListener('dragover', (event) => {
    event.preventDefault();
    zone.classList.add('over');
  });
  zone.addEventListener('dragleave', () => zone.classList.remove('over'));
  zone.addEventListener('drop', (event) => {
    event.preventDefault();
    zone.classList.remove('over');
    const file = event.dataTransfer?.files[0];
    if (file) void choose(role, file);
  });
}
// A file dropped beside a slot must not replace the page with the file.
for (const type of ['dragover', 'drop']) window.addEventListener(type, (event) => event.preventDefault());

element('new-redline').addEventListener('click', () => {
  if (running) return;
  for (const role of ROLES) {
    choiceId[role]++;
    chosen[role] = undefined;
    element(`error-${role}`).hidden = true;
  }
  clearResults();
  setStage('idle');
  say('', '');
});

// ---------- downloads ----------

function save(data: Uint8Array | string, type: string, name: string): void {
  const url = URL.createObjectURL(new Blob([typeof data === 'string' ? data : data.slice().buffer], { type }));
  const link = document.createElement('a');
  link.href = url;
  link.download = name;
  document.body.append(link);
  link.click();
  link.remove();
  setTimeout(() => URL.revokeObjectURL(url), 60_000);
}

const DOWNLOAD_LABEL: Record<Format, string> = { pdf: 'download pdf', static: 'download word · static', tracked: 'download word · tracked' };
const menuItem = (which: Format): HTMLButtonElement => menu.querySelector<HTMLButtonElement>(`[data-format="${which}"]`)!;

/** A format whose checks failed is not offered. The PDF is made from the static redline, so it needs that one's checks too. */
function blocked(which: Format): boolean {
  if (!released) return true;
  if (which === 'tracked') return !released.tracked;
  if (which === 'static') return !released.staticRedline;
  return !released.staticRedline || released.pdfFailed;
}
/** The format the main button downloads: the remembered one, or the first that is offered. */
const offered = (): Format | undefined => (!blocked(format) ? format : FORMATS.find((which) => !blocked(which)));

let makingPdf = false;
function renderExport(): void {
  const current = offered();
  for (const which of FORMATS) {
    const item = menuItem(which);
    const sub = item.querySelector<HTMLElement>('.rl-itemsub')!;
    item.disabled = blocked(which);
    item.setAttribute('aria-checked', String(which === current));
    sub.textContent = blocked(which) ? 'Not offered: a check failed' : sub.dataset.sub!;
  }
  downloadButton.disabled = !current || running;
  setHidden('download-icon', makingPdf);
  setHidden('download-spinner', !makingPdf);
  element('download-label').textContent = makingPdf ? 'making pdf' : DOWNLOAD_LABEL[current ?? format];
  downloadButton.dataset.format = current ?? '';
}

function openMenu(): void {
  menu.hidden = false;
  moreButton.setAttribute('aria-expanded', 'true');
  const items = FORMATS.map(menuItem).filter((item) => !item.disabled);
  (items.find((item) => item.getAttribute('aria-checked') === 'true') ?? items[0])?.focus();
}
function closeMenu(returnFocus: boolean): void {
  if (menu.hidden) return;
  menu.hidden = true;
  moreButton.setAttribute('aria-expanded', 'false');
  if (returnFocus) moreButton.focus();
}
moreButton.addEventListener('click', () => (menu.hidden ? openMenu() : closeMenu(true)));
document.addEventListener('click', (event) => {
  if (!menu.hidden && !menu.contains(event.target as Node) && !moreButton.contains(event.target as Node)) closeMenu(false);
});
menu.addEventListener('keydown', (event) => {
  const items = FORMATS.map(menuItem).filter((item) => !item.disabled);
  const at = items.indexOf(document.activeElement as HTMLButtonElement);
  if (event.key === 'Escape') closeMenu(true);
  else if (event.key === 'ArrowDown') items[(at + 1) % items.length]?.focus();
  else if (event.key === 'ArrowUp') items[(at - 1 + items.length) % items.length]?.focus();
  else if (event.key === 'Home') items[0]?.focus();
  else if (event.key === 'End') items[items.length - 1]?.focus();
  else if (event.key === 'Tab') closeMenu(false);
  else return;
  if (event.key !== 'Tab') event.preventDefault();
});
moreButton.addEventListener('keydown', (event) => {
  if (event.key === 'ArrowDown' && menu.hidden) {
    event.preventDefault();
    openMenu();
  } else if (event.key === 'Escape') closeMenu(true);
});
for (const which of FORMATS) {
  menuItem(which).addEventListener('click', () => {
    if (blocked(which)) return;
    format = which;
    remember('redline.format', which);
    renderExport();
    closeMenu(true);
  });
}

/**
 * The one server step. Send the static redline the Checker passed to be
 * converted, have the Checker compare the PDF with it, and keep the PDF only on a pass.
 */
async function makePdf(): Promise<void> {
  const current = released;
  const staticRedline = current?.staticRedline;
  if (!current || !staticRedline || running || current.pdf || current.pdfFailed) return;
  running = true;
  makingPdf = true;
  render();
  renderExport();
  renderPagesNotice();
  pdfStatus.dataset.state = 'working';
  pdfStatus.textContent = 'Sending the static redline for conversion, then checking the PDF…';
  let converted = false;
  try {
    const reviewed = await runPdf(staticRedline, {
      convert: async (docx) => {
        // The converter is given a copy prepared for it (print-copy.ts); the Checker still holds the PDF to the redline itself.
        const copy = await makePrintCopy(docx);
        const response = await fetch(PDF_ENDPOINT, { method: 'POST', headers: { 'Content-Type': DOCX }, body: new Blob([copy.slice().buffer], { type: DOCX }) });
        if (!response.ok) throw new Error((await response.text()).slice(0, 200) || `the server answered ${response.status}`);
        const pdf = new Uint8Array(await response.arrayBuffer());
        converted = true;
        return pdf;
      },
      check: (input) => checker.checkPdf(input),
    });
    // The comparison may have been replaced while the PDF was being made.
    if (released !== current) return;
    if (!converted) {
      // No PDF came back at all (the server was busy, or the connection dropped): there is nothing to check or withhold, and it can be tried again.
      pdfStatus.dataset.state = 'error';
      pdfStatus.textContent = `${reviewed.report.checks[0]?.failure?.message ?? 'The PDF could not be made.'} You can try again.`;
      return;
    }
    showReport('pdf', reviewed.report);
    setHidden('group-pdf', false);
    reports = [...reports.filter((report) => report.output !== 'pdf'), reviewed.report];
    refreshReportLink();
    renderChecksSummary();
    if (reviewed.file) {
      current.pdf = reviewed.file;
      pdfStatus.dataset.state = 'done';
      pdfStatus.textContent = 'The PDF passed its checks.';
      try {
        pdfPages = await readPdfPages(reviewed.file);
      } catch {
        pdfPages = undefined;
      }
      if (released === current && view === 'pages') applyLayout();
    } else {
      current.pdfFailed = true;
      pdfStatus.dataset.state = 'withheld';
      pdfStatus.textContent = 'The PDF did not pass its checks, so it has been withheld. See the checks.';
    }
  } catch (error) {
    if (released === current) {
      pdfStatus.dataset.state = 'error';
      pdfStatus.textContent = `Something went wrong while making the PDF. (${error instanceof Error ? error.message : String(error)})`;
    }
  } finally {
    running = false;
    makingPdf = false;
    render();
    renderExport();
    renderPagesNotice();
  }
}

downloadButton.addEventListener('click', async () => {
  const which = offered();
  const current = released;
  if (!which || !current || running) return;
  if (which === 'tracked' && current.tracked) save(current.tracked, DOCX, 'redline-tracked-changes.docx');
  else if (which === 'static' && current.staticRedline) save(current.staticRedline, DOCX, 'redline-static.docx');
  else if (which === 'pdf') {
    if (!current.pdf) await makePdf();
    if (released === current && current.pdf) save(current.pdf, 'application/pdf', 'redline.pdf');
  }
});

// ---------- the checks ----------

const TICK = '<svg width="14" height="14" viewBox="0 0 14 14" role="img" aria-label="Passed"><path d="M3 7.25 5.75 10 11 4.25" fill="none" stroke="#2CC285" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round"></path></svg>';
const CROSS = '<svg width="14" height="14" viewBox="0 0 14 14" role="img" aria-label="Failed"><path d="M3.75 3.75l6.5 6.5M10.25 3.75l-6.5 6.5" fill="none" stroke="#FF5A5C" stroke-width="1.7" stroke-linecap="round"></path></svg>';

/** The reports of the current comparison, for the list and for the downloadable check report. */
let reports: CheckReport[] = [];
let reportContext: { originalName: string; modifiedName: string; when: Date } | undefined;

function showReport(kind: 'tracked' | 'static' | 'pdf', report: CheckReport): void {
  const list = element<HTMLUListElement>(`report-${kind}`);
  list.replaceChildren();
  for (const check of report.checks) {
    const item = document.createElement('li');
    item.dataset.check = check.id;
    item.dataset.passed = String(check.passed);
    item.innerHTML = check.passed ? TICK : CROSS;
    const body = document.createElement('div');
    body.className = 'rl-checkbody';
    const title = document.createElement('span');
    title.textContent = check.title;
    body.append(title);
    if (check.failure) {
      const detail = document.createElement('div');
      detail.className = 'failure';
      const lines = [check.failure.message];
      if (check.failure.where) lines.push(`Where: ${check.failure.where}`);
      if (check.failure.expected !== undefined) lines.push(`Expected: ${check.failure.expected}`);
      if (check.failure.found !== undefined) lines.push(`Found: ${check.failure.found}`);
      detail.textContent = lines.join('\n');
      body.append(detail);
    }
    item.append(body);
    list.append(item);
  }
  const failed = report.checks.filter((check) => !check.passed).length;
  const verdict = element(`verdict-${kind}`);
  verdict.dataset.passed = String(report.passed && failed === 0);
  verdict.textContent = failed ? `${failed} failed` : 'all passed';
}

function renderChecksSummary(): void {
  const all = reports.flatMap((report) => report.checks);
  const passed = all.filter((check) => check.passed).length;
  const allPassed = passed === all.length && reports.every((report) => report.passed);
  element('checks-summary').textContent = `${passed} of ${all.length} checks passed`;
  checksToggle.dataset.passed = String(allPassed);
  setHidden('checks-ok', !allPassed);
  setHidden('checks-bad', allPassed);
}

function setChecksOpen(open: boolean): void {
  checksToggle.setAttribute('aria-expanded', String(open));
  setHidden('rl-checklist', !open);
}
checksToggle.addEventListener('click', () => {
  closeMenu(false);
  setChecksOpen(checksToggle.getAttribute('aria-expanded') !== 'true');
});

function refreshReportLink(): void {
  if (!reportContext) return;
  const link = element<HTMLAnchorElement>('download-report');
  if (link.href) URL.revokeObjectURL(link.href);
  link.href = URL.createObjectURL(new Blob([reports.map((report) => reportAsText(report, reportContext!)).join('\n')], { type: 'text/plain;charset=utf-8' }));
  link.hidden = false;
}

// ---------- the viewer: layout, page colour, change navigation ----------

let stopNow = 0;
let ringed = false;
let positions: number[] | undefined;
/** True while the viewer is scrolling because an arrow was pressed, not because the reader scrolled. */
let steering = false;
let steeringTimer = 0;

const pageScrolls = (): boolean => viewer.scrollHeight <= viewer.clientHeight + 1;
const stopElements = (n: number): HTMLElement[] => Array.from(doc.querySelectorAll<HTMLElement>(`[data-chg="${n}"]`));

function renderCounter(): void {
  const total = model?.stops ?? 0;
  counter.textContent = `${total ? stopNow : 0} / ${total}`;
  element<HTMLButtonElement>('prev-change').disabled = !total;
  element<HTMLButtonElement>('next-change').disabled = !total;
}

function renderRing(): void {
  for (const old of Array.from(doc.querySelectorAll('.rl-cur'))) old.classList.remove('rl-cur');
  if (ringed) for (const part of stopElements(stopNow)) part.classList.add('rl-cur');
}

function step(by: number): void {
  const total = model?.stops ?? 0;
  if (!total) return;
  stopNow = ((stopNow - 1 + by + total) % total) + 1;
  ringed = true;
  renderCounter();
  renderRing();
  const first = stopElements(stopNow)[0];
  if (!first) return;
  const smooth = !window.matchMedia('(prefers-reduced-motion: reduce)').matches;
  steering = true;
  window.clearTimeout(steeringTimer);
  steeringTimer = window.setTimeout(() => (steering = false), 1200);
  if (pageScrolls()) {
    window.scrollTo({ top: Math.max(0, first.getBoundingClientRect().top + window.scrollY - window.innerHeight / 3), behavior: smooth ? 'smooth' : 'auto' });
  } else {
    const top = first.getBoundingClientRect().top - viewer.getBoundingClientRect().top + viewer.scrollTop - viewer.clientHeight / 3;
    viewer.scrollTo({ top: Math.max(0, top), behavior: smooth ? 'smooth' : 'auto' });
  }
}
element('prev-change').addEventListener('click', () => step(-1));
element('next-change').addEventListener('click', () => step(1));

/** The reader scrolled by hand: the counter follows to the change nearest the top third of the view. No ring until an arrow is pressed. */
let scrollQueued = false;
function onScroll(): void {
  if (steering || scrollQueued || !model?.stops || viewer.hidden) return;
  scrollQueued = true;
  requestAnimationFrame(() => {
    scrollQueued = false;
    if (steering || !model?.stops) return;
    const total = model.stops;
    const whole = pageScrolls();
    const origin = whole ? -window.scrollY : viewer.getBoundingClientRect().top - viewer.scrollTop;
    positions ??= Array.from({ length: total }, (_, i) => {
      const first = stopElements(i + 1)[0];
      return first ? first.getBoundingClientRect().top - origin : Number.POSITIVE_INFINITY;
    });
    const line = whole ? window.scrollY + window.innerHeight / 3 : viewer.scrollTop + viewer.clientHeight / 3;
    let nearest = 0;
    for (let i = 1; i < total; i++) if (Math.abs(positions[i]! - line) < Math.abs(positions[nearest]! - line)) nearest = i;
    if (nearest + 1 === stopNow && !ringed) return;
    if (nearest + 1 !== stopNow) {
      stopNow = nearest + 1;
      ringed = false;
      renderCounter();
      renderRing();
    }
  });
}
viewer.addEventListener('scroll', onScroll, { passive: true });
window.addEventListener('scroll', onScroll, { passive: true });
for (const target of [viewer, window]) target.addEventListener('scrollend', () => (steering = false));
new ResizeObserver(() => (positions = undefined)).observe(doc);

function renderPagesNotice(): void {
  const wanted = view === 'pages' && !!model && !pdfPages;
  setHidden('pages-notice', !wanted);
  if (!wanted) return;
  const failed = !!released?.pdfFailed || (!!released?.pdf && !pdfPages);
  element('pages-notice-text').textContent = failed
    ? 'Pages are not available for this redline, because its PDF could not be made or did not pass its checks.'
    : makingPdf
      ? 'Making the PDF, so each page breaks where the downloaded file does. The static redline is sent to our server for this, and deleted there straight after conversion.'
      : 'Pages follow the PDF, so each page breaks where the downloaded file does. Making a PDF sends the static redline to our server, where it is deleted straight after conversion.';
  setHidden('pages-make', failed);
  element<HTMLButtonElement>('pages-make').disabled = running;
}
element('pages-make').addEventListener('click', () => void makePdf());

function applyLayout(): void {
  element('view-continuous').setAttribute('aria-pressed', String(view === 'continuous'));
  element('view-pages').setAttribute('aria-pressed', String(view === 'pages'));
  element('page-dark').setAttribute('aria-pressed', String(paper === 'dark'));
  element('page-light').setAttribute('aria-pressed', String(paper === 'light'));
  doc.classList.toggle('rl-light', paper === 'light');
  renderPagesNotice();
  if (!model) return;
  layOut(doc, model, view === 'pages' ? pdfPages : undefined);
  positions = undefined;
  renderRing();
}
for (const [id, value] of [['view-continuous', 'continuous'], ['view-pages', 'pages']] as const) {
  element(id).addEventListener('click', () => {
    view = value;
    remember('redline.view', value);
    applyLayout();
    // Pressing "pages" is a request for the PDF, which sets the page breaks (owner's decision, 2026-11-25): the static
    // redline is sent for conversion at once. A remembered choice of pages does not send anything by itself: after a
    // new comparison the notice and its button are shown, and the PDF is made when the user presses one of them.
    if (value === 'pages' && released?.staticRedline && !released.pdf && !released.pdfFailed) void makePdf();
  });
}
for (const [id, value] of [['page-dark', 'dark'], ['page-light', 'light']] as const) {
  element(id).addEventListener('click', () => {
    paper = value;
    remember('redline.page', value);
    doc.classList.toggle('rl-light', value === 'light');
    element('page-dark').setAttribute('aria-pressed', String(value === 'dark'));
    element('page-light').setAttribute('aria-pressed', String(value === 'light'));
  });
}

/** The document's own layout rules, kept in one style element that each comparison replaces. */
function setDocumentCss(css: string): void {
  let style = document.getElementById('redline-document-style');
  if (!style) {
    style = document.createElement('style');
    style.id = 'redline-document-style';
    document.head.append(style);
  }
  style.textContent = css;
}

// ---------- showing a finished comparison ----------

async function show(result: FlowResult, names: { originalName: string; modifiedName: string }, when: Date): Promise<void> {
  released = { tracked: result.tracked.file, staticRedline: result.staticRedline.file, pdfFailed: false };
  showReport('tracked', result.tracked.report);
  showReport('static', result.staticRedline.report);
  reports = [result.staticRedline.report, result.tracked.report];
  reportContext = { ...names, when };
  refreshReportLink();
  renderChecksSummary();

  // The counts are shown only with a static redline the Checker passed (it has then confirmed them).
  if (result.counts) {
    let total = 0;
    for (const [key, value] of Object.entries(result.counts)) {
      total += value;
      const cell = document.querySelector(`#summary [data-count="${key}"]`);
      if (cell) cell.textContent = String(value);
    }
    element('changes-total').textContent = String(total);
  }

  // The viewer shows the static redline, and only if the Checker passed it.
  model = undefined;
  if (result.staticRedline.file) {
    try {
      model = buildRedline(await producer.renderHtml(result.staticRedline.file));
      setDocumentCss(model.css);
    } catch {
      model = undefined;
      element('withheld-title').textContent = 'redline not shown';
      viewerMessage.textContent = 'The redline passed its checks but could not be displayed here. You can still download it.';
    }
  } else {
    element('withheld-title').textContent = 'redline withheld';
    viewerMessage.textContent = `The ${OUTPUT_NAME['static-redline'].toLowerCase()} did not pass its checks, so it is not shown. The checks on the left say which one failed and where.`;
  }
  stopNow = model?.stops ? 1 : 0;
  ringed = !!model?.stops;
  viewer.setAttribute('aria-label', `Redline of ${chosen.modified ? baseName(chosen.modified) : 'the document'}`);
}

async function create(): Promise<void> {
  const { original, modified } = chosen;
  if (!original || !modified || running || validating > 0) return;
  running = true;
  clearResults();
  setStage('loading');
  say('working', '');
  const started = performance.now();
  let outcome: (() => void) | undefined;

  try {
    const refusal =
      (await refusalFor({ name: original.name, size: original.size, bytes: async () => new Uint8Array(await original.arrayBuffer()) }, 'original')) ??
      (await refusalFor({ name: modified.name, size: modified.size, bytes: async () => new Uint8Array(await modified.arrayBuffer()) }, 'modified'));
    if (refusal) {
      outcome = () => {
        flowError.textContent = refusal;
        flowError.hidden = false;
        say('refused', refusal);
      };
      return;
    }
    const when = new Date();
    const names = { originalName: displayName(original.name), modifiedName: displayName(modified.name) };
    const [originalBytes, modifiedBytes] = await Promise.all([original.arrayBuffer(), modified.arrayBuffer()]);
    // An upload with tracked changes of its own is compared as if they had been accepted, and the user is told.
    const [inOriginal, inModified] = await Promise.all([containsTrackedChanges(new Uint8Array(originalBytes)), containsTrackedChanges(new Uint8Array(modifiedBytes))]);
    const trackedChangesIn: 'original' | 'modified' | 'both' | undefined = inOriginal && inModified ? 'both' : inOriginal ? 'original' : inModified ? 'modified' : undefined;
    const details = { ...names, comparedAt: formatComparedAt(when), trackedChangesIn };
    const result = await runComparison(new Uint8Array(originalBytes), new Uint8Array(modifiedBytes), details, {
      produce: (a, b) => producer.compare(a, b),
      makeStatic: (tracked, source, info) => producer.makeStatic(tracked, source, info),
      checkTracked: (input) => checker.checkTrackedChanges(input),
      checkStatic: (input) => checker.checkStaticRedline(input),
    });
    status.dataset.engineMs = String(Math.round(result.engineMs));
    await show(result, names, when);
    await new Promise((resolve) => setTimeout(resolve, Math.max(0, MIN_CREATING_MS - (performance.now() - started))));
    outcome = () => {
      element('tracked-notice-text').textContent = trackedChangesIn ? TRACKED_NOTICE[trackedChangesIn] : '';
      setHidden('tracked-notice', !trackedChangesIn);
      setHidden('section-changes', !result.counts);
      applyLayout();
      renderCounter();
      renderExport();
      viewer.scrollTop = 0;
      const count = [result.tracked.file, result.staticRedline.file].filter(Boolean).length;
      if (count === 2) say('done', 'Done. Every check passed.');
      else {
        // Said in plain sight, beside the downloads, as well as to screen readers.
        const text = count === 0 ? 'The redline did not pass its checks, so it has been withheld. See the checks.' : 'One of the two Word files did not pass its checks and has been withheld. See the checks.';
        element('withheld-note').textContent = text;
        setHidden('withheld-note', false);
        say('withheld', text);
      }
    };
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    const text = error instanceof ProducerError ? `The documents could not be compared. (${message})` : `Something went wrong. (${message})`;
    outcome = () => {
      flowError.textContent = text;
      flowError.hidden = false;
      say('error', text);
    };
  } finally {
    running = false;
    const finished = !!released;
    if (!finished) clearResults();
    setStage(finished ? 'done' : 'idle');
    outcome?.();
  }
}
createButton.addEventListener('click', () => void create());
createAgain.addEventListener('click', () => void create());

// ⌘+Enter on a Mac, Ctrl+Enter elsewhere, creates the redline when both files are present.
const onMac = /Mac|iPhone|iPad/.test(navigator.platform);
element('create-hint').textContent = onMac ? '⌘ ↵' : 'Ctrl ↵';
document.addEventListener('keydown', (event) => {
  if (event.key === 'Enter' && (onMac ? event.metaKey : event.ctrlKey) && !event.repeat) {
    event.preventDefault();
    void create();
  }
});

render();
renderExport();
applyLayout();
restartBand();
