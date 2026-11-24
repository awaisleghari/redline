/**
 * The page: two documents in, a checked redline out.
 *
 * Nothing here decides whether an output is right. The flow (src/app/flow.ts)
 * hands back a file only when the independent Checker has passed it, and this
 * page offers for download, and shows in the viewer, only files it was handed.
 */
import { ProducerError, runComparison, runPdf, type FlowResult, type Reviewed } from './app/flow';
import { displayName, formatComparedAt } from './app/format';
import { containsTrackedChanges, refusalFor, TRACKED_NOTICE } from './app/validate';
import { OUTPUT_NAME, reportAsText, type CheckReport } from './checker';
import { CheckerClient } from './checker/checker-client';
import { CompareClient } from './producer/compare-client';
import { showPdf } from './app/pdf-viewer';

function element<T extends HTMLElement>(id: string): T {
  const found = document.getElementById(id);
  if (!found) throw new Error(`missing element #${id}`);
  return found as T;
}

const inputs = { original: element<HTMLInputElement>('original'), modified: element<HTMLInputElement>('modified') };
const chosen: { original?: File; modified?: File } = {};
const compareButton = element<HTMLButtonElement>('compare');
const status = element<HTMLParagraphElement>('status');
const progress = element<HTMLDivElement>('progress');
const progressText = element<HTMLSpanElement>('progress-text');
const results = element<HTMLElement>('results');
const summary = element<HTMLDivElement>('summary');
const viewer = element<HTMLIFrameElement>('viewer');
const viewerMessage = element<HTMLParagraphElement>('viewer-message');
const DOCX = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';

const makePdfButton = element<HTMLButtonElement>('make-pdf');
const pdfStatus = element<HTMLParagraphElement>('pdf-status');
const pdfViewer = element<HTMLDivElement>('pdf-viewer');
const tabs = element<HTMLDivElement>('viewer-tabs');
const PDF_ENDPOINT = '/api/pdf';
/** The static redline the Checker passed in the current comparison (the only file ever sent to the server). */
let releasedStatic: Uint8Array | undefined;

const producer = new CompareClient();
const checker = new CheckerClient();
const urls: string[] = [];
/** True while a comparison is running: the files cannot be changed and no second run can start. */
let running = false;

/**
 * The redline as a web page for the viewer, made inert: links in the document
 * cannot be followed (a click would otherwise send a request to wherever the
 * document's author chose), and the page may load nothing from anywhere.
 */
function inert(html: string): string {
  // Only attributes inside tags are removed; the document's words are never touched.
  const withoutLinks = html.replace(/<[a-zA-Z][^<>]*>/g, (tag) => tag.replace(/\s(?:href|action|formaction|ping)\s*=\s*(?:"[^"]*"|'[^']*'|[^\s>]+)/gi, ''));
  const policy =
    '<meta http-equiv="Content-Security-Policy" content="default-src \'none\'; style-src \'unsafe-inline\'; img-src data:; font-src data:">' +
    '<style>a{pointer-events:none;cursor:text;}</style>';
  return /<head[^>]*>/i.test(withoutLinks) ? withoutLinks.replace(/<head[^>]*>/i, (head) => head + policy) : policy + withoutLinks;
}

function say(state: string, text: string): void {
  status.dataset.state = state;
  status.textContent = text;
}

function choose(role: 'original' | 'modified', file: File | undefined): void {
  if (running) return;
  chosen[role] = file;
  element(`name-${role}`).textContent = file ? displayName(file.name) : '';
  compareButton.disabled = !(chosen.original && chosen.modified);
  // Results on screen belong to the files they were made from: a new choice clears them.
  reset();
  say('', '');
}

function setRunning(now: boolean): void {
  running = now;
  inputs.original.disabled = now;
  inputs.modified.disabled = now;
  compareButton.disabled = now || !(chosen.original && chosen.modified);
}

for (const role of ['original', 'modified'] as const) {
  const input = inputs[role];
  const zone = element(`drop-${role}`);
  input.addEventListener('change', () => choose(role, input.files?.[0]));
  zone.addEventListener('dragover', (event) => {
    event.preventDefault();
    zone.classList.add('over');
  });
  zone.addEventListener('dragleave', () => zone.classList.remove('over'));
  zone.addEventListener('drop', (event) => {
    event.preventDefault();
    zone.classList.remove('over');
    const file = event.dataTransfer?.files[0];
    if (file && !running) {
      input.value = '';
      choose(role, file);
    }
  });
}

function objectUrl(data: Uint8Array | string, type: string): string {
  const url = URL.createObjectURL(new Blob([typeof data === 'string' ? data : data.slice().buffer], { type }));
  urls.push(url);
  return url;
}

function showReport(kind: 'tracked' | 'static' | 'pdf', report: CheckReport): void {
  const list = element<HTMLUListElement>(`report-${kind}`);
  list.replaceChildren();
  for (const check of report.checks) {
    const item = document.createElement('li');
    item.dataset.check = check.id;
    item.dataset.passed = String(check.passed);
    const title = document.createElement('span');
    title.textContent = `${check.title}: `;
    const result = document.createElement('span');
    result.className = 'result';
    result.textContent = check.passed ? 'passed' : 'failed';
    item.append(title, result);
    if (check.failure) {
      const detail = document.createElement('div');
      detail.className = 'failure';
      const lines = [check.failure.message];
      if (check.failure.where) lines.push(`Where: ${check.failure.where}`);
      if (check.failure.expected !== undefined) lines.push(`Expected: ${check.failure.expected}`);
      if (check.failure.found !== undefined) lines.push(`Found: ${check.failure.found}`);
      detail.textContent = lines.join('\n');
      item.append(detail);
    }
    list.append(item);
  }
  const verdict = element(`verdict-${kind}`);
  verdict.dataset.passed = String(report.passed);
  verdict.textContent = report.passed ? '(all checks passed)' : '(failed: withheld)';
}

function offer(kind: 'tracked' | 'static' | 'pdf', reviewed: Reviewed): void {
  const link = element<HTMLAnchorElement>(`download-${kind}`);
  const withheld = element(`withheld-${kind}`);
  link.removeAttribute('href');
  link.hidden = true;
  withheld.hidden = true;
  if (reviewed.file) {
    link.href = objectUrl(reviewed.file, kind === 'pdf' ? 'application/pdf' : DOCX);
    link.hidden = false;
  } else {
    withheld.hidden = false;
  }
}

function reset(): void {
  for (const url of urls.splice(0)) URL.revokeObjectURL(url);
  results.hidden = true;
  element<HTMLParagraphElement>('tracked-notice').hidden = true;
  summary.hidden = true;
  viewer.hidden = true;
  viewer.removeAttribute('srcdoc');
  viewerMessage.hidden = true;
  for (const id of ['download-tracked', 'download-static', 'download-report', 'download-pdf']) {
    const link = element<HTMLAnchorElement>(id);
    link.removeAttribute('href');
    link.hidden = true;
  }
  releasedStatic = undefined;
  makePdfButton.disabled = true;
  pdfStatus.textContent = '';
  element('withheld-pdf').hidden = true;
  element('section-pdf').hidden = true;
  pdfViewer.hidden = true;
  pdfViewer.replaceChildren();
  tabs.hidden = true;
  pendingReports = [];
}

/** Which of the two the viewer shows. The PDF tab exists only once a PDF has passed its checks. */
function showTab(which: 'word' | 'pdf'): void {
  viewer.hidden = which !== 'word';
  pdfViewer.hidden = which !== 'pdf';
  element('show-word').setAttribute('aria-pressed', String(which === 'word'));
  element('show-pdf').setAttribute('aria-pressed', String(which === 'pdf'));
}
element('show-word').addEventListener('click', () => showTab('word'));
element('show-pdf').addEventListener('click', () => showTab('pdf'));

/** The reports of the current comparison, for the downloadable check report. */
let pendingReports: CheckReport[] = [];
let reportContext: { originalName: string; modifiedName: string; when: Date } | undefined;
function refreshReportLink(): void {
  if (!reportContext) return;
  const link = element<HTMLAnchorElement>('download-report');
  link.href = objectUrl(pendingReports.map((report) => reportAsText(report, reportContext!)).join('\n'), 'text/plain;charset=utf-8');
  link.hidden = false;
}

async function show(result: FlowResult, names: { originalName: string; modifiedName: string }, when: Date): Promise<void> {
  showReport('tracked', result.tracked.report);
  showReport('static', result.staticRedline.report);
  offer('tracked', result.tracked);
  offer('static', result.staticRedline);

  pendingReports = [result.staticRedline.report, result.tracked.report];
  reportContext = { ...names, when };
  refreshReportLink();
  // A PDF can be asked for only of a static redline that passed its checks.
  releasedStatic = result.staticRedline.file;
  makePdfButton.disabled = !releasedStatic;

  if (result.counts) {
    for (const [key, value] of Object.entries(result.counts)) {
      const cell = summary.querySelector(`[data-count="${key}"]`);
      if (cell) cell.textContent = String(value);
    }
    summary.hidden = false;
  }

  // The viewer shows the static redline, and only if the Checker passed it.
  if (result.staticRedline.file) {
    try {
      viewer.srcdoc = inert(await producer.renderHtml(result.staticRedline.file));
      viewer.hidden = false;
    } catch {
      viewerMessage.textContent = 'The redline passed its checks but could not be displayed here. You can still download it.';
      viewerMessage.hidden = false;
    }
  } else {
    viewerMessage.textContent = `The ${OUTPUT_NAME['static-redline'].toLowerCase()} did not pass its checks, so it is not shown. The check report says which check failed and where.`;
    viewerMessage.hidden = false;
  }
  results.hidden = false;
}

makePdfButton.addEventListener('click', async () => {
  const staticRedline = releasedStatic;
  if (!staticRedline || running) return;
  setRunning(true);
  makePdfButton.disabled = true;
  pdfStatus.textContent = 'Sending the static redline for conversion, then checking the PDF…';
  try {
    const reviewed = await runPdf(staticRedline, {
      convert: async (docx) => {
        const response = await fetch(PDF_ENDPOINT, { method: 'POST', headers: { 'Content-Type': DOCX }, body: new Blob([docx.slice().buffer], { type: DOCX }) });
        if (!response.ok) throw new Error((await response.text()).slice(0, 200) || `the server answered ${response.status}`);
        return new Uint8Array(await response.arrayBuffer());
      },
      check: (input) => checker.checkPdf(input),
    });
    // The comparison may have been replaced while the PDF was being made.
    if (releasedStatic !== staticRedline) return;
    showReport('pdf', reviewed.report);
    element('section-pdf').hidden = false;
    offer('pdf', reviewed);
    pendingReports = [...pendingReports.filter((report) => report.output !== 'pdf'), reviewed.report];
    refreshReportLink();
    if (reviewed.file) {
      await showPdf(pdfViewer, reviewed.file);
      tabs.hidden = false;
      showTab('pdf');
      pdfStatus.textContent = 'The PDF passed its checks.';
    } else {
      pdfStatus.textContent = 'The PDF did not pass its checks, so it has been withheld. See the check report.';
    }
  } catch (error) {
    pdfStatus.textContent = `Something went wrong while making the PDF. (${error instanceof Error ? error.message : String(error)})`;
  } finally {
    setRunning(false);
    makePdfButton.disabled = !releasedStatic || !element<HTMLAnchorElement>('download-pdf').hidden;
  }
});

compareButton.addEventListener('click', async () => {
  const { original, modified } = chosen;
  if (!original || !modified || running) return;
  setRunning(true);
  reset();
  progressText.textContent = 'Reading the documents…';
  progress.hidden = false;
  say('working', '');

  try {
    const refusal =
      (await refusalFor({ name: original.name, size: original.size, bytes: async () => new Uint8Array(await original.arrayBuffer()) }, 'original')) ??
      (await refusalFor({ name: modified.name, size: modified.size, bytes: async () => new Uint8Array(await modified.arrayBuffer()) }, 'modified'));
    if (refusal) {
      say('refused', refusal);
      return;
    }
    progressText.textContent = 'Comparing and checking…';
    const when = new Date();
    const names = { originalName: displayName(original.name), modifiedName: displayName(modified.name) };
    const [originalBytes, modifiedBytes] = await Promise.all([original.arrayBuffer(), modified.arrayBuffer()]);
    // An upload with tracked changes of its own is compared as if they had been accepted, and the user is told.
    const [inOriginal, inModified] = await Promise.all([containsTrackedChanges(new Uint8Array(originalBytes)), containsTrackedChanges(new Uint8Array(modifiedBytes))]);
    const trackedChangesIn: 'original' | 'modified' | 'both' | undefined = inOriginal && inModified ? 'both' : inOriginal ? 'original' : inModified ? 'modified' : undefined;
    const details = { ...names, comparedAt: formatComparedAt(when), trackedChangesIn };
    const trackedNotice = element<HTMLParagraphElement>('tracked-notice');
    trackedNotice.textContent = trackedChangesIn ? TRACKED_NOTICE[trackedChangesIn] : '';
    trackedNotice.hidden = !trackedChangesIn;
    const result = await runComparison(new Uint8Array(originalBytes), new Uint8Array(modifiedBytes), details, {
      produce: (a, b) => producer.compare(a, b),
      makeStatic: (tracked, source, info) => producer.makeStatic(tracked, source, info),
      checkTracked: (input) => checker.checkTrackedChanges(input),
      checkStatic: (input) => checker.checkStaticRedline(input),
    });
    status.dataset.engineMs = String(Math.round(result.engineMs));
    await show(result, names, when);
    const released = [result.tracked.file, result.staticRedline.file].filter(Boolean).length;
    if (released === 2) say('done', 'Done. Every check passed.');
    else say('withheld', released === 0 ? 'The redline did not pass its checks, so it has been withheld. See the check report.' : 'One of the two Word files did not pass its checks and has been withheld. See the check report.');
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    say('error', error instanceof ProducerError ? `The documents could not be compared. (${message})` : `Something went wrong. (${message})`);
  } finally {
    progress.hidden = true;
    setRunning(false);
  }
});
