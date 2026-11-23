/**
 * The page: two documents in, a checked redline out.
 *
 * Nothing here decides whether an output is right. The flow (src/app/flow.ts)
 * hands back a file only when the independent Checker has passed it, and this
 * page offers for download, and shows in the viewer, only files it was handed.
 */
import { ProducerError, runComparison, type FlowResult, type Reviewed } from './app/flow';
import { displayName, formatComparedAt } from './app/format';
import { refusalFor } from './app/validate';
import { OUTPUT_NAME, reportAsText, type CheckReport } from './checker';
import { CheckerClient } from './checker/checker-client';
import { CompareClient } from './producer/compare-client';

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

function showReport(kind: 'tracked' | 'static', report: CheckReport): void {
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

function offer(kind: 'tracked' | 'static', reviewed: Reviewed): void {
  const link = element<HTMLAnchorElement>(`download-${kind}`);
  const withheld = element(`withheld-${kind}`);
  link.removeAttribute('href');
  link.hidden = true;
  withheld.hidden = true;
  if (reviewed.file) {
    link.href = objectUrl(reviewed.file, DOCX);
    link.hidden = false;
  } else {
    withheld.hidden = false;
  }
}

function reset(): void {
  for (const url of urls.splice(0)) URL.revokeObjectURL(url);
  results.hidden = true;
  summary.hidden = true;
  viewer.hidden = true;
  viewer.removeAttribute('srcdoc');
  viewerMessage.hidden = true;
  for (const id of ['download-tracked', 'download-static', 'download-report']) {
    const link = element<HTMLAnchorElement>(id);
    link.removeAttribute('href');
    link.hidden = true;
  }
}

async function show(result: FlowResult, names: { originalName: string; modifiedName: string }, when: Date): Promise<void> {
  showReport('tracked', result.tracked.report);
  showReport('static', result.staticRedline.report);
  offer('tracked', result.tracked);
  offer('static', result.staticRedline);

  const reportText = [result.staticRedline.report, result.tracked.report].map((report) => reportAsText(report, { ...names, when })).join('\n');
  const reportLink = element<HTMLAnchorElement>('download-report');
  reportLink.href = objectUrl(reportText, 'text/plain;charset=utf-8');
  reportLink.hidden = false;

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
    const details = { ...names, comparedAt: formatComparedAt(when) };
    const [originalBytes, modifiedBytes] = await Promise.all([original.arrayBuffer(), modified.arrayBuffer()]);
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
