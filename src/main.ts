/**
 * Bare development page: pick two .docx files, press Compare, see the check
 * report, and (only if every check passed) download the tracked-changes .docx.
 * The real interface comes in Stage 4.
 */
import { ProducerError, runComparison } from './app/flow';
import { reportAsText, type CheckReport } from './checker';
import { CheckerClient } from './checker/checker-client';
import { CompareClient } from './producer/compare-client';

function element<T extends HTMLElement>(id: string): T {
  const found = document.getElementById(id);
  if (!found) throw new Error(`missing element #${id}`);
  return found as T;
}

const originalInput = element<HTMLInputElement>('original');
const modifiedInput = element<HTMLInputElement>('modified');
const compareButton = element<HTMLButtonElement>('compare');
const status = element<HTMLParagraphElement>('status');
const reportList = element<HTMLUListElement>('report');
const download = element<HTMLAnchorElement>('download');
const reportDownload = element<HTMLAnchorElement>('report-download');

const producer = new CompareClient();
const checker = new CheckerClient();
const urls: string[] = [];

function refreshButton(): void {
  compareButton.disabled = !(originalInput.files?.length && modifiedInput.files?.length);
}
originalInput.addEventListener('change', refreshButton);
modifiedInput.addEventListener('change', refreshButton);
refreshButton();

function objectUrl(data: Uint8Array | string, type: string): string {
  const url = URL.createObjectURL(new Blob([typeof data === 'string' ? data : data.slice().buffer], { type }));
  urls.push(url);
  return url;
}

function showReport(report: CheckReport): void {
  reportList.replaceChildren();
  for (const check of report.checks) {
    const item = document.createElement('li');
    item.dataset.check = check.id;
    item.dataset.passed = String(check.passed);
    item.textContent = `${check.title}: ${check.passed ? 'passed' : 'failed'}`;
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
    reportList.append(item);
  }
  reportList.hidden = false;
}

compareButton.addEventListener('click', async () => {
  const original = originalInput.files?.[0];
  const modified = modifiedInput.files?.[0];
  if (!original || !modified) return;

  compareButton.disabled = true;
  download.hidden = true;
  reportDownload.hidden = true;
  reportList.hidden = true;
  download.removeAttribute('href');
  for (const url of urls.splice(0)) URL.revokeObjectURL(url);
  status.dataset.state = 'working';
  status.textContent = 'Comparing and checking…';

  try {
    const [originalBytes, modifiedBytes] = await Promise.all([original.arrayBuffer(), modified.arrayBuffer()]);
    const result = await runComparison(new Uint8Array(originalBytes), new Uint8Array(modifiedBytes), {
      produce: (a, b) => producer.compare(a, b),
      check: (input) => checker.checkTrackedChanges(input),
    });
    showReport(result.report);
    reportDownload.href = objectUrl(
      reportAsText(result.report, { originalName: original.name, modifiedName: modified.name, when: new Date() }),
      'text/plain;charset=utf-8',
    );
    reportDownload.hidden = false;
    status.dataset.engineMs = String(Math.round(result.engineMs));
    if (result.trackedChanges) {
      download.href = objectUrl(result.trackedChanges, 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
      download.hidden = false;
      status.dataset.state = 'done';
      status.textContent = `Done in ${(result.engineMs / 1000).toFixed(1)} seconds. Every check passed.`;
    } else {
      status.dataset.state = 'withheld';
      status.textContent = 'The redline did not pass its checks, so it has been withheld. See the report below.';
    }
  } catch (error) {
    status.dataset.state = 'error';
    const message = error instanceof Error ? error.message : String(error);
    status.textContent = error instanceof ProducerError ? `The comparison failed: ${message}` : `Something went wrong: ${message}`;
  } finally {
    refreshButton();
  }
});
