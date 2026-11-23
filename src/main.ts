/**
 * Bare development page (Stage 1): pick two .docx files, press Compare,
 * download the tracked-changes .docx. The real interface comes in Stage 4.
 */
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
const download = element<HTMLAnchorElement>('download');

const client = new CompareClient();
let downloadUrl: string | undefined;

function refreshButton(): void {
  compareButton.disabled = !(originalInput.files?.length && modifiedInput.files?.length);
}
originalInput.addEventListener('change', refreshButton);
modifiedInput.addEventListener('change', refreshButton);
refreshButton();

compareButton.addEventListener('click', async () => {
  const original = originalInput.files?.[0];
  const modified = modifiedInput.files?.[0];
  if (!original || !modified) return;

  compareButton.disabled = true;
  download.hidden = true;
  if (downloadUrl) URL.revokeObjectURL(downloadUrl);
  downloadUrl = undefined;
  status.dataset.state = 'working';
  status.textContent = 'Comparing…';

  try {
    const result = await client.compare(original, modified);
    const blob = new Blob([result.trackedChanges.slice().buffer], {
      type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    });
    downloadUrl = URL.createObjectURL(blob);
    download.href = downloadUrl;
    download.hidden = false;
    status.dataset.state = 'done';
    status.dataset.engineMs = String(Math.round(result.engineMs));
    status.textContent = `Done in ${(result.engineMs / 1000).toFixed(1)} seconds.`;
  } catch (error) {
    status.dataset.state = 'error';
    status.textContent = `The comparison failed: ${error instanceof Error ? error.message : String(error)}`;
  } finally {
    refreshButton();
  }
});
