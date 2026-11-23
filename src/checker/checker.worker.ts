/// <reference lib="webworker" />
/**
 * The Checker's own Web Worker, separate from the one that produces outputs.
 * It receives the two uploads and an output as raw bytes and returns a report.
 */
import { checkPdf, checkStaticRedline, checkTrackedChanges } from './index';
import type { CheckRequest, CheckResponse, PdfCheckRequest } from './protocol';

declare const self: DedicatedWorkerGlobalScope;

self.addEventListener('message', async (event: MessageEvent<CheckRequest | PdfCheckRequest>) => {
  const request = event.data;
  if (request.type === 'check-pdf') {
    let reply: CheckResponse;
    try {
      reply = { type: 'checked', id: request.id, report: await checkPdf({ staticRedline: new Uint8Array(request.staticRedline), pdf: new Uint8Array(request.pdf) }) };
    } catch (error) {
      reply = { type: 'check-crashed', id: request.id, message: error instanceof Error ? error.message : String(error) };
    }
    self.postMessage(reply);
    return;
  }
  if (request.type !== 'check-tracked-changes' && request.type !== 'check-static-redline') return;
  let response: CheckResponse;
  try {
    const files = { original: new Uint8Array(request.original), modified: new Uint8Array(request.modified), output: new Uint8Array(request.output) };
    const report =
      request.type === 'check-tracked-changes'
        ? await checkTrackedChanges(files)
        : await checkStaticRedline({ ...files, originalName: request.originalName ?? '', modifiedName: request.modifiedName ?? '', comparedAt: request.comparedAt ?? '' });
    response = { type: 'checked', id: request.id, report };
  } catch (error) {
    response = { type: 'check-crashed', id: request.id, message: error instanceof Error ? error.message : String(error) };
  }
  self.postMessage(response);
});
