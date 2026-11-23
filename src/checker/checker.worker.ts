/// <reference lib="webworker" />
/**
 * The Checker's own Web Worker, separate from the one that produces outputs.
 * It receives the two uploads and an output as raw bytes and returns a report.
 */
import { checkTrackedChanges } from './index';
import type { CheckRequest, CheckResponse } from './protocol';

declare const self: DedicatedWorkerGlobalScope;

self.addEventListener('message', async (event: MessageEvent<CheckRequest>) => {
  const request = event.data;
  if (request.type !== 'check-tracked-changes') return;
  let response: CheckResponse;
  try {
    const report = await checkTrackedChanges({
      original: new Uint8Array(request.original),
      modified: new Uint8Array(request.modified),
      output: new Uint8Array(request.output),
    });
    response = { type: 'checked', id: request.id, report };
  } catch (error) {
    response = { type: 'check-crashed', id: request.id, message: error instanceof Error ? error.message : String(error) };
  }
  self.postMessage(response);
});
