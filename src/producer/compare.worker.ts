/// <reference lib="webworker" />
/**
 * The compare Web Worker. The engine (docxodus, a .NET comparer compiled to
 * WebAssembly) runs here, off the main thread, so the page stays responsive.
 * Documents are handed over in memory and never leave the browser.
 */
import { compareDocuments, initialize } from 'docxodus/core';
import { REVISION_AUTHOR, type CompareRequest, type CompareResponse } from './protocol';

declare const self: DedicatedWorkerGlobalScope;

// Before each comparison the engine waits for two animation frames, so that a
// page can paint a "loading" state first. That only makes sense on the main
// thread: a worker has nothing to paint, and browsers may pause animation
// frames while a tab is in the background, which would stall the comparison
// until the user came back. Answer those requests on an ordinary timer instead.
self.requestAnimationFrame = (callback: FrameRequestCallback): number =>
  self.setTimeout(() => callback(performance.now()), 0);
self.cancelAnimationFrame = (handle: number): void => self.clearTimeout(handle);

// The engine's runtime files are served from <site root>/wasm/ (see scripts/copy-wasm.mjs).
const WASM_BASE = new URL(`${import.meta.env.BASE_URL}wasm/`, self.location.origin).href;

let ready: Promise<void> | undefined;
function engineReady(): Promise<void> {
  ready ??= initialize(WASM_BASE).catch((error: unknown) => {
    ready = undefined; // allow a retry on the next request
    throw error;
  });
  return ready;
}

// Listen with addEventListener, never `self.onmessage = …`. The .NET runtime
// looks at `self.onmessage` when it starts: if a handler is already set it
// assumes it is one of its own helper threads and waits for instructions that
// never come, so the engine would never finish loading.
self.addEventListener('message', async (event: MessageEvent<CompareRequest>) => {
  const request = event.data;
  if (request.type !== 'compare') return;
  let response: CompareResponse;
  let transfer: Transferable[] = [];
  try {
    await engineReady();
    const started = performance.now();
    const result = await compareDocuments(new Uint8Array(request.original), new Uint8Array(request.modified), {
      authorName: REVISION_AUTHOR,
    });
    const engineMs = performance.now() - started;
    // Copy into a buffer we own outright so it can be handed to the page without a second copy.
    const trackedChanges = result.slice().buffer;
    response = { type: 'compared', id: request.id, trackedChanges, engineMs };
    transfer = [trackedChanges];
  } catch (error) {
    // Report the engine's own message only; never file names or document contents.
    response = { type: 'failed', id: request.id, message: error instanceof Error ? error.message : String(error) };
  }
  self.postMessage(response, transfer);
});

// Start loading the engine straight away, while the user is still choosing files.
// A failure here is surfaced by the first compare request, which retries.
engineReady().catch(() => undefined);
