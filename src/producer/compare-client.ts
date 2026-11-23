/**
 * Page-side handle on the compare Web Worker: send two .docx files, get back
 * the tracked-changes .docx.
 */
import type { CompareRequest, CompareResponse } from './protocol';

export interface CompareResult {
  /** The tracked-changes Word file. */
  trackedChanges: Uint8Array;
  /** Time the engine itself spent comparing, in milliseconds. */
  engineMs: number;
}

interface Pending {
  resolve: (result: CompareResult) => void;
  reject: (error: Error) => void;
}

export class CompareClient {
  private readonly worker: Worker;
  private readonly pending = new Map<number, Pending>();
  private nextId = 1;

  constructor() {
    this.worker = new Worker(new URL('./compare.worker.ts', import.meta.url), { type: 'module', name: 'redline-compare' });
    this.worker.onmessage = (event: MessageEvent<CompareResponse>) => {
      const response = event.data;
      const waiting = this.pending.get(response.id);
      if (!waiting) return;
      this.pending.delete(response.id);
      if (response.type === 'compared') {
        waiting.resolve({ trackedChanges: new Uint8Array(response.trackedChanges), engineMs: response.engineMs });
      } else {
        waiting.reject(new Error(response.message));
      }
    };
    this.worker.onerror = (event) => {
      const error = new Error(event.message || 'The compare engine stopped unexpectedly.');
      for (const waiting of this.pending.values()) waiting.reject(error);
      this.pending.clear();
    };
  }

  async compare(original: Blob, modified: Blob): Promise<CompareResult> {
    const [originalBytes, modifiedBytes] = await Promise.all([original.arrayBuffer(), modified.arrayBuffer()]);
    const id = this.nextId++;
    const request: CompareRequest = { type: 'compare', id, original: originalBytes, modified: modifiedBytes };
    return new Promise<CompareResult>((resolve, reject) => {
      this.pending.set(id, { resolve, reject });
      this.worker.postMessage(request, [originalBytes, modifiedBytes]);
    });
  }

  terminate(): void {
    this.worker.terminate();
  }
}
