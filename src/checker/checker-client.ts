/** Page-side handle on the Checker's Web Worker. */
import type { CheckInput, CheckReport } from './index';
import type { CheckRequest, CheckResponse } from './protocol';

interface Pending {
  resolve: (report: CheckReport) => void;
  reject: (error: Error) => void;
}

/** A private copy of the bytes in a buffer of its own, safe to hand to a worker. */
const copy = (bytes: Uint8Array): ArrayBuffer => bytes.slice().buffer;

export class CheckerClient {
  private readonly worker: Worker;
  private readonly pending = new Map<number, Pending>();
  private nextId = 1;

  constructor() {
    this.worker = new Worker(new URL('./checker.worker.ts', import.meta.url), { type: 'module', name: 'redline-checker' });
    this.worker.onmessage = (event: MessageEvent<CheckResponse>) => {
      const response = event.data;
      const waiting = this.pending.get(response.id);
      if (!waiting) return;
      this.pending.delete(response.id);
      if (response.type === 'checked') waiting.resolve(response.report);
      else waiting.reject(new Error(response.message));
    };
    this.worker.onerror = (event) => {
      const error = new Error(event.message || 'The checker stopped unexpectedly.');
      for (const waiting of this.pending.values()) waiting.reject(error);
      this.pending.clear();
    };
  }

  checkTrackedChanges(input: CheckInput): Promise<CheckReport> {
    const id = this.nextId++;
    const request: CheckRequest = {
      type: 'check-tracked-changes',
      id,
      original: copy(input.original),
      modified: copy(input.modified),
      output: copy(input.output),
    };
    return new Promise<CheckReport>((resolve, reject) => {
      this.pending.set(id, { resolve, reject });
      this.worker.postMessage(request, [request.original, request.modified, request.output]);
    });
  }

  terminate(): void {
    this.worker.terminate();
  }
}
