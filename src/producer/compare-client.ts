/**
 * Page-side handle on the Producer's Web Worker: make the tracked-changes file
 * from two uploads, and the static redline from the tracked-changes file.
 */
import type { ProducerRequest, ProducerResponse } from './protocol';

export interface CompareResult {
  /** The tracked-changes Word file. */
  trackedChanges: Uint8Array;
  /** Time the engine itself spent comparing, in milliseconds. */
  engineMs: number;
}

export interface StaticResult {
  staticRedline: Uint8Array;
  counts: { insertions: number; deletions: number; moves: number; formattingChanges: number };
}

export interface StaticDetails {
  originalName: string;
  modifiedName: string;
  comparedAt: string;
}

interface Pending {
  resolve: (response: ProducerResponse) => void;
  reject: (error: Error) => void;
}

/** A private copy of the bytes in a buffer of its own, safe to hand to the worker. */
const copy = (bytes: Uint8Array): ArrayBuffer => bytes.slice().buffer;

export class CompareClient {
  private readonly worker: Worker;
  private readonly pending = new Map<number, Pending>();
  private nextId = 1;

  constructor() {
    this.worker = new Worker(new URL('./compare.worker.ts', import.meta.url), { type: 'module', name: 'redline-compare' });
    this.worker.onmessage = (event: MessageEvent<ProducerResponse>) => {
      const response = event.data;
      const waiting = this.pending.get(response.id);
      if (!waiting) return;
      this.pending.delete(response.id);
      if (response.type === 'failed') waiting.reject(new Error(response.message));
      else waiting.resolve(response);
    };
    this.worker.onerror = (event) => {
      const error = new Error(event.message || 'The compare engine stopped unexpectedly.');
      for (const waiting of this.pending.values()) waiting.reject(error);
      this.pending.clear();
    };
  }

  private send(build: (id: number) => ProducerRequest, transfer: (request: ProducerRequest) => Transferable[]): Promise<ProducerResponse> {
    const request = build(this.nextId++);
    return new Promise<ProducerResponse>((resolve, reject) => {
      this.pending.set(request.id, { resolve, reject });
      this.worker.postMessage(request, transfer(request));
    });
  }

  /** The worker receives its own copies; the caller's bytes stay usable (the Checker needs them). */
  async compare(original: Uint8Array, modified: Uint8Array): Promise<CompareResult> {
    const response = await this.send(
      (id) => ({ type: 'compare', id, original: copy(original), modified: copy(modified) }),
      (request) => (request.type === 'compare' ? [request.original, request.modified] : []),
    );
    if (response.type !== 'compared') throw new Error('The compare engine gave an unexpected reply.');
    return { trackedChanges: new Uint8Array(response.trackedChanges), engineMs: response.engineMs };
  }

  async makeStatic(trackedChanges: Uint8Array, original: Uint8Array, details: StaticDetails): Promise<StaticResult> {
    const response = await this.send(
      (id) => ({ type: 'make-static', id, trackedChanges: copy(trackedChanges), original: copy(original), ...details }),
      (request) => (request.type === 'make-static' ? [request.trackedChanges, request.original] : []),
    );
    if (response.type !== 'static-made') throw new Error('The compare engine gave an unexpected reply.');
    return { staticRedline: new Uint8Array(response.staticRedline), counts: response.counts };
  }

  /** The document as a web page, for the in-page viewer. */
  async renderHtml(docx: Uint8Array): Promise<string> {
    const response = await this.send(
      (id) => ({ type: 'render-html', id, docx: copy(docx) }),
      (request) => (request.type === 'render-html' ? [request.docx] : []),
    );
    if (response.type !== 'rendered') throw new Error('The viewer gave an unexpected reply.');
    return response.html;
  }

  terminate(): void {
    this.worker.terminate();
  }
}
