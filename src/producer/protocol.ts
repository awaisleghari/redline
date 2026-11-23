/** Messages exchanged between the page and the Producer's Web Worker. */

export interface CompareRequest {
  type: 'compare';
  /** Echoed back so replies can be matched to requests. */
  id: number;
  original: ArrayBuffer;
  modified: ArrayBuffer;
}

export interface StaticRequest {
  type: 'make-static';
  id: number;
  trackedChanges: ArrayBuffer;
  original: ArrayBuffer;
  originalName: string;
  modifiedName: string;
  comparedAt: string;
}

export interface RenderRequest {
  type: 'render-html';
  id: number;
  docx: ArrayBuffer;
}

export type ProducerRequest = CompareRequest | StaticRequest | RenderRequest;

export type ProducerResponse =
  | { type: 'compared'; id: number; trackedChanges: ArrayBuffer; engineMs: number }
  | { type: 'static-made'; id: number; staticRedline: ArrayBuffer; counts: { insertions: number; deletions: number; moves: number; formattingChanges: number } }
  | { type: 'rendered'; id: number; html: string }
  | { type: 'failed'; id: number; message: string };

/** Author name given to the engine; each kind of change is then relabelled (see label-authors.ts). */
export const REVISION_AUTHOR = 'Redline';
