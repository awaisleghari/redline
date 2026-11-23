/** Messages exchanged between the page and the compare Web Worker. */

export interface CompareRequest {
  type: 'compare';
  /** Echoed back so replies can be matched to requests. */
  id: number;
  original: ArrayBuffer;
  modified: ArrayBuffer;
}

export type CompareResponse =
  | { type: 'compared'; id: number; trackedChanges: ArrayBuffer; engineMs: number }
  | { type: 'failed'; id: number; message: string };

/** Author name stamped on every tracked change the engine writes. */
export const REVISION_AUTHOR = 'Redline';
