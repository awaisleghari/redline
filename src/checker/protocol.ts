/** Messages exchanged between the page and the Checker's Web Worker. */
import type { CheckReport } from './report';

export interface CheckRequest {
  type: 'check-tracked-changes' | 'check-static-redline';
  id: number;
  original: ArrayBuffer;
  modified: ArrayBuffer;
  output: ArrayBuffer;
  /** For the static redline: what the cover page must say. Supplied by the page, not by the Producer. */
  originalName?: string;
  modifiedName?: string;
  comparedAt?: string;
}

export type CheckResponse = { type: 'checked'; id: number; report: CheckReport } | { type: 'check-crashed'; id: number; message: string };
