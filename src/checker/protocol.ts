/** Messages exchanged between the page and the Checker's Web Worker. */
import type { CheckReport } from './report';

export interface CheckRequest {
  type: 'check-tracked-changes';
  id: number;
  original: ArrayBuffer;
  modified: ArrayBuffer;
  output: ArrayBuffer;
}

export type CheckResponse = { type: 'checked'; id: number; report: CheckReport } | { type: 'check-crashed'; id: number; message: string };
