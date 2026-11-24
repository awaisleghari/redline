/**
 * Checks on an uploaded file before anything is compared. Plain-English
 * refusals for the cases version 1 does not handle (see PLAN.md, "Out of scope"),
 * and the notice for uploads that contain tracked changes.
 */
import { unzip } from '../common/zip';
import { partHasTrackedChanges } from '../producer/accept-changes';

export const MAX_FILE_BYTES = 25 * 1024 * 1024;

/** Why a file was refused, in words for the user; undefined when it is acceptable. */
export async function refusalFor(file: { name: string; size: number; bytes: () => Promise<Uint8Array> }, role: 'original' | 'modified'): Promise<string | undefined> {
  const which = role === 'original' ? 'The original' : 'The modified document';
  const name = file.name.toLowerCase();
  if (name.endsWith('.doc')) return `${which} is an old-format Word file (.doc). Please open it in Word, save it as a Word Document (.docx), and upload that.`;
  if (!name.endsWith('.docx')) return `${which} is not a Word document. Only .docx files can be compared.`;
  if (file.size > MAX_FILE_BYTES) return `${which} is larger than 25 MB, which is the limit for a comparison.`;
  if (file.size === 0) return `${which} is empty.`;

  let parts: Map<string, Uint8Array>;
  try {
    parts = await unzip(await file.bytes());
  } catch {
    return `${which} could not be read as a Word document. It may be damaged, or not really a .docx file.`;
  }
  if (!parts.has('[Content_Types].xml') || ![...parts.keys()].some((part) => /^word\/document\d*\.xml$/.test(part))) {
    return `${which} could not be read as a Word document. It may be damaged, or not really a .docx file.`;
  }
  return undefined;
}

/**
 * Whether an upload already contains tracked changes of its own. Such a document
 * is compared as if they had all been accepted, as Word's own Compare does, and
 * the user is told so (owner's decision, 2026-11-24; before that it was refused).
 */
export async function containsTrackedChanges(docx: Uint8Array): Promise<boolean> {
  const decoder = new TextDecoder();
  for (const [part, bytes] of await unzip(docx)) {
    if (/^word\/.+\.xml$/.test(part) && hasTrackedChanges(decoder.decode(bytes))) return true;
  }
  return false;
}

/** What the page says when one or both uploads came with tracked changes. */
export const TRACKED_NOTICE = {
  original: 'The original already contained tracked changes. It was compared as if they had all been accepted.',
  modified: 'The modified document already contained tracked changes. It was compared as if they had all been accepted.',
  both: 'Both documents already contained tracked changes. They were compared as if those had all been accepted.',
} as const;

/** Whether one part of a Word file holds tracked-change markup of any kind (the Producer's own test, so the two cannot disagree). */
export function hasTrackedChanges(partXml: string): boolean {
  return partHasTrackedChanges(partXml);
}
