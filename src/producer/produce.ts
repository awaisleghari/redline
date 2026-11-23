/**
 * Producer: make the tracked-changes Word file. Runs the engine, then gives
 * each kind of change its own author name (see label-authors.ts).
 */
import { compareDocuments } from 'docxodus/core';
import { unzip, zip } from '../common/zip';
import { labelAuthors } from './label-authors';
import { REVISION_AUTHOR } from './protocol';

/** The engine must already be initialised. */
export async function produceTrackedChanges(original: Uint8Array, modified: Uint8Array): Promise<Uint8Array> {
  const result = await compareDocuments(original, modified, { authorName: REVISION_AUTHOR });
  const files = await unzip(result);
  const decoder = new TextDecoder();
  const encoder = new TextEncoder();
  for (const [name, bytes] of files) {
    if (/^word\/[^/]+\.xml$/.test(name)) files.set(name, encoder.encode(labelAuthors(decoder.decode(bytes))));
  }
  return zip(files);
}
