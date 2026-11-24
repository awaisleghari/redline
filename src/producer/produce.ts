/**
 * Producer: make the tracked-changes Word file. Prepares the uploads for the
 * engine (explicit-spaces.ts), runs it, then gives
 * each kind of change its own author name (see label-authors.ts).
 */
import { compareDocuments } from 'docxodus/core';
import { unzip, zip } from '../common/zip';
import { acceptTrackedChanges } from './accept-changes';
import { makeSpacesExplicit } from './explicit-spaces';
import { labelAuthors } from './label-authors';
import { REVISION_AUTHOR } from './protocol';

/** The engine must already be initialised. */
export async function produceTrackedChanges(original: Uint8Array, modified: Uint8Array): Promise<Uint8Array> {
  // The engine loses a document-wide "keep white space" setting; write it onto each piece of text first (explicit-spaces.ts).
  // Tracked changes already in an upload count as accepted (accept-changes.ts).
  const prepare = async (upload: Uint8Array): Promise<Uint8Array> => makeSpacesExplicit(await acceptTrackedChanges(upload));
  const result = await compareDocuments(await prepare(original), await prepare(modified), { authorName: REVISION_AUTHOR });
  const files = await unzip(result);
  const decoder = new TextDecoder();
  const encoder = new TextEncoder();
  for (const [name, bytes] of files) {
    if (/^word\/[^/]+\.xml$/.test(name)) files.set(name, encoder.encode(labelAuthors(decoder.decode(bytes))));
  }
  return zip(files);
}
