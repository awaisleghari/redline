/**
 * Producer: make the tracked-changes Word file. Prepares the uploads for the
 * engine (explicit-spaces.ts), runs it, then gives
 * each kind of change its own author name (see label-authors.ts).
 */
import { compareDocuments } from 'docxodus/core';
import { unzip, zip } from '../common/zip';
import { acceptTrackedChanges } from './accept-changes';
import { freshRelationshipIds } from './fresh-relationship-ids';
import { makeSpacesExplicit } from './explicit-spaces';
import { dropUnusedHeadersAndFooters } from './drop-unused-headers';
import { labelAuthors } from './label-authors';
import { repairNotes } from './repair-notes';
import { restorePictures } from './restore-pictures';
import { unwrapRowControls } from './unwrap-row-controls';
import { repairParagraphProperties } from './repair-paragraphs';
import { refineChanges } from './refine-changes';
import { repairSeparators } from './repair-separators';
import { repairSectionReferences } from './repair-section-references';
import { REVISION_AUTHOR } from './protocol';

/** The engine must already be initialised. */
export async function produceTrackedChanges(original: Uint8Array, modified: Uint8Array): Promise<Uint8Array> {
  // The engine loses a document-wide "keep white space" setting; write it onto each piece of text first (explicit-spaces.ts).
  // Tracked changes already in an upload count as accepted (accept-changes.ts).
  // ...and table rows wrapped in content controls are unwrapped, which the engine cannot handle (unwrap-row-controls.ts).
  // ...and internal reference names of the form the engine itself writes are replaced, or the engine's own collide with them (fresh-relationship-ids.ts).
  const prepare = async (upload: Uint8Array): Promise<Uint8Array> => freshRelationshipIds(await makeSpacesExplicit(await unwrapRowControls(await acceptTrackedChanges(upload))));
  const preparedModified = await prepare(modified);
  const preparedOriginal = await prepare(original);
  const result = await compareDocuments(preparedOriginal, preparedModified, { authorName: REVISION_AUTHOR });
  const files = await unzip(result);
  const decoder = new TextDecoder();
  const encoder = new TextEncoder();
  // The engine's output, part by part: first repaired where the engine is known to go wrong, then relabelled.
  const parts = new Map<string, string>();
  for (const [name, bytes] of files) if (/^word\/[^/]+\.xml$/.test(name)) parts.set(name, decoder.decode(bytes));
  for (const [name, xml] of repairNotes(parts)) parts.set(name, xml);
  const modifiedFiles = await unzip(preparedModified);
  const originalFiles = await unzip(preparedOriginal);
  // Pictures stored as text come out rewritten; they are given back the uploaded bytes (restore-pictures.ts).
  restorePictures(files, [originalFiles, modifiedFiles]);
  const modifiedMain = modifiedFiles.get('word/document.xml');
  const trackedMain = parts.get('word/document.xml');
  if (modifiedMain && trackedMain) {
    const repaired = repairParagraphProperties(trackedMain, decoder.decode(modifiedMain));
    if (repaired) parts.set('word/document.xml', repaired);
  }
  // A change of wording in a footnote or endnote separator, which the engine does not compare, is shown (repair-separators.ts).
  for (const name of ['word/footnotes.xml', 'word/endnotes.xml']) {
    const [mine, before, after] = [parts.get(name), originalFiles.get(name), modifiedFiles.get(name)];
    if (!mine || !before || !after) continue;
    const date = /\bw:date="([^"]+)"/.exec(parts.get('word/document.xml') ?? '')?.[1] ?? new Date().toISOString().replace(/\.\d+Z$/, 'Z');
    const repaired = repairSeparators(mine, decoder.decode(before), decoder.decode(after), date);
    if (repaired) parts.set(name, repaired);
  }
  // Where a section break is shown removed and added, the engine names the wrong headers for later sections (repair-section-references.ts).
  const relsOf = (from: Map<string, Uint8Array>): string => decoder.decode(from.get('word/_rels/document.xml.rels') ?? new Uint8Array());
  const mainNow = parts.get('word/document.xml');
  const originalMain = originalFiles.get('word/document.xml');
  if (mainNow && modifiedMain && originalMain) {
    const repaired = repairSectionReferences(mainNow, relsOf(files), { main: decoder.decode(originalMain), rels: relsOf(originalFiles) }, { main: decoder.decode(modifiedMain), rels: relsOf(modifiedFiles) });
    if (repaired) parts.set('word/document.xml', repaired);
  }
  // Text nobody changed is not shown removed and typed again: over-marked stretches are compared word by word (refine-changes.ts).
  for (const [name, xml] of parts) {
    if (!/^word\/(document|header\d*|footer\d*|footnotes|endnotes)\.xml$/.test(name)) continue;
    const refined = refineChanges(xml);
    if (refined) parts.set(name, refined);
  }
  for (const [name, xml] of parts) files.set(name, encoder.encode(labelAuthors(xml)));
  dropUnusedHeadersAndFooters(files);
  return zip(files);
}
