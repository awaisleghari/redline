/**
 * Shared, deliberately trivial helper: unzip a .docx into its parts.
 * Used by both the Producer and the Checker; it must contain no document logic.
 */
import JSZip from 'jszip';

/** Part name (no leading slash, e.g. "word/document.xml") to raw bytes. */
export type ZipEntries = Map<string, Uint8Array>;

export async function unzip(bytes: Uint8Array): Promise<ZipEntries> {
  const zip = await JSZip.loadAsync(bytes);
  const entries: ZipEntries = new Map();
  for (const [name, file] of Object.entries(zip.files)) {
    if (!file.dir) entries.set(name, await file.async('uint8array'));
  }
  return entries;
}

export async function zip(entries: ZipEntries): Promise<Uint8Array> {
  const out = new JSZip();
  for (const [name, data] of entries) out.file(name, data);
  return out.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
}
