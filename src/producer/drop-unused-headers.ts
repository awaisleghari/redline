/**
 * Producer: a header or footer that no section refers to is never shown by
 * Word (it is left behind in a file when a section is deleted). The engine
 * copies such leftovers into its output. The outputs must hold nothing that is
 * not part of the document, so they are taken out: the part itself, its
 * relationships, the link to it from the main text's relationships, and its
 * content-type entry.
 */
const decoder = new TextDecoder();
const encoder = new TextEncoder();

export function dropUnusedHeadersAndFooters(files: Map<string, Uint8Array>): void {
  const mainName = [...files.keys()].find((name) => /^word\/document\d*\.xml$/.test(name));
  const relsName = mainName && `word/_rels/${mainName.slice('word/'.length)}.rels`;
  if (!mainName || !relsName || !files.has(relsName)) return;
  const main = decoder.decode(files.get(mainName)!);
  const used = new Set([...main.matchAll(/<w:(?:header|footer)Reference\b[^>]*\br:id="([^"]+)"/g)].map((m) => m[1]!));
  let rels = decoder.decode(files.get(relsName)!);
  let types = files.has('[Content_Types].xml') ? decoder.decode(files.get('[Content_Types].xml')!) : '';
  let changed = false;
  for (const relationship of rels.match(/<Relationship\b[^>]*>/g) ?? []) {
    if (!/Type="[^"]*\/(header|footer)"/.test(relationship)) continue;
    const id = /\bId="([^"]+)"/.exec(relationship)?.[1];
    const target = /\bTarget="([^"]+)"/.exec(relationship)?.[1];
    if (!id || !target || used.has(id) || /TargetMode="External"/.test(relationship)) continue;
    const part = target.startsWith('/') ? target.slice(1) : `word/${target}`;
    const name = [...files.keys()].find((existing) => existing.toLowerCase() === part.toLowerCase());
    // Another relationship may point at the same part and be in use: then the part stays, only this link goes.
    const sharedInUse = (rels.match(/<Relationship\b[^>]*>/g) ?? []).some((other) => other !== relationship && /\bTarget="([^"]+)"/.exec(other)?.[1] === target && used.has(/\bId="([^"]+)"/.exec(other)?.[1] ?? ''));
    rels = rels.replace(relationship, '');
    changed = true;
    if (!name || sharedInUse) continue;
    files.delete(name);
    files.delete(name.replace(/([^/]+)$/, '_rels/$1.rels'));
    types = types.replace(new RegExp(`<Override\\b[^>]*PartName="/${name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}"[^>]*/>`, 'i'), '');
  }
  if (!changed) return;
  files.set(relsName, encoder.encode(rels));
  if (types) files.set('[Content_Types].xml', encoder.encode(types));
}
