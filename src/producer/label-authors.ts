/**
 * Producer: give insertions, deletions, moves and formatting changes separate
 * author names in the tracked-changes file.
 *
 * Word chooses the colour of a tracked change by its author (unless the user
 * has set fixed colours), so with a single author every change shows in one
 * colour. Separate author names make Word show each kind in its own colour.
 */
export const AUTHOR = {
  inserted: 'Redline (inserted)',
  deleted: 'Redline (deleted)',
  moved: 'Redline (moved)',
  formatting: 'Redline (formatting)',
} as const;

function authorFor(element: string): string | undefined {
  if (element === 'ins' || element === 'cellIns') return AUTHOR.inserted;
  if (element === 'del' || element === 'cellDel') return AUTHOR.deleted;
  if (/^move(From|To)(RangeStart)?$/.test(element)) return AUTHOR.moved;
  if (/Change$/.test(element) || element === 'cellMerge') return AUTHOR.formatting;
  return undefined;
}

/** Rewrite the author of every tracked change in one part's XML. Nothing else is touched. */
export function labelAuthors(xml: string): string {
  return xml.replace(/<w:([A-Za-z]+)\b([^<>]*?)\bw:author="[^"]*"/g, (whole, element: string, before: string) => {
    const author = authorFor(element);
    return author ? `<w:${element}${before}w:author="${author}"` : whole;
  });
}
