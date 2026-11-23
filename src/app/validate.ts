/**
 * Checks on an uploaded file before anything is compared. Plain-English
 * refusals for the cases version 1 does not handle (see PLAN.md, "Out of scope").
 */
import { unzip } from '../common/zip';

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
  const decoder = new TextDecoder();
  for (const [part, bytes] of parts) {
    if (!/^word\/[^/]+\.xml$/.test(part)) continue;
    if (hasTrackedChanges(decoder.decode(bytes))) {
      return `${which} already contains tracked changes. Please accept or reject the existing changes first, then compare.`;
    }
  }
  return undefined;
}

/**
 * Every kind of tracked change Word can store: inserted, deleted and moved
 * text; tracked changes to formatting, numbering, sections and tables; and
 * tracked table cells. Matched whatever namespace prefix the file uses.
 */
const TRACKED_CHANGE = new RegExp(
  '<(?:[A-Za-z_][\\w.-]*:)?(?:ins|del|moveFrom|moveTo|cellIns|cellDel|cellMerge|numberingChange|' +
    '(?:rPr|pPr|sectPr|tblPr|tblPrEx|tcPr|trPr|tblGrid)Change|' +
    'customXml(?:Ins|Del|MoveFrom|MoveTo)RangeStart)[\\s/>]',
);

export function hasTrackedChanges(partXml: string): boolean {
  return TRACKED_CHANGE.test(partXml);
}
