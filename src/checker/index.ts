/**
 * The Checker: an independent reviewer of every output.
 *
 * It is given the two uploads and one output, re-reads all three from their
 * raw bytes, and decides whether the user may have the output. It trusts
 * nothing the Producer says, and imports nothing from `src/producer/`.
 *
 * It fails closed: anything unexpected is a failed check, never a pass.
 */
import { compareStories } from './compare-text';
import { completenessFailure, coverageNote, unreadText } from './completeness';
import { NS } from '../common/namespaces';
import type { XmlElement } from './xml';
import { DELETED_COLOUR, INSERTED_COLOUR, usesDeletionLook, usesInsertionLook, type ChangeColours } from './formatting';
import { DELETED_LEGEND, INSERTED_LEGEND, countStaticChanges, coverPageFailure, formatMarksFailure, formattingFailure, leftoverTrackedChanges, numberingFailure, oldNumberFailure, otherFormsFailure } from './detail-checks';
import { DocxPackage, PackageError } from './package';
import { PdfError, readPdf } from './pdf';
import { comparePdf, type PdfCheckResult } from './pdf-check';
import { CHECK_TITLES, CHECKS_FOR, type CheckFailure, type CheckId, type CheckReport, type CheckResult, type OutputKind } from './report';
import { readStories, type StoryReading } from './stories';
import { structureProblems } from './structure';

export type { CheckFailure, CheckId, CheckReport, CheckResult, OutputKind } from './report';
export type { StaticCounts } from './detail-checks';
export { CHECK_TITLES, CHECKS_FOR, OUTPUT_NAME, reportAsText } from './report';

export interface CheckInput {
  /** The original upload, exactly as the user provided it. */
  original: Uint8Array;
  /** The modified upload, exactly as the user provided it. */
  modified: Uint8Array;
  /** The tracked-changes Word file to be checked. */
  output: Uint8Array;
}

const pass = (id: CheckId, note?: string): CheckResult => ({ id, title: CHECK_TITLES[id], passed: true, note });
const fail = (id: CheckId, failure: CheckFailure): CheckResult => ({ id, title: CHECK_TITLES[id], passed: false, failure });

function describe(error: unknown): string {
  if (error instanceof PackageError) return error.message;
  return `The check could not be completed (${error instanceof Error ? error.message : String(error)}).`;
}

/** Run a check; any exception is a failure of that check. */
function guarded(id: CheckId, run: () => CheckResult): CheckResult {
  try {
    return run();
  } catch (error) {
    return fail(id, { message: describe(error) });
  }
}

/** Hyperlinks must cover the same words and point to the same place as in the document being reproduced. */
/**
 * Handwriting. An output may hold the picture of a pen mark in place of the strokes (the picture is compared with
 * the text, as every picture is). But strokes it does hold must be the document's own, all of them and unchanged.
 */
function compareInks(expected: string[], found: string[], expectedName: string): CheckFailure | undefined {
  if (found.length === 0 || (expected.length === found.length && expected.every((ink, i) => ink === found[i]))) return undefined;
  return { message: `Handwriting (pen strokes) differs from the ${expectedName}.`, expected: `${expected.length} pen mark(s), unchanged`, found: `${found.length} pen mark(s), not the same` };
}

/** A legacy shape drawn from a template must find its template in the output wherever it did in the document being reproduced. */
function compareShapeTypes(expected: string[], found: string[], expectedName: string): CheckFailure | undefined {
  const lost = found.filter((id) => !expected.includes(id));
  if (lost.length === 0) return undefined;
  return { message: `A drawing has lost the template it is drawn from, which it has in the ${expectedName}, so it may not look the same.`, expected: 'every shape template present', found: `missing: ${lost.join(', ')}` };
}

function compareLinks(expected: string[], found: string[], expectedName: string): CheckFailure | undefined {
  for (let i = 0; i < Math.max(expected.length, found.length); i++) {
    if (expected[i] === found[i]) continue;
    return {
      message: `A hyperlink differs from the ${expectedName} (in the words it covers or in where it points).`,
      where: `hyperlink ${i + 1}`,
      expected: expected[i] ?? '(no further hyperlinks)',
      found: found[i] ?? '(no further hyperlinks)',
    };
  }
  return undefined;
}

interface Opened {
  original: DocxPackage;
  modified: DocxPackage;
  output: DocxPackage;
}

/** Open the three files, or produce the report that says they could not be opened. */
async function open(input: CheckInput, kind: OutputKind): Promise<Opened | CheckReport> {
  const failAll = (message: string): CheckReport => ({ output: kind, passed: false, checks: CHECKS_FOR[kind].map((id) => fail(id, { message })) });
  let original: DocxPackage;
  let modified: DocxPackage;
  try {
    original = await DocxPackage.open(input.original);
    modified = await DocxPackage.open(input.modified);
  } catch (error) {
    return failAll(`An uploaded file could not be read, so nothing could be checked. ${describe(error)}`);
  }
  try {
    return { original, modified, output: await DocxPackage.open(input.output) };
  } catch (error) {
    return failAll(describe(error));
  }
}

const finish = (kind: OutputKind, checks: CheckResult[]): CheckReport => ({ output: kind, passed: checks.length === CHECKS_FOR[kind].length && checks.every((c) => c.passed), checks });
const verdict = (id: CheckId, failure: CheckFailure | undefined, note?: string): CheckResult => (failure ? fail(id, failure) : pass(id, note));
const firstOf = (problems: string[]): CheckFailure | undefined => (problems.length > 0 ? { message: problems[0]! + (problems.length > 1 ? ` (and ${problems.length - 1} more)` : '') } : undefined);

/** Check a tracked-changes Word file (checks 1, 2, 6, 7 and 8 of PLAN.md, and clause numbers). */
export async function checkTrackedChanges(input: CheckInput): Promise<CheckReport> {
  const opened = await open(input, 'tracked-changes');
  if ('checks' in opened) return opened;
  const { original, modified, output } = opened;

  // Each reading is taken lazily and at most once; a reading that throws fails every check that needs it.
  const cache = new Map<string, StoryReading>();
  const reading = (name: string, make: () => StoryReading): StoryReading => {
    let value = cache.get(name);
    if (!value) cache.set(name, (value = make()));
    return value;
  };
  // Uploads are read as Word's own Compare reads them: any tracked changes already in them count as accepted.
  const originalText = (): StoryReading => reading('original', () => readStories(original, 'accepted'));
  const modifiedText = (): StoryReading => reading('modified', () => readStories(modified, 'accepted'));
  const accepted = (): StoryReading => reading('accepted', () => readStories(output, 'accepted'));
  const rejected = (): StoryReading => reading('rejected', () => readStories(output, 'rejected'));

  const checks: CheckResult[] = [
    guarded('accept-all', () => {
      const failure = compareStories(modifiedText().stories, accepted().stories, 'modified document') ?? compareLinks(modifiedText().links, accepted().links, 'modified document') ?? compareInks(modifiedText().inks, accepted().inks, 'modified document') ?? compareShapeTypes(modifiedText().missingShapeTypes, accepted().missingShapeTypes, 'modified document');
      return failure ? fail('accept-all', failure) : pass('accept-all', hadTrackedChanges(modified) ? ACCEPTED_READING.modified : undefined);
    }),
    guarded('reject-all', () => {
      const failure = compareStories(originalText().stories, rejected().stories, 'original document') ?? compareLinks(originalText().links, rejected().links, 'original document') ?? compareInks(originalText().inks, rejected().inks, 'original document') ?? compareShapeTypes(originalText().missingShapeTypes, rejected().missingShapeTypes, 'original document');
      return failure ? fail('reject-all', failure) : pass('reject-all', hadTrackedChanges(original) ? ACCEPTED_READING.original : undefined);
    }),
    guarded('formatting', () => verdict('formatting', formattingFailure(modifiedText().stories, accepted().stories, false))),
    guarded('numbering', () => verdict('numbering', numberingFailure(modifiedText().stories, accepted().stories) ?? otherFormsFailure(modifiedText().otherForms, accepted().otherForms) ?? otherFormsFailure(originalText().otherForms, rejected().otherForms, 'original'))),
    guarded('completeness', () => {
      const problems = [
        ...unreadText(original, [originalText()], 'original document', deletedSide(original), true),
        ...unreadText(modified, [modifiedText()], 'modified document', deletedSide(modified), true),
        ...unreadText(output, [accepted(), rejected()], 'output'),
      ];
      const failure = completenessFailure(problems);
      return failure ? fail('completeness', failure) : pass('completeness', coverageNote(accepted().stories));
    }),
    guarded('structure', () => verdict('structure', firstOf(structureProblems(output)))),
  ];

  return finish('tracked-changes', checks);
}

export interface PdfCheckInput {
  /** The static redline the PDF was made from: the very file the Checker has already passed. */
  staticRedline: Uint8Array;
  /** The PDF to be checked. */
  pdf: Uint8Array;
}

/** Check a PDF against the static redline it was made from (check 9 of PLAN.md). */
export async function checkPdf(input: PdfCheckInput): Promise<CheckReport> {
  const failBoth = (message: string): CheckReport => ({ output: 'pdf', passed: false, checks: CHECKS_FOR.pdf.map((id) => fail(id, { message })) });
  let result: PdfCheckResult;
  try {
    const redline = await DocxPackage.open(input.staticRedline);
    const pdf = await readPdf(input.pdf);
    // Which colour this redline shows insertions in is read from its own legend (the redline itself was checked against the uploads on that point).
    const standard = readStories(redline, 'all', 'static');
    const legendHas = (line: string): boolean => standard.stories.some((story) => story.key === 'body' && story.paragraphs.slice(0, 30).includes(line));
    const colours: ChangeColours = {
      inserted: legendHas(INSERTED_LEGEND.alternative) ? INSERTED_COLOUR.alternative : INSERTED_COLOUR.standard,
      deleted: legendHas(DELETED_LEGEND.alternative) ? DELETED_COLOUR.alternative : DELETED_COLOUR.standard,
    };
    result = comparePdf(colours.inserted !== INSERTED_COLOUR.standard || colours.deleted !== DELETED_COLOUR.standard ? readStories(redline, 'all', 'static', colours) : standard, pdf);
  } catch (error) {
    return failBoth(error instanceof PdfError ? error.message : describe(error));
  }
  const checks: CheckResult[] = [
    verdict('pdf-text', result.text, result.note),
    // Colours can only be vouched for where the text itself was found.
    verdict('pdf-colours', result.colours ?? (result.text ? { message: 'This could not be checked because the text does not match (see the text check).' } : undefined)),
  ];
  return finish('pdf', checks);
}

export interface StaticCheckInput extends CheckInput {
  /** The file names and comparison time the cover page must show, as given to the user. They come from the page, not from the Producer. */
  originalName: string;
  modifiedName: string;
  comparedAt: string;
}

/** For an upload with tracked changes of its own: its reading with those changes rejected, used only to see which notes deleted text refers to. */
const deletedSide = (upload: DocxPackage): StoryReading[] => (hadTrackedChanges(upload) ? [readStories(upload, 'rejected')] : []);

/**
 * Whether a document's own text (main text, headers, footers, notes) carries the standard look of a change:
 * blue with a double underline (an insertion), or red and struck through (a deletion).
 */
function usesLookIn(upload: DocxPackage, usesLook: (root: XmlElement) => boolean): boolean {
  // Headers and footers no section refers to are not document content, and are not carried into the redline.
  const main = upload.mainDocumentPart();
  const referenced = new Set([...upload.relationshipsOfType(main, 'header'), ...upload.relationshipsOfType(main, 'footer')].filter((r) => referencedIn(upload.xml(main), r.id)).map((r) => r.partName));
  for (const partName of upload.partNames) {
    if (!partName.endsWith('.xml') || partName === '[Content_Types].xml') continue;
    const root = upload.xml(partName);
    if (root.ns !== NS.w || !['document', 'hdr', 'ftr', 'footnotes', 'endnotes', 'comments'].includes(root.name)) continue;
    if ((root.name === 'hdr' || root.name === 'ftr') && !referenced.has(partName)) continue;
    // The upload is compared as if its own tracked changes had been accepted: text they delete does not count.
    if (usesLook(withoutDeleted(root))) return true;
  }
  return false;
}

function referencedIn(root: XmlElement, id: string): boolean {
  if (root.ns === NS.w && (root.name === 'headerReference' || root.name === 'footerReference') && [...root.attributes].some(([key, value]) => key.startsWith(`${NS.r}|`) && value === id)) return true;
  return root.children.some((c) => referencedIn(c, id));
}

/** A copy without what the document's own tracked changes delete: deleted and moved-away text, and paragraph marks flagged as deleted. */
function withoutDeleted(element: XmlElement): XmlElement {
  const kept: XmlElement[] = [];
  for (const c of element.children) {
    if (c.ns === NS.w && (c.name === 'del' || c.name === 'moveFrom') && !(element.ns === NS.w && ['rPr', 'trPr', 'tcPr', 'numPr', 'tblPrEx'].includes(element.name))) continue;
    if (c.ns === NS.w && c.name === 'pPr') {
      const mark = c.children.find((x) => x.ns === NS.w && x.name === 'rPr');
      if (mark && mark.children.some((x) => x.ns === NS.w && (x.name === 'del' || x.name === 'moveFrom'))) {
        kept.push({ ...c, children: c.children.filter((x) => x !== mark) });
        continue;
      }
    }
    kept.push(withoutDeleted(c));
  }
  return { ...element, children: kept };
}

/** Whether an upload holds tracked changes of its own (it is then read as if they had all been accepted). */
const hadTrackedChanges = (upload: DocxPackage): boolean => leftoverTrackedChanges(upload).length > 0;
const ACCEPTED_READING = {
  original: 'The original contained tracked changes of its own; it was read as if they had all been accepted.',
  modified: 'The modified document contained tracked changes of its own; it was read as if they had all been accepted.',
} as const;

/**
 * Check a static redline Word file (checks 3, 4, 5, 6, 7 and 8 of PLAN.md,
 * the marking of formatting changes, and clause numbers). Changes are
 * recognised by their formatting alone, exactly as a reader would see them.
 */
export async function checkStaticRedline(input: StaticCheckInput): Promise<CheckReport> {
  const opened = await open(input, 'static-redline');
  if ('checks' in opened) return opened;
  const { original, modified, output } = opened;

  const cache = new Map<string, StoryReading>();
  const reading = (name: string, make: () => StoryReading): StoryReading => {
    let value = cache.get(name);
    if (!value) cache.set(name, (value = make()));
    return value;
  };
  const originalText = (): StoryReading => reading('original', () => readStories(original, 'accepted'));
  const modifiedText = (): StoryReading => reading('modified', () => readStories(modified, 'accepted'));
  // Whether insertions must be shown in the alternative colour is the Checker's own finding, from the uploads themselves.
  const alternativeInsertionColour = usesLookIn(original, usesInsertionLook) || usesLookIn(modified, usesInsertionLook);
  const alternativeDeletionColour = usesLookIn(original, usesDeletionLook) || usesLookIn(modified, usesDeletionLook);
  const colours: ChangeColours = {
    inserted: alternativeInsertionColour ? INSERTED_COLOUR.alternative : INSERTED_COLOUR.standard,
    deleted: alternativeDeletionColour ? DELETED_COLOUR.alternative : DELETED_COLOUR.standard,
  };
  const newSide = (): StoryReading => reading('new', () => readStories(output, 'accepted', 'static', colours));
  const oldSide = (): StoryReading => reading('old', () => readStories(output, 'rejected', 'static', colours));

  const checks: CheckResult[] = [
    guarded('static-new', () => verdict('static-new', compareStories(modifiedText().stories, newSide().stories, 'modified document') ?? compareLinks(modifiedText().links, newSide().links, 'modified document') ?? compareInks(modifiedText().inks, newSide().inks, 'modified document') ?? compareShapeTypes(modifiedText().missingShapeTypes, newSide().missingShapeTypes, 'modified document'), hadTrackedChanges(modified) ? ACCEPTED_READING.modified : undefined)),
    guarded('static-old', () => verdict('static-old', compareStories(originalText().stories, oldSide().stories, 'original document') ?? compareLinks(originalText().links, oldSide().links, 'original document') ?? compareInks(originalText().inks, oldSide().inks, 'original document') ?? compareShapeTypes(originalText().missingShapeTypes, oldSide().missingShapeTypes, 'original document'), hadTrackedChanges(original) ? ACCEPTED_READING.original : undefined)),
    guarded('no-tracked-changes', () => verdict('no-tracked-changes', firstOf(leftoverTrackedChanges(output)))),
    guarded('cover-page', () => {
      const counts = countStaticChanges(output, colours);
      const note = `Counted independently: ${counts.insertions} insertions, ${counts.deletions} deletions, ${counts.moves} moves, ${counts.formattingChanges} formatting changes.`;
      // Whether an upload came with tracked changes of its own is the Checker's own finding, from the upload itself.
      const inOriginal = hadTrackedChanges(original);
      const inModified = hadTrackedChanges(modified);
      const trackedChangesIn = inOriginal && inModified ? 'both' : inOriginal ? 'original' : inModified ? 'modified' : undefined;
      const expectation = { originalName: input.originalName, modifiedName: input.modifiedName, comparedAt: input.comparedAt, trackedChangesIn, alternativeInsertionColour, alternativeDeletionColour } as const;
      return verdict('cover-page', coverPageFailure(newSide().cover, newSide().coverLooks, expectation, counts), note);
    }),
    guarded('formatting', () => verdict('formatting', formattingFailure(modifiedText().stories, newSide().stories, true))),
    guarded('format-marks', () => verdict('format-marks', formatMarksFailure(originalText().stories, oldSide().stories, newSide().stories))),
    guarded('numbering', () => verdict('numbering', numberingFailure(modifiedText().stories, newSide().stories) ?? otherFormsFailure(modifiedText().otherForms, newSide().otherForms) ?? otherFormsFailure(originalText().otherForms, oldSide().otherForms, 'original') ?? oldNumberFailure(originalText().stories, oldSide().stories))),
    guarded('completeness', () => {
      const problems = [
        ...unreadText(original, [originalText()], 'original document', deletedSide(original), true),
        ...unreadText(modified, [modifiedText()], 'modified document', deletedSide(modified), true),
        ...unreadText(output, [newSide(), oldSide()], 'output'),
      ];
      const failure = completenessFailure(problems);
      return failure ? fail('completeness', failure) : pass('completeness', coverageNote(newSide().stories));
    }),
    guarded('structure', () => verdict('structure', firstOf(structureProblems(output)))),
  ];

  return finish('static-redline', checks);
}
