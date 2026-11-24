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
import { countStaticChanges, coverPageFailure, formatMarksFailure, formattingFailure, leftoverTrackedChanges, numberingFailure, oldNumberFailure } from './detail-checks';
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
      const failure = compareStories(modifiedText().stories, accepted().stories, 'modified document') ?? compareLinks(modifiedText().links, accepted().links, 'modified document');
      return failure ? fail('accept-all', failure) : pass('accept-all', hadTrackedChanges(modified) ? ACCEPTED_READING.modified : undefined);
    }),
    guarded('reject-all', () => {
      const failure = compareStories(originalText().stories, rejected().stories, 'original document') ?? compareLinks(originalText().links, rejected().links, 'original document');
      return failure ? fail('reject-all', failure) : pass('reject-all', hadTrackedChanges(original) ? ACCEPTED_READING.original : undefined);
    }),
    guarded('formatting', () => verdict('formatting', formattingFailure(modifiedText().stories, accepted().stories, false))),
    guarded('numbering', () => verdict('numbering', numberingFailure(modifiedText().stories, accepted().stories))),
    guarded('completeness', () => {
      const problems = [
        ...unreadText(original, [originalText()], 'original document', deletedSide(original)),
        ...unreadText(modified, [modifiedText()], 'modified document', deletedSide(modified)),
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
    result = comparePdf(readStories(redline, 'all', 'static'), pdf);
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
  const newSide = (): StoryReading => reading('new', () => readStories(output, 'accepted', 'static'));
  const oldSide = (): StoryReading => reading('old', () => readStories(output, 'rejected', 'static'));

  const checks: CheckResult[] = [
    guarded('static-new', () => verdict('static-new', compareStories(modifiedText().stories, newSide().stories, 'modified document') ?? compareLinks(modifiedText().links, newSide().links, 'modified document'), hadTrackedChanges(modified) ? ACCEPTED_READING.modified : undefined)),
    guarded('static-old', () => verdict('static-old', compareStories(originalText().stories, oldSide().stories, 'original document') ?? compareLinks(originalText().links, oldSide().links, 'original document'), hadTrackedChanges(original) ? ACCEPTED_READING.original : undefined)),
    guarded('no-tracked-changes', () => verdict('no-tracked-changes', firstOf(leftoverTrackedChanges(output)))),
    guarded('cover-page', () => {
      const counts = countStaticChanges(output);
      const note = `Counted independently: ${counts.insertions} insertions, ${counts.deletions} deletions, ${counts.moves} moves, ${counts.formattingChanges} formatting changes.`;
      // Whether an upload came with tracked changes of its own is the Checker's own finding, from the upload itself.
      const inOriginal = hadTrackedChanges(original);
      const inModified = hadTrackedChanges(modified);
      const trackedChangesIn = inOriginal && inModified ? 'both' : inOriginal ? 'original' : inModified ? 'modified' : undefined;
      const expectation = { originalName: input.originalName, modifiedName: input.modifiedName, comparedAt: input.comparedAt, trackedChangesIn } as const;
      return verdict('cover-page', coverPageFailure(newSide().cover, newSide().coverLooks, expectation, counts), note);
    }),
    guarded('formatting', () => verdict('formatting', formattingFailure(modifiedText().stories, newSide().stories, true))),
    guarded('format-marks', () => verdict('format-marks', formatMarksFailure(originalText().stories, oldSide().stories, newSide().stories))),
    guarded('numbering', () => verdict('numbering', numberingFailure(modifiedText().stories, newSide().stories) ?? oldNumberFailure(originalText().stories, oldSide().stories))),
    guarded('completeness', () => {
      const problems = [
        ...unreadText(original, [originalText()], 'original document', deletedSide(original)),
        ...unreadText(modified, [modifiedText()], 'modified document', deletedSide(modified)),
        ...unreadText(output, [newSide(), oldSide()], 'output'),
      ];
      const failure = completenessFailure(problems);
      return failure ? fail('completeness', failure) : pass('completeness', coverageNote(newSide().stories));
    }),
    guarded('structure', () => verdict('structure', firstOf(structureProblems(output)))),
  ];

  return finish('static-redline', checks);
}
