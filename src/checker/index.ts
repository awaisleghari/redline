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
import { DocxPackage, PackageError } from './package';
import { CHECK_TITLES, type CheckFailure, type CheckId, type CheckReport, type CheckResult } from './report';
import { readStories, type StoryReading } from './stories';
import { structureProblems } from './structure';

export type { CheckFailure, CheckId, CheckReport, CheckResult } from './report';
export { CHECK_TITLES, reportAsText } from './report';

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

/** Check a tracked-changes Word file (checks 1, 2, 6 and 7 of PLAN.md). */
export async function checkTrackedChanges(input: CheckInput): Promise<CheckReport> {
  const ids: CheckId[] = ['accept-all', 'reject-all', 'completeness', 'structure'];
  let original: DocxPackage;
  let modified: DocxPackage;
  let output: DocxPackage;
  try {
    original = await DocxPackage.open(input.original);
    modified = await DocxPackage.open(input.modified);
  } catch (error) {
    const failure = { message: `An uploaded file could not be read, so nothing could be checked. ${describe(error)}` };
    return { output: 'tracked-changes', passed: false, checks: ids.map((id) => fail(id, failure)) };
  }
  try {
    output = await DocxPackage.open(input.output);
  } catch (error) {
    const failure = { message: describe(error) };
    return { output: 'tracked-changes', passed: false, checks: ids.map((id) => fail(id, failure)) };
  }

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
      return failure ? fail('accept-all', failure) : pass('accept-all');
    }),
    guarded('reject-all', () => {
      const failure = compareStories(originalText().stories, rejected().stories, 'original document') ?? compareLinks(originalText().links, rejected().links, 'original document');
      return failure ? fail('reject-all', failure) : pass('reject-all');
    }),
    guarded('completeness', () => {
      const problems = [
        ...unreadText(original, [originalText()], 'original document'),
        ...unreadText(modified, [modifiedText()], 'modified document'),
        ...unreadText(output, [accepted(), rejected()], 'output'),
      ];
      const failure = completenessFailure(problems);
      return failure ? fail('completeness', failure) : pass('completeness', coverageNote(accepted().stories));
    }),
    guarded('structure', () => {
      const problems = structureProblems(output);
      return problems.length > 0
        ? fail('structure', { message: problems[0]! + (problems.length > 1 ? ` (and ${problems.length - 1} more)` : '') })
        : pass('structure');
    }),
  ];

  return { output: 'tracked-changes', passed: checks.every((c) => c.passed), checks };
}
