/**
 * The compare flow: produce each output, have the independent Checker review
 * it, and release it only if every one of its checks passes.
 *
 * This is the one place where an output can reach the user, and it has no path
 * that does so without a passing report for that output.
 */
import { CHECK_TITLES, CHECKS_FOR, type CheckReport, type OutputKind } from '../checker';

export interface ChangeCounts {
  insertions: number;
  deletions: number;
  moves: number;
  formattingChanges: number;
}

export interface ComparisonDetails {
  /** File names as shown to the user. */
  originalName: string;
  modifiedName: string;
  /** Date and time of the comparison, formatted for display. */
  comparedAt: string;
}

type Files = { original: Uint8Array; modified: Uint8Array; output: Uint8Array };

export interface FlowDependencies {
  /** The Producer: make the tracked-changes file. */
  produce: (original: Uint8Array, modified: Uint8Array) => Promise<{ trackedChanges: Uint8Array; engineMs: number }>;
  /** The Producer: make the static redline from the tracked-changes file. */
  makeStatic: (trackedChanges: Uint8Array, original: Uint8Array, details: ComparisonDetails) => Promise<{ staticRedline: Uint8Array; counts: ChangeCounts }>;
  /** The Checker: review each output against the two uploads. */
  checkTracked: (input: Files) => Promise<CheckReport>;
  checkStatic: (input: Files & ComparisonDetails) => Promise<CheckReport>;
}

export interface Reviewed {
  report: CheckReport;
  /** The output file. Present only when every check passed. */
  file?: Uint8Array;
}

export interface FlowResult {
  tracked: Reviewed;
  staticRedline: Reviewed;
  /** The Producer's change counts. Present only when the static redline was released (the Checker has then confirmed them). */
  counts?: ChangeCounts;
  engineMs: number;
}

/** Thrown when the Producer cannot compare the documents at all; there is then nothing to check. */
export class ProducerError extends Error {}

/** A report saying an output could not be produced or checked. The output is withheld, exactly as for a failed check. */
export function checkerUnavailable(kind: OutputKind, message: string): CheckReport {
  return {
    output: kind,
    passed: false,
    checks: CHECKS_FOR[kind].map((id) => ({ id, title: CHECK_TITLES[id], passed: false, failure: { message } })),
  };
}

const reason = (error: unknown): string => (error instanceof Error ? error.message : String(error));

/** Release a file only on an explicit, complete pass of every check its kind requires. */
function review(kind: OutputKind, file: Uint8Array, report: CheckReport): Reviewed {
  const required = CHECKS_FOR[kind];
  const complete = report.output === kind && required.every((id) => report.checks.some((c) => c.id === id && c.passed === true));
  const released = report.passed === true && complete && report.checks.every((c) => c.passed === true);
  return released ? { report, file } : { report: { ...report, passed: false } };
}

export interface PdfDependencies {
  /** The one server step: send the static redline for conversion and receive the PDF. */
  convert: (staticRedline: Uint8Array) => Promise<Uint8Array>;
  /** The Checker: compare the PDF with the static redline it was made from. */
  check: (input: { staticRedline: Uint8Array; pdf: Uint8Array }) => Promise<CheckReport>;
}

/**
 * Make the PDF of a static redline that the Checker has already passed, have
 * the Checker compare the PDF with it, and release the PDF only on a pass.
 */
export async function runPdf(staticRedline: Uint8Array, deps: PdfDependencies): Promise<Reviewed> {
  let pdf: Uint8Array;
  try {
    pdf = await deps.convert(staticRedline);
  } catch (error) {
    return { report: checkerUnavailable('pdf', `The PDF could not be made. (${reason(error)})`) };
  }
  let report: CheckReport;
  try {
    report = await deps.check({ staticRedline, pdf });
  } catch (error) {
    report = checkerUnavailable('pdf', `The check could not be run, so the PDF was withheld. (${reason(error)})`);
  }
  return review('pdf', pdf, report);
}

export async function runComparison(original: Uint8Array, modified: Uint8Array, details: ComparisonDetails, deps: FlowDependencies): Promise<FlowResult> {
  let produced: { trackedChanges: Uint8Array; engineMs: number };
  try {
    produced = await deps.produce(original, modified);
  } catch (error) {
    throw new ProducerError(reason(error));
  }

  let trackedReport: CheckReport;
  try {
    trackedReport = await deps.checkTracked({ original, modified, output: produced.trackedChanges });
  } catch (error) {
    trackedReport = checkerUnavailable('tracked-changes', `The check could not be run, so the output was withheld. (${reason(error)})`);
  }
  const tracked = review('tracked-changes', produced.trackedChanges, trackedReport);

  let staticRedline: Reviewed;
  let counts: ChangeCounts | undefined;
  try {
    const made = await deps.makeStatic(produced.trackedChanges, original, details);
    let report: CheckReport;
    try {
      report = await deps.checkStatic({ original, modified, output: made.staticRedline, ...details });
    } catch (error) {
      report = checkerUnavailable('static-redline', `The check could not be run, so the output was withheld. (${reason(error)})`);
    }
    staticRedline = review('static-redline', made.staticRedline, report);
    if (staticRedline.file) counts = made.counts;
  } catch (error) {
    staticRedline = { report: checkerUnavailable('static-redline', `The static redline could not be made. (${reason(error)})`) };
  }

  return { tracked, staticRedline, counts, engineMs: produced.engineMs };
}
