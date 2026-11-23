/**
 * The compare flow: produce an output, have the independent Checker review it,
 * and release it only if every check passes.
 *
 * This is the one place where an output can reach the user, and it has no path
 * that does so without a passing report.
 */
import { CHECK_TITLES, type CheckId, type CheckReport } from '../checker';

export interface FlowDependencies {
  /** The Producer: make the tracked-changes file. */
  produce: (original: Uint8Array, modified: Uint8Array) => Promise<{ trackedChanges: Uint8Array; engineMs: number }>;
  /** The Checker: review that file against the two uploads. */
  check: (input: { original: Uint8Array; modified: Uint8Array; output: Uint8Array }) => Promise<CheckReport>;
}

export interface FlowResult {
  report: CheckReport;
  /** The tracked-changes Word file. Present only when every check passed. */
  trackedChanges?: Uint8Array;
  engineMs: number;
}

/** Thrown when the Producer itself fails; there is then nothing to check. */
export class ProducerError extends Error {}

/** A report saying the Checker could not run. The output is withheld, exactly as for a failed check. */
function checkerUnavailable(message: string): CheckReport {
  const ids: CheckId[] = ['accept-all', 'reject-all', 'completeness', 'structure'];
  return {
    output: 'tracked-changes',
    passed: false,
    checks: ids.map((id) => ({
      id,
      title: CHECK_TITLES[id],
      passed: false,
      failure: { message: `The check could not be run, so the output was withheld. (${message})` },
    })),
  };
}

export async function runComparison(original: Uint8Array, modified: Uint8Array, deps: FlowDependencies): Promise<FlowResult> {
  let produced: { trackedChanges: Uint8Array; engineMs: number };
  try {
    produced = await deps.produce(original, modified);
  } catch (error) {
    throw new ProducerError(error instanceof Error ? error.message : String(error));
  }

  let report: CheckReport;
  try {
    report = await deps.check({ original, modified, output: produced.trackedChanges });
  } catch (error) {
    report = checkerUnavailable(error instanceof Error ? error.message : String(error));
  }

  // Release only on an explicit, complete pass.
  const released = report.passed === true && report.checks.length > 0 && report.checks.every((c) => c.passed === true);
  return { report: released ? report : { ...report, passed: false }, trackedChanges: released ? produced.trackedChanges : undefined, engineMs: produced.engineMs };
}
