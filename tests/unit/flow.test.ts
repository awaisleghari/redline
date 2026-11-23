/**
 * The compare flow must never release an output without a complete, passing
 * report from the Checker.
 */
import { describe, expect, it } from 'vitest';
import { ProducerError, runComparison } from '../../src/app/flow';
import { CHECK_TITLES, CHECKS_FOR, type CheckId, type CheckReport } from '../../src/checker';

const bytes = (text: string): Uint8Array => new TextEncoder().encode(text);
const IDS: CheckId[] = CHECKS_FOR['tracked-changes'];
const report = (failing: CheckId[] = [], claimPassed = failing.length === 0): CheckReport => ({
  output: 'tracked-changes',
  passed: claimPassed,
  checks: IDS.map((id) => ({ id, title: CHECK_TITLES[id], passed: !failing.includes(id), failure: failing.includes(id) ? { message: 'wrong' } : undefined })),
});
const produce = async (): Promise<{ trackedChanges: Uint8Array; engineMs: number }> => ({ trackedChanges: bytes('output'), engineMs: 5 });

describe('compare flow', () => {
  it('releases the output when every check passes', async () => {
    const result = await runComparison(bytes('a'), bytes('b'), { produce, check: async () => report() });
    expect(result.report.passed).toBe(true);
    expect(result.trackedChanges).toEqual(bytes('output'));
  });

  it('gives the Checker the raw uploads and the raw output', async () => {
    let seen: { original: Uint8Array; modified: Uint8Array; output: Uint8Array } | undefined;
    await runComparison(bytes('a'), bytes('b'), {
      produce,
      check: async (input) => {
        seen = input;
        return report();
      },
    });
    expect(seen).toEqual({ original: bytes('a'), modified: bytes('b'), output: bytes('output') });
  });

  it.each(IDS)('withholds the output when the %s check fails', async (id) => {
    const result = await runComparison(bytes('a'), bytes('b'), { produce, check: async () => report([id]) });
    expect(result.report.passed).toBe(false);
    expect(result.trackedChanges).toBeUndefined();
  });

  it('withholds the output when the Checker crashes', async () => {
    const result = await runComparison(bytes('a'), bytes('b'), {
      produce,
      check: async () => {
        throw new Error('worker died');
      },
    });
    expect(result.report.passed).toBe(false);
    expect(result.report.checks.every((c) => !c.passed)).toBe(true);
    expect(result.trackedChanges).toBeUndefined();
  });

  it('withholds the output when a report claims to pass but contains a failed check', async () => {
    const result = await runComparison(bytes('a'), bytes('b'), { produce, check: async () => report(['structure'], true) });
    expect(result.report.passed).toBe(false);
    expect(result.trackedChanges).toBeUndefined();
  });

  it('withholds the output when a report claims to pass but lists no checks', async () => {
    const result = await runComparison(bytes('a'), bytes('b'), {
      produce,
      check: async () => ({ output: 'tracked-changes', passed: true, checks: [] }),
    });
    expect(result.report.passed).toBe(false);
    expect(result.trackedChanges).toBeUndefined();
  });

  it('reports a Producer failure as such and never calls the Checker', async () => {
    let checked = false;
    const run = runComparison(bytes('a'), bytes('b'), {
      produce: async () => {
        throw new Error('engine failed');
      },
      check: async () => {
        checked = true;
        return report();
      },
    });
    await expect(run).rejects.toBeInstanceOf(ProducerError);
    expect(checked).toBe(false);
  });
});
