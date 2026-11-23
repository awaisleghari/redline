/**
 * The compare flow must never release an output without a complete, passing
 * report from the Checker for that output.
 */
import { describe, expect, it } from 'vitest';
import { ProducerError, runComparison, type FlowDependencies } from '../../src/app/flow';
import { CHECK_TITLES, CHECKS_FOR, type CheckId, type CheckReport, type OutputKind } from '../../src/checker';

const bytes = (text: string): Uint8Array => new TextEncoder().encode(text);
const DETAILS = { originalName: 'a.docx', modifiedName: 'b.docx', comparedAt: 'now' };
const COUNTS = { insertions: 1, deletions: 2, moves: 3, formattingChanges: 4 };

const report = (kind: OutputKind, failing: CheckId[] = [], claimPassed = failing.length === 0): CheckReport => ({
  output: kind,
  passed: claimPassed,
  checks: CHECKS_FOR[kind].map((id) => ({ id, title: CHECK_TITLES[id], passed: !failing.includes(id), failure: failing.includes(id) ? { message: 'wrong' } : undefined })),
});

const deps = (overrides: Partial<FlowDependencies> = {}): FlowDependencies => ({
  produce: async () => ({ trackedChanges: bytes('tracked'), engineMs: 5 }),
  makeStatic: async () => ({ staticRedline: bytes('static'), counts: COUNTS }),
  checkTracked: async () => report('tracked-changes'),
  checkStatic: async () => report('static-redline'),
  ...overrides,
});
const run = (overrides: Partial<FlowDependencies> = {}) => runComparison(bytes('a'), bytes('b'), DETAILS, deps(overrides));

describe('compare flow', () => {
  it('releases both outputs, and the counts, when every check passes', async () => {
    const result = await run();
    expect(result.tracked.file).toEqual(bytes('tracked'));
    expect(result.staticRedline.file).toEqual(bytes('static'));
    expect(result.counts).toEqual(COUNTS);
  });

  it('gives the Checker the raw uploads, the raw outputs, and the details from the page', async () => {
    const seen: unknown[] = [];
    await run({
      checkTracked: async (input) => (seen.push(input), report('tracked-changes')),
      checkStatic: async (input) => (seen.push(input), report('static-redline')),
    });
    expect(seen[0]).toEqual({ original: bytes('a'), modified: bytes('b'), output: bytes('tracked') });
    expect(seen[1]).toEqual({ original: bytes('a'), modified: bytes('b'), output: bytes('static'), ...DETAILS });
  });

  it.each(CHECKS_FOR['tracked-changes'])('withholds the tracked-changes file when its %s check fails, without affecting the static redline', async (id) => {
    const result = await run({ checkTracked: async () => report('tracked-changes', [id]) });
    expect(result.tracked.report.passed).toBe(false);
    expect(result.tracked.file).toBeUndefined();
    expect(result.staticRedline.file).toEqual(bytes('static'));
  });

  it.each(CHECKS_FOR['static-redline'])('withholds the static redline and the counts when its %s check fails', async (id) => {
    const result = await run({ checkStatic: async () => report('static-redline', [id]) });
    expect(result.staticRedline.report.passed).toBe(false);
    expect(result.staticRedline.file).toBeUndefined();
    expect(result.counts).toBeUndefined();
    expect(result.tracked.file).toEqual(bytes('tracked'));
  });

  it('withholds an output when the Checker crashes', async () => {
    const crash = async (): Promise<CheckReport> => {
      throw new Error('worker died');
    };
    const result = await run({ checkTracked: crash, checkStatic: crash });
    for (const reviewed of [result.tracked, result.staticRedline]) {
      expect(reviewed.report.passed).toBe(false);
      expect(reviewed.report.checks.every((c) => !c.passed)).toBe(true);
      expect(reviewed.file).toBeUndefined();
    }
  });

  it('withholds an output when a report claims to pass but contains a failed check', async () => {
    const result = await run({ checkTracked: async () => report('tracked-changes', ['structure'], true) });
    expect(result.tracked.report.passed).toBe(false);
    expect(result.tracked.file).toBeUndefined();
  });

  it('withholds an output when a report claims to pass but lists no checks, or not all of them', async () => {
    const none = await run({ checkTracked: async () => ({ output: 'tracked-changes', passed: true, checks: [] }) });
    expect(none.tracked.file).toBeUndefined();
    const partial = report('static-redline');
    partial.checks = partial.checks.filter((c) => c.id !== 'cover-page');
    const some = await run({ checkStatic: async () => partial });
    expect(some.staticRedline.file).toBeUndefined();
  });

  it('withholds an output whose report is for the other kind of output', async () => {
    const result = await run({ checkStatic: async () => report('tracked-changes') });
    expect(result.staticRedline.file).toBeUndefined();
  });

  it('withholds the static redline when it cannot be made, and still releases the tracked-changes file', async () => {
    const result = await run({
      makeStatic: async () => {
        throw new Error('conversion failed');
      },
    });
    expect(result.staticRedline.file).toBeUndefined();
    expect(result.staticRedline.report.passed).toBe(false);
    expect(result.tracked.file).toEqual(bytes('tracked'));
  });

  it('reports a failure to compare as such and never calls the Checker', async () => {
    let checked = false;
    const attempt = run({
      produce: async () => {
        throw new Error('engine failed');
      },
      checkTracked: async () => {
        checked = true;
        return report('tracked-changes');
      },
    });
    await expect(attempt).rejects.toBeInstanceOf(ProducerError);
    expect(checked).toBe(false);
  });
});
