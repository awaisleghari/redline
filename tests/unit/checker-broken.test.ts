/**
 * Stage 2 acceptance, part 2: every deliberately broken output is caught, and
 * by exactly the checks that are meant to catch it.
 */
import { compareDocuments, initialize } from 'docxodus/core';
import { beforeAll, describe, expect, it } from 'vitest';
import { checkTrackedChanges } from '../../src/checker';
import { BROKEN_CASES, breakOutput, expectedFailures } from '../fixtures/broken';
import { buildPair, pairById } from '../fixtures/pairs';

beforeAll(async () => {
  await initialize();
});

describe('Checker on deliberately broken outputs', () => {
  it('has a broken case for each kind of damage PLAN.md lists for this stage', () => {
    const ids = BROKEN_CASES.map((c) => c.id);
    for (const required of ['dropped-word', 'changed-digit-unchanged-text', 'missing-footnote', 'deletion-styled-as-insertion', 'broken-comment-anchor']) {
      expect(ids).toContain(required);
    }
    expect(new Set(ids).size).toBe(ids.length);
  });

  it.each(BROKEN_CASES)('$id is caught ($description)', async (brokenCase) => {
    const pair = await buildPair(pairById(brokenCase.pairId));
    const genuine = await compareDocuments(pair.original, pair.modified, { authorName: 'Redline' });

    // The genuine output passes, so any failure below is caused by the damage alone.
    const clean = await checkTrackedChanges({ original: pair.original, modified: pair.modified, output: genuine });
    expect(clean.passed).toBe(true);

    const broken = await breakOutput(genuine, brokenCase);
    const report = await checkTrackedChanges({ original: pair.original, modified: pair.modified, output: broken });
    expect(report.passed).toBe(false);
    const failed = report.checks.filter((c) => !c.passed);
    expect(failed.map((c) => c.id).sort()).toEqual(expectedFailures(brokenCase));
    // Every failure explains itself in words.
    for (const check of failed) expect(check.failure?.message ?? '').not.toBe('');
  });
});
