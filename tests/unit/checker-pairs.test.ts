/**
 * Stage 2 acceptance, part 1: every generated test pair's tracked-changes
 * output passes the Checker.
 */
import { compareDocuments, initialize } from 'docxodus/core';
import { beforeAll, describe, expect, it } from 'vitest';
import { checkTrackedChanges } from '../../src/checker';
import { ALL_PAIR_SPECS, buildPair } from '../fixtures/pairs';

beforeAll(async () => {
  await initialize();
});

describe('Checker on real engine output', () => {
  it.each(ALL_PAIR_SPECS)('$id: every check passes', async (spec) => {
    const pair = await buildPair(spec);
    const output = await compareDocuments(pair.original, pair.modified, { authorName: 'Redline' });
    const report = await checkTrackedChanges({ original: pair.original, modified: pair.modified, output });
    expect(report.checks.filter((c) => !c.passed)).toEqual([]);
    expect(report.passed).toBe(true);
    expect(report.checks.map((c) => c.id)).toEqual(['accept-all', 'reject-all', 'formatting', 'numbering', 'completeness', 'structure']);
  });
});
