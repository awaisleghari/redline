/**
 * Stage 3 acceptance, part 1: every generated pair's static redline passes
 * every one of its checks.
 */
import { compareDocuments, initialize } from 'docxodus/core';
import { beforeAll, describe, expect, it } from 'vitest';
import { CHECKS_FOR, checkStaticRedline } from '../../src/checker';
import { makeStaticRedline } from '../../src/producer/make-static-redline';
import { ALL_PAIR_SPECS, buildPair } from '../fixtures/pairs';

export const COVER = { originalName: 'original.docx', modifiedName: 'modified.docx', comparedAt: '23 November 2026 at 14:05 (UTC+00:00)' };

beforeAll(async () => {
  await initialize();
});

describe('Checker on the static redline of every generated pair', () => {
  it.each(ALL_PAIR_SPECS)('$id: every check passes', async (spec) => {
    const pair = await buildPair(spec);
    const tracked = await compareDocuments(pair.original, pair.modified, { authorName: 'Redline' });
    const { docx, counts } = await makeStaticRedline(tracked, pair.original, COVER);
    const report = await checkStaticRedline({ original: pair.original, modified: pair.modified, output: docx, ...COVER });
    expect(report.checks.filter((c) => !c.passed)).toEqual([]);
    expect(report.passed).toBe(true);
    expect(report.checks.map((c) => c.id)).toEqual(CHECKS_FOR['static-redline']);
    // Where a pair states what the cover page must count, the Producer's counts must say exactly that.
    if (spec.staticCounts) expect(counts).toEqual(spec.staticCounts);
  });
});
