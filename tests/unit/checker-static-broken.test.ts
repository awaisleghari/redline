/**
 * Stage 3 acceptance, part 2: every deliberately broken static redline is
 * caught, by exactly the checks expected.
 */
import { compareDocuments, initialize } from 'docxodus/core';
import { beforeAll, describe, expect, it } from 'vitest';
import { checkStaticRedline } from '../../src/checker';
import { unzip, zip } from '../../src/common/zip';
import { makeStaticRedline } from '../../src/producer/make-static-redline';
import { BROKEN_STATIC_CASES, expectedStaticFailures } from '../fixtures/broken-static';
import { buildPair, pairById } from '../fixtures/pairs';

const COVER = { originalName: 'original.docx', modifiedName: 'modified.docx', comparedAt: '23 November 2026 at 14:05 (UTC+00:00)' };

beforeAll(async () => {
  await initialize();
});

describe('Checker on deliberately broken static redlines', () => {
  it('has a broken case for each check added in this stage', () => {
    const caught = new Set(BROKEN_STATIC_CASES.flatMap((c) => c.caughtBy));
    for (const id of ['static-new', 'static-old', 'no-tracked-changes', 'cover-page', 'formatting', 'format-marks', 'numbering', 'completeness', 'structure']) expect(caught).toContain(id);
    expect(BROKEN_STATIC_CASES.map((c) => c.id)).toContain('static-wrong-cover-count');
    expect(BROKEN_STATIC_CASES.map((c) => c.id)).toContain('static-leftover-tracked-change');
    expect(new Set(BROKEN_STATIC_CASES.map((c) => c.id)).size).toBe(BROKEN_STATIC_CASES.length);
  });

  it.each(BROKEN_STATIC_CASES)('$id is caught ($description)', async (brokenCase) => {
    const pair = await buildPair(pairById(brokenCase.pairId));
    const tracked = await compareDocuments(pair.original, pair.modified, { authorName: 'Redline' });
    const genuine = (await makeStaticRedline(tracked, pair.original, COVER)).docx;
    const clean = await checkStaticRedline({ original: pair.original, modified: pair.modified, output: genuine, ...COVER });
    expect(clean.passed).toBe(true);

    const files = await unzip(genuine);
    brokenCase.damage(files);
    const report = await checkStaticRedline({ original: pair.original, modified: pair.modified, output: await zip(files), ...COVER });
    expect(report.passed).toBe(false);
    const failed = report.checks.filter((c) => !c.passed);
    expect(failed.map((c) => c.id).sort()).toEqual(expectedStaticFailures(brokenCase));
    for (const check of failed) expect(check.failure?.message ?? '').not.toBe('');
  });
});
