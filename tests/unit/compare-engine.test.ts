/**
 * Stage 1 acceptance: every generated test pair produces the expected tracked
 * changes: the count and the exact text of insertions, deletions and moves, in
 * every part of the file.
 *
 * The engine here is the same WebAssembly build the browser uses, loaded
 * through docxodus's Node entry point. The browser tests in tests/e2e/ repeat
 * the check through the real page and Web Worker.
 */
import { compareDocuments, initialize } from 'docxodus/core';
import { beforeAll, describe, expect, it } from 'vitest';
import { REVISION_AUTHOR } from '../../src/producer/protocol';
import { ALL_PAIR_SPECS, PAIR_SPECS, buildPair } from '../fixtures/pairs';
import { fullExpectation, readTrackedChanges } from '../helpers/tracked-changes';

beforeAll(async () => {
  await initialize();
});

describe('generated test pairs', () => {
  it('cover every kind of change listed in PLAN.md', () => {
    expect(PAIR_SPECS.map((spec) => spec.id)).toEqual([
      'one-word-change',
      'number-change',
      'inserted-paragraph',
      'deleted-paragraph',
      'moved-paragraph',
      'table-cell-change',
      'header-change',
      'footnote-change',
      'numbered-list-change',
    ]);
  });

  it.each(ALL_PAIR_SPECS)('$id: the original and modified files really differ', async (spec) => {
    const pair = await buildPair(spec);
    const original = await readTrackedChanges(pair.original);
    const modified = await readTrackedChanges(pair.modified);
    // The inputs themselves must be clean: no tracked changes of their own.
    expect(original.parts).toEqual(fullExpectation({}));
    expect(modified.parts).toEqual(fullExpectation({}));
    expect(Buffer.from(pair.original).equals(Buffer.from(pair.modified))).toBe(false);
  });
});

describe('compare engine', () => {
  it.each(ALL_PAIR_SPECS)('$id: produces exactly the expected tracked changes', async (spec) => {
    const pair = await buildPair(spec);
    const redline = await compareDocuments(pair.original, pair.modified, { authorName: REVISION_AUTHOR });
    const found = await readTrackedChanges(redline);
    expect(found.parts).toEqual(fullExpectation(spec.expected));
    expect(found.formatChanges).toBe(0);
  });

  it('comparing a document with itself produces no tracked changes', async () => {
    const pair = await buildPair(PAIR_SPECS[0]!);
    const redline = await compareDocuments(pair.original, pair.original, { authorName: REVISION_AUTHOR });
    const found = await readTrackedChanges(redline);
    expect(found.parts).toEqual(fullExpectation({}));
  });
});
