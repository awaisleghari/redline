/** Paragraph-by-paragraph comparison of two readings of a document. */
import type { CheckFailure } from './report';
import type { Story } from './stories';

/** A short excerpt of `text` around position `at`, so a difference can be shown without the whole paragraph. */
function excerpt(text: string, at: number): string {
  const start = Math.max(0, at - 40);
  const end = Math.min(text.length, at + 60);
  return `${start > 0 ? '…' : ''}${text.slice(start, end)}${end < text.length ? '…' : ''}`;
}

function firstDifference(a: string, b: string): number {
  const limit = Math.min(a.length, b.length);
  for (let i = 0; i < limit; i++) if (a[i] !== b[i]) return i;
  return limit;
}

/**
 * Compare what the output yields (`actual`) with what the upload contains
 * (`expected`). Returns the first difference, or undefined when identical.
 * `expectedName` is "modified document" or "original document".
 */
export function compareStories(expected: Story[], actual: Story[], expectedName: string): CheckFailure | undefined {
  const actualByKey = new Map(actual.map((s) => [s.key, s]));
  const expectedKeys = new Set(expected.map((s) => s.key));

  for (const want of expected) {
    const got = actualByKey.get(want.key);
    if (!got) {
      return {
        message: `The ${want.label} of the ${expectedName} is missing from the output.`,
        where: want.label,
        expected: excerpt(want.paragraphs.find((p) => p !== '') ?? '', 0),
        found: '(nothing)',
      };
    }
    const count = Math.min(want.paragraphs.length, got.paragraphs.length);
    for (let i = 0; i < count; i++) {
      const a = want.paragraphs[i]!;
      const b = got.paragraphs[i]!;
      if (a === b) continue;
      const at = firstDifference(a, b);
      return {
        message: `The text differs from the ${expectedName}.`,
        where: `${want.label}, paragraph ${i + 1}`,
        expected: excerpt(a, at),
        found: excerpt(b, at),
      };
    }
    if (want.paragraphs.length !== got.paragraphs.length) {
      const longer = want.paragraphs.length > got.paragraphs.length ? want : got;
      const extra = longer.paragraphs[count] ?? '';
      return {
        message:
          want.paragraphs.length > got.paragraphs.length
            ? `The output has fewer paragraphs than the ${expectedName} (${got.paragraphs.length} instead of ${want.paragraphs.length}).`
            : `The output has more paragraphs than the ${expectedName} (${got.paragraphs.length} instead of ${want.paragraphs.length}).`,
        where: `${want.label}, paragraph ${count + 1}`,
        expected: want.paragraphs.length > got.paragraphs.length ? excerpt(extra, 0) : '(no further paragraphs)',
        found: want.paragraphs.length > got.paragraphs.length ? '(no further paragraphs)' : excerpt(extra, 0),
      };
    }
  }

  for (const got of actual) {
    if (!expectedKeys.has(got.key)) {
      return {
        message: `The output has a ${got.label} that the ${expectedName} does not have.`,
        where: got.label,
        expected: '(nothing)',
        found: excerpt(got.paragraphs.find((p) => p !== '') ?? '', 0),
      };
    }
  }
  return undefined;
}
