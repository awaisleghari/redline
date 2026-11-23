/** What the Checker reports. Plain data, safe to pass between workers and to show to the user. */

export type CheckId = 'accept-all' | 'reject-all' | 'completeness' | 'structure';

export interface CheckFailure {
  /** Plain-English explanation of what is wrong. */
  message: string;
  /** Where it went wrong, e.g. "main text, paragraph 4". */
  where?: string;
  /** What the text should have been (an excerpt around the first difference). */
  expected?: string;
  /** What the output contains instead. */
  found?: string;
}

export interface CheckResult {
  id: CheckId;
  /** The plain-English statement being checked, e.g. "Accepting every change reproduces the modified document". */
  title: string;
  passed: boolean;
  /** One line of supporting detail when passed (what was covered). */
  note?: string;
  failure?: CheckFailure;
}

export interface CheckReport {
  /** Which output this report is about. */
  output: 'tracked-changes';
  /** True only if every check passed. The output may be released to the user only then. */
  passed: boolean;
  checks: CheckResult[];
}

export const CHECK_TITLES: Record<CheckId, string> = {
  'accept-all': 'Accepting every change reproduces the modified document',
  'reject-all': 'Rejecting every change reproduces the original document',
  completeness: 'Every part of the documents that holds text was checked',
  structure: 'The Word file is structurally sound',
};

/** Render a report as the short text file a user can keep with the redline. */
export function reportAsText(report: CheckReport, context: { originalName: string; modifiedName: string; when: Date }): string {
  const lines = [
    'Redline check report',
    `Original: ${context.originalName}`,
    `Modified: ${context.modifiedName}`,
    `Checked: ${context.when.toISOString()}`,
    `Output: Word file with tracked changes`,
    `Overall: ${report.passed ? 'all checks passed' : 'FAILED - this output was withheld'}`,
    '',
  ];
  for (const check of report.checks) {
    lines.push(`${check.title}: ${check.passed ? 'passed' : 'FAILED'}`);
    if (check.passed && check.note) lines.push(`  ${check.note}`);
    if (check.failure) {
      lines.push(`  ${check.failure.message}`);
      if (check.failure.where) lines.push(`  Where: ${check.failure.where}`);
      if (check.failure.expected !== undefined) lines.push(`  Expected: ${check.failure.expected}`);
      if (check.failure.found !== undefined) lines.push(`  Found:    ${check.failure.found}`);
    }
  }
  return lines.join('\n') + '\n';
}
