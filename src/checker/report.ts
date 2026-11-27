/** What the Checker reports. Plain data, safe to pass between workers and to show to the user. */

export type CheckId =
  | 'accept-all'
  | 'reject-all'
  | 'static-new'
  | 'static-old'
  | 'no-tracked-changes'
  | 'cover-page'
  | 'formatting'
  | 'format-marks'
  | 'numbering'
  | 'completeness'
  | 'structure'
  | 'pdf-text'
  | 'pdf-colours';

export type OutputKind = 'tracked-changes' | 'static-redline' | 'pdf';

/** The checks run on each kind of output, in the order they are reported. */
export const CHECKS_FOR: Record<OutputKind, CheckId[]> = {
  'tracked-changes': ['accept-all', 'reject-all', 'formatting', 'numbering', 'completeness', 'structure'],
  'static-redline': ['static-new', 'static-old', 'no-tracked-changes', 'cover-page', 'formatting', 'format-marks', 'numbering', 'completeness', 'structure'],
  pdf: ['pdf-text', 'pdf-colours'],
};

export const OUTPUT_NAME: Record<OutputKind, string> = {
  'tracked-changes': 'Word file with tracked changes',
  'static-redline': 'Word file with the static redline',
  pdf: 'PDF of the static redline',
};

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
  output: OutputKind;
  /** True only if every check passed. The output may be released to the user only then. */
  passed: boolean;
  checks: CheckResult[];
}

export const CHECK_TITLES: Record<CheckId, string> = {
  'accept-all': 'Accepting every change reproduces the modified document',
  'reject-all': 'Rejecting every change reproduces the original document',
  'static-new': 'Taking out the deleted and moved-from text leaves the modified document',
  'static-old': 'Taking out the inserted and moved-to text leaves the original document',
  'no-tracked-changes': 'No tracked changes are left in the static redline',
  'cover-page': 'The cover page names the right files and time, and its counts match the redline',
  formatting: "The text has the modified document's bold, italic, underline, capitals and hidden-text settings",
  'format-marks': 'Formatting and link changes are shaded exactly where they happened, and nowhere else',
  numbering: 'Clause numbers match the modified document',
  completeness: 'Every part of the documents that holds text was checked',
  structure: 'The Word file is structurally sound',
  'pdf-text': "The PDF's text matches the static redline, complete and in order",
  'pdf-colours': 'Inserted, deleted and moved text is in the right colours, struck through or underlined, in the PDF',
};

/** Render a report as the short text file a user can keep with the redline. */
export function reportAsText(report: CheckReport, context: { originalName: string; modifiedName: string; when: Date }): string {
  const lines = [
    'Redline check report',
    `Original: ${context.originalName}`,
    `Modified: ${context.modifiedName}`,
    `Checked: ${context.when.toISOString()}`,
    `Output: ${OUTPUT_NAME[report.output]}`,
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
