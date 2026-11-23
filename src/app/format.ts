/**
 * Formatting of values shown both on the page and on the cover page. This is
 * app code, not Producer code: the app formats the comparison time once and
 * hands the same text to the Producer (to print) and to the Checker (to verify).
 */

const MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];

/** e.g. "23 November 2026 at 14:05 (UTC+01:00)", in the user's own time zone. */
export function formatComparedAt(when: Date): string {
  const two = (n: number): string => String(n).padStart(2, '0');
  const offset = -when.getTimezoneOffset();
  const sign = offset < 0 ? '-' : '+';
  const zone = `UTC${sign}${two(Math.floor(Math.abs(offset) / 60))}:${two(Math.abs(offset) % 60)}`;
  return `${when.getDate()} ${MONTHS[when.getMonth()]} ${when.getFullYear()} at ${two(when.getHours())}:${two(when.getMinutes())} (${zone})`;
}

/** A file name safe to print: no control characters, and not absurdly long. */
export function displayName(fileName: string): string {
  // eslint-disable-next-line no-control-regex
  const clean = fileName.replace(/[\u0000-\u001f\u007f]/g, '').trim();
  return clean.length > 120 ? `${clean.slice(0, 117)}...` : clean || 'untitled.docx';
}
