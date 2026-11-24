/**
 * A public, read-only report on the PDF converter, so that "it keeps nothing"
 * can be checked on the live site and not only in the tests.
 *
 * For each copy of the converter it gives either "asleep" (not running: its
 * disk does not exist) or three counts: conversions attempted since it started,
 * files it is holding that were not there when it started, and LibreOffice
 * processes running. It never reports a name, a size or any content.
 */
export const STATE_PATH = '/api/converter-state';

/** One line per copy of the converter. */
export type StateReader = (detail: boolean) => Promise<string[]>;

// Numbers are kept short and lines bounded: this route is the only way anything from inside the converter
// reaches the outside, so it must not be able to carry more than a few counts.
const SHAPE = /^(asleep|unavailable|conversions=\d{1,9} files=\d{1,9} office_processes=\d{1,9})$/;
const MAX_DETAIL_LENGTH = 600;
/**
 * With ?detail: where the counted files are, as a count per system folder (never a file's name).
 * The folders are the ones the converter found writable when it started, before any document arrived,
 * so their names cannot come from a document.
 */
const DETAIL_SHAPE = /^(asleep|unavailable|before_first=(none|\d{1,9}) since_first=\d{1,9} where=(\/[A-Za-z0-9._\/-]{0,80}:\d{1,9},?){0,12})$/;

/** The reply to a request for STATE_PATH. Never throws. */
export async function handleStateRequest(request: Request, read: StateReader): Promise<Response> {
  const headers = { 'Cache-Control': 'no-store', 'X-Content-Type-Options': 'nosniff' };
  if (request.method !== 'GET') return new Response('Ask with GET.', { status: 405, headers: { ...headers, Allow: 'GET', 'Content-Type': 'text/plain; charset=utf-8' } });
  let converters: string[];
  try {
    // Only the three counts can pass: whatever else a converter might say is not repeated.
    const detail = new URL(request.url).searchParams.has('detail');
    converters = (await read(detail)).map((line) => (line.length <= MAX_DETAIL_LENGTH && (detail ? DETAIL_SHAPE : SHAPE).test(line) ? line : 'unavailable'));
  } catch {
    converters = [];
  }
  return Response.json({ converters }, { headers });
}
