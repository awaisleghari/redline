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
export type StateReader = () => Promise<string[]>;

const SHAPE = /^(asleep|unavailable|conversions=\d+ files=\d+ office_processes=\d+)$/;

/** The reply to a request for STATE_PATH. Never throws. */
export async function handleStateRequest(request: Request, read: StateReader): Promise<Response> {
  const headers = { 'Cache-Control': 'no-store', 'X-Content-Type-Options': 'nosniff' };
  if (request.method !== 'GET') return new Response('Ask with GET.', { status: 405, headers: { ...headers, Allow: 'GET', 'Content-Type': 'text/plain; charset=utf-8' } });
  let converters: string[];
  try {
    // Only the three counts can pass: whatever else a converter might say is not repeated.
    converters = (await read()).map((line) => (SHAPE.test(line) ? line : 'unavailable'));
  } catch {
    converters = [];
  }
  return Response.json({ converters }, { headers });
}
