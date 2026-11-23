/**
 * The one server-side step in the product: turn a static redline into a PDF.
 *
 * This module is plain web-standard code (Request in, Response out) so that the
 * same logic runs on Cloudflare (worker/index.ts), in local development and in
 * the tests. It keeps nothing and logs nothing: there is no console call here
 * or anywhere on the path.
 */

export const PDF_PATH = '/api/pdf';
export const DOCX_TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
/** A little above the page's 25 MB upload limit: a static redline is slightly larger than the documents it is made from. */
export const MAX_DOCX_BYTES = 30 * 1024 * 1024;

/** Sends a Word file to the converter and returns its reply. */
export type Converter = (docx: ArrayBuffer) => Promise<Response>;

const plain = (status: number, message: string, extra: Record<string, string> = {}): Response =>
  new Response(message, { status, headers: { 'Content-Type': 'text/plain; charset=utf-8', 'Cache-Control': 'no-store', 'X-Content-Type-Options': 'nosniff', ...extra } });

/** The reply to a request for PDF_PATH. Never throws. */
export async function handlePdfRequest(request: Request, convert: Converter): Promise<Response> {
  if (request.method !== 'POST') return plain(405, 'Send the document with POST.', { Allow: 'POST' });

  // Only this site's own page may use the converter.
  const origin = request.headers.get('Origin');
  if (origin !== null && origin !== new URL(request.url).origin) return plain(403, 'The converter can only be used from this site.');

  if ((request.headers.get('Content-Type') ?? '').split(';')[0]!.trim() !== DOCX_TYPE) return plain(415, 'Only Word documents (.docx) can be converted.');
  const declared = Number(request.headers.get('Content-Length') ?? NaN);
  if (Number.isFinite(declared) && declared > MAX_DOCX_BYTES) return plain(413, 'The file is too large to convert.');

  let docx: ArrayBuffer;
  try {
    docx = await request.arrayBuffer();
  } catch {
    return plain(400, 'The file could not be read.');
  }
  if (docx.byteLength === 0) return plain(400, 'No file was sent.');
  if (docx.byteLength > MAX_DOCX_BYTES) return plain(413, 'The file is too large to convert.');
  const head = new Uint8Array(docx, 0, 2);
  if (head[0] !== 0x50 || head[1] !== 0x4b) return plain(400, 'The file is not a Word document.');

  let converted: Response;
  try {
    converted = await convert(docx);
  } catch {
    return plain(502, 'The PDF converter could not be reached. Please try again in a moment.');
  }
  if (!converted.ok) {
    // Pass on the converter's own plain refusals; anything else becomes a general message.
    const known = [400, 413, 415, 422, 503].includes(converted.status);
    return plain(known ? converted.status : 502, known ? await converted.text() : 'The document could not be converted to PDF.');
  }
  const pdf = await converted.arrayBuffer();
  const start = new Uint8Array(pdf, 0, Math.min(5, pdf.byteLength));
  if (new TextDecoder().decode(start) !== '%PDF-') return plain(502, 'The document could not be converted to PDF.');
  return new Response(pdf, { status: 200, headers: { 'Content-Type': 'application/pdf', 'Cache-Control': 'no-store', 'X-Content-Type-Options': 'nosniff' } });
}
