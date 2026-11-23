/** The server-side PDF route: what it accepts, what it refuses, and that it passes nothing else through. */
import { describe, expect, it } from 'vitest';
import { DOCX_TYPE, handlePdfRequest, MAX_DOCX_BYTES, PDF_PATH } from '../../worker/pdf-route';

const SITE = 'https://redline.example';
const docx = new Uint8Array([0x50, 0x4b, 3, 4, 1, 2, 3]);
const pdfReply = (): Response => new Response(new TextEncoder().encode('%PDF-1.7 ...'), { status: 200, headers: { 'Content-Type': 'application/pdf' } });
const request = (init: RequestInit & { headers?: Record<string, string> } = {}): Request =>
  new Request(SITE + PDF_PATH, { method: 'POST', body: docx, ...init, headers: { 'Content-Type': DOCX_TYPE, ...init.headers } });

describe('PDF route', () => {
  it('sends the file to the converter and returns the PDF, uncacheable', async () => {
    let received: ArrayBuffer | undefined;
    const response = await handlePdfRequest(request({ headers: { Origin: SITE } }), async (body) => {
      received = body;
      return pdfReply();
    });
    expect(new Uint8Array(received!)).toEqual(docx);
    expect(response.status).toBe(200);
    expect(response.headers.get('Content-Type')).toBe('application/pdf');
    expect(response.headers.get('Cache-Control')).toBe('no-store');
    expect(await response.text()).toMatch(/^%PDF-/);
  });

  it('refuses anything but POST, other sites, other file types, empty and oversized files, without calling the converter', async () => {
    let called = 0;
    const never = async (): Promise<Response> => {
      called++;
      return pdfReply();
    };
    expect((await handlePdfRequest(new Request(SITE + PDF_PATH), never)).status).toBe(405);
    expect((await handlePdfRequest(request({ headers: { Origin: 'https://elsewhere.example' } }), never)).status).toBe(403);
    expect((await handlePdfRequest(request({ headers: { 'Content-Type': 'application/pdf' } }), never)).status).toBe(415);
    expect((await handlePdfRequest(request({ body: new Uint8Array(0) }), never)).status).toBe(400);
    expect((await handlePdfRequest(request({ body: new TextEncoder().encode('not a zip') }), never)).status).toBe(400);
    expect((await handlePdfRequest(request({ headers: { 'Content-Length': String(MAX_DOCX_BYTES + 1) } }), never)).status).toBe(413);
    const big = new Uint8Array(MAX_DOCX_BYTES + 1);
    big.set([0x50, 0x4b]);
    expect((await handlePdfRequest(request({ body: big }), never)).status).toBe(413);
    expect(called).toBe(0);
  });

  it("passes on the converter's plain refusals and hides anything else", async () => {
    const refused = await handlePdfRequest(request(), async () => new Response('The document could not be converted to PDF.', { status: 422 }));
    expect(refused.status).toBe(422);
    expect(await refused.text()).toBe('The document could not be converted to PDF.');
    const crashed = await handlePdfRequest(request(), async () => new Response('Traceback: /tmp/c-abc/in.docx ...', { status: 500 }));
    expect(crashed.status).toBe(502);
    expect(await crashed.text()).toBe('The document could not be converted to PDF.');
    const down = await handlePdfRequest(request(), async () => {
      throw new Error('connection refused');
    });
    expect(down.status).toBe(502);
    expect(await down.text()).not.toContain('connection refused');
    const notPdf = await handlePdfRequest(request(), async () => new Response('<html>oops</html>', { status: 200 }));
    expect(notPdf.status).toBe(502);
  });
});
