/**
 * The Cloudflare Worker: serves the site's static files, and passes PDF
 * requests to the converter container. (Deployment settings, wrangler.jsonc,
 * are Stage 6.)
 *
 * The container gets no access to the internet: it receives a file, converts
 * it and replies.
 */
import { Container, getRandom } from '@cloudflare/containers';
import { DOCX_TYPE, handlePdfRequest, PDF_PATH } from './pdf-route';

export class PdfConverter extends Container {
  defaultPort = 8080;
  /** Shut the converter down after a short idle spell; it starts again on demand. */
  sleepAfter = '5m';
  enableInternet = false;
}

interface Env {
  PDF_CONVERTER: DurableObjectNamespace<PdfConverter>;
  ASSETS: Fetcher;
}

/** How many converter instances requests are spread over. */
const INSTANCES = 2;

export default {
  async fetch(request: Request, env: Env): Promise<Response> {
    if (new URL(request.url).pathname !== PDF_PATH) return env.ASSETS.fetch(request);
    return handlePdfRequest(request, async (docx) => {
      const converter = await getRandom(env.PDF_CONVERTER, INSTANCES);
      return converter.fetch(new Request('http://converter/convert', { method: 'POST', headers: { 'Content-Type': DOCX_TYPE, 'Content-Length': String(docx.byteLength) }, body: docx }));
    });
  },
} satisfies ExportedHandler<Env>;
