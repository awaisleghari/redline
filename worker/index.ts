/**
 * The Cloudflare Worker. Cloudflare serves the site's static files itself
 * (wrangler.jsonc, "assets"); this code handles the two addresses that are not
 * files: the PDF route, which passes a static redline to the converter
 * container, and the converter's state report.
 *
 * The container gets no access to the internet: it receives a file, converts
 * it and replies. Nothing is logged here (there is no console call), and
 * Cloudflare's own request logs are switched off in wrangler.jsonc.
 */
import { Container } from '@cloudflare/containers';
import { DOCX_TYPE, handlePdfRequest, PDF_PATH } from './pdf-route';
import { handleStateRequest, STATE_PATH } from './state-route';

// The containers library needs this exported to enforce "no internet access".
export { ContainerProxy } from '@cloudflare/containers';

const CONVERTER_PORT = 8080;

export class PdfConverter extends Container {
  defaultPort = CONVERTER_PORT;
  /** Shut the converter down after a short idle spell (its disk goes with it); it starts again on demand. */
  sleepAfter = '5m';
  enableInternet = false;

  /** The converter's three counts, or "asleep" when it is not running. Asking never starts it or keeps it awake. */
  async converterState(): Promise<string> {
    const container = this.ctx.container;
    if (!container?.running) return 'asleep';
    try {
      const reply = await container.getTcpPort(CONVERTER_PORT).fetch('http://converter/state', { signal: AbortSignal.timeout(5000) });
      return reply.ok ? await reply.text() : 'unavailable';
    } catch {
      return 'unavailable';
    }
  }
}

interface Env {
  PDF_CONVERTER: DurableObjectNamespace<PdfConverter>;
  ASSETS: Fetcher;
}

/** How many copies of the converter requests are spread over (wrangler.jsonc allows no more than this to run). */
const INSTANCES = 2;
const converter = (env: Env, index: number): DurableObjectStub<PdfConverter> => env.PDF_CONVERTER.get(env.PDF_CONVERTER.idFromName(`instance-${index}`));

export default {
  async fetch(request: Request, env: Env): Promise<Response> {
    const path = new URL(request.url).pathname;
    if (path === PDF_PATH) {
      return handlePdfRequest(request, (docx) =>
        converter(env, Math.floor(Math.random() * INSTANCES)).fetch(
          new Request('http://converter/convert', { method: 'POST', headers: { 'Content-Type': DOCX_TYPE, 'Content-Length': String(docx.byteLength) }, body: docx }),
        ),
      );
    }
    if (path === STATE_PATH) {
      return handleStateRequest(request, () => Promise.all(Array.from({ length: INSTANCES }, (_, index) => converter(env, index).converterState())));
    }
    return env.ASSETS.fetch(request);
  },
} satisfies ExportedHandler<Env>;
