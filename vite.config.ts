import type { IncomingMessage, ServerResponse } from 'node:http';
import type { Plugin } from 'vite';
import { defineConfig } from 'vitest/config';
import { DOCX_TYPE, handlePdfRequest, PDF_PATH } from './worker/pdf-route';

/**
 * Local stand-in for the Cloudflare Worker's PDF route: the same request
 * handling (worker/pdf-route.ts), with the converter reached at a local
 * address instead of through Cloudflare. Start the converter with
 * `docker run --rm -p 127.0.0.1:8791:8080 redline-pdf` (after `npm run build:pdf-image`).
 */
function localPdfRoute(): Plugin {
  const converter = process.env.PDF_CONVERTER_URL ?? 'http://127.0.0.1:8791';
  const middleware = async (req: IncomingMessage, res: ServerResponse, next: () => void): Promise<void> => {
    if (req.url?.split('?')[0] !== PDF_PATH) return next();
    const chunks: Buffer[] = [];
    for await (const chunk of req) chunks.push(chunk as Buffer);
    const headers = new Headers();
    for (const [name, value] of Object.entries(req.headers)) if (typeof value === 'string') headers.set(name, value);
    const body = Buffer.concat(chunks);
    const request = new Request(`http://${req.headers.host ?? 'localhost'}${PDF_PATH}`, { method: req.method, headers, body: req.method === 'POST' ? new Uint8Array(body) : undefined });
    const response = await handlePdfRequest(request, (docx) => fetch(`${converter}/convert`, { method: 'POST', headers: { 'Content-Type': DOCX_TYPE }, body: new Blob([docx]) }));
    res.statusCode = response.status;
    response.headers.forEach((value, name) => res.setHeader(name, value));
    res.end(Buffer.from(await response.arrayBuffer()));
  };
  return {
    name: 'redline-local-pdf-route',
    configureServer: (server) => void server.middlewares.use(middleware),
    configurePreviewServer: (server) => void server.middlewares.use(middleware),
  };
}

/**
 * The page's content-security policy lets the development server's live-reload
 * connection through (ws://localhost). The built site must not carry that.
 */
function productionPolicy(): Plugin {
  return {
    name: 'redline-production-policy',
    apply: 'build',
    transformIndexHtml(html) {
      const built = html.replace(" ws://localhost:* ws://127.0.0.1:*", '');
      if (built === html || built.includes('ws://')) throw new Error('index.html: the development-only part of the content-security policy was not found to remove');
      return built;
    },
  };
}

export default defineConfig({
  plugins: [localPdfRoute(), productionPolicy()],
  server: { port: 5173, strictPort: true },
  preview: { port: 4173, strictPort: true },
  worker: { format: 'es' },
  test: {
    include: ['tests/unit/**/*.test.ts'],
    environment: 'node',
    testTimeout: 120_000,
    hookTimeout: 120_000,
  },
});
