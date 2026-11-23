# Redline

An open-source web app that compares two Word documents (an original and a
modified version) and produces a legal-style redline, in the manner of Litera
Compare. Comparison happens in your browser: the documents are not uploaded
anywhere to be compared. No AI model is used anywhere in the product.

The project is under construction. `PLAN.md` holds the design, the stages of
work and a running status log.

## Running it locally

You need [Node.js](https://nodejs.org) 22 or later.

```bash
npm install
```

```bash
npm run dev
```

Then open the address it prints (normally http://localhost:5173). Drop in an
original and a modified `.docx` and press Compare. The page shows the redline,
a summary of the changes and a check report, and offers two Word downloads: one
with tracked changes, one with the static (Litera-style) redline. Each is
offered only if every check on it passed.

The comparison engine is [docxodus](https://www.npmjs.com/package/docxodus), a
.NET engine compiled to WebAssembly. `npm run dev` and `npm run build` copy its
runtime files into `public/wasm/` automatically.

## Tests

```bash
npm test
```

runs the unit tests (Vitest). These include the Checker's tests: every
generated document pair must pass it, and every deliberately broken output in
`tests/fixtures/broken.ts` must be caught.

```bash
npm run test:schema
```

validates every generated output against the Office Open XML schema with
Microsoft's Open XML SDK. It needs the .NET 10 SDK
(`apt-get install -y dotnet-sdk-10.0` on Ubuntu).

```bash
npm run test:e2e
```

runs the browser tests (Playwright, in Chromium). The first time, install the
browser with `npx playwright install chromium`.

```bash
npm run test:pdf
```

builds and runs the PDF converter container and checks its output. It needs
Docker. So do the browser tests, because they include making a PDF.

```bash
npm run test:worker
```

runs the real Cloudflare Worker locally (`wrangler dev`, with the converter
container behind it) and tests the site as it is deployed: the page, the PDF
route, and that the converter holds nothing afterwards. It needs Docker.

```bash
npm run test:all
```

runs the type check and all five.

## The PDF converter

Making the PDF is the one step that happens on a server. `pdf-container/` holds
a small image (Ubuntu, LibreOffice, fonts) that turns a static redline into a
PDF and keeps nothing. To use it locally:

```bash
npm run build:pdf-image
```

```bash
docker run --rm -p 127.0.0.1:8791:8080 redline-pdf
```

With that running, `npm run dev` can make PDFs. `worker/` holds the Cloudflare
Worker that does the same job in production.

## How to deploy

The site runs on Cloudflare. It needs the Workers Paid plan, because the PDF
converter is a Cloudflare Container. `wrangler.jsonc` holds every deployment
setting; no setting lives only in the Cloudflare dashboard, and the repository
holds no password, token or key.

What gets set up on Cloudflare:

- **A Worker called `redline`**, at `https://redline.<your-account>.workers.dev`.
  Cloudflare serves the built site (`dist/`) directly. The Worker's own code
  (`worker/`) answers two addresses only: `/api/pdf` (the PDF step) and
  `/api/converter-state` (below).
- **The PDF converter**: the image in `pdf-container/`, at most two copies. A
  copy starts when a PDF is asked for and is shut down after five idle minutes,
  which also discards its disk. It has no access to the internet.
- **No storage of any kind** (no database, no file store) and **no logs**:
  Cloudflare's request logs are switched off for the Worker and the converter.
  `tests/unit/deploy-config.test.ts` fails if any of that changes.

### Automatic deployment (how the live site is run)

The GitHub repository is connected to Cloudflare ("Workers Builds"). Whenever
`main` changes, Cloudflare builds the site and the converter image and deploys
them. To set that up once, in the Cloudflare dashboard: **Workers & Pages**,
**Create application**, choose **GitHub**, pick this repository, and set:

| Setting | Value |
|---|---|
| Project name | `redline` (must match `name` in `wrangler.jsonc`) |
| Production branch | `main` |
| Build command | `npm run build` |
| Deploy command | `npx wrangler deploy` |
| Preview builds for other branches | off |

### Deploying by hand

With Docker running and a Cloudflare login (`npx wrangler login`):

```bash
npm run deploy
```

### Checking the live site

```bash
npm run build
```

```bash
npm run test:live -- https://redline.<your-account>.workers.dev
```

makes real PDFs through the live converter, runs the Checker on each, and
confirms the refusals and the security headers.

```bash
REDLINE_URL=https://redline.<your-account>.workers.dev npm run test:e2e
```

runs the whole browser suite against the live site (Docker is still needed
locally: a few tests make deliberately wrong PDFs to prove they are withheld).

`/api/converter-state` is a public, read-only report: for each copy of the
converter, either `asleep` or three counts (conversions attempted since it
started, files it holds that were not there when it started, LibreOffice
processes running). Between conversions the last two must be 0. It never
reports a name, a size or any content.

## How it is built

The product has two halves that do not share code (a test enforces this):

- `src/producer/` makes the outputs (the compare engine, in a Web Worker).
- `src/checker/` is an independent reviewer. It re-reads the two uploads and
  each output from their raw bytes, in its own Web Worker, and decides whether
  the output may be released. `src/app/flow.ts` is the only route by which an
  output reaches the user, and it releases one only when every check passes.

`scripts/early-checks/` holds the one-off scripts behind the engine findings
recorded in `PLAN.md` (same output as the .NET build, speed on a 100-page
contract, viewer fidelity).

## License

MIT. See `LICENSE`.
