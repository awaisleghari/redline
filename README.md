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

Then open the address it prints (normally http://localhost:5173). At this
stage it is a bare development page: choose an original and a modified `.docx`,
press Compare, read the check report, and download the result as a Word file
with tracked changes (offered only if every check passed).

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
npm run test:all
```

runs the type check and all three.

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
