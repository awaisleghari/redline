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

Then open the address it prints (normally http://localhost:5173).

## Tests

```bash
npm test
```

runs the unit tests (Vitest).

```bash
npm run test:e2e
```

runs the browser tests (Playwright, in Chromium). The first time, install the
browser with `npx playwright install chromium`.

```bash
npm run test:all
```

runs the type check and both test suites.

## License

MIT. See `LICENSE`.
