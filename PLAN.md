# Redline — plan

## Goal

A user opens the website, uploads an original and a modified Word document, and presses Compare. The page shows the redline in a built-in viewer, and the user can download it as:

1. **Word, tracked changes**: native Word revisions the user can accept or reject.
2. **Word, static redline**: Litera style, with changes baked in as formatting (spec below).
3. **PDF**: the static redline as a PDF.

The output must be as faithful as Litera Compare or Word Compare on ordinary commercial contracts. Correctness is proven by automated checks, not judged by eye.

## Out of scope for version 1

- PDF *inputs* (later; needs text extraction and possibly OCR).
- `.doc` (old binary Word) inputs: reject with a clear message.
- Documents that already contain tracked changes: detect them and refuse with "accept or reject existing changes first, then compare".
- Accounts, saved history, server-side storage of documents.
- Any AI model.

## Architecture

The app is hosted on **Cloudflare**. Most of the work happens **inside the user's browser**, so documents never leave the user's computer for comparing, viewing or Word downloads.

| Part | Where it runs | Technology |
|---|---|---|
| Web page and viewer | Browser, served by Cloudflare Workers (static assets) | TypeScript, Vite, plain HTML/CSS (no heavy framework unless clearly needed) |
| Compare engine | Browser (WebAssembly, in a Web Worker so the page stays responsive) | `docxodus` npm package (MIT): a maintained fork of Microsoft's Open-XML-PowerTools `WmlComparer`, with move detection |
| Static redline, cover page (Producer) | Browser (same Web Worker) | TypeScript, JSZip, the browser's XML parser |
| Checker (independent reviewer of every output) | Browser (its own Web Worker) | TypeScript; separate code from the Producer (see below) |
| In-page viewer | Browser | `docxodus` DOCX-to-HTML conversion for the redline; PDF.js (or the browser's built-in viewer) once a PDF exists |
| PDF conversion | **Server**: Cloudflare Container | LibreOffice headless in a Docker image, called from a Worker |

Expected `docxodus` usage (verify against the current README and record the real API here):

```ts
import { initialize, compareDocuments, convertDocxToHtml } from 'docxodus';
await initialize('/path/to/wasm/');
const redlinedDocx = await compareDocuments(originalFile, modifiedFile, { authorName: 'Redline' });
const html = await convertDocxToHtml(redlinedDocx);
```

Early checks to do in Stage 1 and record in the decisions log:
- Does the npm build produce the same comparison as the .NET `WmlComparer` (moves, tables, format changes)?
- How does it perform on a 100-page contract in a normal laptop browser?
- Does `convertDocxToHtml` render tracked changes and the static redline's formatting faithfully enough for the viewer?

**Fallback.** If the browser engine is unsuitable, run the Docxodus command-line tool inside the same Cloudflare Container used for PDFs. This means documents are uploaded for comparison, which changes the privacy promise to users, so ask the owner before switching.

**Container base image.** Use an Ubuntu base image for the PDF container. The build environment's network allowlist includes the Ubuntu package mirrors (`archive.ubuntu.com`, `security.ubuntu.com`) but not Debian's, so Debian-based images will fail to install LibreOffice.

**PDF step and privacy.** The PDF is the only step that sends a file to a server. The page must say so in one plain line before the user clicks. The container processes the file in a temporary folder, returns the PDF, deletes everything, and logs no file names or contents. Install metric-compatible fonts in the image so layout matches Word: `fonts-liberation`, `fonts-crosextra-carlito` (Calibri), `fonts-crosextra-caladea` (Cambria).

**Cloudflare costs.** Containers need the Workers Paid plan (USD 5 per month, with included usage). The owner must approve and set this up; see "Owner actions".

## Licensing

Any open-source license is acceptable, including GPL and AGPL; this is a non-commercial open-source project. Start the project under MIT. If an AGPL or GPL dependency is added, relicense the project to AGPL-3.0 at that point, record it in the decisions log, and keep the repository public (AGPL requires offering source to users of the hosted site; the public repo satisfies that). Link the source code from the site's footer.

## Producer and Checker (most important design rule)

The product has two separate halves:

- **Producer**: makes the outputs (the compare engine, the static redline, the cover page, the PDF).
- **Checker**: an independent built-in reviewer that inspects every output against the two original uploads and decides whether the user may have it. The Producer never grades its own work.

### Independence rules
- Checker code lives in `src/checker/` and **must not import anything from the Producer** (`src/producer/`). Shared low-level helpers (unzipping a .docx, XML namespaces) go in `src/common/` and must stay trivial. Enforce this with an automated test that fails if `src/checker/` imports from `src/producer/`.
- The Checker re-reads everything itself from the raw files: the two uploads and each output. It never trusts counts, labels or metadata produced by the Producer.
- Where the Checker needs an operation the Producer also performs (such as walking tracked changes), it uses its own implementation, written separately.

### What the Checker verifies

For the **tracked-changes Word file**:
1. **Accept all** changes (Checker's own implementation over the XML): the text must equal the modified upload's text.
2. **Reject all** changes: the text must equal the original upload's text.

For the **static redline Word file**:
3. Removing red-strikethrough and green-double-strikethrough text gives the modified text; removing blue and green double-underlined text gives the original text.
4. No tracked changes remain (`w:ins`, `w:del`, `w:moveFrom`, `w:moveTo`, `w:rPrChange`, `w:pPrChange`).
5. The cover-page counts equal the counts the Checker computes itself.

For **every Word output**:
6. **Completeness**: every part that holds text is checked (body, tables, headers, footers, footnotes, endnotes, text boxes). Nothing from either upload is silently missing.
7. **Structure**: the file unzips, required parts exist, every relationship points to a part that exists, comment anchors are balanced and every footnote reference resolves, and the file validates against the Office Open XML schema (find a validator usable from the test suite, or a small .NET tool using the Open XML SDK's validator run in CI).
8. **Unchanged text keeps its formatting**: bold, italic, underline and capitals on text that did not change match the modified upload.

For the **PDF**:
9. The PDF's extracted text, in order, matches the static redline's text, and changed text carries the right colours.

Text-comparison rules: compare normalised text paragraph by paragraph. Normalise only what is safe to ignore, such as how text is split into runs. Never ignore real characters: spaces, punctuation, digits.

### What the user sees
- A **check report** panel next to the viewer: each check with a pass or fail mark, in plain English ("Accepting every change reproduces the modified document: passed").
- A download is offered only when all of its checks pass. If any check fails, the user sees which check failed and the first paragraph where it went wrong, and that output is withheld.
- The check report can be downloaded as a short text file to keep with the redline.

### Checking the Checker
A checker that never fails proves nothing. Build a test suite of **deliberately broken outputs**: a dropped word, a changed digit, a missing footnote, a deletion styled as an insertion, a wrong cover-page count, a leftover tracked change, a broken comment anchor. Every one must be caught by the Checker. Add a new broken case whenever a real bug is found.

### Real Word
Microsoft Word cannot run in automated tests, so representative outputs become owner checks: ask the owner in the chat to open them in real Word.

## Static redline: style spec

| Change | Formatting |
|---|---|
| Inserted text | Blue (`0000FF`), double underline |
| Deleted text | Red (`FF0000`), strikethrough |
| Moved from (original location) | Green (`008000`), double strikethrough |
| Moved to (new location) | Green (`008000`), double underline |
| Formatting-only change | Keep the new formatting; count it on the cover page |

Implementation outline:
- `w:ins`: unwrap the runs; add `w:color` and `w:u w:val="double"` to each run's `w:rPr`.
- `w:del`: unwrap the runs; turn `w:delText` into `w:t` (keep `xml:space="preserve"`); add `w:color` and `w:strike`.
- Moves: same idea for `w:moveFrom` and `w:moveTo`; remove the move range markers.
- Remove `w:rPrChange` and `w:pPrChange` after counting them.
- Apply to every part: body, headers, footers, footnotes, endnotes.
- Keep comments anchored: never drop `w:commentRangeStart`, `w:commentRangeEnd` or `w:commentReference`.

**Cover page** (first page of the static redline and the PDF): "Comparison of", both file names, date and time, a legend showing each style, and counts of insertions, deletions, moves and formatting changes. Plain and professional, in the document's own body font.

## The web interface

One page, simple and calm. Legal users, not designers.
- Two clearly labelled drop zones: **Original** and **Modified**. Accept only `.docx`; reject anything else with a plain message. Size limit 25 MB per file.
- A **Compare** button, a progress indicator, and errors in plain English.
- After comparing: the **viewer** fills most of the page and shows the static redline (cover page first). A summary bar shows the change counts, and the **check report** panel shows the Checker's results.
- Download buttons: Word (tracked changes), Word (static redline), PDF. The PDF button carries the one-line notice that the file is sent to the server for conversion and deleted immediately.
- Version 1.1 (after Stage 7): a side panel listing each change, clickable to jump to it in the viewer.
- Accessible: keyboard usable, readable contrast, works in current Chrome, Safari, Edge and Firefox.
- Footer: link to the source code and a one-line statement that comparison happens in the browser.

## Known hard areas

- **Automatic numbering.** In a static redline, deleted numbered paragraphs still take a number, so clause numbers shift. Investigate freezing list numbers into literal text before conversion. Set out the options and ask the owner in the chat before building.
- Tables (rows added or removed, merged cells).
- Moved clauses: whether they are detected and how they display.
- Headers, footers and footnotes.
- Fields and cross-references ("Section 7.2" references, tables of contents).
- Large documents (100+ pages) in the browser: memory and speed.

## Testing

- Unit tests: Vitest.
- Browser tests: Playwright (upload two files, compare, view, download).
- Generate test pairs in code (for example with the `docx` npm package): a one-word change; a number change (`$300,000` to `$30,000`); an inserted paragraph; a deleted paragraph; a moved paragraph; a table cell change; a header change; a footnote change; a numbered-list change.
- Every bug found gets a regression test built from a generated (non-client) document.

## Stages and acceptance criteria

Work in order. A stage is done only when every criterion passes and the independent review finds no real gaps.

**Stage 0: Setup**
- TypeScript + Vite project, Vitest and Playwright running, MIT `LICENSE`, `README.md`, `.gitignore` (excludes `node_modules`, `dist`, `tests/fixtures/private/`, `.env*`, `.dev.vars`).
- Done when: one placeholder test passes in each runner and the first commit is pushed.

**Stage 1: Compare engine in the browser**
- A bare development page: pick two files, press Compare, download the tracked-changes `.docx`. The engine runs in a Web Worker.
- Generated test pairs as listed under Testing.
- Answer the three early checks under Architecture and record them in the decisions log.
- Done when: every test pair produces the expected tracked changes, checked in code (count and text of insertions, deletions and moves).

**Stage 2: Checker (tracked-changes output)**
- Build `src/checker/` with checks 1, 2, 6 and 7, the import-independence test, and the first set of deliberately broken outputs. Wire it into the compare flow so a failed check withholds the file.
- Done when: every test pair passes the Checker, every deliberately broken output is caught, and the independence test passes.

**Stage 3: Static redline and cover page**
- Producer: implement the style spec and the cover page.
- Checker: add checks 3, 4, 5 and 8, with broken-output cases for each.
- Done when: every test pair's static redline passes all its checks, every broken case is caught, and the owner has checked representative files in Word (asked in the chat).

**Stage 4: Interface, viewer and check report**
- Build the interface described above, running locally (`wrangler dev` or Vite), including the check report panel and withheld downloads on failure.
- Done when: the Playwright test covers upload, compare, viewer showing the redline, the check report, and both Word downloads; refusals for `.doc` files, oversized files and files with existing tracked changes are tested; a forced Checker failure withholds the download in the UI; the owner has reviewed screenshots of the page (asked in the chat).

**Stage 5: PDF conversion**
- Dockerfile with LibreOffice and fonts; a Worker route that sends the static redline to the container and returns the PDF. Build and test the container with Docker inside the cloud session first. Checker: add check 9.
- Done when: the PDF passes check 9, files are deleted after conversion, nothing is logged, and the viewer can show the PDF.
- Running it on Cloudflare needs the paid plan (see Owner actions). Build and test everything you can before that.

**Stage 6: Deploy to Cloudflare**
- `wrangler.jsonc`, deployment script, and a short "how to deploy" section in `README.md`.
- The cloud session cannot reach Cloudflare's API by default. Set out the deployment options for the owner and ask which to use: (a) Cloudflare's own GitHub integration, which deploys automatically when `main` changes (check that it supports Containers); (b) a GitHub Actions workflow with a Cloudflare API token stored as a GitHub secret; (c) allowing `api.cloudflare.com` in the cloud environment and storing the token as a network secret. Never put a token in the repository or in a plain environment variable.
- Done when: the owner has approved the first public deployment, and the live site passes the Playwright test run against its URL.

**Stage 7: Real-world testing**
- Real contracts must never be committed, and cloud sessions see only what is committed. So this stage runs on the owner's own computer: he pulls the session down with `claude --teleport` (or opens the repository in local Claude Code) and places anonymised contract pairs in `tests/fixtures/private/` (git-ignored). Before this stage, give the owner plain step-by-step instructions for that switch.
- Compare results against Word Compare or Litera output for the same pairs; log every discrepancy in the status log, fix the causes, and add a regression test for each fix.
- Done when: the owner signs off on the private test set.

## Owner actions

These need the owner. Ask for them in the chat as early as they are foreseeable, grouped with other pending questions, so they never block work for long.
- Add domains to the cloud environment's allowlist when a needed download is blocked (for example Playwright's browser downloads; see the setup notes in the status log).
- Create a Cloudflare account and approve the Workers Paid plan (needed for Containers), then complete whichever deployment option he chooses at Stage 6.
- Approve the first public deployment.
- Provide anonymised real contract pairs for Stage 7.
- Answer owner checks in the chat (opening files in Word, looking at screenshots).

## Decisions log

Record each significant decision: date, decision, reason.

- 2026-11-22: Hosting on Cloudflare; compare and viewer run in the browser via `docxodus` (WASM); PDF via LibreOffice in a Cloudflare Container. Reason: Cloudflare cannot run the engine or LibreOffice outside a container, and in-browser comparison keeps documents on the user's computer.
- 2026-11-22: GPL and AGPL dependencies are permitted (non-commercial open-source project).
- 2026-11-22: Producing the redline and checking it are separate: an independent built-in Checker verifies every output against the uploads, and failed outputs are withheld. Reason: the producing code must never grade its own work.

## Status log

Update at the start and end of every session: current stage, what was done, what is next, any blockers.

- 2026-11-22: Plan written. Built in Claude Code cloud sessions. Environment set to Custom network access with the default package-manager list plus Playwright's download hosts (`cdn.playwright.dev`, `playwright.download.prss.microsoft.com`, `playwright.azureedge.net`). Next: Stage 0.
- 2026-11-23 (Stage 0, plan): Scaffold a TypeScript + Vite project; Vitest for unit tests (`tests/unit/`), Playwright for browser tests (`tests/e2e/`, Chromium, run against the production build); MIT `LICENSE`, `README.md`, `.gitignore`; one placeholder test per runner; commit and push.
- 2026-11-23 (Stage 0, done): Scaffold in place. `npm run test:all` runs the type check, the unit tests and the browser tests; one placeholder test passes in each runner. Setup notes: the cloud image ships Chromium with Playwright 1.63.0, so `@playwright/test` is pinned to exactly 1.63.0 (a different version would look for a browser that is not installed). Next: Stage 1.
