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
| Compare engine | Browser (WebAssembly, in a Web Worker so the page stays responsive) | `docxodus` npm package (MIT): a .NET document-comparison engine compiled to WebAssembly, with move detection. It began as a fork of Microsoft's Open-XML-PowerTools `WmlComparer`; since its version 11 it uses its own engine (see decisions log, 2026-11-23) |
| Static redline, cover page (Producer) | Browser (same Web Worker) | TypeScript, JSZip, the browser's XML parser |
| Checker (independent reviewer of every output) | Browser (its own Web Worker) | TypeScript; separate code from the Producer (see below) |
| In-page viewer | Browser | `docxodus` DOCX-to-HTML conversion for the redline; PDF.js (or the browser's built-in viewer) once a PDF exists |
| PDF conversion | **Server**: Cloudflare Container | LibreOffice headless in a Docker image, called from a Worker |

Real `docxodus` usage (verified against version 16.1.0 in Stage 1; see the decisions log for the findings behind it):

```ts
// In the compare Web Worker (src/producer/compare.worker.ts)
import { initialize, compareDocuments, convertDocxToHtml } from 'docxodus/core'; // engine only, no editor UI
await initialize('https://<site>/wasm/');           // folder that contains _framework/ (copied by scripts/copy-wasm.mjs)
const redlinedDocx = await compareDocuments(originalBytes, modifiedBytes, { authorName: 'Redline' }); // Uint8Array in, Uint8Array out
const html = await convertDocxToHtml(redlinedDocx, {
  renderTrackedChanges: true,        // off by default: without it changes are not shown at all
  renderHeadersAndFooters: true,     // off by default
  renderFootnotesAndEndnotes: true,  // off by default
});
```

Two rules for the worker, both learned the hard way in Stage 1: register the message handler with `addEventListener('message', …)`, never `self.onmessage = …` (the .NET runtime refuses to start in a worker that has `onmessage` set); and answer `requestAnimationFrame` with a timer inside the worker (the engine waits two animation frames before each comparison).

Early checks (answered in Stage 1; details in the decisions log):
- Does the npm build produce the same comparison as the .NET engine (moves, tables, format changes)? **Yes, byte for byte.**
- How does it perform on a 100-page contract in a normal laptop browser? **About 2.4 seconds.**
- Does `convertDocxToHtml` render tracked changes and the static redline's formatting faithfully enough for the viewer? **Yes, with the three options above switched on.**

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
- 2026-11-23: **Engine version and entry point.** Use `docxodus` 16.1.0 through its `docxodus/core` entry (engine only), in our own Web Worker (`src/producer/compare.worker.ts`), not the package's ready-made worker. Reason: the static redline (Stage 3) is to run in the same worker, and our own worker keeps the message protocol under our control. Finding: the plan assumed the engine was Microsoft's `WmlComparer`; that engine was removed from Docxodus in its version 11 and replaced by Docxodus's own comparison engine. `compareDocuments` is the documented front door to it, in both the .NET and the npm builds.
- 2026-11-23: **Early check 1 (same result as .NET?): yes.** The npm (WebAssembly) build and the .NET build (NuGet `Docxodus` 16.1.0, `DocxCompare.Compare`) were run on the nine test pairs, a formatting-only change, a table row insertion and the 100-page contract. Every part of every output file was byte-for-byte identical between the two. Because `WmlComparer` no longer exists in any current Docxodus release, the comparison is against the current .NET engine, which is the same source code. Reproduce with `scripts/early-checks/` (needs the .NET 10 SDK: `apt-get install dotnet-sdk-10.0`).
- 2026-11-23: **Early check 2 (100-page contract in the browser): fast enough.** A generated contract of 1,000 paragraphs and about 46,000 words (about 100 pages) with 195 edits compared in about 2.4 seconds inside the Web Worker in Chromium, about 2.9 seconds from click to finished, and the page's own thread never paused for more than about 60 ms. First-time start-up of the engine adds under a second locally; over the internet the first visit also downloads the engine's files, about 27 MB before compression (measured: 26.9 MB, 40 files, all fetched on first use), which the browser then caches; the web host's compression should cut this substantially, to be measured at Stage 6. Caveats: measured in headless Chromium on the cloud machine, not on a real laptop, and on generated text; real contracts with heavy formatting, large tables or images may be slower. A browser test (`tests/e2e/large-document.spec.ts`) now guards this with a 30-second ceiling. Real-document timing belongs to Stage 7.
- 2026-11-23: **Early check 3 (viewer fidelity): good enough.** With `renderTrackedChanges`, `renderHeadersAndFooters` and `renderFootnotesAndEndnotes` switched on (all are off by default), `convertDocxToHtml` shows insertions and deletions as `<ins>`/`<del>` in the body, tables, headers and footnotes, and shows moves distinctly (`rev-move-from` / `rev-move-to`). The HTML does not say which moved-from belongs to which moved-to (its `data-move-id` is just each mark's own id), so the Stage 4 change list must pair moves itself from the Word file. In a numbered list, the HTML also shows each renumbered item's old and new number as a deletion and an insertion (five marks for one real inserted item), so change counts must never be taken from the HTML. The static redline's formatting converts exactly: blue double underline, red strikethrough, green double strikethrough and green double underline each become the matching CSS. Checked on generated documents by reading the HTML and by screenshots; the owner's eye on real output is an owner check in Stages 3 and 4.
- 2026-11-23: **Engine behaviours to remember for later stages** (observed, not yet acted on). (a) A formatting-only change to part of a paragraph makes the engine write a formatting-change record on every run of that paragraph, including runs whose formatting did not change (three records for one real change), so Stage 3 must count only records where old and new formatting differ, and the Checker must count the same way independently. (b) When a list item is inserted, the engine also records the old number of each renumbered item (`w:numberingChange`); relevant to the automatic-numbering question in Stage 3. (c) Within a replaced phrase the engine writes the insertion before the deletion. (d) The engine silently accepts any tracked changes already in the uploads before comparing; the plan still requires detecting and refusing such files (Stage 4), so that check must be ours. (e) All tracked changes are dated 2000-01-01 by default (the engine pins dates so output is repeatable); decide in Stage 3 whether to stamp the real date.
- 2026-11-23: **Checker reads XML with its own parser.** Web Workers have no built-in XML parser (`DOMParser` exists only on the page), so the plan's "the browser's XML parser" is not available where the Checker runs. The Checker uses `saxes` (ISC licence), a strict streaming parser, feeding its own small read-only tree (`src/checker/xml.ts`). The Producer will use a different parser (`@xmldom/xmldom`) in Stage 3, so a parser bug cannot affect both sides the same way. The only shared code is `src/common/` (unzip/zip via JSZip, and namespace names), kept under 80 lines by a test.
- 2026-11-23: **How the Checker compares text** (checks 1 and 2). Each file is read into "stories": main text, each header and footer (through the section that refers to it), each footnote and endnote (in the order the text refers to them, so renumbered ids do not matter), and text boxes. Each story is a list of paragraphs; comparison is exact, paragraph by paragraph, after joining runs. Tabs, line breaks, non-breaking hyphens and empty paragraphs all count. Footnote marks, endnote marks and embedded objects (pictures, shapes) leave a marker in the text so their loss is noticed. Tracked paragraph marks, table rows and table cells are honoured (an inserted row vanishes on reject, and so on), and a row left behind empty counts as a difference. Uploads are read as Word's own Compare reads them (any tracked changes already in them count as accepted).
- 2026-11-23: **What the Checker does not compare, deliberately.** (a) Reviewer comments: their text is not compared, because a modified draft normally has different comments from the original and the redline can only carry one set; comment anchors are verified structurally instead (check 7). (b) (Superseded: field codes are now compared too, see "Stage 2 review findings".) (c) A template's glossary/building-blocks part: never shown in the document. (d) Formatting: that is check 8, Stage 3. (Corrected after the Stage 2 review, which showed this sentence was at first untrue: see the entry "Stage 2 review findings" below.) Text in charts and diagrams beside the document is now read and compared; a Word-markup part holding text that nothing refers to, an embedded document, or text inside an element the Checker does not know, fails check 6, in the uploads as well as in the output.
- 2026-11-23: **Schema validation runs in the test suite, not in the browser** (as the plan allows). `npm run test:schema` validates every generated output with Microsoft's Open XML SDK validator through a small .NET tool (`scripts/validate-docx/`), and proves the validator works by requiring it to reject a deliberately invalid file. It needs the .NET 10 SDK (`apt-get install -y dotnet-sdk-10.0`). A full schema validator is not practical inside the browser, so for each user's actual output the in-browser Checker performs the structural checks only (well-formed XML, content types, relationships, comment anchors, note references, move pairing). Result today: all 19 generated outputs, including the 100-page contract, validate.
- 2026-11-23: **Checker rule corrected: move pairing.** The first version of check 7 required each move name to be used exactly once. It then rejected the engine's genuine output for the 100-page contract. Investigation: when several consecutive paragraphs move together, the engine gives each paragraph its own origin/destination markers under one shared name. The Open XML SDK validator accepts this and accept-all/reject-all reproduce both documents exactly, so the rule was stricter than the file format. It now requires what matters: every move name has as many destinations as origins, and none is missing. Regression material added: a three-paragraph block-move pair, and two broken cases (one destination renamed; one destination dropped), both caught. How Word displays such a move is an owner check.
- 2026-11-23: **Checker tightened: content types.** A broken case (the footnotes part declared as generic XML) passed the first version of check 7. The Checker now requires parts with a defined role (main document, headers, footers, footnotes, endnotes, comments, styles, numbering, settings) to be declared as exactly that.
- 2026-11-23: **Engine behaviours seen in Stage 2.** It does not compare inside text boxes: a changed text box is shown as the whole old box deleted and the whole new box inserted (correct, but coarse). It treats a hyphen as a word boundary ("exclusive" to "non-exclusive" is shown as "non-" inserted). It keeps the modified document's comments and their anchors.
- 2026-11-23: **Stage 2 review findings (Checker strengthened).** The independent review attacked the Checker with 33 new kinds of damage and found six kinds of wrong output that passed. Each is fixed in the Checker and kept as a broken case: (1) a run that lost its "preserve spaces" flag (Word then drops the space and two words run together): the Checker now trims text exactly as Word does; (2) table structure (a table dissolved into paragraphs, rows merged, text shifted to the next cell): tables, rows and cells are now part of what is compared; (3) text added to a footnote separator: separators are now read; (4) content stored in two alternative forms that say different things: now refused unless both forms say the same; text inside an element the Checker does not know is refused too; (5) an embedded HTML or other document that Word would import: refused; (6) text outside Word's own text elements (watermarks, drawing text, equations, chart and diagram parts): now read and compared. Also added on the reviewer's prompting: page breaks and column breaks are told apart from line breaks, and field codes are compared (so a cross-reference cannot be silently repointed). (How charts and diagrams are read was reworked in the second review round, next entry.)
- 2026-11-23: **Second review round (Checker strengthened again, and one false alarm removed).** The re-check confirmed the six fixes and that no test had been weakened, then found: (a) a false alarm: any document containing a chart was always withheld, because the engine copies each chart into a new part and leaves the old one behind unused, and the Checker was reading every chart part in the package; (b) further wrong outputs that passed. Fixed: charts, diagrams, pictures and embedded files are now read at the place the document anchors them, by following the link (a chart's text is compared; a picture or embedded file is compared by a fingerprint of its bytes), and parts nothing links to are ignored because Word never shows them; where content is stored in several alternative forms, every form must read the same, including which footnote it points to; anything wrapped in an element the Checker does not know is refused, not only text; separate text boxes are kept separate; a hyperlink's destination is compared (the confirmation pass then showed my first version of this wrongly withheld documents in which a link's address had simply changed; the engine tracks such a change as the old link's words deleted and the new link's words inserted, so the Checker now compares each view's links, ignoring a link whose words are all removed in that view, and a changed-address pair was added to the test pairs); deleted-text markup outside a deletion, and ordinary text markup directly inside one, fail the structure check. The reviewer's 21 "ordinary Word document" cases (text boxes, charts, fields, equations, watermarks, embedded Excel and Word objects, bound cover-page controls, nested tables) all pass. Correction to an earlier entry: a redline is withheld only when a chart, diagram or picture actually differs between the two documents and the engine has not shown it as a change.
- 2026-11-23: **Known limits of the Checker after Stage 2** (deliberate, declared, to be revisited). The Checker does not yet verify: hidden text (`w:vanish`), added to check 8 in Stage 3; automatic numbering (a clause number changing because numbering was altered), which gets its own check in Stage 3; page and table layout that changes what is visible without changing text (margins, "different first page", a table row marked hidden, cells merged so that one cell's text is covered); content controls bound to document properties (Word may refresh their text from the property when the file is opened; the Checker compares the text as stored); settings that make Word recalculate fields on opening; document properties and custom XML data stores; reviewer comments' text. Content stored in a single alternative form with no fallback is read as written.
- 2026-11-23: **How the current docxodus engine differs from Microsoft's original `WmlComparer`** (researched at the owner's request before Stage 3; sources: the Docxodus repository's `docs/architecture/ir_diff_engine.md` and `CHANGELOG.md`, and the original `WmlComparer.cs` source in Open-Xml-PowerTools; the claims about the current engine are the Docxodus project's own, except where marked "we saw").
  - *History.* Docxodus began as a fork of `WmlComparer`. Its own engine ("DocxDiff") became the default in its version 8 and the old engine was removed entirely in version 11; it cannot be switched back on. The new engine was developed against the old one as a reference, and is tuned to match Word's own Compare, not `WmlComparer`.
  - *Moves.* The original has none: its source lists move markup among elements it refuses to process, and its change types are only "inserted" and "deleted", so a moved clause comes out as a deletion plus an insertion. The current engine writes native Word moves, including a moved clause with a word or two changed, a block of several paragraphs, moved tables and reordered rows. We saw: single and three-paragraph moves are detected; a multi-paragraph move shares one name across its paragraphs.
  - *Tables.* The original marks inserted and deleted rows, compares rows pairwise when the tables line up, and otherwise (notably with merged cells) replaces the whole table: old table deleted, new table inserted. The current engine compares inside cells, tracks individual cells added or removed (`w:cellIns` / `w:cellDel`, which Word's own Compare never writes) and tracks changes to cell, row and table properties; it falls back to whole-table replacement when it cannot line the tables up. We saw: cell edits and row insertions and deletions are tracked precisely.
  - *Formatting.* The original records no formatting changes at all. The current engine records character-level and paragraph, table and section-level formatting changes. By default it compares only the formatting properties it models, so a change in something it does not model (its documentation gives run shading as an example) is missed. We saw defects here: it writes formatting-change records on text whose formatting did not change, and the "old formatting" stored in a record is empty even when the old text was bold, so rejecting a formatting change in Word does not restore the old formatting. The engine's formatting records are therefore not a sound basis for counting.
  - *Headers, footers, notes.* The original ignores header and footer differences; the current engine compares them (we saw this working). Both compare footnotes and endnotes.
  - *Granularity.* The original has a `DetailThreshold` setting (how finely to split a heavily rewritten passage); the current engine has no equivalent and deliberately shows a rewritten passage as one deletion and one insertion. Both split words on a configurable list of separators (we saw the hyphen splitting "non-exclusive").
  - *Uploads that already contain tracked changes.* Both accept them before comparing, as Word's Compare does.
  - *Not established.* How faithfully either matches Word's Compare on real contracts: all such claims are the Docxodus project's own, and only Stage 7 (real documents against Word or Litera output) can test them.
  - *What this means for us.* Nothing in the plan depended on `WmlComparer` specifically, and on moves, tables, headers and footers the current engine does more than the original did. The one area to distrust is formatting-change records, which is why the options for formatting-only changes go to the owner.

## Status log

Update at the start and end of every session: current stage, what was done, what is next, any blockers.

- 2026-11-22: Plan written. Built in Claude Code cloud sessions. Environment set to Custom network access with the default package-manager list plus Playwright's download hosts (`cdn.playwright.dev`, `playwright.download.prss.microsoft.com`, `playwright.azureedge.net`). Next: Stage 0.
- 2026-11-23 (Stage 0, plan): Scaffold a TypeScript + Vite project; Vitest for unit tests (`tests/unit/`), Playwright for browser tests (`tests/e2e/`, Chromium, run against the production build); MIT `LICENSE`, `README.md`, `.gitignore`; one placeholder test per runner; commit and push.
- 2026-11-23 (Stage 0, done): Scaffold in place. `npm run test:all` runs the type check, the unit tests and the browser tests; one placeholder test passes in each runner. Setup notes: the cloud image ships Chromium with Playwright 1.63.0, so `@playwright/test` is pinned to exactly 1.63.0 (a different version would look for a browser that is not installed). Next: Stage 1.
- 2026-11-23 (Stage 0, review): Independent review found the tests passing and two gaps: the commit was not on GitHub (the push was refused; see blocker below) and `@playwright/test` was pinned loosely (`^1.63.0`) although this log said "exactly". The pin is now exact.
- 2026-11-23 (Stage 1, plan): Verify the real `docxodus` API; run the engine in our own Web Worker; bare page (two file pickers, Compare, download); generate the nine test pairs in code with the `docx` package, each with the tracked changes a correct comparison must produce; a test-only reader that pulls insertions, deletions and moves out of the raw XML; unit tests (engine in Node) and browser tests (real page, real worker, real download); answer the three early checks with scripts kept in `scripts/early-checks/`.
- 2026-11-23 (Stage 1, done): All of the above built. All nine pairs produce exactly the expected tracked changes (count and text of insertions, deletions and moves, in body, header, footer, footnotes and endnotes), both in Node and through the real page in Chromium. The three early checks are answered in the decisions log: the engine is suitable, so the fallback (server-side comparison) is not needed. Two engine start-up problems were found and fixed at the root (recorded under Architecture). Tests at that point: 23 unit, 13 browser, all passing (`npm run test:all`); 26 unit after the review fixes below. Next: Stage 2 (Checker).
- 2026-11-23 (blocker, owner action): `git push` is refused with "Claude doesn't have GitHub access to awaisleghari/redline" (reading works, writing does not). Until the owner connects GitHub / installs the Claude GitHub App on the repository, work is committed locally in the session only and no pull request can be opened.
- 2026-11-23 (Stage 1, review): Independent review re-ran everything (23 unit and 13 browser tests passing; .NET comparison identical; timings confirmed; three deliberate wrong expectations each caught) and found five real gaps, all fixed: (1) the engine download size was understated (about 27 MB, not 18 MB); (2) the log wrongly said the viewer pairs moves; (3) the log omitted that the viewer shows renumbered list items as changes; (4) the tests could not tell a move's direction; (5) the test reader ignored tracked paragraph marks, tracked table rows and changes inside hyperlinks or content controls. The reader now records move order, paragraph marks and rows and looks inside those containers, and the pairs assert them. Re-checked by a second independent review: no real gaps; it confirmed each fix by deliberately breaking the code or an expectation and seeing the matching test fail.
- 2026-11-23 (blocker resolved): The owner connected GitHub; the branch is pushed and a pull request to `main` is open. The session cannot reach the Cloudflare API and no deployment has been attempted.
- 2026-11-23 (end of session): Stages 0 and 1 complete. 26 unit tests and 13 browser tests pass. No owner checks arose in these stages (nothing yet needs opening in Word). Next: Stage 2 (Checker for the tracked-changes output). No blockers.
- 2026-11-23 (Stage 2, plan): Build `src/checker/` with its own XML reader, its own package reader and its own accept-all / reject-all reading; checks 1, 2, 6 and 7; run it in its own Web Worker; a flow module that is the only route by which an output reaches the user and that releases it only on a complete pass; the import-independence test; a suite of deliberately broken outputs, each required to fail exactly the checks meant to catch it; schema validation with the Open XML SDK in the test suite; more generated pairs for the parts the Checker must read (footer, endnote, text box, comment, table rows, formatted text, several changes at once).
- 2026-11-23 (Stage 2, done): All built. 18 generated pairs (the plan's nine plus nine more) and the 100-page contract pass all four checks. 44 deliberately broken outputs (55 after the review, below) are each caught, by exactly the expected checks. The independence test passes and proves its own detector against 15 ways of importing. In the browser, a tampered output is withheld (no link, no address behind it) and the report names the failed check and the first paragraph that went wrong. The Checker reads the 100-page contract in about 0.1 seconds. Two findings on the way, both in the decisions log: one Checker rule was too strict (move names) and one too lax (content types); each was fixed in the Checker with new broken cases. Tests: 166 unit, 29 browser, schema validation of 19 outputs.
- 2026-11-23 (Stage 2, review): The independent review confirmed the tests, the independence of the Checker, the flow (no route to the user without a pass) and the browser withholding test, and then found six kinds of wrong output the Checker passed, plus one untrue sentence in this file. All fixed in the Checker (never in the tests) and recorded as 11 further broken cases; details in the decisions log. Tests now: 185 unit (55 broken outputs), 29 browser, 19 outputs schema-valid.
- 2026-11-23 (Stage 2, second review round): The re-check confirmed the fixes and found one false alarm (documents with charts were always withheld) and several more exotic wrong outputs. All fixed in the Checker; the reviewer's own probe was re-run and shows each one caught and every ordinary-document case passing. What remains unverified is listed under "Known limits" in the decisions log. Tests now: 212 unit (59 broken outputs), 29 browser, 38 outputs schema-valid.
- 2026-11-23 (Stage 2, confirmation): A third independent pass found the acceptance criteria met and no wrong output passing; every remaining "passes" item in the reviewer's probe is a declared known limit. It found one more false alarm (ordinary hyperlink edits were withheld), now fixed and re-checked with the reviewer's own script: all 8 ordinary-document picture and link cases pass and all 7 picture attacks are caught. Final tests: 216 unit (60 broken outputs, 19 pairs), 30 browser, 40 outputs schema-valid, all passing. Stage 2 complete; pull request merged. Next: Stage 3.
