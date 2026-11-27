# The PDF download on real contracts: the problem, in full

Written 27 November 2026. Self-contained: it can be handed to an outside engineer as it stands.

## 1. What the product does

Redline (useredline.dev) compares two Word files in the browser and produces a legal-style redline. It is built as two halves that share no code:

- **Producer**: makes the outputs. The compare engine (`docxodus` 17, WebAssembly, in the browser) writes a Word file with tracked changes; our code turns that into a **static redline** (`.docx` with insertions in blue double underline, deletions in red strikethrough, moves in green, formatting changes shaded).
- **Checker**: independently re-reads every output and proves it against the two uploads. Nothing is offered for download unless the Checker passes it. This is the product's whole promise: a redline that is wrong is withheld, never shown.

The **PDF** is the one step done on a server: the static redline (already passed by the Checker) is sent to a container running **LibreOffice 24.2 headless on Ubuntu 24.04** (`soffice --convert-to pdf:writer_pdf_Export`), which keeps and logs nothing. The PDF comes back to the browser, where the Checker's **check 9** reads it with PDF.js and compares it with the static redline. If check 9 fails, the PDF is withheld and the user is told so; the two Word files are still offered.

## 2. The symptom

On generated test documents the PDF passes (162 PDF tests, including 88 deliberately wrong PDFs that must be, and are, refused). On **real contracts it is almost always withheld**: of 20 public procurement contracts and bidding documents (Irish public works contracts, Islamic Development Bank bidding documents, New Zealand templates; 15 to 300 pages), **1 PDF was released**. The same contracts' Word outputs are released 87 times out of 87.

The converter never fails. It is always check 9 that refuses.

## 3. What check 9 demands

For the PDF to be released, **every character the PDF prints must be accounted for by the static redline, in order, and every character of the redline must be printed**:

1. Text. The redline's text (main text, tables cell by cell, footnotes, endnotes, headers and footers page by page, text boxes) must appear in the PDF complete and in order. Things a page adds (clause numbers, bullets, footnote marks, tab leader dots, page numbers) are allowed only in exact places and are verified (the clause numbers against the Checker's own numbering).
2. Nothing extra. Any character of the PDF not claimed by step 1 fails the check.
3. Marks. Each changed character must be in its change's colour and carry its mark, read from the lines the PDF draws (one line through deleted text, two under inserted text, and so on); unchanged text must carry none.
4. Place. Cells of a row side by side in order; rows one below the other; header above and footer below the page's text; text running left to right, top to bottom.

To read a PDF in "order" at all, the Checker uses the structure LibreOffice writes into its PDFs (a "tagged PDF": paragraphs, tables, rows, cells, list labels, notes), plus the position of every character on the page.

## 4. Why real contracts fail: two different groups

Each failing contract was opened and the first failure traced to its cause.

### Group A: the PDF really does not show what the Word file holds (about a third)

Here check 9 is **right**. LibreOffice is not Word; it lays the document out with its own engine and sometimes prints something different:

- **Text cut off by a fixed row height.** A table row set to "exactly 0.5 cm" holding two lines of text. Word clips the second line on screen and in print, though it is in the file. LibreOffice clips it too. The words are in the redline and not in the PDF.
- **A dropped character.** A closing bracket at the very end of a full line is not printed by LibreOffice (seen on two contracts).
- **Equations** are printed as pictures, so their characters are not text in the PDF.
- **Cross-reference fields** are recalculated by LibreOffice on load. Where it cannot resolve one, it prints "Error: Reference source not found" in place of the clause number Word had stored.
- One more contract where words are nowhere in the PDF; cause not yet examined.

### Group B: the words are all in the PDF, but not where the Checker's reading expects them (about two thirds)

Here check 9 is **wrong to refuse**: the PDF is a faithful print.

- **Floating frames and text boxes.** The PDF's structure lists a floating box where it sits on the page; the Checker reads it where it is anchored in the text. The streams fall out of step.
- **The matcher then loses its place.** These documents repeat whole clauses (standard forms). Having lost step, the matcher finds the same wording pages later and resumes there, so the failure is reported far from its cause and everything between is counted as "extra".
- **Footers of several lines**, or with the page number set apart in a frame or table (8 of 55 contracts).
- **Symbol-font bullets** (Wingdings and the like) in front of paragraphs, which arrive in the PDF as characters the Checker does not expect (6 of 55).

## 5. The stumbling block

Two things, and they are different in kind.

**(a) The PDF is made by a different layout engine from the one the document was written for.** LibreOffice re-flows a Word document with its own rules for line breaking, fonts, tables, fields and floating objects. Most of the time the words come out the same and only the page breaks differ; sometimes (Group A) content differs. No amount of checking fixes this; it can only be avoided (by changing what is sent to the converter) or removed (by changing the converter).

**(b) Proving a PDF correct means reconstructing reading order from a format that does not have one.** A PDF is a set of drawing instructions per page. Check 9 demands one ordered stream of characters equal to the document's stream. That is a much stronger demand than "the PDF shows this document", and every layout feature that takes text out of the main flow (floats, multi-column footers, frames, repeated table headings, rows split across pages) needs its own reconstruction rule. We have written those rules one at a time; real contracts keep producing the next one.

**Why Litera does not have this problem.** Litera runs inside Microsoft Word on the user's computer. Its PDF is Word's own "Save as PDF" of the redline: the same engine that laid the document out prints it, so (a) cannot arise. And Litera does not independently verify its PDF against the Word file at all, so (b) cannot arise. We chose a browser-only product with no Word installed and an independent check on every output; the PDF step is where that choice costs most.

## 6. Ways forward

They are not exclusive.

1. **Send LibreOffice a copy prepared for it** (attacks Group A). Before conversion, make a PDF-only copy of the static redline in which the things LibreOffice is known to get wrong are neutralised: fields frozen to the text Word stored (no recalculation, so no "Reference source not found"), fixed row heights turned into minimum heights (nothing clipped; the PDF would then show text Word hides, which for a redline is arguably the better result), equations left as they are and declared. The Checker would still compare the PDF with the released static redline's words, so a mistake in preparing the copy is caught like any other.
2. **Match by structure, not as one stream** (attacks Group B at the root). The PDF's own structure already divides it into paragraphs, cells, notes, frames. Instead of one ordered stream for the whole document, match each structural piece of the PDF to one paragraph or cell of the redline, and require: every piece matched, every paragraph matched, pieces of the main flow in order, floating pieces anywhere on the page their anchor is on. Repeated clauses can no longer be confused, because a piece is matched whole and neighbours pin it down. This is a rewrite of the matching core of check 9 (the parts that verify colours, marks, numbers and positions stay).
3. **A different converter** (removes cause (a)). Word itself is the only engine that prints a Word file exactly as Word does. Options: Microsoft's own cloud conversion (Graph API), or a commercial library (for example Aspose.Words). Both cost money, and the cloud route sends the document to a third party, which the product currently promises not to do beyond our own container. Check 9 would still be needed, and Group B would remain.
4. **Offer the PDF with a plain warning when only Group A differences remain.** Not recommended as the main route: it weakens the promise that nothing unverified is offered.

**Recommended:** 1 and 2 together, in that order of cheapness: 1 is days, 2 is the real fix and larger. Measure after each on the same 87 contracts. Consider 3 only if, after 1 and 2, LibreOffice's own faults (the dropped bracket, for instance) still withhold a share of real contracts that is not acceptable.

## 7. Where the code is

- Converter: `pdf-container/` (Dockerfile, `server.py`).
- PDF reader: `src/checker/pdf.ts` (288 lines). Check 9: `src/checker/pdf-check.ts` (1,025 lines).
- Tests: `tests/pdf/` run with `npm run test:pdf` (needs Docker); deliberately wrong PDFs in `tests/fixtures/broken-pdf.ts`.
- Tools for one contract at a time: `scripts/corpus/debug/pdf-one.ts`, `pdf-look.ts`, `pdf-look-all.ts`, `pdf-dump.ts` (usage in `PLAN.md`, handover).
- Test contracts: public documents listed in `tests/corpus/sources-heldout-2.json` (downloaded, never committed).
