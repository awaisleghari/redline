/**
 * Shows a PDF in the page by drawing each page onto a canvas with PDF.js.
 * (App code: neither Producer nor Checker. The PDF shown here has already
 * passed the Checker.)
 */
import * as pdfjs from 'pdfjs-dist';
import workerUrl from 'pdfjs-dist/build/pdf.worker.min.mjs?url';

pdfjs.GlobalWorkerOptions.workerSrc = workerUrl;

/** Draw every page of the PDF into `container`, replacing what was there. Returns the number of pages. */
export async function showPdf(container: HTMLElement, pdf: Uint8Array): Promise<number> {
  container.replaceChildren();
  const task = pdfjs.getDocument({ data: pdf.slice() });
  const document_ = await task.promise;
  const scale = Math.max(1, Math.min(2, window.devicePixelRatio || 1)) * 1.3;
  for (let pageNumber = 1; pageNumber <= document_.numPages; pageNumber++) {
    const page = await document_.getPage(pageNumber);
    const viewport = page.getViewport({ scale });
    const canvas = document.createElement('canvas');
    canvas.width = Math.floor(viewport.width);
    canvas.height = Math.floor(viewport.height);
    canvas.style.width = `${Math.floor(viewport.width / scale * 1.3)}px`;
    canvas.setAttribute('role', 'img');
    canvas.setAttribute('aria-label', `Page ${pageNumber} of ${document_.numPages}`);
    container.append(canvas);
    await page.render({ canvas, canvasContext: canvas.getContext('2d')!, viewport }).promise;
    page.cleanup();
  }
  return document_.numPages;
}
