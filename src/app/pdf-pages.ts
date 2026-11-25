/**
 * Reads where a PDF's pages begin, so the viewer's "pages" layout can break the
 * redline at the same places as the file the user downloads.
 * (App code: neither Producer nor Checker. The PDF read here has already passed the Checker.)
 */
import * as pdfjs from 'pdfjs-dist';
import workerUrl from 'pdfjs-dist/build/pdf.worker.min.mjs?url';
import type { PdfPages } from './redline-view';

pdfjs.GlobalWorkerOptions.workerSrc = workerUrl;

/** The shape of the PDF's first page and the text of every page, in order. */
export async function readPdfPages(pdf: Uint8Array): Promise<PdfPages> {
  const task = pdfjs.getDocument({ data: pdf.slice() });
  try {
    const document_ = await task.promise;
    const pages: string[] = [];
    let ratio = 8.5 / 11;
    for (let pageNumber = 1; pageNumber <= document_.numPages; pageNumber++) {
      const page = await document_.getPage(pageNumber);
      if (pageNumber === 1) {
        const { width, height } = page.getViewport({ scale: 1 });
        if (width > 0 && height > 0) ratio = width / height;
      }
      const content = await page.getTextContent();
      pages.push(content.items.map((item) => ('str' in item ? item.str : '')).join(''));
      page.cleanup();
    }
    return { ratio, pages };
  } finally {
    await task.destroy();
  }
}
