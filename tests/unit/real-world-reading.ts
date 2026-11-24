/** The Checker's own reading of a document's text, for comparing two files that should read alike. */
import { DocxPackage } from '../../src/checker/package';
import { readStories as read } from '../../src/checker/stories';

export const openPackage = (docx: Uint8Array): Promise<DocxPackage> => DocxPackage.open(docx);
/** Every story's key and paragraphs, as the uploads are read. */
export const readStories = (pkg: DocxPackage): { key: string; paragraphs: string[] }[] => read(pkg, 'accepted').stories.map((story) => ({ key: story.key, paragraphs: story.paragraphs }));
