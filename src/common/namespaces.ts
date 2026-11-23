/** Shared, deliberately trivial: XML namespace names used in Word files. */
export const NS = {
  w: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
  r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  mc: 'http://schemas.openxmlformats.org/markup-compatibility/2006',
  rels: 'http://schemas.openxmlformats.org/package/2006/relationships',
  contentTypes: 'http://schemas.openxmlformats.org/package/2006/content-types',
} as const;

/** Relationship type suffixes (the part after ".../relationships/"). */
export const REL = {
  officeDocument: 'officeDocument',
  header: 'header',
  footer: 'footer',
  footnotes: 'footnotes',
  endnotes: 'endnotes',
  comments: 'comments',
} as const;
