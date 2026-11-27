// The lazily loaded half of `export-pdf-entry.ts`. Importing the export by name from this module,
// rather than `import('ooxml-core/docx/pdf-export')` there, lets a bundler keep the writer out of
// the editor's initial download: Print Layout imports the page marks from the same entry.
export { exportDocxPdf } from 'ooxml-core/docx/pdf-export';
