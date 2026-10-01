// Document format detection and loading: DOCX (OPC zip) and Word 97-2003 .doc (OLE2 compound file).
// The legacy reader inlines the shared `@christophervr/ole2` codecs at build time (see tsup.config.ts).
export { detectDocumentFormat, loadDocument, type DocumentFormat } from './detect.js';
export { LegacyDocError, loadLegacyDoc } from './legacy-doc.js';
