// Opens a PDF as an editable Word document, like Word's PDF Reflow: the text becomes flowing
// paragraphs with their character and paragraph formatting, repeated heads and feet become the
// header and footer, and pictures stay pictures. The result is approximate; the report says per
// page what was approximated or left out. DOM-free and loaded on demand
// (`ooxml-core/docx/pdf-import`; `ooxml-core/docx/pdf-convert` is the conversion alone).
export * from '../pdf-convert/index';
export { importPdf, type ImportPdfResult } from './load';
