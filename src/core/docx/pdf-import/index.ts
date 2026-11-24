// Opens a PDF as an editable Word document, like Word's PDF Reflow: the text becomes flowing
// paragraphs with their character and paragraph formatting, repeated heads and feet become the
// header and footer, and pictures stay pictures. The result is approximate; the report says per
// page what was approximated or left out. DOM-free and loaded on demand
// (`ooxml-core/docx/pdf-import`).
export {
	convertPdf,
	importPdf,
	PDF_IMPORT_NOTICE,
	type ConvertPdfResult,
	type ImportPdfOptions,
	type ImportPdfResult,
} from './import';
export type { PdfImportReport, PdfOcrProvider, PdfOcrWord } from './report';
export { isPdf, PdfPasswordError, type PdfPermissions } from '../../pdf/index';
