// `ooxml-core/docx/pdf-convert`: the light half of `ooxml-core/docx/pdf-import`: the conversion
// to .docx bytes without the .docx reader, for a host that already loads .docx files (the Word
// editor loads it only when a PDF is opened).
export {
	convertPdf,
	PDF_IMPORT_NOTICE,
	type ConvertPdfResult,
	type ImportPdfOptions,
} from '../pdf-import/import';
export type { PdfImportReport, PdfOcrProvider, PdfOcrWord } from '../pdf-import/report';
export { isPdf, PdfPasswordError, type PdfPermissions } from '../../pdf/index';
