// `ooxml-core/docx/pdf-export`: Word's "Save as PDF" without the editor. DOM-free, so it runs in
// Node: `exportDocxPdf` lays the document out and writes real text and vector graphics, measuring
// with the supplied font files (`createFontFileMeasurer`) unless a measurer is given.
// `ooxml-ui/docx`'s `exportDocxPdf` wraps it with the canvas measurer Print Layout uses.
export {
	exportDocxPdf,
	type DecodedPicture,
	type DocxPdfBookmarks,
	type DocxPdfConformance,
	type DocxPdfFontMeasurement,
	type DocxPdfMarkupBalloons,
	type DocxPdfMedia,
	type DocxPdfOmittedKind,
	type DocxPdfOptions,
	type DocxPdfPictureDecoder,
	type DocxPdfReport,
	type DocxPdfResult,
	type DocxPdfUaConformance,
} from './export-pdf';
export {
	createEstimatedMeasurer,
	createFontFileMeasurer,
	type FontFileMeasurer,
	type FontFileMeasurerOptions,
} from './font-measurer';
export { DocxPdfFontFiles, type DocxPdfFontMatch, type DocxPdfFontSource } from './font-files';
export type { PdfFontFile } from '../../pdf/write/pdf-embedded-font';
export * from './page-marks';
