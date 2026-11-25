// The PDF reader: file structure, encryption (the standard security handler), stream filters,
// fonts and content streams, giving what each page shows (`PdfPageContent`). DOM-free; it never
// runs JavaScript or actions in the file and never fetches anything.
export { openPdf, pageMatrix, type PdfDocumentInfo, type PdfReader } from './reader';
export { isPdf, type PdfLimits, type PdfOpenOptions } from './document';
export { PdfPasswordError, type PdfPermissions } from './crypt-types';
export { PdfLimitError } from './filters';
export { familyName } from './fonts';
export { decodeCcitt, type CcittOptions } from './ccitt';
export type {
	PdfFontInfo,
	PdfGradientFill,
	PdfOutlineItem,
	PdfPageContent,
	PdfPageImage,
	PdfPageLink,
	PdfPagePath,
	PdfPageRule,
	PdfPathSegment,
	PdfPatternFill,
	PdfTextSpan,
} from './page-content';
