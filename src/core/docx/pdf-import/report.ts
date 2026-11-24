import type { PdfPermissions } from '../../pdf/index';

/** What the conversion did with each page and with the file. */
export interface PdfImportReport {
	pages: { index: number; notes: string[]; scanned?: true }[];
	/** Font families the text uses (a missing one is substituted when the document is shown). */
	fonts: string[];
	encrypted: boolean;
	/** The file's permission flags, when it is encrypted. */
	permissions?: PdfPermissions;
	/** The PDF was signed: a converted document cannot keep its signature. */
	signed: boolean;
	/** Notes on the whole conversion. */
	notes: string[];
}

/** A word an OCR engine found on a scanned page: its text and box in page points (y down). */
export interface PdfOcrWord {
	text: string;
	x: number;
	y: number;
	width: number;
	height: number;
}

/**
 * Optional OCR for scanned pages, supplied by the host: given the page's picture (PNG or JPEG
 * bytes) and the page size in points, it returns the words with their positions.
 */
export type PdfOcrProvider = (page: {
	index: number;
	width: number;
	height: number;
	image: {
		data: Uint8Array;
		format: 'png' | 'jpeg';
		x: number;
		y: number;
		width: number;
		height: number;
	};
}) => Promise<PdfOcrWord[]>;
