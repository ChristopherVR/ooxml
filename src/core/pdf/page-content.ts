/**
 * What a page shows, in reading coordinates: points from the page's top-left corner, y down, after
 * the crop box and the page rotation. The document reconstruction (`ooxml-core/docx/pdf-import`)
 * reads only these types, so another PDF engine (PDF.js, PDFium) can supply them instead.
 */
import type { PdfFontInfo } from './fonts';

export type { PdfFontInfo };

/** A run of text drawn in one font, size and colour, on one baseline. */
export interface PdfTextSpan {
	text: string;
	/** Left edge of the first glyph. */
	x: number;
	/** The baseline. */
	y: number;
	/** Advance of the whole span. */
	width: number;
	/** Font size as drawn (the text matrix and the transformation scale included). */
	fontSize: number;
	font: PdfFontInfo;
	/** `#rrggbb`. */
	color: string;
	/** Baseline direction in radians, counter-clockwise (0 for ordinary horizontal text). */
	angle: number;
	/** Text rise (super- and subscripts) as drawn, in points; positive raises. */
	rise: number;
	/** Text rendering mode 3: invisible, as an OCR layer over a scanned page. */
	invisible?: true;
}

/** A picture drawn on the page, with its encoded bytes when the reader could take them. */
export interface PdfPageImage {
	x: number;
	y: number;
	width: number;
	height: number;
	pixelWidth: number;
	pixelHeight: number;
	/** JPEG bytes as they are, or a PNG the reader encoded from the samples. */
	format: 'jpeg' | 'png' | 'unsupported';
	data?: Uint8Array;
	/** Why the picture could not be read (an unsupported codec or colour space). */
	reason?: string;
}

/** A straight line or a rectangle (filled or stroked): table rules, borders, underlines. */
export interface PdfPageRule {
	x0: number;
	y0: number;
	x1: number;
	y1: number;
	kind: 'line' | 'rect';
	stroked: boolean;
	filled: boolean;
	color: string;
	lineWidth: number;
}

/** A link annotation: the area and its target. */
export interface PdfPageLink {
	x: number;
	y: number;
	width: number;
	height: number;
	uri?: string;
	/** A page index for an internal link. */
	page?: number;
}

export interface PdfPageContent {
	index: number;
	width: number;
	height: number;
	spans: PdfTextSpan[];
	images: PdfPageImage[];
	rules: PdfPageRule[];
	links: PdfPageLink[];
	/** What the reader could not take from the page. */
	notes: string[];
}
