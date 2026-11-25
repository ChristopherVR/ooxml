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
	/** Drawn as an artifact (page furniture such as running heads, in a tagged file). */
	artifact?: true;
	/**
	 * In a tagged file, the standard structure type of the block the text belongs to (`P`, `H1`
	 * to `H6`, `LI`, `LBody`, `TD`, `Caption`...), after the role map.
	 */
	role?: string;
	/** In a tagged file, an id shared by the spans of one structure element (one paragraph). */
	element?: number;
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
	/** Alternate text, from the figure's structure element or its marked content. */
	alt?: string;
}

/** A linear or radial gradient fill read from a PDF shading or shading pattern. */
export interface PdfGradientFill {
	kind: 'linear' | 'radial';
	/** Clockwise angle in degrees (DrawingML convention: 0 left-to-right, 90 top-to-bottom). */
	angle?: number;
	stops: { pos: number; color: string }[];
}

/** A tiling pattern fill read from a PDF PatternType 1 dictionary. */
export interface PdfPatternFill {
	preset: string;
	fg: string;
	bg?: string;
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
	/** The outline colour of a rectangle both filled and stroked (`color` is then the fill). */
	strokeColor?: string;
	/** Opacity below 1 (the graphics state's `ca` or `CA`). */
	alpha?: number;
	/** The dash pattern of a stroke, in points. */
	dash?: number[];
	/** Painted with a pattern or shading: `color` is not what it shows. */
	pattern?: true;
	/** Gradient fill from a PDF shading or shading pattern. */
	gradient?: PdfGradientFill;
	/** Tiling pattern fill from a PDF PatternType 1 resource. */
	patternFill?: PdfPatternFill;
	/** A clipping path cuts it (or a clip that is not a rectangle is active). */
	clipped?: true;
}

/** A piece of a painted path, in reading coordinates. */
export type PdfPathSegment =
	| { op: 'M' | 'L'; x: number; y: number }
	| { op: 'C'; x1: number; y1: number; x2: number; y2: number; x: number; y: number }
	| { op: 'Z' };

/** A painted path that is not a straight rule or rectangle: curves, diagonals, polygons. */
export interface PdfPagePath {
	segments: PdfPathSegment[];
	/** The bounds of the segments' points (control points included). */
	x0: number;
	y0: number;
	x1: number;
	y1: number;
	/** `#rrggbb` when the path is filled. */
	fill?: string;
	/** Gradient fill from a PDF shading or shading pattern. */
	gradient?: PdfGradientFill;
	/** Tiling pattern fill from a PDF PatternType 1 resource. */
	patternFill?: PdfPatternFill;
	/** `#rrggbb` when the path is stroked. */
	stroke?: string;
	lineWidth: number;
	fillAlpha?: number;
	strokeAlpha?: number;
	dash?: number[];
	/** Filled with the even-odd rule. */
	evenOdd?: true;
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
	/** Where on that page the link goes (reading coordinates), when the destination says. */
	targetY?: number;
	/** The named destination the link goes to, if it uses one. */
	destination?: string;
}

/** An entry of the document outline (bookmarks), in document order. */
export interface PdfOutlineItem {
	title: string;
	/** Nesting depth, 0 for a top-level entry. */
	level: number;
	page?: number;
	y?: number;
	destination?: string;
}

export interface PdfPageContent {
	index: number;
	width: number;
	height: number;
	spans: PdfTextSpan[];
	images: PdfPageImage[];
	rules: PdfPageRule[];
	/** Painted paths other than rules (absent when the engine does not supply them). */
	paths?: PdfPagePath[];
	links: PdfPageLink[];
	/** What the reader could not take from the page. */
	notes: string[];
}
