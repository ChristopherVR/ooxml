import type { VisioHyperlink, VisioShapeData } from './shape-metadata.js';

/** Visio's internal distance unit is the inch; angles are radians. */
export type VisioMatrix = readonly [number, number, number, number, number, number];
export interface VisioDiagnostic {
	code: string;
	severity: 'info' | 'warning';
	message: string;
	part?: string;
	pageId?: string;
	shapeId?: string;
}
export interface VisioGeometry {
	/** SVG path data in local, y-up, inch coordinates. No untrusted markup. */
	path: string;
	fill: boolean;
	stroke: boolean;
}
export interface VisioLinearGradient {
	type: 'linear';
	/** Gradient endpoints in local y-up inches, for SVG userSpaceOnUse. */
	start: readonly [number, number];
	end: readonly [number, number];
	stops: { offset: number; color: string; opacity: number }[];
}
export interface VisioStyle {
	fill: string;
	/** Supported normalized gradient; fill remains the solid fallback color. */
	fillGradient?: VisioLinearGradient;
	lineColor: string;
	lineWidth: number;
	/** Normalized SVG-compatible cap; absent means the effective cap is unresolved. */
	lineCap?: 'round' | 'butt' | 'square';
	/** Effective ShapeSheet line pattern; 0 is no line, 1 is solid (also invalid-cache fallback). */
	linePattern: number;
	/** Alternating dash/gap lengths in stroke-width multiples, not inches.
	 * Cached built-ins 2-23 use inferred compatibility spacing (2-6 entries, each 0-27).
	 * Absent for solid, invisible or unresolved patterns. Consumers scale by lineWidth.
	 */
	lineDash?: readonly number[];
	fillOpacity: number;
	lineOpacity: number;
	/** Visio arrowhead codes. Consumers must report unsupported arrow variants. */
	startArrow: number;
	endArrow: number;
	startArrowSize?: number;
	endArrowSize?: number;
}
export interface VisioTextRun {
	text: string;
	fontFamily: string;
	fontSize: number;
	color: string;
	bold: boolean;
	italic: boolean;
	underline: boolean;
}
export interface VisioParagraph {
	/** UTF-16 offsets into plainText. End excludes the paragraph's newline. */
	start: number;
	end: number;
	horizontalAlign: 'left' | 'center' | 'right' | 'justify' | 'distributed';
	indentLeft: number;
	indentRight: number;
	indentFirst: number;
	spaceBefore: number;
	spaceAfter: number;
	/** Exact values are inches; multiples scale the largest font size on the line. */
	lineSpacing: { kind: 'multiple' | 'exact'; value: number };
	direction: 'ltr' | 'rtl';
	bullet?: { text: string; fontFamily: string; fontSize: number; offset: number };
}
export interface VisioText {
	plainText: string;
	/** Optional backdrop behind laid-out text, not the complete shape rectangle. */
	backgroundColor?: string;
	backgroundOpacity?: number;
	runs: VisioTextRun[];
	paragraphs?: VisioParagraph[];
	fontFamily: string;
	fontSize: number;
	color: string;
	horizontalAlign: 'left' | 'center' | 'right';
	verticalAlign: 'top' | 'middle' | 'bottom';
	/** Local text box transform (y-up), dimensions and margins in inches. */
	transform: VisioMatrix;
	width: number;
	height: number;
	margins: { left: number; right: number; top: number; bottom: number };
}
export interface VisioImage {
	mimeType: 'image/png' | 'image/jpeg' | 'image/gif';
	/** Validated embedded raster bytes shared between instances. Consumers must not modify them. */
	bytes: Uint8Array;
	pixelWidth: number;
	pixelHeight: number;
	opacity?: number;
	/** Optional placement in local y-up inches, defaulting to the shape bounds. */
	x?: number;
	y?: number;
	width?: number;
	height?: number;
}
export interface VisioLayer {
	id: string;
	name: string;
	visible: boolean;
	printable: boolean;
	locked: boolean;
	color?: string;
	colorOpacity?: number;
}
/** Summary of normalized layer Print flags, not a shape printing decision or mixed-layer policy. */
export type VisioLayerPrintSummary =
	| 'unlayered'
	| 'all-enabled'
	| 'all-disabled'
	| 'mixed'
	| 'unknown';
/** Saved reasons retained independently so display overrides cannot reveal intrinsic hiding. */
export interface VisioShapeVisibility {
	layerHidden: boolean;
	/** Guides are hidden by this renderer, but can be printable in Visio. */
	guide: boolean;
	/** Legacy shape-level NoShow compatibility flag, not Geometry.NoShow. Missing means unknown. */
	noShow?: boolean;
	layerPrintSummary: VisioLayerPrintSummary;
	/**
	 * Cached NonPrinting only. Missing or unusable caches remain unknown.
	 * https://learn.microsoft.com/en-us/office/client-developer/visio/nonprinting-cell-miscellaneous-section
	 */
	nonPrinting?: boolean;
}
export interface VisioShape {
	id: string;
	name: string;
	kind: 'shape' | 'group' | 'connector' | 'foreign';
	/** Cached group data ordering: 0 hides own data, 1 behind members, 2 in front. */
	groupDisplayMode?: 0 | 1 | 2;
	width: number;
	height: number;
	/** Local y-up coordinates to parent y-up coordinates. */
	transform: VisioMatrix;
	geometry: VisioGeometry[];
	style: VisioStyle;
	text: VisioText;
	/** Saved display suppression of this shape and its subtree; retained for existing consumers. */
	hidden: boolean;
	/** Additive metadata. Older or manually constructed scenes may omit it. */
	visibility?: VisioShapeVisibility;
	children: VisioShape[];
	masterId?: string;
	layerIds?: string[];
	image?: VisioImage;
	shapeData?: VisioShapeData[];
	hyperlinks?: VisioHyperlink[];
}
export interface VisioConnection {
	fromShapeId: string;
	toShapeId: string;
	fromCell: string;
	toCell: string;
	fromPart?: number;
	toPart?: number;
}
export interface VisioPage {
	id: string;
	name: string;
	width: number;
	height: number;
	isBackground: boolean;
	backgroundPageId?: string;
	shapes: VisioShape[];
	connectors: VisioConnection[];
	layers?: VisioLayer[];
}
export interface VisioDocument {
	format: 'vsdx';
	pages: VisioPage[];
	diagnostics: VisioDiagnostic[];
}
