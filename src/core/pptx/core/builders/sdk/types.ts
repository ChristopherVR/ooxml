/**
 * SDK input types for the headless PPTX builder API.
 *
 * These types provide a simplified, ergonomic interface for creating
 * presentation elements programmatically. Internal types (PptxElement,
 * ShapeStyle, TextStyle) are the canonical model; these types are
 * convenience wrappers that get mapped to the internal model by
 * the builder functions.
 *
 * @module sdk/types
 */

import type { PptxAnimationPreset, PptxAnimationTrigger } from '../../types/animation';
import type {
	PptxChartBoxWhiskerOptions,
	PptxChartHistogramOptions,
	PptxChartRegionMapOptions,
	PptxChartTreemapOptions,
	PptxChartWaterfallOptions,
} from '../../types/chart';
import type { PptxThemeColorRef } from '../../types/color-ref';
import type { StrokeDashType, ConnectorArrowType } from '../../types/common';
import type { PptxTransitionType } from '../../types/transition';
import type { ChartGradientInput } from './chart-gradient-input';
import type { ChartAreaFormatInput } from './chart-layout-input';
import type { ChartAxisEdit, PptxChartAxisType } from './chart-operations';
import type {
	GlowInput,
	InnerShadowInput,
	LineEndInput,
	ReflectionInput,
	RelativeRectInput,
} from './shape-style-input-types';

export type { ChartGradientInput, ChartGradientStopInput } from './chart-gradient-input';
export type {
	GlowInput,
	InnerShadowInput,
	LineEndInput,
	ReflectionInput,
	RelativeRectInput,
} from './shape-style-input-types';

// ---------------------------------------------------------------------------
// Position & layout
// ---------------------------------------------------------------------------

/** Position and size in pixels. Converted to EMU internally when needed. */
export interface ElementPosition {
	x: number;
	y: number;
	width: number;
	height: number;
	rotation?: number;
}

/** Text-capable ST_PlaceholderType values (ECMA-376 19.7.10). */
export type TextPlaceholderType =
	| 'title'
	| 'ctrTitle'
	| 'subTitle'
	| 'body'
	| 'obj'
	| 'dt'
	| 'ftr'
	| 'hdr'
	| 'sldNum';

/** Target a layout slot, optionally overriding some or all of its frame. */
export interface PlaceholderOptions {
	/** Layout placeholder index. Required to disambiguate repeated content slots. */
	idx?: number;
	x?: number;
	y?: number;
	width?: number;
	height?: number;
}

// ---------------------------------------------------------------------------
// Fill, stroke, shadow helpers
// ---------------------------------------------------------------------------

export type FillInput =
	| {
			type: 'solid';
			color: string;
			opacity?: number;
			/**
			 * A theme colour to use instead of a plain hex. When set, the shape
			 * saves as `<a:schemeClr>` (e.g. `{ scheme: 'accent1', lumMod: 0.8 }`
			 * for "Accent 1, Lighter 80%") so it keeps following the theme after a
			 * later theme change; `color` still supplies the immediate resolved
			 * hex for renderers that read it directly.
			 */
			themeColorRef?: PptxThemeColorRef;
	  }
	| {
			type: 'gradient';
			/**
			 * Gradient direction in the OOXML `a:lin/@ang` convention: degrees
			 * clockwise from the positive x-axis, pointing from the first stop
			 * towards the last (`0` = left to right, `90` = top to bottom). This is
			 * what lands in `ShapeStyle.fillGradientAngle` and what is written back
			 * to the file, NOT a CSS `linear-gradient()` angle.
			 */
			angle?: number;
			gradientType?: 'linear' | 'radial';
			/**
			 * A path gradient (`a:path/@path`) instead of a linear one. Setting it
			 * makes the gradient radial; `'circle'` is what `gradientType: 'radial'`
			 * alone gives.
			 */
			path?: 'circle' | 'rect' | 'shape';
			/** Where a path gradient reaches its last stop (`a:fillToRect`). */
			focus?: RelativeRectInput;
			/** The rectangle one gradient tile occupies (`a:tileRect`). */
			tile?: RelativeRectInput;
			/** Whether a linear gradient scales with the shape (`a:lin/@scaled`). */
			scaled?: boolean;
			/** Whether the gradient rotates with the shape (`@rotWithShape`). */
			rotWithShape?: boolean;
			stops: Array<{ color: string; position: number; opacity?: number }>;
	  }
	| {
			type: 'pattern';
			preset: string;
			/** Pattern colour (`a:fgClr`). */
			foreground?: string;
			/** Colour behind the pattern (`a:bgClr`). */
			background?: string;
	  }
	| { type: 'image'; url: string; mode?: 'stretch' | 'tile' }
	| { type: 'none' };

export interface StrokeInput {
	/** Outline colour; `'transparent'` writes an explicit empty outline (`<a:ln><a:noFill/></a:ln>`). */
	color?: string;
	width?: number;
	/**
	 * A preset dash, or an exact pattern of dash and gap lengths in px
	 * (`[dash, gap, dash, gap, ...]`, repeated once when odd, as in SVG's
	 * `stroke-dasharray`), written as `a:custDash` relative to `width`.
	 */
	dash?: StrokeDashType | number[];
	opacity?: number;
	join?: 'round' | 'bevel' | 'miter';
	/**
	 * Miter limit as a ratio of the line width (e.g. 8). Implies `join: 'miter'`
	 * when `join` is omitted; ignored with `'round'` or `'bevel'`, which have no
	 * limit in DrawingML.
	 */
	miterLimit?: number;
	cap?: 'flat' | 'rnd' | 'sq';
	/** Compound line type (`a:ln/@cmpd`). */
	compound?: 'sng' | 'dbl' | 'thickThin' | 'thinThick' | 'tri';
	/** Pen alignment (`a:ln/@algn`): centred on the outline or inside it. */
	align?: 'ctr' | 'in';
	/** Decoration on the start of the outline (`a:headEnd`). */
	headEnd?: LineEndInput;
	/** Decoration on the end of the outline (`a:tailEnd`). */
	tailEnd?: LineEndInput;
	/** A theme colour for the outline; see {@link FillInput}'s `themeColorRef`. */
	themeColorRef?: PptxThemeColorRef;
}

export interface ShadowInput {
	color?: string;
	blur?: number;
	offsetX?: number;
	offsetY?: number;
	opacity?: number;
}

// ---------------------------------------------------------------------------
// Text input
// ---------------------------------------------------------------------------

export interface TextStyleInput {
	fontSize?: number;
	fontFamily?: string;
	bold?: boolean;
	italic?: boolean;
	underline?: boolean;
	strikethrough?: boolean;
	color?: string;
	/** A theme colour for the run; see {@link FillInput}'s `themeColorRef`. */
	themeColorRef?: PptxThemeColorRef;
	alignment?: 'left' | 'center' | 'right' | 'justify';
	verticalAlignment?: 'top' | 'middle' | 'bottom';
	lineSpacing?: number;
	spaceBefore?: number;
	spaceAfter?: number;
}

export interface TextSegmentInput {
	text: string;
	style?: Partial<TextStyleInput>;
}

export interface TextOptions extends Partial<ElementPosition> {
	fontSize?: number;
	fontFamily?: string;
	bold?: boolean;
	italic?: boolean;
	underline?: boolean;
	strikethrough?: boolean;
	color?: string;
	alignment?: 'left' | 'center' | 'right' | 'justify';
	verticalAlignment?: 'top' | 'middle' | 'bottom';
	lineSpacing?: number;
	fill?: FillInput;
	stroke?: StrokeInput;
	shadow?: ShadowInput;
	innerShadow?: InnerShadowInput;
	glow?: GlowInput;
	/** Soft-edge radius in px (`a:softEdge`). */
	softEdge?: number;
	reflection?: ReflectionInput;
	opacity?: number;
}

// ---------------------------------------------------------------------------
// Shape input
// ---------------------------------------------------------------------------

export interface ShapeOptions extends Partial<ElementPosition> {
	fill?: FillInput;
	stroke?: StrokeInput;
	text?: string;
	textStyle?: Partial<TextStyleInput>;
	adjustments?: Record<string, number>;
	shadow?: ShadowInput;
	innerShadow?: InnerShadowInput;
	glow?: GlowInput;
	/** Soft-edge radius in px (`a:softEdge`). */
	softEdge?: number;
	reflection?: ReflectionInput;
	opacity?: number;
}

// ---------------------------------------------------------------------------
// Image input
// ---------------------------------------------------------------------------

export interface ImageOptions extends Partial<ElementPosition> {
	altText?: string;
	cropLeft?: number;
	cropTop?: number;
	cropRight?: number;
	cropBottom?: number;
	opacity?: number;
}

// ---------------------------------------------------------------------------
// Table input
// ---------------------------------------------------------------------------

export interface TableInput {
	rows: TableRowInput[];
	/**
	 * Column widths, relative to each other (fractions, or px that add up to
	 * the table's width), one per column. Omitted: equal columns. Widths that
	 * are not one finite value of 0 or more per column, or that add up to 0,
	 * throw a `RangeError`.
	 */
	columnWidths?: number[];
	/**
	 * Table style id (`a:tableStyleId`). Omitted: PowerPoint's default
	 * (Medium Style 2 - Accent 1). `'none'` (or `''`): no table style, so only
	 * the cells' own formatting shows.
	 */
	style?: string;
	/**
	 * Opt-in: `rows[].cells` list only the real cells, and the slots that a
	 * `gridSpan` / `rowSpan` covers are added as empty `hMerge` / `vMerge`
	 * cells. Without it the cells are taken as given, one per grid slot. A
	 * span that is not a whole number from 1, reaches past the last row or
	 * overlaps another merge throws a `RangeError`.
	 */
	fillMergedCells?: boolean;
	bandRows?: boolean;
	bandColumns?: boolean;
	firstRow?: boolean;
	lastRow?: boolean;
	firstCol?: boolean;
	lastCol?: boolean;
}

export interface TableRowInput {
	cells: TableCellInput[];
	height?: number;
}

export interface TableCellInput {
	text: string;
	style?: Partial<TextStyleInput>;
	/**
	 * The cell's own fill, written in `a:tcPr` over the table style's. Solid
	 * (no opacity below 1), gradient (no `tile`, `scaled: false` or
	 * `rotWithShape`), pattern and `none` (an explicit `a:noFill`). A fill a
	 * cell cannot carry, such as an image, throws a `RangeError`.
	 */
	fill?: FillInput;
	/** The cell's own edge lines, written over the table style's. */
	borders?: TableCellBordersInput;
	gridSpan?: number;
	rowSpan?: number;
}

/**
 * A table cell's edge lines: `all` for the four edges, and each edge on its
 * own (`a:lnT` / `a:lnR` / `a:lnB` / `a:lnL`), which overrides `all`.
 */
export interface TableCellBordersInput {
	all?: TableCellBorderInput;
	top?: TableCellBorderInput;
	right?: TableCellBorderInput;
	bottom?: TableCellBorderInput;
	left?: TableCellBorderInput;
}

/** One edge line of a table cell. */
export interface TableCellBorderInput {
	color?: string;
	/** Line width in px. */
	width?: number;
	/**
	 * A preset dash, or an exact pattern of dash and gap lengths in px
	 * (`[dash, gap, dash, gap, ...]`, repeated once when odd, as in SVG's
	 * `stroke-dasharray`), written as `a:custDash` relative to `width` (1 px
	 * when `width` is omitted or 0). An empty pattern, or a length that is
	 * not a finite number above 0, throws a `RangeError`.
	 */
	dash?: Exclude<StrokeDashType, 'custom'> | number[];
}

export interface TableOptions extends Partial<ElementPosition> {}

// ---------------------------------------------------------------------------
// Chart input
// ---------------------------------------------------------------------------

export interface ChartSeriesInput {
	name: string;
	values: number[];
	color?: string;
	/**
	 * Gradient fill for the whole series, written as `c:ser/c:spPr/a:gradFill`.
	 * Takes precedence over `color`. Area-filled chart types only (bar, area,
	 * pie, doughnut, ofPie, bubble, surface and their 3-D variants);
	 * `createChartElement` throws for line-drawn and ChartEx types.
	 */
	gradientFill?: ChartGradientInput;
	boxWhiskerOptions?: PptxChartBoxWhiskerOptions;
	histogramOptions?: PptxChartHistogramOptions;
	waterfallOptions?: PptxChartWaterfallOptions;
	regionMapOptions?: PptxChartRegionMapOptions;
	treemapOptions?: PptxChartTreemapOptions;
}

export interface ChartInput {
	series: ChartSeriesInput[];
	categories: string[];
	/** ChartEx hierarchy levels in leaf-to-root XML order. */
	categoryLevels?: string[][];
	title?: string;
	hasLegend?: boolean;
	legendPosition?: 't' | 'b' | 'l' | 'r' | 'tr';
	grouping?: 'clustered' | 'stacked' | 'percentStacked';
	/** Bar series direction (`c:barDir`): vertical columns (default) or horizontal bars. */
	barDirection?: 'col' | 'bar';
	/** Bar cluster gap, % of bar width (`c:gapWidth`, 0..500; default 150). Bar, bar3D, ofPie. */
	gapWidth?: number;
	/** Bar overlap within a cluster, % (`c:overlap`, -100..100). 2-D bar only. */
	overlap?: number;
	/** First slice angle in degrees (`c:firstSliceAng`, 0..360). Pie, doughnut. */
	firstSliceAngle?: number;
	/** Doughnut hole size, % (`c:holeSize`, 1..90; default 50). Doughnut only. */
	holeSize?: number;
	/** Chart-area fill, border and rounded corners (`c:chartSpace/c:spPr`, `c:roundedCorners`). */
	chartArea?: ChartAreaFormatInput;
	/** Plot-area fill and border (`c:plotArea/c:spPr`). */
	plotArea?: ChartAreaFormatInput;
	/** Axis edits by axis type, as `setChartAxis` takes them (e.g. `{ valAx: { visible: false } }`). */
	axes?: Partial<Record<PptxChartAxisType, ChartAxisEdit>>;
}

export interface ChartOptions extends Partial<ElementPosition> {}

// ---------------------------------------------------------------------------
// Connector input
// ---------------------------------------------------------------------------

export interface ConnectorOptions extends Partial<ElementPosition> {
	type?: 'straight' | 'bent' | 'curved';
	/** The outline; an omitted `color` is black and an omitted `width` 1 px. */
	stroke?: StrokeInput;
	/** Arrow on the start; replaces `stroke.headEnd`, its size included. */
	startArrow?: ConnectorArrowType;
	/** Arrow on the end; replaces `stroke.tailEnd`, its size included. */
	endArrow?: ConnectorArrowType;
	from?: { elementId: string; site: number };
	to?: { elementId: string; site: number };
}

// ---------------------------------------------------------------------------
// Media input
// ---------------------------------------------------------------------------

export interface MediaOptions extends Partial<ElementPosition> {
	autoPlay?: boolean;
	loop?: boolean;
	volume?: number;
	trimStartMs?: number;
	trimEndMs?: number;
	posterFrame?: string;
}

// ---------------------------------------------------------------------------
// Group input
// ---------------------------------------------------------------------------

export interface GroupOptions extends Partial<ElementPosition> {}

// ---------------------------------------------------------------------------
// Slide-level inputs
// ---------------------------------------------------------------------------

export type BackgroundInput =
	| { type: 'solid'; color: string }
	| {
			type: 'gradient';
			/**
			 * A CSS `linear-gradient()` angle, as the slide's CSS
			 * `backgroundGradient` has always used: degrees clockwise from "to
			 * top" (`90` = left to right, `180` = top to bottom). Defaults to
			 * `180`. Saved as `a:lin/@ang` (a quarter turn less). Not with `path`.
			 */
			angle?: number;
			/**
			 * A path gradient (`a:path/@path`) growing from `focus` outwards
			 * instead of a linear one.
			 */
			path?: 'circle' | 'rect' | 'shape';
			/** Where a path gradient starts (`a:fillToRect`). Default: the slide centre. */
			focus?: RelativeRectInput;
			/**
			 * At least two stops: `color` as `#RRGGBB`, `position` 0-1, `opacity`
			 * 0-1 (default opaque).
			 */
			stops: Array<{ color: string; position: number; opacity?: number }>;
	  }
	| { type: 'image'; source: string };

export interface TransitionInput {
	type: PptxTransitionType;
	duration?: number;
	direction?: string;
	advanceAfterMs?: number;
}

export interface AnimationInput {
	preset: PptxAnimationPreset;
	trigger?: PptxAnimationTrigger;
	duration?: number;
	delay?: number;
}

// ---------------------------------------------------------------------------
// Presentation-level inputs
// ---------------------------------------------------------------------------

export interface PresentationOptions {
	/** Slide width in EMU. Default: 12192000 (16:9 widescreen). */
	width?: number;
	/** Slide height in EMU. Default: 6858000 (16:9 widescreen). */
	height?: number;
	/** Theme configuration. */
	theme?: PresentationThemeInput;
	/** Presentation title (stored in docProps/core.xml). */
	title?: string;
	/** Presentation author. */
	creator?: string;
	/**
	 * Number of blank slides to include in the initial presentation.
	 * Default: 0 (no slides). Slides use the "Blank" layout.
	 */
	initialSlideCount?: number;
	/**
	 * `p:presentation/@firstSlideNum` (ECMA-376 §19.2.1.26): the number a
	 * slide-number field shows on the first slide. Omitted when undefined, so
	 * the spec default (`1`) applies.
	 */
	firstSlideNumber?: number;
}

export interface PresentationThemeInput {
	name?: string;
	colors?: {
		dk1?: string;
		lt1?: string;
		dk2?: string;
		lt2?: string;
		accent1?: string;
		accent2?: string;
		accent3?: string;
		accent4?: string;
		accent5?: string;
		accent6?: string;
		hlink?: string;
		folHlink?: string;
	};
	fonts?: {
		majorFont?: string;
		minorFont?: string;
	};
}
