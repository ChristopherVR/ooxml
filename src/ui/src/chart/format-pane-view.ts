/**
 * What a chart Format pane reads: a neutral view of the selected element
 * (colours as `#RRGGBB`, transparency and sizes in the pane's own units) and
 * of its options, supplied by a {@link ChartFormatSource}. `chartSpaceFormatSource`
 * reads a `ChartSpace` (`ooxml-core/chart`); a product whose chart model is
 * not a `ChartSpace` yet supplies its own source. Effects, 3-D and text views
 * are optional extension sections: a source that cannot read or write them
 * leaves them out and the pane does not offer them. The shapes match
 * PowerPoint's `ChartElementFormatView`, which they were taken from.
 *
 * @module chart/format-pane-view
 */
import type { ChartFormatTarget, ChartGradientFill } from 'ooxml-core/chart';
import type { ChartTrendline, ChartErrorBars } from 'ooxml-core/chart';

import type {
	ChartFormatChoice,
	ChartFormatTargetOption,
	ChartFormatText,
} from './format-pane-types';

/** The sections an element supports. */
export interface ChartFormatSupport {
	fill: boolean;
	gradient: boolean;
	pattern: boolean;
	picture: boolean;
	threeD: boolean;
	text: boolean;
	/** The Border (Line) section; absent means supported. */
	line?: boolean;
	/** Gradient lines, compound, join and arrow types; absent means supported. */
	lineDetails?: boolean;
	/** The Effects tab; absent means supported. */
	effects?: boolean;
	secondaryAxis?: boolean;
}

export type ChartFormatFillKind =
	| 'none'
	| 'solid'
	| 'gradient'
	| 'pattern'
	| 'picture'
	| 'automatic';

export interface ChartFormatFillView {
	kind: ChartFormatFillKind;
	/** `#RRGGBB`. */
	color?: string;
	/** 0 (opaque) through 100 (clear). */
	transparency?: number;
	gradient?: ChartGradientFill;
	pattern?: { preset: string; foreground: string; background: string };
	picture?: { src?: string; relId?: string; tile?: boolean; transparency?: number };
}

export interface ChartFormatLineView {
	kind: 'none' | 'solid' | 'gradient' | 'automatic';
	color?: string;
	gradient?: ChartGradientFill;
	transparency?: number;
	/** Points. */
	width?: number;
	dash?: string;
	compound?: string;
	cap?: string;
	join?: string;
	beginArrow?: string;
	endArrow?: string;
}

export interface ChartFormatShadowView {
	color: string;
	transparency?: number;
	blur: number;
	distance: number;
	angle: number;
}

export interface ChartFormatGlowView {
	color: string;
	transparency?: number;
	radius: number;
}

export interface ChartFormatEffectsView {
	shadow?: ChartFormatShadowView | null;
	glow?: ChartFormatGlowView | null;
	softEdge?: number | null;
}

export interface ChartFormatBevelView {
	type: string;
	width: number;
	height: number;
}

export interface ChartFormatThreeDView {
	bevelTop?: ChartFormatBevelView | null;
	bevelBottom?: ChartFormatBevelView | null;
	depth?: { size: number; color?: string } | null;
	contour?: { size: number; color?: string } | null;
	material?: string | null;
	lighting?: string | null;
}

export interface ChartFormatTextView {
	fill?: {
		kind: 'none' | 'solid' | 'gradient' | 'automatic';
		color?: string;
		transparency?: number;
		gradient?: ChartGradientFill;
	};
	outline?: {
		kind: 'none' | 'solid' | 'automatic';
		color?: string;
		transparency?: number;
		width?: number;
		dash?: string;
	};
	effects?: { shadow?: ChartFormatShadowView | null; glow?: ChartFormatGlowView | null };
	rotation?: number | null;
	direction?: string | null;
	anchor?: string | null;
}

/** The element's shape and text formatting. */
export interface ChartFormatElementView {
	support: ChartFormatSupport;
	fill: ChartFormatFillView;
	line: ChartFormatLineView;
	effects?: ChartFormatEffectsView;
	threeD?: ChartFormatThreeDView;
	text?: ChartFormatTextView;
	series?: ChartFormatSeriesOptionsView;
}

/**
 * A chart type in PowerPoint's names, which decide the options a series and its labels offer:
 * `bar`, `bar3D`, `pie`, `pie3D`, `doughnut`, `ofPie`, `line`, `scatter`, `bubble` are recognised,
 * any other name gets the options every type has.
 */
export type ChartFormatChartType = string;

export interface ChartFormatSeriesOptionsView {
	/** Undefined for a combination chart. */
	chartType?: ChartFormatChartType;
	/** On the secondary axis; undefined when the source cannot move a series between axes. */
	secondaryAxis?: boolean;
	overlap?: number;
	gapWidth?: number;
	firstSliceAngle?: number;
	holeSize?: number;
	/** Draw sliders beside the spacing boxes. */
	sliders?: boolean;
	marker?: { symbol?: string; size?: number };
	analysis?: {
		trendline?: ChartTrendline;
		errorBars?: ChartErrorBars;
		multipleTrendlines: boolean;
		multipleErrorBars: boolean;
	};
}

export interface ChartFormatAxisView {
	title?: string;
	baseTimeUnit?: string;
	majorTimeUnit?: string;
	minorTimeUnit?: string;
	timeUnitsEditable?: boolean;
	numberFormatSourceLinked?: boolean;
	kind: 'cat' | 'val' | 'date' | 'ser';
	min?: number;
	max?: number;
	majorUnit?: number;
	minorUnit?: number;
	crosses?: string;
	crossesAt?: number;
	displayUnits?: string;
	logScale?: boolean;
	logBase?: number;
	reverse: boolean;
	majorTickMark?: string;
	minorTickMark?: string;
	tickLabelPosition?: string;
	/** A format code not linked to the source; empty when the axis follows its data. */
	numberFormat?: string;
}

export interface ChartFormatLegendView {
	position?: string;
	overlay?: boolean;
	/** Whether the overlay switch is offered; absent means yes. */
	overlaySupported?: boolean;
}

export interface ChartFormatLabelsView {
	chartType?: ChartFormatChartType;
	/** The options of one series' labels (each flag falling back to the chart's). */
	seriesScoped: boolean;
	showSeriesName?: boolean;
	showCategory?: boolean;
	showValue?: boolean;
	showPercent?: boolean;
	showLegendKey?: boolean;
	position?: string;
}

/** How an axis reads in a caption ("Secondary Vertical (Value) Axis"). */
export interface ChartFormatAxisInfo {
	kind: 'cat' | 'val' | 'date' | 'ser';
	vertical: boolean;
	secondary: boolean;
}

/** Everything a Format pane reads from one chart. */
export interface ChartFormatSource {
	/** The element picker's entries, in order; the first is the fallback target. */
	readonly targets: readonly ChartFormatTargetOption[];
	/** A valid target the picker does not list, or undefined. */
	resolve?(key: string): ChartFormatTargetOption | undefined;
	element(target: ChartFormatTarget): ChartFormatElementView;
	seriesOptions?(seriesIndex: number): ChartFormatSeriesOptionsView | undefined;
	axisOptions?(axisIndex: number): ChartFormatAxisView | undefined;
	legendOptions?(): ChartFormatLegendView | undefined;
	labelOptions?(seriesIndex: number | undefined): ChartFormatLabelsView | undefined;
	/** Pattern presets the pattern fill offers. */
	readonly patternPresets?: readonly ChartFormatChoice[];
	/** Sections the host draws itself, by section id, mapped to the pane's slot names. */
	readonly slots?: Readonly<Partial<Record<'fill', string>>>;
	/** Whether the Text Options toggle is offered (disabled where an element has no text); default yes. */
	readonly textOptions?: boolean;
	/** Show every section on one page, without tabs. */
	readonly flat?: boolean;
	/** Caption of the element picker (default "Chart Options"). */
	readonly picker?: ChartFormatText;
}
