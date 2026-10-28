// The chart summary the SVG painter draws: the first chart group's type and series with their
// cached data and DrawingML paint, plus the legend, title and formatting metadata. Spreadsheet
// charts extend it with their anchor and legacy colours; Word and PowerPoint build it from a
// `ChartSpace` (`chartSummaryFromSpace`).
import type { ChartMarkerSummary } from './chart-view-markers';
import type { ChartManualLayout } from '../manual-layout';
import type { ChartAxis } from '../model';
import type { DrawingColor, DrawingFill, DrawingLine } from '../../drawingml/types';
import type { DataLabelFlagLevel } from '../data-label-text';
import type { ChartErrorAmountOptions } from '../error-bar-amount';
import type { ChartErrorBarDirection, ChartErrorBarType } from '../model-analysis';
import type { ChartTrendlineFitOptions } from '../trendline-fit';
import type { BubbleRadiusOptions } from '../bubble-radius';
import type { ParentLabelLayout } from '../hierarchy';
import type { ChartShapeProperties } from '../model-series';
import type { HistogramOptions } from '../histogram-binning';
import type { ChartStyleDefinition } from '../style-definition';
import type { WaterfallOptions } from '../waterfall';

/** The chart families a summary records; the painter draws all but surface. */
export type ChartSummaryType =
	| 'bar'
	| 'column'
	| 'line'
	| 'pie'
	| 'doughnut'
	| 'area'
	| 'scatter'
	| 'radar'
	| 'bubble'
	| 'stock'
	| 'surface'
	| 'waterfall'
	| 'histogram'
	| 'treemap'
	| 'sunburst';

/** One cascade level of a series' data labels: its content switches, placement and format. */
export interface ChartSummaryLabelLevel extends DataLabelFlagLevel {
	layout?: ChartManualLayout;
	showLeaderLines?: boolean;
	leaderLine?: DrawingLine;
	/** `c:delete`: the label (or, at series and group level, every label) is suppressed. */
	deleted?: boolean;
	/** `c:dLblPos`: `ctr`, `inEnd`, `inBase`, `outEnd`, `bestFit`, `t`, `b`, `l`, `r`. */
	position?: string;
	/** Number format code of the label. */
	numberFormat?: string;
	/** Custom label text (`c:dLbl/c:tx/c:rich`), which replaces the built label. */
	text?: string;
}

/** The data labels of a series by cascade level; absent when the series shows none. */
export interface ChartSummaryDataLabels {
	chart?: ChartSummaryLabelLevel;
	series?: ChartSummaryLabelLevel;
	/** Per-point labels by `c:idx`. */
	points?: Record<number, ChartSummaryLabelLevel>;
}

export interface ChartSummarySeries extends ChartMarkerSummary {
	smooth?: boolean;
	showLine?: boolean;
	line?: DrawingLine;
	name?: string;
	/** Formula reference for the series name (`Sheet1!$B$1`). */
	nameRef?: string;
	categoriesRef?: string;
	valuesRef?: string;
	/** Cached categories and values from the chart part, used when references cannot resolve. */
	categories: (string | number)[];
	values: (number | null)[];
	/** Cached source value format, used when data labels declare no number format. */
	numberFormat?: string;
	/** Bubble charts: the `c:bubbleSize` values, parallel to `values`. */
	bubbleSizes?: (number | null)[];
	/** Formula reference for the bubble sizes. */
	bubbleSizesRef?: string;
	/** DrawingML fill/line color, including theme transforms. */
	drawingColor?: DrawingColor;
	/** Explicit colors of individual points (indexed by c:dPt/c:idx). */
	pointColors?: Record<number, DrawingColor>;
	/** Non-solid primary DrawingML fill/line paint. Solid fills use drawingColor. */
	fill?: DrawingFill;
	/** Non-solid individual point fills; solid point overrides use pointColors. */
	pointFills?: Record<number, DrawingFill>;
	/** Opaque imported DrawingML effect list, retained during chart regeneration. */
	effectsXml?: string;
	/** Data labels the painter draws at each point (`c:dLbls`). */
	dataLabels?: ChartSummaryDataLabels;
	/** Trendlines the painter fits and draws over the series (`c:trendline`). */
	trendlines?: ChartSummaryTrendline[];
	/** Error bars the painter draws at each point (`c:errBars`). */
	errorBars?: ChartSummaryErrorBars[];
}

/** A trendline of a summary series: its fit and its outline (`c:trendline/c:spPr/a:ln`). */
export interface ChartSummaryTrendline extends ChartTrendlineFitOptions {
	name?: string;
	line?: DrawingLine;
}

/** Error bars of a summary series, with their amounts resolved to numbers. */
export interface ChartSummaryErrorBars extends ChartErrorAmountOptions {
	direction: ChartErrorBarDirection;
	barType: ChartErrorBarType;
	noEndCap?: boolean;
	line?: DrawingLine;
}

/** Chart families a combination chart can layer over its first group. */
export type ChartSummaryGroupType = 'bar' | 'column' | 'line' | 'area' | 'stock';

/** The formatting of a stock group: its `c:hiLowLines`, `c:upBars`, `c:downBars` and bar gap. */
export interface StockSummaryStyle {
	hiLowLines?: ChartShapeProperties;
	upBars?: ChartShapeProperties;
	downBars?: ChartShapeProperties;
	/** `c:upDownBars/c:gapWidth` in percent. */
	gapWidth?: number;
}

/**
 * A further chart group of a combination chart (`c:lineChart` beside a `c:barChart`): it shares the
 * category axis of the first group and plots on its value axis or on the secondary one.
 */
export interface ChartSummaryGroup {
	chartType: ChartSummaryGroupType;
	grouping?: 'clustered' | 'stacked' | 'percentStacked' | 'standard';
	barGapWidth?: number;
	barOverlap?: number;
	series: ChartSummarySeries[];
	/** Stock groups: their line and bar formatting. */
	stock?: StockSummaryStyle;
	/** Plotted against its own (secondary) value axis, drawn at the right edge. */
	secondary: boolean;
}

/** The binning of a histogram summary: `cx:binning` plus whether bars aggregate by category. */
export interface HistogramSummaryOptions extends HistogramOptions {
	aggregateByCategory?: boolean | undefined;
}

/** Bubble display options of `c:bubbleChart`: scale (percent), what the size represents, negatives. */
export interface BubbleSummaryOptions extends BubbleRadiusOptions {
	showNegativeBubbles?: boolean | undefined;
}

/**
 * The category hierarchy of a treemap or sunburst: the `cx:lvl` levels leaf first (the first level
 * holds the leaf labels), each parent label written once on the first row it covers.
 */
export interface HierarchySummaryOptions {
	levels: string[][];
	parentLabelLayout?: ParentLabelLayout | undefined;
	/** Whether the series writes `cx:dataLabels`; Office draws no labels without them. Default true. */
	dataLabels?: boolean | undefined;
}

export interface ChartSummary<S extends ChartSummarySeries = ChartSummarySeries> {
	chartType: ChartSummaryType;
	/** Native axes and the primary plot group IDs, retained for shared rendering. */
	date1904?: boolean;
	plotLayout?: ChartManualLayout;
	axes?: ChartAxis[];
	axisIds?: number[];
	/** Bar/column grouping. */
	grouping?: 'clustered' | 'stacked' | 'percentStacked' | 'standard';
	/** Gap between category clusters, as a percentage of one bar's width (0..500). */
	barGapWidth?: number;
	/** Overlap of adjacent series, in percent (-100..100). */
	barOverlap?: number;
	title?: string;
	series: S[];
	/** Combination chart: the groups drawn over the first one, in document order. */
	extraGroups?: ChartSummaryGroup[];
	/** Waterfall: the subtotal marks and connector visibility of its `cx:layoutPr`. */
	waterfall?: WaterfallOptions;
	/** Histogram: the binning of its `cx:layoutPr`. */
	histogram?: HistogramSummaryOptions;
	/** Stock: the formatting of its `c:stockChart`. */
	stock?: StockSummaryStyle;
	/** Bubble: the display options of its `c:bubbleChart`. */
	bubble?: BubbleSummaryOptions;
	/** Treemap and sunburst: the category levels of its `cx:strDim`. */
	hierarchy?: HierarchySummaryOptions;
	showLegend: boolean;
	legendPosition?: 'r' | 'l' | 't' | 'b' | 'tr';
	/** Office color-style id (ChartColor). Manual series/point colors retain precedence. */
	colorPalette?: number;
	/** Imported Office chart-style metadata. Style authoring is not yet supported. */
	styleDefinition?: ChartStyleDefinition;
	/** Direct formatting. Axis flags and fills serialize; other metadata edits remain unsupported. */
	formatting?: ChartStyleDefinition;
}

/** A document theme resolved for chart painting. */
export interface ChartThemePalette {
	/** 12 colours in theme order (lt1, dk1, lt2, dk2, accent1-6, hlink, folHlink) as `RRGGBB`. */
	colors: string[];
	majorFont: string;
	minorFont: string;
}
