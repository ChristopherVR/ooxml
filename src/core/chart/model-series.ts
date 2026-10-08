// Format-neutral model of a DrawingML chart part (`c:chartSpace`, ECMA-376 Part 1, 21.2): series,
// data sources, text, labels and shape properties. One model for every Office format: xlsx builds
// its chart objects from it, docx attaches it to `w:drawing` charts, and pptx is meant to alias it
// (docs/agnostic-core-plan.md, step 5). Values are kept as written (raw enumerations, cached point
// text) so clients choose their own defaults; the parser reports what it does not model as issues,
// and extension lists stay raw XML for round-trip.
import type { DrawingFill, DrawingLine, DrawingTextBody } from '../drawingml/types';
import type { ChartManualLayout } from './manual-layout';

/** Something in a chart part the parser did not model or could not read. */
export interface ChartParseIssue {
	code: string;
	message: string;
}

/** Shape properties (`c:spPr`) of a chart element, read through `drawingml`. */
export interface ChartShapeProperties {
	/** The fill; absent when the element declares none (inherit). */
	fill?: DrawingFill;
	/** The outline (`a:ln`). */
	line?: DrawingLine;
	/** `a:effectLst` as written, kept for round-trip. */
	effectsXml?: string;
	/**
	 * The children of `c:spPr` as written, with the part's `c`, `a` and `r` declarations left to the
	 * root. The writer re-emits them while `fill`, `line` and `effectsXml` still read back from them
	 * and patches only the changed pieces otherwise, so unmodelled geometry and 3-D settings survive.
	 */
	sourceXml?: string;
}

/**
 * A DrawingML text body of a chart (`c:rich`, `c:txPr`). `sourceXml` holds its children as written
 * (same declaration rule as {@link ChartShapeProperties.sourceXml}); the writer re-emits them while
 * the modelled fields still read back from them, and writes the modelled fields otherwise.
 */
export interface ChartTextBody extends DrawingTextBody {
	sourceXml?: string;
}

/**
 * A manual layout (`c:layout`): `{}` for an empty (automatic) layout element. `sourceXml` keeps
 * the children as written, so values Excel wrote with 17 digits or an exponent come back as written
 * while they still read back the same.
 */
export interface ChartLayout extends ChartManualLayout {
	sourceXml?: string;
}

/** Chart lines (`c:majorGridlines`, `c:leaderLines`, `c:hiLowLines`...): only their shape. */
export interface ChartLines {
	spPr?: ChartShapeProperties;
}

/** A number format (`c:numFmt`). */
export interface ChartNumberFormat {
	formatCode: string;
	/** Whether the format follows the source cells. */
	sourceLinked?: boolean;
}

/** Whether a cache holds numbers (`c:numCache`, `c:numLit`) or text (`c:strCache`, `c:strLit`, `c:lvl`). */
export type ChartDataCacheType = 'number' | 'string';

/** One cached point (`c:pt`): its index and its text exactly as written. */
export interface ChartCachePoint {
	/** `c:pt/@idx`; absent when the attribute is missing or not an index. */
	index?: number;
	value: string;
	/** Point-level `@formatCode` (number caches). */
	formatCode?: string;
}

/** A point cache or literal (`c:numCache`, `c:strCache`, `c:numLit`, `c:strLit`, `c:lvl`). */
export interface ChartDataCache {
	type: ChartDataCacheType;
	/** `c:ptCount`; absent when not written. */
	pointCount?: number;
	/** `c:formatCode` of a number cache. */
	formatCode?: string;
	points: ChartCachePoint[];
}

/** How a data source is written. */
export type ChartDataSourceKind = 'numRef' | 'strRef' | 'multiLvlStrRef' | 'numLit' | 'strLit';

/**
 * A data source (`c:cat`, `c:val`, `c:xVal`, `c:yVal`, `c:bubbleSize`, a `c:tx/c:strRef`): the
 * formula and the values the producer cached. `kind` is absent for an empty source element.
 */
export interface ChartDataSource {
	kind?: ChartDataSourceKind;
	/** `c:f`, when non-empty (`Sheet1!$B$2:$B$5`). */
	formula?: string;
	/** The cache of a reference, or the literal values. */
	cache?: ChartDataCache;
	/** The levels of a multi-level category cache (`c:multiLvlStrCache/c:lvl`), innermost first. */
	levels?: ChartDataCache[];
	/** `c:ptCount` of the multi-level cache. */
	levelPointCount?: number;
}

/**
 * Chart text (`c:tx`): rich text, a reference with its cached text, or a literal value (series
 * names only). `text` is the flattened string a client shows when it draws no runs.
 */
export interface ChartText {
	rich?: ChartTextBody;
	reference?: ChartDataSource;
	/** `c:v` of a series name. */
	value?: string;
	/** `value`, else the rich text (paragraphs joined by newlines), else the first cached point. */
	text?: string;
}

/** A title (`c:title`) of the chart or an axis. */
export interface ChartTitle {
	tx?: ChartText;
	/** Flattened text of `tx`; absent when the title has no text (an automatic title). */
	text?: string;
	overlay?: boolean;
	layout?: ChartLayout;
	spPr?: ChartShapeProperties;
	/** Text properties (`c:txPr`). */
	txPr?: ChartTextBody;
	extLst?: string;
}

/** Data-label switches shared by `c:dLbls` and `c:dLbl`. */
export interface ChartDataLabelOptions {
	deleted?: boolean;
	showLegendKey?: boolean;
	showValue?: boolean;
	showCategoryName?: boolean;
	showSeriesName?: boolean;
	showPercent?: boolean;
	showBubbleSize?: boolean;
	separator?: string;
	/** `c:dLblPos` (`ctr`, `outEnd`, `bestFit`...). */
	position?: string;
	numberFormat?: ChartNumberFormat;
	spPr?: ChartShapeProperties;
	txPr?: ChartTextBody;
	extLst?: string;
}

/** One point's label (`c:dLbl`). */
export interface ChartDataLabel extends ChartDataLabelOptions {
	index?: number;
	/** Custom label text. */
	tx?: ChartText;
	layout?: ChartLayout;
}

/** Data labels of a group or series (`c:dLbls`). */
export interface ChartDataLabels extends ChartDataLabelOptions {
	labels: ChartDataLabel[];
	showLeaderLines?: boolean;
	/** `c:leaderLines`, present when written. */
	leaderLines?: ChartLines;
}

/** A marker (`c:marker`). */
export interface ChartMarker {
	/** `circle`, `square`, `none`, `auto`... */
	symbol?: string;
	size?: number;
	spPr?: ChartShapeProperties;
	extLst?: string;
}

/** A data point override (`c:dPt`). */
export interface ChartDataPoint {
	/** `c:idx`; absent when missing or not a non-negative integer. */
	index?: number;
	invertIfNegative?: boolean;
	bubble3D?: boolean;
	explosion?: number;
	marker?: ChartMarker;
	/** Present whenever `c:spPr` is written, even empty. */
	spPr?: ChartShapeProperties;
	extLst?: string;
}

/** One series (`c:ser`) of a chart group. */
export interface ChartSpaceSeries {
	index?: number;
	order?: number;
	/** Series name (`c:tx`). */
	tx?: ChartText;
	/** Present whenever `c:spPr` is written, even empty. */
	spPr?: ChartShapeProperties;
	marker?: ChartMarker;
	dataLabels?: ChartDataLabels;
	dataPoints: ChartDataPoint[];
	/** `c:cat`, present whenever the element is written. */
	categories?: ChartDataSource;
	/** `c:val`. */
	values?: ChartDataSource;
	/** `c:xVal` (scatter and bubble). */
	xValues?: ChartDataSource;
	/** `c:yVal` (scatter and bubble). */
	yValues?: ChartDataSource;
	/** `c:bubbleSize`. */
	bubbleSizes?: ChartDataSource;
	smooth?: boolean;
	invertIfNegative?: boolean;
	explosion?: number;
	bubble3D?: boolean;
	/** `c:shape` of a 3-D bar series. */
	shape?: string;
	extLst?: string;
}
